/**
 * 3.2 DOM overlay text editing (run-level rich text) — a contentEditable stacked over the text
 * box takes over input (caret/selection/IME for free). Bold/italic/underline are triggered by the
 * ribbon font group (execCommand on the selection). On exit, walk the DOM to extract the
 * paragraph/run structure (each run's format preserved independently) and go through IPC editText.
 */
import React, { useEffect, useRef } from 'react'
import type { GlyphRun, ShapeRenderNode, TextLine } from '@genoffice/pptx-render'
import type { EditParagraph, EditRun, LinkTargetOp } from '../shared/ipc'
import { decodeLinkTarget, encodeLinkTarget } from '../shared/run-link'
import { displayFontFamily } from './konva-adapter'
import { ZOOM_PREVIEW_EVENT } from './zoom-preview'
import { FONT_SIZES } from './components/ribbon-shared'

interface Props {
  node: ShapeRenderNode
  /** Viewport scale (RenderSlide.scale) — editing renders in viewport px; divided back out when committing to pt */
  scale: number
  onCommit: (paragraphs: EditParagraph[]) => void
  onCancel: () => void
  /** Tab/Shift+Tab (for table cell editing): commit current content and jump to the next/previous cell.
   * paragraphs=null means content unchanged (the host may skip committing and only jump). */
  onTabNav?: (paragraphs: EditParagraph[] | null, dir: 1 | -1) => void
  /** Viewport coordinates of the double-click: select the word there when entering editing; defaults to caret at end */
  caretPoint?: { x: number; y: number }
  /** Entered by typing directly on a selected shape: select all, then replace the whole content with that character */
  replaceWith?: string
  /** ⌘/Ctrl+click on a linked run follows the link (slide jump / external url) */
  onFollowLink?: (target: LinkTargetOp) => void
  /** Edit-frame color (matches the canvas selection chrome: white on dark slide backgrounds) */
  frameColor?: string
  /** Canvas CSS zoom: the outline divides by it to keep a constant on-screen weight */
  zoom?: number
}

/** Layout lines → paragraph grouping (paraStart marks wrap boundaries; missing means an independent paragraph, backward compatible). */
function groupLinesToParagraphs(lines: TextLine[]): TextLine[][] {
  const paras: TextLine[][] = []
  for (const line of lines) {
    if (line.paraStart === false && paras.length) paras[paras.length - 1]!.push(line)
    else paras.push([line])
  }
  return paras
}

/** Preserve the layout engine's glyph fragments so the editor uses the same measured advances as canvas.
 * Extraction merges adjacent fragments back into source model runs by srcRunIdx. Trailing spaces swallowed
 * on wrap are restored at the previous fragment's tail; <a:br/> soft breaks are restored as "\n" sentinels. */
function editorParaRuns(
  paraLines: TextLine[],
): Array<{ run?: GlyphRun; srcRun?: number; text: string }> {
  const segs: Array<{ run?: GlyphRun; srcRun?: number; text: string }> = []
  paraLines.forEach((line, li) => {
    if (li > 0 && paraLines[li - 1]!.trailingSpace && segs.length) {
      segs[segs.length - 1]!.text += ' '
    }
    // Canvas consumes visual bidi order; contentEditable must receive logical source order and
    // lets Chromium perform bidi shaping/reordering itself.
    const logicalRuns = [...line.runs].sort(
      (a, b) =>
        (a.logicalOrder ?? Number.MAX_SAFE_INTEGER) - (b.logicalOrder ?? Number.MAX_SAFE_INTEGER),
    )
    for (const run of logicalRuns) {
      if (run.isBullet || run.text === '') continue
      segs.push({ run, srcRun: run.srcRunIdx, text: run.text })
    }
    if (line.softBreakAfter != null) {
      segs.push({ srcRun: line.softBreakAfter, text: '\n' })
    }
  })
  return segs
}

/** Browser inline-box metrics of a font (Chromium: an inline text box is exactly
 * ascent+descent tall, and a zero-size inline-block probe's offsetTop is the baseline).
 * Cached per font. Returns zero heights in layout-less environments (jsdom) — callers
 * skip compensation there. */
const fontBoxCache = new Map<string, { height: number; ascent: number }>()
function browserFontBox(
  family: string,
  sizePx: number,
  bold?: boolean,
  italic?: boolean,
): { height: number; ascent: number } {
  const key = `${family}|${Math.round(sizePx * 10)}|${bold ? 'b' : ''}${italic ? 'i' : ''}`
  const hit = fontBoxCache.get(key)
  if (hit) return hit
  const host = document.createElement('div')
  host.style.cssText =
    'position:absolute;left:-9999px;top:0;visibility:hidden;line-height:normal;white-space:pre'
  host.style.fontFamily = family
  host.style.fontSize = `${sizePx}px`
  host.style.fontWeight = bold ? 'bold' : 'normal'
  host.style.fontStyle = italic ? 'italic' : 'normal'
  const text = document.createElement('span')
  text.textContent = 'Hg'
  const probe = document.createElement('span')
  probe.style.cssText = 'display:inline-block;width:0;height:0'
  host.append(text, probe)
  document.body.appendChild(host)
  const hostTop = host.getBoundingClientRect().top
  const m = {
    height: text.getBoundingClientRect().height,
    // zero-size inline-block: its box top sits exactly on the baseline (fractional,
    // unlike offsetTop which rounds to whole px)
    ascent: probe.getBoundingClientRect().top - hostTop,
  }
  host.remove()
  fontBoxCache.set(key, m)
  return m
}

/** Where the canvas actually draws the baseline, relative to the engine's baselineY.
 * The Konva adapter positions Text by top = baselineY − 0.8em, and Konva paints with
 * canvas2d 'middle' semantics — so the visible baseline lands at top + emHeightAscent,
 * i.e. engineBaseline + (emHeightAscent − 0.8em). ≈0 for Latin sans (the 0.8 was tuned
 * for it) but several px for CJK/serif metrics; the editor must match the pixels. */
let baselineDropCtx: CanvasRenderingContext2D | null = null
const baselineDropCache = new Map<string, number>()
function konvaBaselineDrop(
  family: string,
  sizePx: number,
  bold?: boolean,
  italic?: boolean,
): number {
  const key = `${family}|${Math.round(sizePx * 10)}|${bold ? 'b' : ''}${italic ? 'i' : ''}`
  const hit = baselineDropCache.get(key)
  if (hit !== undefined) return hit
  baselineDropCtx ??= document.createElement('canvas').getContext('2d')
  if (!baselineDropCtx) return 0
  const ctx = baselineDropCtx
  ctx.font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePx}px ${family}`
  // Same ink measured from both baselines: the ascent difference IS the exact distance
  // from the 'middle' anchor down to the alphabetic baseline for this font (measuring
  // beats modelling — Chromium derives 'middle' from per-font metrics)
  ctx.textBaseline = 'alphabetic'
  const a1 = ctx.measureText('Hg').actualBoundingBoxAscent
  ctx.textBaseline = 'middle'
  const a2 = ctx.measureText('Hg').actualBoundingBoxAscent
  const drop = a1 > 0 || a2 > 0 ? sizePx / 2 + (a1 - a2) - 0.8 * sizePx : 0
  baselineDropCache.set(key, drop)
  return drop
}

/**
 * Layout lines → the editor's initial DOM: one <div> per model paragraph (data-src-para records
 * the source paragraph index, with paragraph alignment); wrap/word-split fragments merged back
 * into their runs by srcRunIdx (data-src-run), one span per model run; bullet glyphs injected by
 * layout are skipped (not body text; committing them would turn them into text).
 * Font sizes use layout viewport px directly (including scale/autofit fontScale); on commit
 * they're divided by norm back to pt.
 * Line/paragraph spacing is derived back from layout lines (lnSpc/lineExact/lnSpcReduction baked
 * into height, spcBef/spcAft baked into gaps in line tops), so editing's vertical metrics match the canvas.
 * anchorDy = the whole offset that middle/bottom anchoring bakes into line tops (editing implements
 * anchoring with flex, so it must be removed from the first paragraph's top or the offset doubles).
 * Exported so tests can do "layout → DOM → extractParagraphs" round-trip assertions.
 */
export function populateEditorDom(
  div: HTMLElement,
  lines: TextLine[],
  anchorDy = 0,
  innerW?: number,
): void {
  div.innerHTML = ''
  delete div.dataset.layoutReleased
  // Root/strut font: the paragraph divs inherit it and it participates in every line box
  const rootCs = div.isConnected ? window.getComputedStyle(div) : null
  const strutFont = rootCs ? { family: rootCs.fontFamily, size: parseFloat(rootCs.fontSize) } : null
  // Widest laid-out line: in a nowrap box the block grows to max-content, so centered/right
  // paragraphs align within this width instead of the box (the canvas splits overflow to
  // both sides) — compensated per paragraph below. Bullet glyphs are skipped: they never
  // enter the editor DOM and the engine's alignment width excludes them too.
  const maxLineW = Math.max(
    0,
    ...lines.map((ln) => ln.runs.reduce((acc, r) => acc + (r.isBullet ? 0 : r.widthPx), 0)),
  )
  let prevEnd = anchorDy
  groupLinesToParagraphs(lines).forEach((paraLines, pi) => {
    const p = document.createElement('div')
    p.dataset.srcPara = String(pi)
    const first = paraLines[0]!
    const last = paraLines[paraLines.length - 1]!
    p.style.lineHeight = `${first.height}px`
    const gap = first.top - prevEnd
    if (Math.abs(gap) > 0.01) p.style.marginTop = `${gap}px`
    prevEnd = last.top + last.height
    // RTL paragraphs (Arabic/Hebrew) align in editing as on canvas: the browser sets direction by the first strong character
    p.dir = 'auto'
    const align = paraLines[0]?.align
    if (align) p.style.textAlign = align
    // Body text starts at marL, exactly like the canvas (lists/indent used to snap to the
    // inset edge on entering edit); first-line indent applies only without a bullet
    const marL = first.marLPx ?? 0
    if (marL) p.style.marginLeft = `${marL}px`
    const indentPx = first.indentPx ?? 0
    if (indentPx && !first.runs.some((r) => r.isBullet)) p.style.textIndent = `${indentPx}px`
    // ── Glyph-position fidelity vs the canvas renderer ──
    // Vertical: the canvas draws the dominant run's baseline at
    // lineTop + engineAscent + konvaBaselineDrop (the adapter's 0.8em top approximation
    // filtered through Konva's 'middle'-baseline painting); CSS puts the DOM baseline at
    // half-leading + browser ascent. The difference is several px on CJK/serif or
    // lnSpc ≠ 100% text and reads as the text jumping when editing starts. Measure both
    // sides and cancel the difference with a relative offset (flow is unaffected).
    let engineAscent = 0
    let domBaseline = 0
    let dominant: GlyphRun | null = null
    for (const r of first.runs) {
      const a = r.ascentPx ?? r.fontSizePx * 0.8
      if (a > engineAscent) {
        engineAscent = a
        dominant = r
      }
    }
    const drop = dominant
      ? konvaBaselineDrop(
          displayFontFamily(dominant.fontFamily ?? ''),
          dominant.fontSizePx,
          dominant.bold,
          dominant.italic,
        )
      : 0
    const canvasBaseline = engineAscent + drop
    const participants: Array<{ family: string; size: number; bold?: boolean; italic?: boolean }> =
      first.runs.map((r) => ({
        family: displayFontFamily(r.fontFamily ?? ''),
        size: r.fontSizePx,
        bold: r.bold,
        italic: r.italic,
      }))
    if (strutFont) participants.push({ family: strutFont.family, size: strutFont.size })
    for (const f of participants) {
      if (!f.family || !f.size) continue
      const m = browserFontBox(f.family, f.size, f.bold, f.italic)
      if (!m.height) continue
      domBaseline = Math.max(domBaseline, (first.height - m.height) / 2 + m.ascent)
    }
    const dyFix = canvasBaseline > 0 && domBaseline > 0 ? canvasBaseline - domBaseline : 0
    // Horizontal: nowrap overflow — the canvas centers/right-aligns within the box and
    // spills both ways; the DOM block is max-content wide and anchored at the box's left.
    // The difference is a constant per box (zero when the content fits or the box wraps).
    let dxFix = 0
    if (innerW != null && innerW > 0 && (align === 'center' || align === 'right')) {
      const over = Math.max(0, maxLineW - innerW)
      dxFix = align === 'center' ? -over / 2 : -over
    }
    if (Math.abs(dyFix) > 0.1 || Math.abs(dxFix) > 0.1) {
      p.style.position = 'relative'
      if (Math.abs(dyFix) > 0.1) p.style.top = `${dyFix}px`
      if (Math.abs(dxFix) > 0.1) p.style.left = `${dxFix}px`
    }
    const level = paraLines[0]?.level ?? 0
    if (level) {
      p.dataset.level = String(level)
      // Visual indent hint only when the real marL isn't known (the canvas lays out by marL)
      if (!marL) p.style.marginLeft = `${level * 24}px`
    }
    // Original bullet kind, for the ribbon's toggle-off semantics while editing
    const bulletRun = first.runs.find((r) => r.isBullet)
    if (bulletRun) p.dataset.hadBullet = /^\d/.test(bulletRun.text) ? 'number' : 'char'
    for (const { run, srcRun, text } of editorParaRuns(paraLines)) {
      if (run) {
        const prev = p.lastElementChild as HTMLElement | null
        const src = srcRun != null ? String(srcRun) : undefined
        const reuse = prev?.dataset.runContainer === 'true' && prev.dataset.srcRun === src
        // Linked runs become <a href> so execCommand('unlink') and extraction see them natively
        const span = reuse ? prev! : document.createElement(run.link ? 'a' : 'span')
        if (!reuse) {
          span.dataset.runContainer = 'true'
          if (src) span.dataset.srcRun = src
          if (run.link) span.setAttribute('href', run.link)
          if (run.bold) span.style.fontWeight = 'bold'
          if (run.italic) span.style.fontStyle = 'italic'
          const deco = [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(
            Boolean,
          )
          if (deco.length) span.style.textDecoration = deco.join(' ')
          else if (run.link) span.style.textDecoration = 'none' // suppress the UA <a> underline
          // Super/subscript: the initial DOM must restore it (otherwise extraction sends explicit 0/false and wipes the original format)
          if (run.baselinePct) span.style.verticalAlign = run.baselinePct > 0 ? 'super' : 'sub'
          span.style.fontSize = `${run.fontSizePx}px`
          if (run.fontFamily) {
            const display = displayFontFamily(run.fontFamily)
            span.style.fontFamily = display
            // Record what was baked in for display: on extraction, if the first item of the stack is
            // still this, the user didn't change the font — commit the model's original name (data-font),
            // or nothing at all when the model run has no explicit font (run.fontFamily is then a layout
            // default / missing-font substitution like Arial, which must never be written into the file).
            span.dataset.displayFont = firstFontFamily(display)
            if (run.srcFontFamily) span.dataset.font = run.srcFontFamily
          }
          span.style.color = normalizeCss(run.color)
          p.appendChild(span)
        }
        const fragment = document.createElement('span')
        fragment.dataset.layoutFragment = 'true'
        fragment.textContent = text
        // Each fragment occupies the exact advance measured by the layout engine. CJK is normally
        // one grapheme per fragment; Latin/SEA keep their script-aware token boundaries. RTL stays
        // in normal inline flow so Chromium can preserve joining and bidirectional shaping.
        if (!run.rtl) {
          fragment.style.display = 'inline-block'
          fragment.style.width = `${run.widthPx}px`
        }
        // Keep the browser editor visually aligned with the canvas renderer. This is display-only:
        // extraction intentionally preserves the source run's PPT letter spacing through srcRun.
        if (run.letterSpacingPx) fragment.style.letterSpacing = `${run.letterSpacingPx}px`
        span.appendChild(fragment)
      } else {
        const span = document.createElement('span')
        span.textContent = text
        if (srcRun != null) span.dataset.srcRun = String(srcRun)
        p.appendChild(span)
      }
    }
    if (!p.childNodes.length) p.appendChild(document.createElement('br')) // Empty paragraph placeholder
    div.appendChild(p)
  })
}

/** Fixed fragment advances align the untouched editor with canvas, but become stale after typing.
 * Release them so inserted/deleted text can reflow naturally. */
export function releaseEditorLayoutConstraints(root: HTMLElement): void {
  if (root.dataset.layoutReleased === 'true') return
  root.dataset.layoutReleased = 'true'
  root.querySelectorAll<HTMLElement>('[data-layout-fragment]').forEach((fragment) => {
    fragment.style.display = ''
    fragment.style.width = ''
  })
}

/** Release one fragment's fixed advance (stale once its text or format changes). */
function releaseFragment(el: Element | null | undefined): void {
  if (!(el instanceof HTMLElement) || el.dataset.layoutFragment !== 'true') return
  el.style.display = ''
  el.style.width = ''
}

/** Nearest enclosing layout fragment of a DOM point (bounded by the editor root). */
function fragmentAround(node: Node | null | undefined, root: HTMLElement): HTMLElement | null {
  let el = node instanceof HTMLElement ? node : (node?.parentElement ?? null)
  while (el && el !== root) {
    if (el.dataset.layoutFragment === 'true') return el
    el = el.parentElement
  }
  return null
}

/** Release only the fragments an edit touches (the given target ranges plus the current
 * selection; a collapsed caret also frees its neighbor fragments — Backspace/Delete at a
 * fragment edge mutates them). Untouched fragments keep their canvas-measured advances, so
 * the rest of the text stays put: releasing everything on the first keystroke re-measured
 * and re-wrapped the whole box with browser rules (natural advances + CJK line-break
 * prohibitions the canvas engine doesn't apply) and made all the text visibly jump.
 * Exported for tests. */
export function releaseFragmentsAtEdit(
  root: HTMLElement,
  targetRanges: readonly AbstractRange[] = [],
): void {
  const ranges: AbstractRange[] = [...targetRanges]
  const sel = window.getSelection()
  if (sel?.rangeCount && root.contains(sel.anchorNode)) ranges.push(sel.getRangeAt(0))
  if (!ranges.length) return
  const frags = [...root.querySelectorAll<HTMLElement>('[data-layout-fragment]')]
  for (const r of ranges) {
    if (r.collapsed) {
      const f = fragmentAround(r.startContainer, root)
      if (!f) continue
      const i = frags.indexOf(f)
      releaseFragment(f)
      releaseFragment(frags[i - 1])
      releaseFragment(frags[i + 1])
      continue
    }
    // Ranged edit (selection replace/delete, execCommand format): free every intersecting fragment
    const live =
      r instanceof Range
        ? r
        : (() => {
            const x = document.createRange()
            x.setStart(r.startContainer, r.startOffset)
            x.setEnd(r.endContainer, r.endOffset)
            return x
          })()
    for (const f of frags) if (live.intersectsNode(f)) releaseFragment(f)
  }
}

export function TextEditOverlay({
  node,
  scale,
  onCommit,
  onCancel,
  onTabNav,
  caretPoint,
  replaceWith,
  onFollowLink,
  frameColor = '#232425',
  zoom = 1,
}: Props) {
  const ref = useRef<HTMLDivElement>(null)
  const frameRef = useRef<HTMLDivElement>(null)

  // The edit frame lives inside the CSS-scaled stage: during a zoom gesture only the
  // transform advances, so the zoom-compensated outline would thicken and then snap at
  // commit. Counter-scale it per previewed frame, same as the selection chrome.
  useEffect(() => {
    const onPreview = (e: Event) => {
      const z = (e as CustomEvent<number>).detail
      const el = frameRef.current
      if (typeof z !== 'number' || !el) return
      el.style.outlineWidth = `${2 / (globalThis.devicePixelRatio || 1) / Math.max(z, 0.1)}px`
    }
    window.addEventListener(ZOOM_PREVIEW_EVENT, onPreview)
    return () => window.removeEventListener(ZOOM_PREVIEW_EVENT, onPreview)
  }, [])

  const box = node.box
  const insets = node.text?.insets ?? { l: 0, t: 0, r: 0, b: 0 }
  const anchor = node.text?.anchor ?? 'top'
  // Editing uses layout viewport px directly (including viewport scale and autofit fontScale): layout height =
  // visual height, so the edit box isn't inflated; on commit extractParagraphs divides by norm back to model pt
  const norm = scale * (node.text?.fontScale ?? 1) || 1
  const wrap = node.text?.wrap !== false
  const firstRun =
    node.text?.lines[0]?.runs.find((r) => !r.isBullet) ?? node.text?.lines[0]?.runs[0]
  // Fallback = the layout engine's 18pt default in px, so typing into an empty body
  // matches the canvas line height (18*norm would be px-as-pt: 13.5pt)
  const baseFontSize = firstRun?.fontSizePx ?? ((18 * 96) / 72) * norm
  const baseColor = normalizeCss(firstRun?.color ?? '#000')
  const baseFont = displayFontFamily(firstRun?.fontFamily ?? 'Calibri')

  // Initial content: see populateEditorDom; snapshot the initial extraction, unchanged commits go through onCancel
  // (no empty undo step / no dirty flag — opening to look and clicking away isn't an edit)
  const initialRef = useRef<string>('')
  useEffect(() => {
    const div = ref.current
    if (!div) return
    div.dataset.norm = String(norm) // The ribbon helpers for font size increase/decrease/set take the conversion factor from here
    // Same anchor offset as the engine (the dy text-layout bakes into line tops), removed back during populate
    const extraH = Math.max(box.h - insets.t - insets.b, 1) - (node.text?.contentHeight ?? 0)
    const anchorDy = anchor === 'middle' ? extraH / 2 : anchor === 'bottom' ? extraH : 0
    populateEditorDom(div, node.text?.lines ?? [], anchorDy, box.w - insets.l - insets.r)
    initialRef.current = JSON.stringify(extractParagraphs(div, norm))
    div.focus()
    const sel = window.getSelection()
    if (sel && replaceWith) {
      // Type-to-replace: select all, then replace the whole content with the first typed character
      const range = document.createRange()
      range.selectNodeContents(div)
      sel.removeAllRanges()
      sel.addRange(range)
      document.execCommand('insertText', false, replaceWith)
      return
    }
    if (sel) {
      // Entering by double-click: select the word at the click; without coordinates/no hit, caret to end
      const hit = caretPoint ? document.caretRangeFromPoint(caretPoint.x, caretPoint.y) : null
      if (hit && div.contains(hit.startContainer)) {
        sel.removeAllRanges()
        sel.addRange(hit)
        const s = sel as Selection & {
          modify?: (alter: string, dir: string, granularity: string) => void
        }
        s.modify?.('move', 'backward', 'word')
        s.modify?.('extend', 'forward', 'word')
      } else {
        const range = document.createRange()
        range.selectNodeContents(div)
        range.collapse(false)
        sel.removeAllRanges()
        sel.addRange(range)
      }
    }
  }, [node, norm, caretPoint, replaceWith])

  const commit = () => {
    savedSel = null
    const div = ref.current
    if (!div) return onCancel()
    const paras = extractParagraphs(div, norm)
    if (JSON.stringify(paras) === initialRef.current) return onCancel()
    onCommit(paras)
  }

  // Targeted layout release (native listeners: React's onBeforeInput synthetic event does
  // not map to the real `beforeinput`). beforeinput sees the pre-mutation target ranges
  // (Backspace at a fragment edge deletes into the neighbor); the input listener covers
  // execCommand formatting (bold/font — width-changing, fires input without beforeinput).
  useEffect(() => {
    const div = ref.current
    if (!div) return
    const onBeforeInput = (ev: InputEvent) =>
      releaseFragmentsAtEdit(div, ev.getTargetRanges?.() ?? [])
    const onInput = () => releaseFragmentsAtEdit(div)
    div.addEventListener('beforeinput', onBeforeInput)
    div.addEventListener('input', onInput)
    return () => {
      div.removeEventListener('beforeinput', onBeforeInput)
      div.removeEventListener('input', onInput)
    }
  }, [])

  // While focus is parked on a keep-edit control the editor is already blurred, so its
  // onBlur can't fire again — a press anywhere else must still commit instead of silently dropping the edit
  const commitRef = useRef(commit)
  commitRef.current = commit
  useEffect(() => {
    const onDown = (ev: MouseEvent) => {
      const div = ref.current
      if (!div || div.contains(document.activeElement)) return
      const t = ev.target instanceof HTMLElement ? ev.target : null
      if (t && (div.parentElement?.contains(t) || t.closest('[data-keep-edit]'))) return
      commitRef.current()
    }
    document.addEventListener('mousedown', onDown, true)
    return () => {
      document.removeEventListener('mousedown', onDown, true)
      savedSel = null
    }
  }, [])

  return (
    // Outer layer = the whole text box (the edit-frame border is drawn here);
    // inner contentEditable edits in place with a transparent background (the canvas already hides this node's text), flex implements the vertical anchor.
    // Height fixed to the shape box: overflowing content shows past it, the edit box doesn't grow with content;
    // border uses outline (takes no layout space) so the inner usable size matches canvas layout exactly
    <div
      ref={frameRef}
      style={{
        position: 'absolute',
        left: box.x,
        top: box.y,
        width: box.w,
        height: box.h,
        zIndex: 20,
        display: 'flex',
        flexDirection: 'column',
        justifyContent:
          anchor === 'middle' ? 'center' : anchor === 'bottom' ? 'flex-end' : 'flex-start',
        // 2 device px, zoom-compensated: the canvas is CSS-scaled, and 2 CSS px reads
        // twice as heavy on retina displays
        outline: `${2 / (globalThis.devicePixelRatio || 1) / Math.max(zoom, 0.1)}px solid ${frameColor}`,
      }}
    >
      <div
        ref={ref}
        contentEditable
        spellCheck
        suppressContentEditableWarning
        onBlur={(e) => {
          // Keep-edit controls (font size input, native color picker) take focus without committing;
          // they save/restore the selection and apply to it instead of element-level
          const to = e.relatedTarget instanceof HTMLElement ? e.relatedTarget : null
          if (to?.closest('[data-keep-edit]')) return
          commit()
        }}
        onClick={(e) => {
          // ⌘/Ctrl+click follows a run link (plain clicks keep editing, matching PowerPoint)
          if (!(e.metaKey || e.ctrlKey) || !onFollowLink) return
          const a = e.target instanceof Node ? linkAround(e.target) : null
          const target = a && decodeLinkTarget(a.getAttribute('href'))
          if (target) {
            e.preventDefault()
            onFollowLink(target)
          }
        }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            // Esc = commit the text and return to shape-selected state (input not lost);
            // when unchanged, commit internally goes through onCancel and produces no history step
            e.preventDefault()
            commit()
          } else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
            e.preventDefault()
            commit()
          } else if (e.key === 'Tab' && onTabNav) {
            // Table cells: Tab commits and jumps to the next cell / Shift+Tab previous;
            // block the default focus move (otherwise blur falsely triggers commit-and-exit)
            e.preventDefault()
            const div = ref.current
            if (div) {
              const paras = extractParagraphs(div, norm)
              const changed = JSON.stringify(paras) !== initialRef.current
              onTabNav(changed ? paras : null, e.shiftKey ? -1 : 1)
            }
          } else if (e.key === 'Tab') {
            // Multi-level lists: Tab/⇧Tab adjust the caret paragraph's indent level (lvl written on commit;
            // editing only shows a marginLeft visual hint, real indentation is laid out by the canvas per master styles)
            e.preventDefault()
            const selNow = window.getSelection()
            let blk: HTMLElement | null =
              selNow?.anchorNode instanceof HTMLElement
                ? selNow.anchorNode
                : (selNow?.anchorNode?.parentElement ?? null)
            while (blk && blk !== ref.current && blk.tagName !== 'DIV') blk = blk.parentElement
            if (blk && blk !== ref.current) {
              const cur = parseInt(blk.dataset.level ?? '0', 10) || 0
              const next = Math.max(0, Math.min(8, cur + (e.shiftKey ? -1 : 1)))
              if (next) {
                blk.dataset.level = String(next)
                blk.style.marginLeft = `${next * 24}px`
              } else {
                delete blk.dataset.level
                blk.style.marginLeft = ''
              }
            }
          } else if (e.key === 'Enter' && e.shiftKey) {
            // Shift+Enter = in-paragraph soft break (<a:br/>). The default behavior inserts <br>,
            // and execCommand('insertText','\n') splits the div in Chromium — both become paragraph splits,
            // so the only way is manually inserting a "\n" text node via Range (displayed in place as a line break under pre-wrap)
            e.preventDefault()
            const sel = window.getSelection()
            if (sel && sel.rangeCount) {
              const range = sel.getRangeAt(0)
              range.deleteContents()
              const tn = document.createTextNode('\n')
              range.insertNode(tn)
              range.setStartAfter(tn)
              range.collapse(true)
              sel.removeAllRanges()
              sel.addRange(range)
            }
          }
        }}
        style={{
          // wrap=none: no wrapping (width follows content, same as the canvas overflow behavior); soft-break \n still breaks via pre.
          // insets use padding not margin: paragraph divs' marginTop (paragraph spacing) must not collapse with the root node
          width: wrap ? Math.max(box.w, 40 + insets.l + insets.r) : 'max-content',
          // nowrap keeps max-content growth for overflow, but never below the box width:
          // a narrower block would defeat per-paragraph text-align (centered titles would
          // visually snap left on entering edit) — the canvas centers within the box
          minWidth: wrap ? undefined : Math.max(box.w, 40 + insets.l + insets.r),
          // Only an empty body needs a synthetic height (so typing matches the canvas line
          // height); inflating a laid-out body distorts the flex vertical anchor — a
          // middle-anchored single line with tight spacing sat a few px too high in edit
          minHeight: node.text?.lines.length ? undefined : baseFontSize * 1.2 + insets.t + insets.b,
          padding: `${insets.t}px ${insets.r}px ${insets.b}px ${insets.l}px`,
          fontSize: baseFontSize,
          fontFamily: baseFont,
          color: baseColor,
          lineHeight: 1.2,
          outline: 'none',
          background: 'transparent',
          caretColor: baseColor,
          boxSizing: 'border-box',
          whiteSpace: wrap ? 'pre-wrap' : 'pre',
        }}
      />
    </div>
  )
}

/** css text-align → paragraph align (start/empty treated as unspecified, main process falls back to the original value). */
function cssAlign(v: string): EditParagraph['align'] | undefined {
  if (v === 'left' || v === 'center' || v === 'right' || v === 'justify') return v
  return undefined
}

/** Whether adjacent runs share source and format (mergeable losslessly). When both srcRun are undefined, merge newly typed text by format. */
function sameRunFormat(a: EditRun, b: EditRun): boolean {
  return (
    a.srcRun === b.srcRun &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.baseline === b.baseline &&
    a.fontSize === b.fontSize &&
    a.fontFamily === b.fontFamily &&
    a.color === b.color &&
    (a.link ? encodeLinkTarget(a.link) : '') === (b.link ? encodeLinkTarget(b.link) : '')
  )
}

/** Merge adjacent same-source same-format runs: stitch fragments split by CJK per-char/latin per-word/execCommand
 * back into whole runs, keeping model run boundaries stable (required for the lossless in-place patch path on save). */
function mergeAdjacentRuns(runs: EditRun[]): EditRun[] {
  const out: EditRun[] = []
  for (const r of runs) {
    const last = out[out.length - 1]
    if (last && sameRunFormat(last, r)) last.text += r.text
    else out.push({ ...r })
  }
  return out
}

/** Split "\n" (soft breaks) inside run text into standalone sentinel runs: at the model/save layer a soft
 * break = a standalone "\n" run (maps to <a:br/>). Spans that are entirely "\n" (initially rendered sentinels) keep their srcRun as is. */
function splitSoftBreaks(runs: EditRun[]): EditRun[] {
  const out: EditRun[] = []
  for (const r of runs) {
    if (r.text === '\n' || !r.text.includes('\n')) {
      out.push(r)
      continue
    }
    r.text.split('\n').forEach((part, i) => {
      if (i > 0)
        out.push({
          text: '\n',
          bold: false,
          italic: false,
          underline: false,
          strike: false,
          baseline: 0,
        })
      if (part) out.push({ ...r, text: part })
    })
  }
  return out
}

/** Walk the contentEditable DOM → paragraph/run structure. <br>/<div> split paragraphs, span styles → run format,
 * div/root text-align → paragraph alignment (product of execCommand justify*).
 * data-src-para/data-src-run carry back source indexes (browser Enter splits divs copying data attributes,
 * so both halves inherit the same source).
 * bold/italic/underline are committed as explicit booleans (the DOM is the authoritative state), avoiding
 * the main process's ?? fallback inheriting wrongly.
 * norm = viewport scale × autofit fontScale: DOM font sizes are viewport px, divided by norm to model pt. */
export function extractParagraphs(root: HTMLElement, norm: number): EditParagraph[] {
  const paragraphs: EditParagraph[] = []
  let cur: EditRun[] = []
  const rootAlign = cssAlign(root.style.textAlign)
  let curAlign = rootAlign
  let curSrcPara: number | undefined
  let curLevel = 0
  let curFmt: Partial<EditParagraph> = {}
  const pushPara = () => {
    paragraphs.push({
      runs: cur.length
        ? splitSoftBreaks(mergeAdjacentRuns(cur))
        : [{ text: '', bold: false, italic: false, underline: false, strike: false, baseline: 0 }],
      ...(curAlign ? { align: curAlign } : {}),
      level: curLevel,
      ...(curSrcPara != null ? { srcPara: curSrcPara } : {}),
      ...curFmt,
    })
    cur = []
  }

  const walk = (node: Node, inherited: Partial<EditRun>) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? ''
      if (text) {
        cur.push({
          text,
          bold: !!inherited.bold,
          italic: !!inherited.italic,
          underline: !!inherited.underline,
          strike: !!inherited.strike,
          baseline: inherited.baseline ?? 0,
          ...(inherited.fontSize != null ? { fontSize: inherited.fontSize } : {}),
          ...(inherited.fontFamily ? { fontFamily: inherited.fontFamily } : {}),
          ...(inherited.color ? { color: inherited.color } : {}),
          ...(inherited.srcRun != null ? { srcRun: inherited.srcRun } : {}),
          link: inherited.link ?? null, // explicit null = no link (the DOM is authoritative here)
        })
      }
      return
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return
    const el = node as HTMLElement
    if (el.tagName === 'BR') {
      pushPara()
      return
    }
    const isBlock = el.tagName === 'DIV' || el.tagName === 'P'
    if (isBlock && cur.length) {
      // Some browsers use <div> for line breaks
      pushPara()
    }
    if (isBlock) {
      curAlign = cssAlign(el.style.textAlign) ?? rootAlign
      const sp = el.dataset ? parseInt(el.dataset.srcPara ?? '', 10) : NaN
      curSrcPara = Number.isNaN(sp) ? undefined : sp
      const lv = el.dataset ? parseInt(el.dataset.level ?? '0', 10) : 0
      curLevel = Number.isNaN(lv) ? 0 : lv
      curFmt = {}
      const ds = el.dataset
      if (ds?.bullet === 'char' || ds?.bullet === 'number' || ds?.bullet === 'none') {
        curFmt.bullet = ds.bullet
        if (ds.bullet === 'char' && ds.bulletChar) curFmt.bulletChar = ds.bulletChar
      }
      const num = (v: string | undefined) => {
        const n = parseFloat(v ?? '')
        return Number.isNaN(n) ? undefined : n
      }
      const ls = num(ds?.lineSpacingPct)
      if (ls != null) curFmt.lineSpacingPct = ls
      const sb = num(ds?.spaceBeforePt)
      if (sb != null) curFmt.spaceBeforePt = sb
      const sa = num(ds?.spaceAfterPt)
      if (sa != null) curFmt.spaceAfterPt = sa
    }
    const style = el.style
    const cs = window.getComputedStyle(el)
    const next: Partial<EditRun> = { ...inherited }
    const sr = el.dataset ? parseInt(el.dataset.srcRun ?? '', 10) : NaN
    if (!Number.isNaN(sr)) next.srcRun = sr
    if (el.tagName === 'A') {
      const link = decodeLinkTarget(el.getAttribute('href'))
      if (link) next.link = link
    }
    if (
      style.fontWeight === 'bold' ||
      cs.fontWeight === '700' ||
      el.tagName === 'B' ||
      el.tagName === 'STRONG'
    )
      next.bold = true
    if (style.fontStyle === 'italic' || el.tagName === 'I' || el.tagName === 'EM')
      next.italic = true
    if ((style.textDecoration || cs.textDecorationLine).includes('underline') || el.tagName === 'U')
      next.underline = true
    if (
      (style.textDecoration || cs.textDecorationLine).includes('line-through') ||
      el.tagName === 'S' ||
      el.tagName === 'STRIKE' ||
      el.tagName === 'DEL'
    )
      next.strike = true
    // Super/subscript: <sub>/<sup> (execCommand product) or vertical-align (for initial DOM restoration)
    if (el.tagName === 'SUP' || style.verticalAlign === 'super') next.baseline = 30
    else if (el.tagName === 'SUB' || style.verticalAlign === 'sub') next.baseline = -25
    if (style.color) next.color = rgbToHex(style.color)
    if (style.fontSize) {
      const px = parseFloat(style.fontSize)
      if (!Number.isNaN(px)) next.fontSize = pxToPt(px, norm)
    }
    if (style.fontFamily) {
      const fam = firstFontFamily(style.fontFamily)
      if (fam) {
        const orig = el.dataset?.font
        const baked = el.dataset?.displayFont
        if (baked && baked.toLowerCase() === fam.toLowerCase()) {
          // Display font unchanged since populate: commit the model's original name; a run without
          // an explicit model font commits none (the layout's Arial default / missing-font
          // substitution baked into the DOM is not a user change)
          if (orig) next.fontFamily = orig
          else delete next.fontFamily
        } else {
          next.fontFamily =
            orig && firstFontFamily(displayFontFamily(orig)).toLowerCase() === fam.toLowerCase()
              ? orig
              : fam
        }
      }
    }
    for (const child of Array.from(el.childNodes)) walk(child, next)
    if (isBlock) {
      // Block end = paragraph end (alignment/source paragraph doesn't leak into following siblings)
      if (cur.length) pushPara()
      curAlign = rootAlign
      curLevel = 0
      curSrcPara = undefined
      curFmt = {}
    }
  }

  for (const child of Array.from(root.childNodes)) walk(child, {})
  pushPara()
  // Drop trailing empty paragraphs (unless everything is empty)
  while (paragraphs.length > 1 && paragraphs[paragraphs.length - 1]!.runs.every((r) => !r.text)) {
    paragraphs.pop()
  }
  return paragraphs
}

/**
 * Selection handoff for ribbon controls that must take real focus (font size input, native color
 * picker): save the editor's Range before focus leaves, restore it (re-focusing the
 * editor) right before applying, so the command hits the original selection instead of element-level.
 */
let savedSel: { root: HTMLElement; range: Range } | null = null

export function saveEditSelection(): void {
  const root = document.activeElement
  const sel = window.getSelection()
  if (!(root instanceof HTMLElement) || !root.isContentEditable || !sel?.rangeCount) return
  savedSel = { root, range: sel.getRangeAt(0).cloneRange() }
}

export function restoreEditSelection(): boolean {
  if (!savedSel?.root.isConnected) return false
  savedSel.root.focus()
  const sel = window.getSelection()
  if (!sel) return false
  sel.removeAllRanges()
  sel.addRange(savedSel.range)
  return true
}

/** Nearest <a href> wrapping a node (bounded by the contentEditable root). */
function linkAround(node: Node | null): HTMLAnchorElement | null {
  const el = node instanceof HTMLElement ? node : node?.parentElement
  const a = el?.closest('a[href]')
  return a instanceof HTMLAnchorElement && el?.closest('[contenteditable="true"]') ? a : null
}

/**
 * The editor selection's current hyperlink (nearest <a> around the selection start; the saved
 * ribbon-handoff selection is consulted when focus already left the editor). For dialog echo-back.
 */
export function selectionLink(): LinkTargetOp | null {
  const sel = window.getSelection()
  const node = sel?.rangeCount ? sel.getRangeAt(0).startContainer : savedSel?.range.startContainer
  const a = linkAround(node ?? null)
  return a ? decodeLinkTarget(a.getAttribute('href')) : null
}

/**
 * Set/clear a hyperlink on the editor selection (restoring the saved selection first — the link
 * dialog took focus). A collapsed caret inside an existing link expands to the whole link, matching
 * PowerPoint. Returns false when there is no usable selection to apply to.
 */
export function applySelectionLink(target: LinkTargetOp | null): boolean {
  if (!restoreEditSelection()) return false
  const sel = window.getSelection()
  if (!sel?.rangeCount) return false
  const range = sel.getRangeAt(0)
  if (range.collapsed) {
    const a = linkAround(range.startContainer)
    if (!a) return false
    range.selectNodeContents(a)
    sel.removeAllRanges()
    sel.addRange(range)
  }
  if (target) document.execCommand('createLink', false, encodeLinkTarget(target))
  else document.execCommand('unlink')
  return true
}

/**
 * Per-paragraph format while editing: mark the paragraph divs covered by the
 * caret/selection; extractParagraphs carries the marks to the main process on commit.
 * Toggle-off (clicking the active bullet kind again) resolves against the first covered
 * paragraph's current state. Returns false when no editor selection is available
 * (the caller falls back to the element-level op).
 */
export function applySelectionParagraphFormat(patch: {
  bullet?: 'char' | 'number' | 'none'
  bulletChar?: string
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
}): boolean {
  const sel = window.getSelection()
  if (!sel?.rangeCount) return false
  const range = sel.getRangeAt(0)
  // Editor root resolved from the selection itself (works when a keep-edit control holds focus)
  const startEl =
    range.startContainer instanceof HTMLElement
      ? range.startContainer
      : range.startContainer.parentElement
  const root =
    startEl?.closest('[data-src-para]')?.parentElement ??
    (startEl?.querySelector('[data-src-para]') ? startEl : null) ??
    (savedSel?.root.isConnected ? savedSel.root : null)
  if (!root) return false
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement =>
      el instanceof HTMLElement && el.tagName === 'DIV' && range.intersectsNode(el),
  )
  if (!blocks.length) return false
  let bullet = patch.bullet
  if (bullet && bullet !== 'none' && !patch.bulletChar) {
    const cur = blocks[0]!.dataset.bullet ?? blocks[0]!.dataset.hadBullet
    if (cur === bullet) bullet = 'none'
  }
  for (const b of blocks) {
    if (bullet) {
      b.dataset.bullet = bullet
      if (bullet === 'char' && patch.bulletChar) b.dataset.bulletChar = patch.bulletChar
      else if (bullet !== 'char') delete b.dataset.bulletChar
    }
    if (patch.lineSpacingPct != null) {
      b.dataset.lineSpacingPct = String(patch.lineSpacingPct)
      // Rough live preview; the canvas re-lays out with the real metrics on commit
      b.style.lineHeight = String(patch.lineSpacingPct / 100)
    }
    if (patch.spaceBeforePt != null) b.dataset.spaceBeforePt = String(patch.spaceBeforePt)
    if (patch.spaceAfterPt != null) b.dataset.spaceAfterPt = String(patch.spaceAfterPt)
  }
  return true
}

/** Bullet-gallery highlight while editing: union of the live paragraph marks across the edit
 * root ('' = none, '#num' = numbered, glyph = char bullet). `undefined` = no uncommitted
 * paragraph-format change, the render tree is still accurate; `null` = unknowable (mixed, or a
 * re-toggled char bullet whose glyph lives only in the engine's uncommitted state). */
export function liveBulletChar(): string | null | undefined {
  const root = document.querySelector('[data-src-para]')?.parentElement
  if (!root) return undefined
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement => el instanceof HTMLElement && el.tagName === 'DIV',
  )
  if (!blocks.length || !blocks.some((b) => b.dataset.bullet)) return undefined
  const found = new Set<string>()
  for (const b of blocks) {
    const kind = b.dataset.bullet ?? b.dataset.hadBullet
    if (!kind || kind === 'none') {
      found.add('')
      continue
    }
    if (kind === 'number') {
      found.add('#num')
      continue
    }
    // char: explicit glyph from the gallery, engine default ('•') for a fresh bullet; a
    // paragraph whose original glyph never reached the DOM stays unknowable
    const glyph =
      b.dataset.bulletChar ??
      (b.dataset.bullet === 'char' && b.dataset.hadBullet == null ? '•' : null)
    if (glyph == null) return null
    found.add(glyph)
  }
  return found.size === 1 ? [...found][0]! : null
}

/** Paragraph alignment at the editing selection, read from the overlay DOM (execCommand
 * justify* products live only there until commit). Blocks intersecting the selection count —
 * the caret's block when collapsed; a block with no inline text-align falls back to the
 * root's, then 'left' (the engine default, so some alignment is always current).
 * undefined = no overlay mounted; null = mixed. */
export function liveAlign(): 'left' | 'center' | 'right' | 'justify' | null | undefined {
  const root = document.querySelector('[data-src-para]')?.parentElement
  if (!(root instanceof HTMLElement)) return undefined
  const rootAlign = cssAlign(root.style.textAlign)
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement => el instanceof HTMLElement && el.tagName === 'DIV',
  )
  if (!blocks.length) return rootAlign ?? 'left'
  const sel = window.getSelection()
  const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null
  const found = new Set<NonNullable<ReturnType<typeof cssAlign>>>()
  for (const b of blocks) {
    if (range && !range.intersectsNode(b)) continue
    found.add(cssAlign(b.style.textAlign) ?? rootAlign ?? 'left')
  }
  // selection outside the overlay (e.g. focus stolen by ribbon chrome): read all blocks
  if (!found.size)
    for (const b of blocks) found.add(cssAlign(b.style.textAlign) ?? rootAlign ?? 'left')
  return found.size === 1 ? [...found][0]! : null
}

/**
 * Font size increase/decrease while editing: execCommand('fontSize', 7) wraps the selection as a
 * placeholder, then <font size="7"> is replaced with a span whose size steps along the PowerPoint
 * ladder from the original px (inherited from the parent computed style) — extractParagraphs
 * reads back style.fontSize.
 */
export function resizeSelectionFont(dir: 1 | -1): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  const norm = parseFloat(root.dataset.norm ?? '') || 1
  document.execCommand('styleWithCSS', false, 'false')
  document.execCommand('fontSize', false, '7')
  const spans: HTMLElement[] = []
  root.querySelectorAll('font[size="7"]').forEach((f) => {
    const font = f as HTMLElement
    const basePx = parseFloat(window.getComputedStyle(font.parentElement ?? root).fontSize) || 18
    const pt = stepFontSizePt(pxToPt(basePx, norm), dir)
    const span = document.createElement('span')
    span.style.fontSize = `${(pt * 96 * norm) / 72}px`
    while (font.firstChild) span.appendChild(font.firstChild)
    font.replaceWith(span)
    // The resized text's canvas-measured advances are stale: free the enclosing/contained fragments
    releaseFragment(span.closest('[data-layout-fragment]'))
    span.querySelectorAll<HTMLElement>('[data-layout-fragment]').forEach(releaseFragment)
    spans.push(span)
  })
  reselectSpans(spans)
}

/** Next/previous ladder size; beyond the ladder ±10pt, clamped to 8~400 */
function stepFontSizePt(cur: number, dir: 1 | -1): number {
  const max = FONT_SIZES[FONT_SIZES.length - 1]!
  if (dir > 0) return cur >= max ? Math.min(400, cur + 10) : FONT_SIZES.find((s) => s > cur)!
  if (cur > max) return Math.max(max, cur - 10)
  for (let i = FONT_SIZES.length - 1; i >= 0; i--) if (FONT_SIZES[i]! < cur) return FONT_SIZES[i]!
  return FONT_SIZES[0]!
}

/** replaceWith kills the live selection — re-select the new spans so the highlight and repeated grow/shrink clicks survive */
function reselectSpans(spans: HTMLElement[]): void {
  if (!spans.length) return
  const sel = window.getSelection()
  if (!sel) return
  const range = document.createRange()
  range.setStartBefore(spans[0]!)
  range.setEndAfter(spans[spans.length - 1]!)
  sel.removeAllRanges()
  sel.addRange(range)
  saveEditSelection()
}

/** css font-family list → first family name (unquoted). Save/display only care about the preferred font. */
export function firstFontFamily(cssList: string): string {
  return (cssList.split(',')[0] ?? '').trim().replace(/^['"]|['"]$/g, '')
}

/**
 * Change the selection's font while editing: execCommand('fontName') on the selection, then unify
 * the product into span style.fontFamily = the display stack with CJK fallbacks (extractParagraphs
 * takes only the first item on commit).
 */
export function applySelectionFontFamily(family: string): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  document.execCommand('styleWithCSS', false, 'true')
  document.execCommand('fontName', false, family)
  // Handle both products: <font face="…"> (path where styleWithCSS doesn't apply) and span style
  root.querySelectorAll('font[face]').forEach((f) => {
    const font = f as HTMLElement
    const span = document.createElement('span')
    span.style.fontFamily = displayFontFamily(font.getAttribute('face') || family)
    while (font.firstChild) span.appendChild(font.firstChild)
    font.replaceWith(span)
  })
  root.querySelectorAll('span').forEach((s) => {
    if (s.style.fontFamily && firstFontFamily(s.style.fontFamily) === family)
      s.style.fontFamily = displayFontFamily(family)
  })
}

/** Set an absolute font size (pt) on the selection while editing: fontSize=7 placeholder then replaced by a px span (same as resizeSelectionFont). */
export function setSelectionFontSizePt(pt: number): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  const norm = parseFloat(root.dataset.norm ?? '') || 1
  const px = Math.min(400, Math.max(8, (pt * 96) / 72)) * norm
  document.execCommand('styleWithCSS', false, 'false')
  document.execCommand('fontSize', false, '7')
  const spans: HTMLElement[] = []
  root.querySelectorAll('font[size="7"]').forEach((f) => {
    const font = f as HTMLElement
    const span = document.createElement('span')
    span.style.fontSize = `${px}px`
    while (font.firstChild) span.appendChild(font.firstChild)
    font.replaceWith(span)
    spans.push(span)
  })
  reselectSpans(spans)
}

// Viewport px font size → model pt (divide back by viewport scale × autofit fontScale).
// Half-pt resolution, matching the ribbon size box (commitSizeDraft).
function pxToPt(px: number, norm: number): number {
  return Math.round(((px * 72) / (96 * norm)) * 2) / 2
}

function normalizeCss(c: string): string {
  if (/^#?[0-9A-Fa-f]{8}$/.test(c)) return `#${c.replace(/^#/, '').slice(0, 6)}`
  return c.startsWith('#') ? c : `#${c}`
}

function rgbToHex(rgb: string): string | undefined {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(rgb)
  if (!m) return undefined
  const h = (n: string) => parseInt(n, 10).toString(16).padStart(2, '0')
  return `#${h(m[1]!)}${h(m[2]!)}${h(m[3]!)}`.toUpperCase()
}
