/**
 * 3.2 DOM overlay text editing (run-level rich text) — a contentEditable stacked over the text
 * box takes over input (caret/selection/IME for free). Bold/italic/underline are triggered by the
 * ribbon font group (execCommand on the selection). On exit, walk the DOM to extract the
 * paragraph/run structure (each run's format preserved independently) and go through IPC editText.
 */
import React, { useEffect, useRef } from 'react'
import type { EditCaret } from './action-context'
import { formatAutoNum, DEFAULT_INSETS_EMU, emuToPx, isWideChar } from '@genoffice/pptx-render'
import type { GlyphRun, ShapeRenderNode, TextLine } from '@genoffice/pptx-render'
import type { EditParagraph, EditRun, LinkTargetOp } from '../shared/ipc'
import { decodeLinkTarget, encodeLinkTarget } from '../shared/run-link'
import { displayFontFamily, konvaBaselineDrop } from './konva-adapter'
import { ZOOM_PREVIEW_EVENT } from './zoom-preview'
import {
  applyFontSizeStep,
  DEFAULT_FONT_SIZE_PT,
  type FontSizeStep,
} from '@genoffice/pptx-ops/font-size'
import { bulletRunText } from './bullet-presets'

interface Props {
  node: ShapeRenderNode
  /** Viewport scale (RenderSlide.scale) — editing renders in viewport px; divided back out when committing to pt */
  scale: number
  onCommit: (paragraphs: EditParagraph[]) => void
  onCancel: () => void
  /** Tab/Shift+Tab (for table cell editing): commit current content and jump to the next/previous cell.
   * paragraphs=null means content unchanged (the host may skip committing and only jump). */
  onTabNav?: (paragraphs: EditParagraph[] | null, dir: 1 | -1) => void
  /** Ctrl+Enter (PowerPoint for Windows): commit and move to the next placeholder; the host
   * only passes it on non-mac platforms, where Ctrl+Enter falls back to a plain commit. */
  onNextPlaceholder?: (paragraphs: EditParagraph[] | null) => void
  /** Viewport coordinates of the double-click: select the word there when entering editing; defaults to caret at end */
  caretPoint?: EditCaret
  /** Entered by typing directly on a selected shape: select all, then replace the whole content with that character */
  replaceWith?: string
  /** Open with the whole text selected (F2, WordArt placeholder) */
  selectAll?: boolean
  /** ⌘/Ctrl+click on a linked run follows the link (slide jump / external url) */
  onFollowLink?: (target: LinkTargetOp) => void
  /** Edit-frame color (matches the canvas selection chrome: white on dark slide backgrounds) */
  frameColor?: string
  /** Canvas CSS zoom: the outline divides by it to keep a constant on-screen weight */
  zoom?: number
  /** Right-click inside the text: open the text context menu (the edit session stays alive) */
  onContextMenu?: (x: number, y: number, collapsed: boolean) => void
  /** Left press on the frame around the text (not on a line box) commits the edit and hands the press over as a shape drag */
  onFrameDrag?: (ev: MouseEvent) => void
}

/**
 * PowerPoint/WPS keep the click split while editing: a press on a laid-out line places the caret,
 * a press on the frame around the text grabs the shape. Line boxes come from the live DOM (the
 * layout under edit may already differ from the canvas), expanded by a small screen-px pad.
 * An empty body counts as text over the whole frame; an empty paragraph row is a caret line.
 */
export function pressOnEditFrame(
  editor: HTMLElement,
  target: EventTarget | null,
  x: number,
  y: number,
  padPx = 4,
): boolean {
  if (!editor.textContent?.trim()) return false
  if (!(target instanceof Node) || target === editor || !editor.contains(target)) return true
  let para: Node = target
  while (para.parentNode && para.parentNode !== editor) para = para.parentNode
  // An empty paragraph only has a <br> (zero-width rect): the whole row stays a caret line
  if (!para.textContent?.trim()) return false
  const range = document.createRange()
  range.selectNodeContents(para)
  const rects = Array.from(range.getClientRects()).filter((r) => r.width > 0)
  if (!rects.length) return false
  return !rects.some(
    (r) =>
      x >= r.left - padPx && x <= r.right + padPx && y >= r.top - padPx && y <= r.bottom + padPx,
  )
}

/** First-strong-character inference over a paragraph's logical text (mirrors what dir="auto" does). */
function inferParaRtl(paraLines: TextLine[]): boolean {
  const runs = paraLines
    .flatMap((l) => l.runs)
    .filter((r) => !r.isBullet)
    .sort((a, b) => (a.logicalOrder ?? 0) - (b.logicalOrder ?? 0))
  for (const r of runs) {
    for (const ch of r.text) {
      if (/[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]/.test(ch)) return true
      // eslint-disable-next-line no-misleading-character-class -- broad strong-LTR ranges; combining marks inside are irrelevant for a per-char strong-direction probe
      if (/[A-Za-z\u00c0-\u058f\u0900-\ud7ff\uf900-\ufdcf]/.test(ch)) return false
    }
  }
  return false
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
 * on wrap are restored after the previous fragment; <a:br/> soft breaks are restored as "\n" sentinels. */
interface EditorSeg {
  run?: GlyphRun
  srcRun?: number
  text: string
  /** Stacked cell of a vertical column (eaVert/wordArtVert): the engine's advance to the next cell (px) */
  stackAdvPx?: number
  /** Wrap-swallowed space restored between lines: must flow naturally, never get a fixed advance.
   * Its run's widthPx excludes this space (the engine strips it before measuring the line), so a
   * fixed-width fragment would clip it once an edit reflows it into the middle of a line. Under
   * pre-wrap it hangs invisibly at the original wrap point, so the untouched editor still matches
   * the canvas. */
  natural?: boolean
}

function editorParaRuns(paraLines: TextLine[], vertical = false): EditorSeg[] {
  const segs: EditorSeg[] = []
  paraLines.forEach((line, li) => {
    if (li > 0 && paraLines[li - 1]!.trailingSpace && segs.length) {
      const prev = segs[segs.length - 1]!
      const restored = paraLines[li - 1]!.trailingText ?? ' '
      if (!vertical && prev.run && !prev.run.rtl) {
        segs.push({ run: prev.run, srcRun: prev.srcRun, text: restored, natural: true })
      } else {
        prev.text += restored
      }
    }
    // Canvas consumes visual bidi order; contentEditable must receive logical source order and
    // lets Chromium perform bidi shaping/reordering itself.
    const logicalRuns = [...line.runs].sort(
      (a, b) =>
        (a.logicalOrder ?? Number.MAX_SAFE_INTEGER) - (b.logicalOrder ?? Number.MAX_SAFE_INTEGER),
    )
    for (const run of logicalRuns) {
      if (run.isBullet || run.text === '') continue
      const seg: EditorSeg = { run, srcRun: run.srcRunIdx, text: run.text }
      // Upright cells advance by the font's line box, not the glyph's 1em: read the pitch back from the column
      if (vertical && !run.rotate90 && !run.rotate270) {
        const cellTop = (r: GlyphRun) => r.baselineY - (r.ascentPx ?? r.fontSizePx * 0.8)
        const next = line.runs[line.runs.indexOf(run) + 1]
        seg.stackAdvPx = (next ? cellTop(next) : line.top + line.height) - cellTop(run)
      }
      segs.push(seg)
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

/**
 * Bullet preview while editing: the canvas hides this node's text (bullets included), so the
 * paragraph div draws its bullet as a ::before (see .slide-text-editor in styles.css) that
 * fills the hanging indent the layout reserved. Not part of the DOM, so extraction and the
 * caret never see it.
 */
function setEditorBullet(
  p: HTMLElement,
  b: {
    text: string
    image?: string
    font: string
    sizePx: number
    color: string
    bold?: boolean
    widthPx: number
    /** marL − bullet x: how far the ::before hangs into the left margin (widthPx − hangPx = push past marL) */
    hangPx: number
    /** Picture bullet box relative to the paragraph div (px); the ::before only reserves widthPx */
    imageBox?: { x: number; y: number; w: number; h: number }
  },
): void {
  p.style.setProperty('--bullet-w', `${b.widthPx}px`)
  p.style.setProperty('--bullet-hang', `${b.hangPx}px`)
  p.style.setProperty('--bullet-font', displayFontFamily(b.font))
  p.style.setProperty('--bullet-size', `${b.sizePx}px`)
  p.style.setProperty('--bullet-weight', b.bold ? 'bold' : 'normal')
  if (b.image && b.imageBox) {
    // An invisible glyph reserves the indent in the line (an empty inline-block collapses
    // inside the contentEditable); the ::after paints the image over it
    p.dataset.bulletText = '\u00a0'
    p.dataset.bulletImg = '1'
    p.style.setProperty('--bullet-color', 'transparent')
    p.style.setProperty('--bullet-img', `url("${b.image}")`)
    p.style.setProperty('--bullet-x', `${b.imageBox.x}px`)
    p.style.setProperty('--bullet-y', `${b.imageBox.y}px`)
    p.style.setProperty('--bullet-img-w', `${b.imageBox.w}px`)
    p.style.setProperty('--bullet-h', `${b.imageBox.h}px`)
  } else {
    p.dataset.bulletText = b.text
    delete p.dataset.bulletImg
    p.style.setProperty('--bullet-color', b.color)
  }
}

/** Re-derive the ::before after a ribbon bullet toggle on a paragraph div (marks set by applySelectionParagraphFormat). */
function refreshEditorBullet(b: HTMLElement, root: HTMLElement): void {
  const kind = b.dataset.bullet ?? b.dataset.hadBullet
  if (!kind || kind === 'none') {
    delete b.dataset.bulletText
    delete b.dataset.bulletImg
    return
  }
  // Toggled back on without a glyph pick: the original bullet preview still applies
  if (kind === 'char' && !b.dataset.bulletChar && b.dataset.bulletText && !b.dataset.bulletImg)
    return
  if (kind === 'blip' && !b.dataset.bulletImgSrc && b.dataset.bulletImg) return
  const sample = (b.querySelector('span, a') as HTMLElement | null) ?? b
  const cs = window.getComputedStyle(sample)
  const sizePx = parseFloat(b.style.getPropertyValue('--bullet-size')) || parseFloat(cs.fontSize)
  let widthPx = parseFloat(b.style.getPropertyValue('--bullet-w'))
  if (!widthPx) {
    // Fresh bullet: the engine will write PowerPoint's 0.3125" hanging indent
    widthPx = 22.5 * (parseFloat(root.dataset.norm ?? '') || 1)
    b.style.marginLeft = `${widthPx}px`
  }
  const hangPx = parseFloat(b.style.getPropertyValue('--bullet-hang')) || widthPx
  let text = '•'
  let font = cs.fontFamily
  if (kind === 'number') {
    // Same counting as the layout: consecutive numbered siblings of one scheme continue from
    // the first one's start number
    const schemeOf = (el: HTMLElement) =>
      el.dataset.numType ?? el.dataset.hadNumType ?? 'arabicPeriod'
    const scheme = schemeOf(b)
    let n = 0
    let first = b
    for (
      let prev = b.previousElementSibling as HTMLElement | null;
      prev &&
      (prev.dataset.bullet ?? prev.dataset.hadBullet) === 'number' &&
      schemeOf(prev) === scheme;
      prev = prev.previousElementSibling as HTMLElement | null
    ) {
      n++
      first = prev
    }
    const start = parseInt(first.dataset.startAt ?? first.dataset.hadStartAt ?? '', 10) || 1
    text = formatAutoNum(start + n, scheme)
  } else if (kind === 'blip' && b.dataset.bulletImgSrc) {
    // Fresh picture bullet: a cap-height square until the canvas lays it out on commit
    const h = sizePx * 0.75
    setEditorBullet(b, {
      text: '',
      image: b.dataset.bulletImgSrc,
      font,
      sizePx,
      color: cs.color,
      widthPx,
      hangPx,
      imageBox: { x: -hangPx, y: sizePx * 0.3, w: h, h },
    })
    return
  } else if (b.dataset.bulletChar) {
    text = bulletRunText(b.dataset.bulletChar, b.dataset.bulletFont)
    if (b.dataset.bulletFont) font = b.dataset.bulletFont
  }
  setEditorBullet(b, { text, font, sizePx, color: cs.color, widthPx, hangPx })
}

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
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
 * vertical = bodyPr vert editing (the contentEditable is in writing-mode: vertical-rl): layout
 * "lines" are columns whose top/height/advance are column metrics, so every horizontal-flow
 * baking (line-height, paragraph gaps, baseline/alignment compensation, fragment advances) is
 * skipped and the browser lays the text out vertically itself.
 * Exported so tests can do "layout → DOM → extractParagraphs" round-trip assertions.
 */
export function populateEditorDom(
  div: HTMLElement,
  lines: TextLine[],
  anchorDy = 0,
  innerW?: number,
  vertical = false,
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
    if (!vertical) {
      p.style.lineHeight = `${first.advance ?? first.height}px`
      const gap = first.top - prevEnd
      if (Math.abs(gap) > 0.01) p.style.marginTop = `${gap}px`
    }
    // The DOM block ends one advance below the last line top (external leading renders
    // inside the block, unlike the canvas) — margins of following paragraphs compensate
    prevEnd = last.top + (last.advance ?? last.height)
    // RTL paragraphs (Arabic/Hebrew) align in editing as on canvas: the browser sets direction
    // by the first strong character. An explicit a:pPr rtl can disagree with that inference
    // (TextLine.rtl carries the effective base) — then the browser needs an explicit dir
    const effRtl = first.rtl === true
    p.dir = effRtl === inferParaRtl(paraLines) ? 'auto' : effRtl ? 'rtl' : 'ltr'
    const align = paraLines[0]?.align
    if (align) p.style.textAlign = align
    // Body text starts at marL, exactly like the canvas (lists/indent used to snap to the
    // inset edge on entering edit); first-line indent applies only without a bullet
    const marL = (!vertical && first.marLPx) || 0
    if (marL) p.style.marginLeft = `${marL}px`
    const indentPx = (!vertical && first.indentPx) || 0
    if (indentPx && !first.runs.some((r) => r.isBullet)) p.style.textIndent = `${indentPx}px`
    // ── Glyph-position fidelity vs the canvas renderer ──
    // Vertical: the canvas draws the dominant run's baseline at
    // lineTop + engineAscent + konvaBaselineDrop (0 for resolved faces, the fallback
    // font's offset from the legacy 0.8em rule otherwise — see the adapter); CSS puts
    // the DOM baseline at half-leading + browser ascent. The difference is several px
    // on fallback-drawn CJK/serif or lnSpc ≠ 100% text and reads as the text jumping
    // when editing starts. Measure both sides and cancel the difference with a
    // relative offset (flow is unaffected).
    let engineAscent = 0
    let domBaseline = 0
    let dominant: GlyphRun | null = null
    for (const r of first.runs) {
      if (r.isBullet) continue // the glyph can be far larger than the text (buSzPct) and is not in the DOM flow
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
    // lnSpc>100%: the canvas pins glyphs to the slot bottom (leadAbove below the line top)
    const canvasBaseline = (first.leadAbove ?? 0) + engineAscent + drop
    const participants: Array<{ family: string; size: number; bold?: boolean; italic?: boolean }> =
      first.runs
        .filter((r) => !r.isBullet)
        .map((r) => ({
          family: displayFontFamily(r.fontFamily ?? ''),
          size: r.fontSizePx,
          bold: r.bold,
          italic: r.italic,
        }))
    if (strutFont) participants.push({ family: strutFont.family, size: strutFont.size })
    // Half-leading distributes over the CSS line-height, which is the advance
    // (box + external leading) when the font has an hhea lineGap
    const cssLineH = first.advance ?? first.height
    for (const f of participants) {
      if (!f.family || !f.size) continue
      const m = browserFontBox(f.family, f.size, f.bold, f.italic)
      if (!m.height) continue
      domBaseline = Math.max(domBaseline, (cssLineH - m.height) / 2 + m.ascent)
    }
    const dyFix =
      !vertical && canvasBaseline > 0 && domBaseline > 0 ? canvasBaseline - domBaseline : 0
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
      if (!marL && !vertical) p.style.marginLeft = `${level * 24}px`
    }
    // Original bullet kind, for the ribbon's toggle-off semantics while editing
    const bulletRun = first.runs.find((r) => r.isBullet)
    if (bulletRun) {
      p.dataset.hadBullet = bulletRun.numType ? 'number' : bulletRun.image ? 'blip' : 'char'
      if (bulletRun.numType) p.dataset.hadNumType = bulletRun.numType
      if (bulletRun.startAt != null) p.dataset.hadStartAt = String(bulletRun.startAt)
      const textRun = first.runs.find((r) => !r.isBullet)
      const textX = textRun ? textRun.x : bulletRun.x + bulletRun.widthPx
      // Mirrored (RTL) bullets sit right of the text and get no preview
      if (!vertical && textX > bulletRun.x) {
        setEditorBullet(p, {
          text: bulletRun.text,
          image: bulletRun.image,
          font: bulletRun.fontFamily,
          sizePx: bulletRun.fontSizePx,
          color: normalizeCss(bulletRun.color),
          bold: bulletRun.bold,
          widthPx: textX - bulletRun.x,
          hangPx: marL - bulletRun.x,
          ...(bulletRun.image
            ? {
                imageBox: {
                  x: bulletRun.x - marL,
                  y: bulletRun.baselineY - first.top - (bulletRun.ascentPx ?? bulletRun.fontSizePx),
                  w: bulletRun.widthPx,
                  h: bulletRun.ascentPx ?? bulletRun.fontSizePx,
                },
              }
            : {}),
        })
      }
    }
    for (const { run, srcRun, text, stackAdvPx, natural } of editorParaRuns(paraLines, vertical)) {
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
          // Text highlight: display-only (extraction never reads it back; the patch path keeps <a:highlight>)
          if (run.highlight) span.style.backgroundColor = normalizeCss(run.highlight)
          p.appendChild(span)
        }
        const fragment = document.createElement('span')
        fragment.dataset.layoutFragment = 'true'
        fragment.textContent = text
        // Each fragment occupies the exact advance measured by the layout engine. CJK is normally
        // one grapheme per fragment; Latin/SEA keep their script-aware token boundaries. RTL stays
        // in normal inline flow so Chromium can preserve joining and bidirectional shaping.
        // Vertical editing skips fixed advances entirely: the engine's widthPx is a horizontal
        // measure, and inline-block cells would break writing-mode glyph orientation. Restored
        // wrap-swallowed spaces (natural) also flow free: their run's widthPx excludes them.
        if (!run.rtl && !vertical && !natural) {
          fragment.style.display = 'inline-block'
          fragment.style.width = `${run.widthPx}px`
          // The UA sheet gives contentEditable overflow-wrap: break-word, and a fixed-width
          // cell is its own wrapping container: a token the browser draws 1px wider than the
          // engine measured ("APP" in a variable-weight CJK face) folded inside its cell into
          // stacked letters. Cells never wrap; overflow spills right like the canvas.
          fragment.style.whiteSpace = 'pre'
        }
        // Keep the browser editor visually aligned with the canvas renderer. This is display-only:
        // extraction intentionally preserves the source run's PPT letter spacing through srcRun.
        const ls = stackAdvPx != null ? stackAdvPx - run.fontSizePx : run.letterSpacingPx
        if (ls && Math.abs(ls) > 0.05) fragment.style.letterSpacing = `${ls}px`
        // CSS pins the ideographic em box (0.88em above the baseline) to the cell top; the canvas draws at cell top + ascent
        if (stackAdvPx != null && isWideChar(text.codePointAt(0) ?? 0)) {
          const drop = (run.ascentPx ?? run.fontSizePx * 0.8) - run.fontSizePx * 0.88
          if (Math.abs(drop) > 0.05) {
            fragment.style.position = 'relative'
            fragment.style.top = `${drop}px`
          }
        }
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
    fragment.style.whiteSpace = ''
  })
}

/** Release one fragment's fixed advance (stale once its text or format changes). */
function releaseFragment(el: Element | null | undefined): void {
  if (!(el instanceof HTMLElement) || el.dataset.layoutFragment !== 'true') return
  el.style.display = ''
  el.style.width = ''
  el.style.whiteSpace = ''
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
  onNextPlaceholder,
  caretPoint,
  replaceWith,
  selectAll,
  onFollowLink,
  frameColor = '#232425',
  zoom = 1,
  onFrameDrag,
  onContextMenu,
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
  // A shape with no text body yet: preview the body setText is about to create so the
  // content does not jump between typing and commit. Every fresh body gets the standard
  // bodyPr insets (createTextBody writes them unconditionally); only an autoshape also
  // gets the centered anchor/alignment — keep both halves in step with setText.
  const freshBody = !node.text && node.type === 'shape'
  const fresh = freshBody && !node.placeholder && !node.txBox
  const insets =
    node.text?.insets ??
    (freshBody
      ? {
          l: emuToPx(DEFAULT_INSETS_EMU.l, scale),
          t: emuToPx(DEFAULT_INSETS_EMU.t, scale),
          r: emuToPx(DEFAULT_INSETS_EMU.r, scale),
          b: emuToPx(DEFAULT_INSETS_EMU.b, scale),
        }
      : { l: 0, t: 0, r: 0, b: 0 })
  const anchor = node.text?.anchor ?? (fresh ? 'middle' : 'top')
  // Editing uses layout viewport px directly (including viewport scale and autofit fontScale): layout height =
  // visual height, so the edit box isn't inflated; on commit extractParagraphs divides by norm back to model pt
  const norm = scale * (node.text?.fontScale ?? 1) || 1
  const wrap = node.text?.wrap !== false
  // spAutoFit boxes follow their content live (the commit writes the same size back into the
  // model): a nowrap body grows in width, a wrapping one in height — PowerPoint's click-to-type box
  const growH = node.text?.autofit === 'resize'
  const growW = growH && !wrap
  // bodyPr vert: edit in a CSS vertical writing mode matching the canvas engine —
  // eaVert: vertical-rl/mixed (CJK upright, Latin rotated, columns right→left);
  // wordArtVert: vertical-lr/upright (every glyph upright, stacked, columns left→right);
  // vert: vertical-rl/sideways (whole block rotated 90° cw, CJK included);
  // vert270: sideways-lr (whole block rotated 90° ccw, lines flow left→right)
  const vertMode = node.text?.vert
  const vertText = !!vertMode
  // Modes whose line stacking runs left→right (the frame flexes as a plain row there)
  const vertLtr = vertMode === 'vert270' || vertMode === 'wordArtVert'
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
    // Same anchor offset as the engine (the dy text-layout bakes into line tops), removed back
    // during populate — the engine anchors against the glyph extent (inkBottom), so mirror it
    const extraH =
      Math.max(box.h - insets.t - insets.b, 1) -
      (node.text?.inkBottom ?? node.text?.contentHeight ?? 0)
    const anchorDy = anchor === 'middle' ? extraH / 2 : anchor === 'bottom' ? extraH : 0
    // Vertical: anchoring/overflow compensation are horizontal-flow corrections — the
    // row-reverse frame flex implements the (right-edge-anchored) flow instead
    populateEditorDom(
      div,
      node.text?.lines ?? [],
      vertText ? 0 : anchorDy,
      vertText ? undefined : box.w - insets.l - insets.r,
      vertText,
    )
    initialRef.current = JSON.stringify(extractParagraphs(div, norm))
    div.focus()
    const sel = window.getSelection()
    if (sel && (replaceWith || selectAll)) {
      // Select all; type-to-replace then swaps the whole content for the first typed character
      const range = document.createRange()
      range.selectNodeContents(div)
      sel.removeAllRanges()
      sel.addRange(range)
      if (replaceWith) document.execCommand('insertText', false, replaceWith)
      return
    }
    if (sel) {
      // Caret at the click point (double-click selects the word there); no coordinates/no hit → caret to end
      const hit = caretPoint ? document.caretRangeFromPoint(caretPoint.x, caretPoint.y) : null
      if (hit && div.contains(hit.startContainer)) {
        sel.removeAllRanges()
        sel.addRange(hit)
        if (caretPoint?.select === 'word') {
          const s = sel as Selection & {
            modify?: (alter: string, dir: string, granularity: string) => void
          }
          s.modify?.('move', 'backward', 'word')
          s.modify?.('extend', 'forward', 'word')
        }
      } else {
        const range = document.createRange()
        range.selectNodeContents(div)
        range.collapse(false)
        sel.removeAllRanges()
        sel.addRange(range)
      }
    }
  }, [node, norm, caretPoint, replaceWith, selectAll])

  const commit = () => {
    savedSel = null
    const div = ref.current
    if (!div) return onCancel()
    const paras = extractParagraphs(div, norm)
    if (JSON.stringify(paras) === initialRef.current) return onCancel()
    onCommit(paras)
  }

  /** Current content for a navigate-away hand-off; null when unchanged so the host can skip the commit */
  const changedParagraphs = (): EditParagraph[] | null | undefined => {
    const div = ref.current
    if (!div) return undefined
    const paras = extractParagraphs(div, norm)
    return JSON.stringify(paras) === initialRef.current ? null : paras
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
    // Height fixed to the shape box (unless spAutoFit grows it): overflowing content shows past it;
    // border uses outline (takes no layout space) so the inner usable size matches canvas layout exactly
    <div
      ref={frameRef}
      style={{
        position: 'absolute',
        left: box.x,
        top: box.y,
        width: growW && !vertText ? 'max-content' : box.w,
        minWidth: growW && !vertText ? box.w : undefined,
        height: growH && !vertText ? 'auto' : box.h,
        minHeight: growH && !vertText ? box.h : undefined,
        // Rotation only — the canvas counter-flips text in flipped shapes (NodeBody), so the editor must not mirror
        transform: `rotate(${box.rotationDeg ?? 0}deg)`,
        transformOrigin: 'center center',
        zIndex: 20,
        display: 'flex',
        // Vertical text: the row direction keeps the anchor mapping (top = the
        // flow-start edge, like the engine's anchoring)
        flexDirection: vertText ? (vertLtr ? 'row' : 'row-reverse') : 'column',
        justifyContent:
          anchor === 'middle' ? 'center' : anchor === 'bottom' ? 'flex-end' : 'flex-start',
        // 2 device px, zoom-compensated: the canvas is CSS-scaled, and 2 CSS px reads
        // twice as heavy on retina displays
        outline: `${2 / (globalThis.devicePixelRatio || 1) / Math.max(zoom, 0.1)}px solid ${frameColor}`,
      }}
      onMouseMove={(e) => {
        const frame = frameRef.current
        const div = ref.current
        if (!frame || !div || !onFrameDrag) return
        frame.style.cursor = pressOnEditFrame(div, e.target, e.clientX, e.clientY) ? 'move' : ''
      }}
      onMouseDown={(e) => {
        if (!onFrameDrag || e.button !== 0 || e.shiftKey || e.metaKey || e.ctrlKey) return
        const div = ref.current
        if (!div || !pressOnEditFrame(div, e.target, e.clientX, e.clientY)) return
        e.preventDefault()
        commit()
        onFrameDrag(e.nativeEvent)
      }}
    >
      <div
        ref={ref}
        className="slide-text-editor"
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
        onContextMenu={(e) => {
          if (!onContextMenu) return
          e.preventDefault()
          e.stopPropagation()
          saveEditSelection()
          onContextMenu(e.clientX, e.clientY, window.getSelection()?.isCollapsed ?? true)
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
          if (e.key === 'Escape' || e.key === 'F2') {
            // Esc / F2 = commit the text and return to shape-selected state (input not lost);
            // when unchanged, commit internally goes through onCancel and produces no history step
            e.preventDefault()
            commit()
          } else if (e.key === 'Enter' && e.ctrlKey && onNextPlaceholder) {
            e.preventDefault()
            const paras = changedParagraphs()
            if (paras !== undefined) onNextPlaceholder(paras)
          } else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
            e.preventDefault()
            commit()
          } else if (e.key === 'Tab' && onTabNav) {
            // Table cells: Tab commits and jumps to the next cell / Shift+Tab previous;
            // block the default focus move (otherwise blur falsely triggers commit-and-exit)
            e.preventDefault()
            const paras = changedParagraphs()
            if (paras !== undefined) onTabNav(paras, e.shiftKey ? -1 : 1)
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
          // visually snap left on entering edit) — the canvas centers within the box.
          // A growing box takes the box width alone: the 40px floor would inflate the
          // 19px click-to-type box and snap back on commit
          minWidth: wrap ? undefined : growW ? box.w : Math.max(box.w, 40 + insets.l + insets.r),
          // Only an empty body needs a synthetic height (so typing matches the canvas line
          // height); inflating a laid-out body distorts the flex vertical anchor — a
          // middle-anchored single line with tight spacing sat a few px too high in edit
          minHeight: node.text?.lines.length ? undefined : baseFontSize * 1.2 + insets.t + insets.b,
          // extractParagraphs reads the root alignment back, so the fresh-shape preview
          // is also what gets committed
          ...(fresh ? { textAlign: 'center' as const } : {}),
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
          // Vertical editing: the browser lays out real vertical text (upright CJK, rotated
          // Latin). The block axis is horizontal, so width follows content (columns) and the
          // row flex stretch pins the height to the box so wrap breaks columns at box height.
          ...(vertText
            ? {
                writingMode:
                  vertMode === 'vert270'
                    ? ('sideways-lr' as const)
                    : vertMode === 'wordArtVert'
                      ? ('vertical-lr' as const)
                      : ('vertical-rl' as const),
                textOrientation:
                  vertMode === 'vert'
                    ? ('sideways' as const)
                    : vertMode === 'wordArtVert'
                      ? ('upright' as const)
                      : ('mixed' as const),
                width: 'max-content',
                minWidth: undefined,
                minHeight: undefined,
                // Keep natural column width: overflow spills left (the flow direction), it
                // must not compress into extra column breaks
                flexShrink: 0,
              }
            : {}),
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
      if (
        ds?.bullet === 'char' ||
        ds?.bullet === 'number' ||
        ds?.bullet === 'blip' ||
        ds?.bullet === 'none'
      ) {
        curFmt.bullet = ds.bullet
        if (ds.bullet === 'char' && ds.bulletChar) {
          curFmt.bulletChar = ds.bulletChar
          if (ds.bulletFont) curFmt.bulletFont = ds.bulletFont
        }
        if (ds.bullet === 'number') {
          if (ds.numType) curFmt.numType = ds.numType
          const sa = parseInt(ds.startAt ?? '', 10)
          if (!Number.isNaN(sa)) curFmt.startAt = sa
        }
        if (ds.bullet === 'blip' && ds.bulletImgSrc && ds.bulletImgExt) {
          const comma = ds.bulletImgSrc.indexOf(',')
          curFmt.bulletImage = { base64: ds.bulletImgSrc.slice(comma + 1), ext: ds.bulletImgExt }
        }
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
      if (ds?.rtl === '1' || ds?.rtl === '0') curFmt.rtl = ds.rtl === '1'
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
  bullet?: 'char' | 'number' | 'blip' | 'none'
  bulletChar?: string
  bulletFont?: string
  numType?: string
  startAt?: number
  bulletImage?: { base64: string; ext: string }
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
  rtl?: boolean
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
  if (bullet && bullet !== 'none' && !patch.bulletChar && !patch.numType && !patch.bulletImage) {
    const cur = blocks[0]!.dataset.bullet ?? blocks[0]!.dataset.hadBullet
    if (cur === bullet) bullet = 'none'
  }
  for (const b of blocks) {
    if (bullet) {
      b.dataset.bullet = bullet
      if (bullet === 'char' && patch.bulletChar) {
        b.dataset.bulletChar = patch.bulletChar
        if (patch.bulletFont) b.dataset.bulletFont = patch.bulletFont
        else delete b.dataset.bulletFont
      } else if (bullet !== 'char') {
        delete b.dataset.bulletChar
        delete b.dataset.bulletFont
      }
      if (bullet === 'number') {
        if (patch.numType) b.dataset.numType = patch.numType
        if (patch.startAt != null) b.dataset.startAt = String(patch.startAt)
      } else {
        delete b.dataset.numType
        delete b.dataset.startAt
      }
      if (bullet === 'blip' && patch.bulletImage) {
        const { base64, ext } = patch.bulletImage
        b.dataset.bulletImgSrc = `data:${IMAGE_MIME[ext.toLowerCase()] ?? 'image/png'};base64,${base64}`
        b.dataset.bulletImgExt = ext
      } else if (bullet !== 'blip') {
        delete b.dataset.bulletImgSrc
        delete b.dataset.bulletImgExt
      }
      refreshEditorBullet(b, root)
    } else if (
      (patch.numType != null || patch.startAt != null) &&
      (b.dataset.bullet ?? b.dataset.hadBullet) === 'number'
    ) {
      // Standalone scheme / start number: only numbered paragraphs, committed as a number mark
      b.dataset.bullet = 'number'
      if (patch.numType) b.dataset.numType = patch.numType
      if (patch.startAt != null) b.dataset.startAt = String(patch.startAt)
      refreshEditorBullet(b, root)
    }
    if (patch.lineSpacingPct != null) {
      b.dataset.lineSpacingPct = String(patch.lineSpacingPct)
      // Rough live preview; the canvas re-lays out with the real metrics on commit
      b.style.lineHeight = String(patch.lineSpacingPct / 100)
    }
    if (patch.spaceBeforePt != null) b.dataset.spaceBeforePt = String(patch.spaceBeforePt)
    if (patch.spaceAfterPt != null) b.dataset.spaceAfterPt = String(patch.spaceAfterPt)
    if (patch.rtl != null) {
      b.dataset.rtl = patch.rtl ? '1' : '0'
      b.dir = patch.rtl ? 'rtl' : 'ltr' // live preview; the canvas re-lays out on commit
    }
  }
  const touched = new Set(blocks)
  // Start at belongs to the list, not the paragraph: PowerPoint stamps it on every consecutive
  // numbered paragraph of the scheme (the layout counter restarts on a differing startAt), so a
  // caret or partial selection extends it over the surrounding run
  if (patch.startAt != null) {
    const isNum = (el: Element | null): el is HTMLElement =>
      el instanceof HTMLElement &&
      el.tagName === 'DIV' &&
      (el.dataset.bullet ?? el.dataset.hadBullet) === 'number'
    const schemeOf = (el: HTMLElement) =>
      el.dataset.numType ?? el.dataset.hadNumType ?? 'arabicPeriod'
    const seeds = blocks.filter(isNum)
    if (seeds.length) {
      const scheme = schemeOf(seeds[0]!)
      const extend = (from: HTMLElement, dir: 'previousElementSibling' | 'nextElementSibling') => {
        for (let el = from[dir]; isNum(el) && schemeOf(el) === scheme; el = el[dir]) {
          if (touched.has(el)) continue
          el.dataset.bullet = 'number'
          el.dataset.startAt = String(patch.startAt)
          touched.add(el)
          refreshEditorBullet(el, root)
        }
      }
      extend(seeds[0]!, 'previousElementSibling')
      extend(seeds[seeds.length - 1]!, 'nextElementSibling')
      // Each preview counts from the run's first sibling, which the walk may have just stamped
      for (const el of touched) if (isNum(el)) refreshEditorBullet(el, root)
    }
  }
  // Numbers count from earlier siblings, so a kind/scheme/start change re-numbers the
  // numbered paragraphs below the selection too — only those that already preview a
  // ::before (mirrored RTL / vertical paragraphs never got one and must not gain a margin)
  if (bullet || patch.numType != null || patch.startAt != null) {
    for (const el of Array.from(root.children)) {
      if (!(el instanceof HTMLElement) || el.tagName !== 'DIV' || touched.has(el)) continue
      if ((el.dataset.bullet ?? el.dataset.hadBullet) === 'number' && el.dataset.bulletText != null)
        refreshEditorBullet(el, root)
    }
  }
  return true
}

/** Effective base direction at the editing selection, read from the overlay DOM (computed
 * direction covers dir="auto" inference and explicit toggles alike). undefined = no overlay
 * mounted; null = mixed. */
export function liveRtl(): boolean | null | undefined {
  const root = document.querySelector('[data-src-para]')?.parentElement
  if (!(root instanceof HTMLElement)) return undefined
  const blocks = Array.from(root.children).filter(
    (el): el is HTMLElement => el instanceof HTMLElement && el.tagName === 'DIV',
  )
  if (!blocks.length) return false
  const sel = window.getSelection()
  const range = sel && sel.rangeCount ? sel.getRangeAt(0) : null
  const found = new Set<boolean>()
  for (const b of blocks) {
    if (range && !range.intersectsNode(b)) continue
    found.add(window.getComputedStyle(b).direction === 'rtl')
  }
  if (!found.size) for (const b of blocks) found.add(window.getComputedStyle(b).direction === 'rtl')
  return found.size === 1 ? [...found][0]! : null
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
      found.add(`#num:${b.dataset.numType ?? b.dataset.hadNumType ?? 'arabicPeriod'}`)
      continue
    }
    if (kind === 'blip') {
      found.add('#img')
      continue
    }
    // char: explicit glyph from the gallery, engine default ('•') for a fresh bullet; a
    // paragraph whose original glyph never reached the DOM stays unknowable
    const glyph = b.dataset.bulletChar
      ? bulletRunText(b.dataset.bulletChar, b.dataset.bulletFont)
      : b.dataset.bullet === 'char' && b.dataset.hadBullet == null
        ? '•'
        : null
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
 * ladder (or by one point) from the original px (inherited from the parent computed style) —
 * extractParagraphs reads back style.fontSize.
 */
export function resizeSelectionFont(dir: 1 | -1, mode: FontSizeStep['mode'] = 'ladder'): void {
  const root = document.activeElement
  if (!(root instanceof HTMLElement) || !root.isContentEditable) return
  const norm = parseFloat(root.dataset.norm ?? '') || 1
  // When the selection exactly covers a sized span (our own product from the
  // previous click), execCommand replaces that span with the <font> wrapper —
  // the current size is gone before we can read it, and the parent computed
  // style still holds the pre-step size, so every further click would restep
  // from the same base. Snapshot each selected text node's size first: the
  // wrap moves text nodes intact, so they key the snapshot across the reflow.
  const selRange = (() => {
    const sel = window.getSelection()
    return sel && sel.rangeCount ? sel.getRangeAt(0) : null
  })()
  const preSizes = new Map<Node, number>()
  if (selRange) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    while (walker.nextNode()) {
      const node = walker.currentNode
      if (!selRange.intersectsNode(node) || !node.parentElement) continue
      const px = parseFloat(window.getComputedStyle(node.parentElement).fontSize)
      if (Number.isFinite(px)) preSizes.set(node, px)
    }
  }
  document.execCommand('styleWithCSS', false, 'false')
  document.execCommand('fontSize', false, '7')
  const spans: HTMLElement[] = []
  root.querySelectorAll('font[size="7"]').forEach((f) => {
    const font = f as HTMLElement
    let basePx: number | undefined
    const walker = document.createTreeWalker(font, NodeFilter.SHOW_TEXT)
    while (walker.nextNode()) {
      const snap = preSizes.get(walker.currentNode)
      if (snap !== undefined) {
        basePx = snap
        break
      }
    }
    basePx ??=
      parseFloat(window.getComputedStyle(font.parentElement ?? root).fontSize) ||
      (DEFAULT_FONT_SIZE_PT * 96) / 72
    const pt = applyFontSizeStep(pxToPt(basePx, norm), { dir, mode })
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
