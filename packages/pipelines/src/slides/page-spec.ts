/**
 * Local single-page generation: a structured JSON slide spec (written by an
 * LLM through the app's own AI transport) is built directly into a one-slide
 * PPTX with pptx-engine primitives — no HTML intermediate, no conversion step.
 *
 * The spec's element model mirrors what an editable deck needs (and what
 * Genspark's gen_pptx capture emits): absolutely positioned shapes, images
 * (center-cropped to their frame) and text runs on a fixed px canvas.
 *
 * Host facilities (network fetch, image decoding, font metrics) are injected so
 * this module stays testable in plain Node and usable from the genoffice CLI.
 */
import {
  addElement,
  addPicture,
  createBlankPptx,
  editPictureSrcRect,
  openPptx,
  promoteSlideBackground,
  savePptx,
  type Paragraph,
  type TextElement,
  type TextRun,
} from '@genoffice/pptx-engine'
import { buildRenderSlide, EMU_PER_PX_96, type FontMetricsProvider } from '@genoffice/pptx-render'
import { coverCropFractions } from './cover-crop'

export const SPEC_CANVAS_W = 1280
export const SPEC_CANVAS_H = 720

const MAX_ELEMENTS = 48
const MAX_IMAGES = 8
const MAX_TEXT_LEN = 4000

/** Preset geometries the spec may use; unknown kinds fall back to rect instead of emitting invalid prst XML. */
const SHAPE_KINDS = new Set([
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'rightArrow',
  'leftArrow',
  'upArrow',
  'downArrow',
  'chevron',
  'diamond',
  'parallelogram',
  'trapezoid',
  'hexagon',
  'pentagon',
  'pie',
  'donut',
  'star5',
  'heart',
  'cloud',
  'line',
  'lineArrow',
])

export interface SpecRun {
  text: string
  sizePt?: number
  bold?: boolean
  italic?: boolean
  color?: string
  font?: string
}

export interface SpecParagraph {
  runs: SpecRun[]
  align?: 'left' | 'center' | 'right' | 'justify'
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
  bullet?: boolean
}

interface SpecBase {
  x: number
  y: number
  w: number
  h: number
}

export interface SpecShape extends SpecBase {
  type: 'shape'
  shape: string
  fill?: string
  stroke?: { color: string; widthPt: number }
  paragraphs?: SpecParagraph[]
  valign?: 'top' | 'middle' | 'bottom'
}

export interface SpecText extends SpecBase {
  type: 'text'
  paragraphs: SpecParagraph[]
  valign?: 'top' | 'middle' | 'bottom'
}

export interface SpecImage extends SpecBase {
  type: 'image'
  url: string
}

export type SpecElement = SpecShape | SpecText | SpecImage

export interface PageSpec {
  background?: string
  elements: SpecElement[]
}

const EMU_PER_PT = 12700

function asRecord(v: unknown): Record<string, unknown> {
  return v && typeof v === 'object' ? (v as Record<string, unknown>) : {}
}

/** #RGB / #RRGGBB / #RRGGBBAA → normalized #RRGGBB(AA), else undefined */
function normColor(v: unknown): string | undefined {
  if (typeof v !== 'string') return undefined
  let hex = v.trim().replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{3}$/.test(hex)) hex = [...hex].map((c) => c + c).join('')
  return /^[0-9A-F]{6}([0-9A-F]{2})?$/.test(hex) ? `#${hex}` : undefined
}

function num(v: unknown): number | undefined {
  const n = typeof v === 'string' ? Number(v) : v
  return typeof n === 'number' && Number.isFinite(n) ? n : undefined
}

/**
 * Extracts and validates the spec from raw LLM output. Tolerant of fences and
 * junk around the JSON; invalid elements are dropped with a warning rather
 * than failing the page. Returns an error only when nothing usable remains,
 * phrased so it can be fed back to the model for a corrected attempt.
 */
export interface ParseSpecOptions {
  /** Accept local file paths and data: URLs as image sources (the CLI); default http(s) only. */
  localImages?: boolean
}

export function parsePageSpec(
  raw: string,
  canvasW = SPEC_CANVAS_W,
  canvasH = SPEC_CANVAS_H,
  opts: ParseSpecOptions = {},
): { ok: true; spec: PageSpec; warnings: string[] } | { ok: false; error: string } {
  const text = String(raw ?? '')
  const start = text.indexOf('{')
  const end = text.lastIndexOf('}')
  if (start < 0 || end <= start) return { ok: false, error: 'no JSON object found in the output' }
  let parsed: unknown
  try {
    parsed = JSON.parse(text.slice(start, end + 1))
  } catch (e) {
    return { ok: false, error: `invalid JSON: ${e instanceof Error ? e.message : String(e)}` }
  }
  return parsePageSpecObject(parsed, canvasW, canvasH, opts)
}

/** Same validation as parsePageSpec for an already-parsed JSON value (one page of a deck spec). */
export function parsePageSpecObject(
  parsed: unknown,
  canvasW = SPEC_CANVAS_W,
  canvasH = SPEC_CANVAS_H,
  opts: ParseSpecOptions = {},
): { ok: true; spec: PageSpec; warnings: string[] } | { ok: false; error: string } {
  const root = asRecord(parsed)
  const rawEls = Array.isArray(root.elements) ? root.elements : []
  if (rawEls.length === 0) return { ok: false, error: 'the "elements" array is missing or empty' }

  const warnings: string[] = []
  const elements: SpecElement[] = []
  let images = 0

  const parseParagraphs = (v: unknown): SpecParagraph[] => {
    if (!Array.isArray(v)) return []
    const out: SpecParagraph[] = []
    for (const p of v) {
      const pr = asRecord(p)
      const runsRaw = Array.isArray(pr.runs) ? pr.runs : []
      const runs: SpecRun[] = []
      for (const r of runsRaw) {
        const rr = asRecord(r)
        const t = typeof rr.text === 'string' ? rr.text.slice(0, MAX_TEXT_LEN) : ''
        const sizePt = num(rr.sizePt)
        runs.push({
          text: t,
          ...(sizePt ? { sizePt: Math.min(Math.max(sizePt, 6), 160) } : {}),
          ...(rr.bold === true ? { bold: true } : {}),
          ...(rr.italic === true ? { italic: true } : {}),
          ...(normColor(rr.color) ? { color: normColor(rr.color) } : {}),
          ...(typeof rr.font === 'string' && rr.font.trim()
            ? { font: rr.font.trim().slice(0, 80) }
            : {}),
        })
      }
      if (runs.length === 0) runs.push({ text: '' })
      const align = pr.align
      const lineSpacingPct = num(pr.lineSpacingPct)
      const spaceBeforePt = num(pr.spaceBeforePt)
      const spaceAfterPt = num(pr.spaceAfterPt)
      out.push({
        runs,
        ...(align === 'left' || align === 'center' || align === 'right' || align === 'justify'
          ? { align }
          : {}),
        ...(lineSpacingPct ? { lineSpacingPct: Math.min(Math.max(lineSpacingPct, 60), 300) } : {}),
        ...(spaceBeforePt !== undefined
          ? { spaceBeforePt: Math.min(Math.max(spaceBeforePt, 0), 96) }
          : {}),
        ...(spaceAfterPt !== undefined
          ? { spaceAfterPt: Math.min(Math.max(spaceAfterPt, 0), 96) }
          : {}),
        ...(pr.bullet === true ? { bullet: true } : {}),
      })
    }
    return out
  }

  for (const [i, rawEl] of rawEls.entries()) {
    if (elements.length >= MAX_ELEMENTS) {
      warnings.push(`element cap ${MAX_ELEMENTS} reached; the rest were dropped`)
      break
    }
    const el = asRecord(rawEl)
    const x = num(el.x)
    const y = num(el.y)
    const w = num(el.w)
    const h = num(el.h)
    if (x === undefined || y === undefined || w === undefined || h === undefined) {
      warnings.push(`element ${i}: missing/non-numeric x/y/w/h, dropped`)
      continue
    }
    // Clamp into the canvas; drop elements whose origin already lies outside
    if (x >= canvasW || y >= canvasH || x + w <= 0 || y + h <= 0) {
      warnings.push(`element ${i}: outside the ${canvasW}x${canvasH} canvas, dropped`)
      continue
    }
    const cx = Math.max(0, Math.min(x, canvasW - 1))
    const cy = Math.max(0, Math.min(y, canvasH - 1))
    const cw = Math.max(0, Math.min(w - (cx - x), canvasW - cx))
    const ch = Math.max(0, Math.min(h - (cy - y), canvasH - cy))
    if (cw < 1 || ch < 1) {
      warnings.push(`element ${i}: outside the ${canvasW}x${canvasH} canvas, dropped`)
      continue
    }
    const base = { x: cx, y: cy, w: cw, h: ch }
    const type = el.type

    if (type === 'image') {
      const url = typeof el.url === 'string' ? el.url.trim() : ''
      if (!/^https?:\/\//.test(url) && !(opts.localImages && url)) {
        warnings.push(`element ${i}: image url must be http(s), dropped`)
        continue
      }
      if (images >= MAX_IMAGES) {
        warnings.push(`element ${i}: image cap ${MAX_IMAGES} reached, dropped`)
        continue
      }
      images += 1
      elements.push({ type: 'image', url, ...base })
      continue
    }

    if (type === 'text') {
      const paragraphs = parseParagraphs(el.paragraphs)
      if (paragraphs.length === 0 || paragraphs.every((p) => !p.runs.some((r) => r.text.trim()))) {
        warnings.push(`element ${i}: text element without any text, dropped`)
        continue
      }
      const valign = el.valign
      elements.push({
        type: 'text',
        ...base,
        paragraphs,
        ...(valign === 'top' || valign === 'middle' || valign === 'bottom' ? { valign } : {}),
      })
      continue
    }

    if (type === 'shape') {
      let shape = typeof el.shape === 'string' ? el.shape.trim() : 'rect'
      if (!SHAPE_KINDS.has(shape)) {
        warnings.push(`element ${i}: unknown shape "${shape}", using rect`)
        shape = 'rect'
      }
      const fill = normColor(el.fill)
      const strokeRec = asRecord(el.stroke)
      const strokeColor = normColor(strokeRec.color)
      const strokeWidth = num(strokeRec.widthPt)
      const stroke = strokeColor
        ? { color: strokeColor, widthPt: Math.min(Math.max(strokeWidth ?? 1, 0.25), 24) }
        : undefined
      const isLine = shape === 'line' || shape === 'lineArrow'
      if (!fill && !stroke && !isLine) {
        warnings.push(`element ${i}: shape without fill or stroke, dropped`)
        continue
      }
      const paragraphs = parseParagraphs(el.paragraphs)
      const valign = el.valign
      elements.push({
        type: 'shape',
        shape,
        ...base,
        ...(fill ? { fill } : {}),
        ...(stroke ? { stroke } : {}),
        ...(paragraphs.some((p) => p.runs.some((r) => r.text.trim())) ? { paragraphs } : {}),
        ...(valign === 'top' || valign === 'middle' || valign === 'bottom' ? { valign } : {}),
      })
      continue
    }

    warnings.push(`element ${i}: unknown type "${String(type)}", dropped`)
  }

  if (elements.length === 0) {
    return {
      ok: false,
      error: `no valid elements (${warnings.join('; ') || 'all dropped'})`,
    }
  }
  warnings.push(...nearDuplicateTextWarnings(rawEls))
  return {
    ok: true,
    spec: {
      ...(normColor(root.background) ? { background: normColor(root.background) } : {}),
      elements,
    },
    warnings,
  }
}

/** Text of a raw spec element for the duplicate check, or null when it has none. */
function rawElementText(el: unknown): string | null {
  const rec = asRecord(el)
  if (rec.type !== 'text' && rec.type !== 'shape') return null
  if (!Array.isArray(rec.paragraphs)) return null
  const text = rec.paragraphs
    .flatMap((p) => (Array.isArray(asRecord(p).runs) ? (asRecord(p).runs as unknown[]) : []))
    .map((r) => (typeof asRecord(r).text === 'string' ? (asRecord(r).text as string) : ''))
    .join('')
  return text.trim() ? text : null
}

/** Lower-cased text without spaces or punctuation, so "A · B" and "A（B）" compare equal. */
function normalizeText(text: string): string {
  return text.toLowerCase().replace(/[\s\p{P}\p{S}]+/gu, '')
}

const DUP_MIN_CHARS = 12
const DUP_MIN_PREFIX = 10
const DUP_PREFIX_SHARE = 0.6

/**
 * Two text boxes on one page that say (nearly) the same thing are almost
 * always an authoring slip — a subtitle restating the chart caption, a
 * heading pasted twice. Compared on normalized text: identical, or sharing a
 * long common prefix that covers most of the shorter one. Advice only.
 */
export function nearDuplicateTextWarnings(rawEls: unknown[]): string[] {
  const texts: { i: number; text: string; norm: string }[] = []
  for (const [i, el] of rawEls.entries()) {
    const text = rawElementText(el)
    if (text === null) continue
    const norm = normalizeText(text)
    if (norm.length >= DUP_MIN_CHARS) texts.push({ i, text, norm })
  }
  const out: string[] = []
  for (let a = 0; a < texts.length; a++) {
    for (let b = a + 1; b < texts.length; b++) {
      const x = texts[a]!
      const y = texts[b]!
      const shorter = Math.min(x.norm.length, y.norm.length)
      let prefix = 0
      while (prefix < shorter && x.norm[prefix] === y.norm[prefix]) prefix++
      const same = x.norm === y.norm
      if (!same && (prefix < DUP_MIN_PREFIX || prefix < shorter * DUP_PREFIX_SHARE)) continue
      const quote = (t: string) => (t.length > 40 ? `${t.slice(0, 40)}…` : t)
      out.push(
        `elements ${x.i} and ${y.i}: ${same ? 'identical' : 'near-duplicate'} text ("${quote(x.text)}" / "${quote(y.text)}"); merge them or make one say something else`,
      )
    }
  }
  return out
}

export interface BuildPageDeps {
  /** Downloads an image; null on failure (page continues without it) */
  fetchImage: (url: string) => Promise<{ bytes: Uint8Array; ext: string } | null>
  /** Decodes natural pixel size for cover-cropping; null skips the crop */
  imageDims?: (bytes: Uint8Array) => { width: number; height: number } | null
  /** Font metrics for the post-build text measurement; absent skips the box-height fix */
  fontMetrics?: FontMetricsProvider
}

/**
 * The LLM sizes text boxes from a rough chars-per-line heuristic, which routinely
 * undersizes big CJK titles; the box has no autofit, so the canvas draws the overflow
 * past the selection frame (and PowerPoint past the shape). Re-measure every text box
 * with the real layout engine — on the reopened (parsed) model, the exact input the
 * landed page will render from — and grow too-short boxes to their content height.
 * Grow-only, plain text boxes only (shape label boxes are design intent); middle/bottom
 * anchored boxes shift up so the rendered glyphs stay exactly where they were.
 * Returns the re-saved bytes, or null when every box already fits.
 */
async function growTextBoxesToContent(
  bytes: Uint8Array,
  metrics: FontMetricsProvider,
): Promise<Uint8Array | null> {
  const opened = await openPptx(bytes)
  const slide = opened.deck.slides[0]
  if (!slide) return null
  const baseWidthPx = opened.deck.size.cx / EMU_PER_PX_96 // native px → vp.scale = 1
  const rendered = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: baseWidthPx, metrics })
  let changed = false
  for (const node of rendered.nodes) {
    if (node.type !== 'text' || !node.text) continue
    const el = slide.elements.find((e) => e.id === node.sourceId)
    if (el?.type !== 'text') continue
    const t = node.text
    const needH = Math.max(t.contentHeight, t.inkBottom ?? 0) + t.insets.t + t.insets.b
    const growPx = needH - node.box.h
    if (growPx < 0.5) continue
    const tel = el as TextElement
    const offset = { ...tel.transform.offset, cy: Math.max(1, Math.round(needH * EMU_PER_PX_96)) }
    if (t.anchor === 'middle') offset.y -= Math.round((growPx / 2) * EMU_PER_PX_96)
    else if (t.anchor === 'bottom') offset.y -= Math.round(growPx * EMU_PER_PX_96)
    tel.transform = { ...tel.transform, offset }
    tel.dirtyTransform = true
    changed = true
  }
  return changed ? savePptx(opened) : null
}

function toEngineParagraphs(paragraphs: SpecParagraph[]): Paragraph[] {
  return paragraphs.map((p) => {
    const runs: TextRun[] = p.runs.map((r) => ({
      text: r.text,
      ...(r.sizePt ? { fontSize: r.sizePt } : {}),
      ...(r.bold ? { bold: true } : {}),
      ...(r.italic ? { italic: true } : {}),
      ...(r.color ? { color: r.color } : {}),
      ...(r.font ? { fontFamily: r.font, latinFont: r.font, eaFont: r.font } : {}),
    }))
    return {
      runs,
      ...(p.align ? { align: p.align } : {}),
      ...(p.lineSpacingPct ? { lineHeight: p.lineSpacingPct } : {}),
      ...(p.spaceBeforePt !== undefined ? { spaceBefore: p.spaceBeforePt } : {}),
      ...(p.spaceAfterPt !== undefined ? { spaceAfter: p.spaceAfterPt } : {}),
      ...(p.bullet
        ? { bullet: { type: 'char' as const, char: '•' }, marL: 228600, indent: -228600 }
        : {}),
    }
  })
}

/**
 * Builds a one-slide PPTX from the spec. Elements are added in array order
 * (spec order = z-order); a background color becomes a full-bleed rect that
 * the landing pipeline's promoteSlideBackground lifts to the slide background.
 */
export async function buildPagePptx(
  spec: PageSpec,
  deps: BuildPageDeps,
  canvasW = SPEC_CANVAS_W,
): Promise<{ bytes: Uint8Array; imageFailures: string[]; imageWarnings: string[] }> {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const scale = opened.deck.size.cx / canvasW
  const toEmu = (px: number) => Math.round(px * scale)
  const anchorOf = (
    v: 'top' | 'middle' | 'bottom' | undefined,
    dflt: 't' | 'ctr',
  ): 't' | 'ctr' | 'b' => (v === 'top' ? 't' : v === 'middle' ? 'ctr' : v === 'bottom' ? 'b' : dflt)
  // Zero insets: the spec's boxes are exact; PowerPoint's default 0.1in/0.05in
  // insets would shift every text off its planned spot.
  const zeroInsets = { l: 0, t: 0, r: 0, b: 0 }

  const imageFailures: string[] = []
  const imageWarnings: string[] = []
  const fetched = new Map<string, { bytes: Uint8Array; ext: string } | null>()
  await Promise.all(
    [
      ...new Set(spec.elements.filter((e) => e.type === 'image').map((e) => (e as SpecImage).url)),
    ].map(async (url) => {
      try {
        fetched.set(url, await deps.fetchImage(url))
      } catch {
        fetched.set(url, null)
      }
    }),
  )

  if (spec.background) {
    addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: opened.deck.size.cx, cy: opened.deck.size.cy },
      fillColor: spec.background,
    })
  }

  for (const el of spec.elements) {
    const offset = {
      x: toEmu(el.x),
      y: toEmu(el.y),
      cx: Math.max(1, toEmu(el.w)),
      cy: Math.max(1, toEmu(el.h)),
    }
    if (el.type === 'image') {
      const img = fetched.get(el.url)
      if (!img) {
        imageFailures.push(el.url)
        continue
      }
      const pic = addPicture(opened, slide, { bytes: img.bytes, ext: img.ext, offset })
      if (!pic) {
        imageFailures.push(el.url)
        continue
      }
      const dims = deps.imageDims?.(img.bytes) ?? null
      if (dims) {
        const crop = coverCropFractions(dims.width, dims.height, el.w, el.h)
        if (crop) {
          editPictureSrcRect(slide, pic.id, crop)
          const warning = heavyCropWarning(el, dims, crop)
          if (warning) imageWarnings.push(warning)
        }
      }
      continue
    }
    if (el.type === 'text') {
      addElement(slide, {
        kind: 'textbox',
        offset,
        paragraphs: toEngineParagraphs(el.paragraphs),
        bodyPr: { wrap: 'square', anchor: anchorOf(el.valign, 't'), insetsEmu: zeroInsets },
      })
      continue
    }
    addElement(slide, {
      kind: el.shape,
      offset,
      ...(el.fill ? { fillColor: el.fill } : {}),
      ...(el.stroke
        ? {
            stroke: {
              color: el.stroke.color,
              widthEmu: Math.round(el.stroke.widthPt * EMU_PER_PT),
            },
          }
        : {}),
      ...(el.paragraphs ? { paragraphs: toEngineParagraphs(el.paragraphs) } : {}),
      ...(el.paragraphs
        ? { bodyPr: { wrap: 'square', anchor: anchorOf(el.valign, 'ctr'), insetsEmu: zeroInsets } }
        : {}),
    })
  }

  promoteSlideBackground(slide, opened.deck.size)
  let bytes = await savePptx(opened)
  if (deps.fontMetrics) bytes = (await growTextBoxesToContent(bytes, deps.fontMetrics)) ?? bytes
  return { bytes, imageFailures, imageWarnings }
}

/** Below this share of the source left visible, the center crop is reported (a square photo in a 16:9 banner keeps 56% and passes). */
export const HEAVY_CROP_KEEP = 0.5

/**
 * A box whose aspect is far from the picture's throws most of the picture
 * away: a 1200×1500 product shot in a 314×96 strip shows a quarter of it, cut
 * through the middle. The build still lands it (cover-crop never distorts),
 * but the author should hear about it.
 */
function heavyCropWarning(
  el: SpecImage,
  dims: { width: number; height: number },
  crop: { l: number; t: number; r: number; b: number },
): string | null {
  const kept = (1 - crop.l - crop.r) * (1 - crop.t - crop.b)
  if (kept >= HEAVY_CROP_KEEP) return null
  const axis = crop.t > 0 ? 'top and bottom' : 'left and right'
  return `image ${el.url}: the ${Math.round(el.w)}×${Math.round(el.h)} box shows only ${Math.round(kept * 100)}% of the ${dims.width}×${dims.height} picture (center crop, ${axis} cut off); size the box nearer the picture's aspect ratio or crop the file first`
}
