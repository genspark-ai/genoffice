import { existsSync, readFileSync } from 'node:fs'
import { extname } from 'node:path'
import { assertAllowed, type PathContext } from '../fs'
import { isAbsolute, resolve } from 'node:path'
import {
  builtinTableStyleName,
  createBlankPptx,
  getSlideLinks,
  getSlideNotes,
  listSlideLayouts,
  openPptx,
  savePptx,
  type LinkTarget,
  type OpenedPptx,
  type Paragraph,
  type Slide,
  type SlideElement,
} from '@genoffice/pptx-engine'
import {
  elementDurableId,
  listSlideAnimations,
  runTxn,
  slideDurableId,
  type AnimationEntry,
  type Op,
  type TxnResult,
} from '@genoffice/pptx-ops'
import { CliError, EXIT } from '../result'
import { clipText } from '../preview'
import { fromEmu, type ReadUnit } from '../length-units'

const EMU_PER_INCH = 914400
/** Image/media ops that take bytes, with the (possibly nested) field each one reads. A local file path there is read for the caller. */
const BYTES_FIELDS: Record<string, string[]> = {
  addPicture: ['bytes'],
  replacePicture: ['bytes'],
  addMedia: ['bytes', 'poster.bytes'],
  setImageFill: ['source.bytes'],
  setBackground: ['source.bytes'],
}

export async function openDeck(bytes: Uint8Array): Promise<OpenedPptx> {
  return openPptx(bytes)
}

export async function blankDeck(): Promise<OpenedPptx> {
  return openPptx(await createBlankPptx())
}

export function saveDeck(opened: OpenedPptx): Promise<Uint8Array> {
  return savePptx(opened)
}

export function parseOps(text: string, source: string): Op[] {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new CliError(
      EXIT.usage,
      `${source}: not valid JSON (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  const ops = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!ops) {
    throw new CliError(
      EXIT.usage,
      `${source}: expected an array of ops or { "ops": [...] }`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  const bad = ops.findIndex(
    (op) =>
      !op ||
      typeof op !== 'object' ||
      Array.isArray(op) ||
      typeof (op as { op?: unknown }).op !== 'string',
  )
  if (bad !== -1) {
    throw new CliError(
      EXIT.usage,
      `${source}: ops[${bad}] must be an object with a string "op"`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  return ops as Op[]
}

/** Resolves `bytes`-style fields that name a local file into the bytes the ops expect. */
export function inlineLocalFiles(ops: Op[], ctx: PathContext): Op[] {
  return ops.map((op) => {
    const fields = Object.hasOwn(BYTES_FIELDS, op.op) ? BYTES_FIELDS[op.op] : undefined
    if (!fields) return op
    const next = structuredClone(op) as Op
    for (const field of fields) {
      const value = getPath(next, field)
      if (typeof value !== 'string' || value.startsWith('data:')) continue
      const path = isAbsolute(value) ? value : resolve(ctx.cwd, value)
      if (!existsSync(path)) continue
      assertAllowed(path, ctx.env, 'read')
      setPath(next, field, new Uint8Array(readFileSync(path)))
      // the op's sibling "ext" is implied by the file name
      const extField = field.replace(/bytes$/, 'ext')
      if (!getPath(next, extField)) setPath(next, extField, extname(path).slice(1).toLowerCase())
    }
    return next
  })
}

function getPath(obj: Record<string, unknown>, dotted: string): unknown {
  return dotted.split('.').reduce<unknown>((cur, key) => {
    return cur && typeof cur === 'object' ? (cur as Record<string, unknown>)[key] : undefined
  }, obj)
}

function setPath(obj: Record<string, unknown>, dotted: string, value: unknown): void {
  const keys = dotted.split('.')
  let cur = obj
  for (const key of keys.slice(0, -1)) cur = cur[key] as Record<string, unknown>
  cur[keys[keys.length - 1]!] = value
}

export interface ApplyOptions {
  isolation?: 'atomic' | 'per_op'
  dryRun?: boolean
}

export function applyOps(opened: OpenedPptx, ops: Op[], opts: ApplyOptions): TxnResult {
  return runTxn(opened, { ops, isolation: opts.isolation, dryRun: opts.dryRun })
}

/** Displayed style of the first run and the inheritance layer each value came from. */
export interface EffectiveStyle {
  fontSizePt?: number
  fontFamily?: string
  color?: string
  bold: boolean
  italic: boolean
  align?: string
  src: Record<string, string>
}

export interface ElementSummary {
  id: string | null
  type: string
  kind?: string
  name?: string
  placeholder?: string
  box: { x: number; y: number; cx: number; cy: number; unit?: ReadUnit }
  text?: string
  /** the text was clipped to the preview length; `--full` or `--max-chars` returns the rest */
  truncated?: true
  effective?: EffectiveStyle
  /** whole-element hyperlink: url, slide jump or named show action (`setLink` writes it) */
  link?: LinkTarget
  rows?: number
  cols?: number
  style?: { id: string; name: string | null; flags: string[] }
  children?: ElementSummary[]
}

export interface LayoutSummary {
  index: number
  name: string
  type: string
  placeholders: Array<{ type: string; idx: string }>
}
export const PREVIEW_CHARS = 300

export interface SlideSummary {
  index: number
  id: string
  elements: ElementSummary[]
  /** the page timeline in play order; absent when the slide has no animations */
  animations?: AnimationEntry[]
  notes?: string
}

export interface DeckSummary {
  slides: number
  size: { cx: number; cy: number; unit?: ReadUnit; inches: { width: number; height: number } }
  emu_per_inch: number
  layouts?: LayoutSummary[]
  pages: SlideSummary[]
}

/**
 * Structure an agent needs to target ops: durable ids, geometry in EMU, text previews.
 * `full` drops the clipping (whole text, every table row) and adds speaker notes, for
 * reading a deck as source material rather than targeting it.
 */
export function describeDeck(
  opened: OpenedPptx,
  only?: number,
  full = false,
  maxChars = full ? Infinity : PREVIEW_CHARS,
  layouts = false,
  unit: ReadUnit = 'emu',
): DeckSummary {
  const { deck } = opened
  const pages = deck.slides
    .map((slide, index) => ({ slide, index }))
    .filter(({ index }) => only === undefined || index === only)
    .map(({ slide, index }) => summarizeSlide(opened, slide, index, full, maxChars))
  if (unit !== 'emu')
    for (const page of pages) for (const el of page.elements) convertBoxes(el, unit)
  return {
    slides: deck.slides.length,
    size: {
      cx: fromEmu(deck.size.cx, unit),
      cy: fromEmu(deck.size.cy, unit),
      ...(unit !== 'emu' ? { unit } : {}),
      inches: {
        width: round(deck.size.cx / EMU_PER_INCH),
        height: round(deck.size.cy / EMU_PER_INCH),
      },
    },
    emu_per_inch: EMU_PER_INCH,
    ...(layouts ? { layouts: describeLayouts(opened) } : {}),
    pages,
  }
}

/** Layouts an `addSlideWithLayout` / `setSlideLayout` op can name, with their content placeholders. */
export function describeLayouts(opened: OpenedPptx): LayoutSummary[] {
  return listSlideLayouts(opened.archive).map((l, index) => ({
    index,
    name: l.name,
    type: l.layoutType,
    placeholders: l.placeholders.map((ph) => ({ type: ph.type || 'body', idx: ph.idx })),
  }))
}

function convertBoxes(el: ElementSummary, unit: ReadUnit): void {
  const { x, y, cx, cy } = el.box
  el.box = {
    x: fromEmu(x, unit),
    y: fromEmu(y, unit),
    cx: fromEmu(cx, unit),
    cy: fromEmu(cy, unit),
    unit,
  }
  for (const child of el.children ?? []) convertBoxes(child, unit)
}

function summarizeSlide(
  opened: OpenedPptx,
  slide: Slide,
  index: number,
  full: boolean,
  maxChars: number,
): SlideSummary {
  const links = new Map(getSlideLinks(opened, index).map((l) => [l.elementId, l.target]))
  const out: SlideSummary = {
    index,
    id: slideDurableId(slide),
    elements: slide.elements.map((el) => summarizeElement(el, full, maxChars, links)),
  }
  const animations = listSlideAnimations(slide)
  if (animations.length) out.animations = animations
  if (full) {
    const notes = getSlideNotes(opened.archive, slide.path)
    if (notes) out.notes = notes
  }
  return out
}

function summarizeElement(
  el: SlideElement,
  full: boolean,
  maxChars: number,
  links: Map<string, LinkTarget>,
): ElementSummary {
  const { offset } = el.transform
  const out: ElementSummary = {
    id: elementDurableId(el),
    type: el.type,
    box: { x: offset.x, y: offset.y, cx: offset.cx, cy: offset.cy },
  }
  if (el.name) out.name = el.name
  if (el.placeholder) out.placeholder = el.placeholder
  const link = links.get(el.id)
  if (link) out.link = link
  switch (el.type) {
    case 'text':
    case 'shape': {
      const { text, truncated } = clipText(paragraphsText(el.text?.paragraphs ?? []), maxChars)
      if (text) out.text = text
      if (truncated) out.truncated = true
      const effective = effectiveStyle(el.text?.paragraphs ?? [])
      if (effective) out.effective = effective
      break
    }
    case 'table': {
      out.rows = el.rows.length
      out.cols = el.colWidths.length
      if (el.styleId) {
        out.style = {
          id: el.styleId,
          name: builtinTableStyleName(el.styleId) ?? null,
          flags: Object.entries(el.styleFlags ?? {})
            .filter(([, on]) => on)
            .map(([k]) => k),
        }
      }
      let clipped = !full && el.rows.length > 3
      out.text = (full ? el.rows : el.rows.slice(0, 3))
        .map((r) =>
          r
            .map((c) => {
              const cell = clipText(paragraphsText(c.text?.paragraphs ?? []), maxChars)
              clipped ||= cell.truncated
              return cell.text
            })
            .join(' | '),
        )
        .join('\n')
      if (clipped) out.truncated = true
      break
    }
    case 'group':
      out.children = el.children.map((child) => summarizeElement(child, full, maxChars, links))
      break
    case 'passthrough':
      out.kind = el.kind
      break
  }
  return out
}

function effectiveStyle(paragraphs: Paragraph[]): EffectiveStyle | undefined {
  const paragraph = paragraphs.find((p) => p.runs.some((r) => r.text))
  const run = paragraph?.runs.find((r) => r.text)
  if (!paragraph || !run) return undefined
  const src: Record<string, string> = { ...run.styleSrc }
  if (paragraph.alignSrc) src.align = paragraph.alignSrc
  return {
    ...(run.fontSize != null ? { fontSizePt: run.fontSize } : {}),
    ...(run.fontFamily ? { fontFamily: run.fontFamily } : {}),
    ...(run.color ? { color: run.color } : {}),
    bold: !!run.bold,
    italic: !!run.italic,
    ...(paragraph.align ? { align: paragraph.align } : {}),
    src,
  }
}

function paragraphsText(paragraphs: { runs: { text: string }[] }[]): string {
  return paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
}

function round(n: number): number {
  return Math.round(n * 100) / 100
}
