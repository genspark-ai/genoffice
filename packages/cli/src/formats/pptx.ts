import { existsSync, readFileSync } from 'node:fs'
import { extname } from 'node:path'
import { assertAllowed, type PathContext } from '../fs'
import { isAbsolute, resolve } from 'node:path'
import {
  createBlankPptx,
  getSlideNotes,
  openPptx,
  savePptx,
  type OpenedPptx,
  type Slide,
  type SlideElement,
} from '@genoffice/pptx-engine'
import {
  elementDurableId,
  runTxn,
  slideDurableId,
  type Op,
  type TxnResult,
} from '@genoffice/pptx-ops'
import { CliError, EXIT } from '../result'

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
    throw new CliError(EXIT.usage, `${source}: not valid JSON (${(err as Error).message})`)
  }
  const ops = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!ops) {
    throw new CliError(EXIT.usage, `${source}: expected an array of ops or { "ops": [...] }`)
  }
  const bad = ops.findIndex(
    (op) =>
      !op ||
      typeof op !== 'object' ||
      Array.isArray(op) ||
      typeof (op as { op?: unknown }).op !== 'string',
  )
  if (bad !== -1) {
    throw new CliError(EXIT.usage, `${source}: ops[${bad}] must be an object with a string "op"`)
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

export interface ElementSummary {
  id: string | null
  type: string
  kind?: string
  name?: string
  placeholder?: string
  box: { x: number; y: number; cx: number; cy: number }
  text?: string
  rows?: number
  cols?: number
  children?: ElementSummary[]
}

export interface SlideSummary {
  index: number
  id: string
  elements: ElementSummary[]
  notes?: string
}

export interface DeckSummary {
  slides: number
  size: { cx: number; cy: number; inches: { width: number; height: number } }
  emu_per_inch: number
  pages: SlideSummary[]
}

/**
 * Structure an agent needs to target ops: durable ids, geometry in EMU, text previews.
 * `full` drops the clipping (whole text, every table row) and adds speaker notes, for
 * reading a deck as source material rather than targeting it.
 */
export function describeDeck(opened: OpenedPptx, only?: number, full = false): DeckSummary {
  const { deck } = opened
  const pages = deck.slides
    .map((slide, index) => ({ slide, index }))
    .filter(({ index }) => only === undefined || index === only)
    .map(({ slide, index }) => summarizeSlide(opened, slide, index, full))
  return {
    slides: deck.slides.length,
    size: {
      cx: deck.size.cx,
      cy: deck.size.cy,
      inches: {
        width: round(deck.size.cx / EMU_PER_INCH),
        height: round(deck.size.cy / EMU_PER_INCH),
      },
    },
    emu_per_inch: EMU_PER_INCH,
    pages,
  }
}

function summarizeSlide(
  opened: OpenedPptx,
  slide: Slide,
  index: number,
  full: boolean,
): SlideSummary {
  const out: SlideSummary = {
    index,
    id: slideDurableId(slide),
    elements: slide.elements.map((el) => summarizeElement(el, full)),
  }
  if (full) {
    const notes = getSlideNotes(opened.archive, slide.path)
    if (notes) out.notes = notes
  }
  return out
}

function summarizeElement(el: SlideElement, full: boolean): ElementSummary {
  const { offset } = el.transform
  const out: ElementSummary = {
    id: elementDurableId(el),
    type: el.type,
    box: { x: offset.x, y: offset.y, cx: offset.cx, cy: offset.cy },
  }
  if (el.name) out.name = el.name
  if (el.placeholder) out.placeholder = el.placeholder
  switch (el.type) {
    case 'text':
    case 'shape': {
      const text = paragraphsText(el.text?.paragraphs ?? [], full)
      if (text) out.text = text
      break
    }
    case 'table':
      out.rows = el.rows.length
      out.cols = el.colWidths.length
      out.text = (full ? el.rows : el.rows.slice(0, 3))
        .map((r) => r.map((c) => paragraphsText(c.text?.paragraphs ?? [], full)).join(' | '))
        .join('\n')
      break
    case 'group':
      out.children = el.children.map((child) => summarizeElement(child, full))
      break
    case 'passthrough':
      out.kind = el.kind
      break
  }
  return out
}

function paragraphsText(paragraphs: { runs: { text: string }[] }[], full: boolean): string {
  const text = paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
  return !full && text.length > 300 ? text.slice(0, 297) + '...' : text
}

function round(n: number): number {
  return Math.round(n * 100) / 100
}
