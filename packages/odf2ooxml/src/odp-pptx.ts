/**
 * .odp → .pptx import filter (Tier 2). Each `draw:page` becomes one slide at
 * the deck size the ODF master page declares; every `draw:frame` with a text
 * box or an image lands at its declared coordinates as a plain textbox /
 * picture. ODF's layout/placeholder model (master inheritance, presentation
 * classes) is not the pptx one, so the rebuild is deliberately flat: real
 * boxes at real positions, no placeholder linkage — an import filter that
 * keeps what it can prove rather than a half-mapped layout system.
 */

import {
  addElement,
  addPicture,
  createBlankPptx,
  insertBlankSlide,
  openPptx,
  savePptx,
  setSlideSize,
  type OpenedPptx,
  type Paragraph,
  type Slide,
  type TextRun,
} from '../../pptx-engine/src/index'
import type { OdfPackage } from './odf'
import { childElements, loadOdfPackage, odfBody } from './odf'
import { OdfConvertError } from './odt-docx'
import { OdfStyles } from './styles'
import {
  Drops,
  firstDescendant,
  imageRefOfFrame,
  runsFromParagraph,
  type InlineImageRef,
  type NeutralRun,
  type WalkContext,
} from './text'
import { odfLengthToEmu, sniffImageMime } from './units'
import { XNode, attr } from './xml'

export { OdfConvertError }

export interface OdpToPptxResult {
  pptx: Uint8Array
  slides: number
  /** constructs the filter could not represent, with how often each occurred */
  dropped: Array<{ what: string; count: number }>
}

/** 16:9 — the fallback when the ODF file declares no page size */
const DEFAULT_DECK_EMU = { widthEmu: 12_192_000, heightEmu: 6_858_000 }

/** frames without coordinates get these boxes, as a fraction of the slide */
const DEFAULT_FRAME_BOX: Record<string, { x: number; y: number; w: number; h: number }> = {
  title: { x: 0.08, y: 0.04, w: 0.84, h: 0.14 },
  subtitle: { x: 0.08, y: 0.22, w: 0.84, h: 0.1 },
  outline: { x: 0.08, y: 0.22, w: 0.84, h: 0.7 },
  text: { x: 0.08, y: 0.22, w: 0.84, h: 0.7 },
}

export async function convertOdpToPptx(odp: Uint8Array): Promise<OdpToPptxResult> {
  const pack = await loadOdfPackage(odp)
  if (!pack) throw new OdfConvertError('unreadable')
  const presentation = odfBody(pack, 'office:presentation')
  if (!presentation) throw new OdfConvertError('not-a-presentation')

  const styles = new OdfStyles(pack.contentRoots, pack.stylesRoots)
  const drops = new Drops()
  const ctx: WalkContext = {
    styles,
    families: ['presentation', 'graphic', 'paragraph', 'text'],
    drops,
    images: [],
  }

  const deckSize = styles.pageSizeEmu(pack.stylesRoots) ?? DEFAULT_DECK_EMU
  const opened = await openPptx(await createBlankPptx())
  setSlideSize(opened, deckSize.widthEmu, deckSize.heightEmu)

  const pages = childElements(presentation).filter(({ name }) => name === 'draw:page')
  const imageTasks: Array<Promise<void>> = []
  let slideCount = 0
  for (const page of pages) {
    // page 1 lands in the blank template's own slide; the rest append after it
    const slide =
      slideCount === 0
        ? opened.deck.slides[0]
        : insertBlankSlide(opened, opened.deck.slides.length - 1)
    if (!slide) break
    slideCount++
    for (const { name, node } of childElements(page.node)) {
      if (name === 'draw:frame') addFrame(opened, slide, node, ctx, deckSize, pack, imageTasks)
      else if (name === 'presentation:notes') drops.add('notes')
      else if (SHAPES.has(name)) drops.add('shapes')
      // anything else on the page (thumbnails, settings) is page furniture
    }
  }

  // picture bytes resolve asynchronously out of the ODF zip — the save must
  // not race them
  await Promise.all(imageTasks)
  const pptx = await savePptx(opened)
  return { pptx, slides: slideCount, dropped: drops.list() }
}

const SHAPES = new Set([
  'draw:rect',
  'draw:line',
  'draw:circle',
  'draw:ellipse',
  'draw:custom-shape',
  'draw:connector',
  'draw:control',
  'draw:object',
  'draw:page-thumbnail',
  'math:math',
])

function addFrame(
  opened: OpenedPptx,
  slide: Slide,
  frame: XNode,
  ctx: WalkContext,
  deck: { widthEmu: number; heightEmu: number },
  pack: OdfPackage,
  imageTasks: Array<Promise<void>>,
): void {
  if (attr(frame, 'presentation:class') === 'notes') {
    ctx.drops.add('notes')
    return
  }
  const box = frameBox(frame, deck)
  if (!box) {
    ctx.drops.add('frames-without-geometry')
    return
  }
  const textBox = firstDescendant(frame, 'draw:text-box')
  if (textBox) {
    const paragraphs: Paragraph[] = []
    collectParagraphs(textBox, ctx, paragraphs, 0)
    addElement(slide, {
      kind: 'textbox',
      offset: box,
      paragraphs: paragraphs.length > 0 ? paragraphs : [{ runs: [{ text: '' }] }],
    })
    return
  }
  const ref = imageRefOfFrame(frame)
  if (ref) {
    imageTasks.push(addImage(opened, slide, ref, box, ctx, pack))
    return
  }
  if (firstDescendant(frame, 'table:table')) {
    ctx.drops.add('tables')
    return
  }
  ctx.drops.add('shapes')
}

function collectParagraphs(
  container: XNode,
  ctx: WalkContext,
  out: Paragraph[],
  listDepth: number,
): void {
  for (const { name, node } of childElements(container)) {
    if (name === 'text:p') {
      out.push(toParagraph(node, ctx, listDepth))
    } else if (name === 'text:h') {
      out.push(toParagraph(node, ctx, listDepth))
    } else if (name === 'text:list') {
      collectParagraphs(node, ctx, out, listDepth + 1)
    } else if (name === 'text:list-item') {
      for (const inner of childElements(node)) {
        if (inner.name === 'text:p' || inner.name === 'text:h') {
          out.push(toParagraph(inner.node, ctx, listDepth))
        } else if (inner.name === 'text:list') {
          collectParagraphs(inner.node, ctx, out, listDepth + 1)
        }
      }
    }
  }
}

function toParagraph(node: XNode, ctx: WalkContext, listDepth: number): Paragraph {
  const imagesBefore = ctx.images.length
  const runs = runsFromParagraph(node, ctx)
  // a picture anchored inside a run flow has no textbox representation
  if (ctx.images.length > imagesBefore) {
    ctx.drops.add('images')
    ctx.images.length = imagesBefore
  }
  const paragraph: Paragraph = { runs: runs.map((run) => toPptxRun(run)) }
  const props = ctx.styles.paraProps(attr(node, 'text:style-name'), ctx.families)
  if (props?.align) paragraph.align = props.align
  if (props?.spaceBeforeTw !== undefined) paragraph.spaceBefore = props.spaceBeforeTw / 20
  if (props?.spaceAfterTw !== undefined) paragraph.spaceAfter = props.spaceAfterTw / 20
  if (listDepth > 0) {
    // a flat bullet per level: ODF's per-level markers live in list styles
    // whose full OOXML bullet model is out of the first cut's scope
    paragraph.bullet = { type: 'char', char: '•' }
    paragraph.marL = 457_200 * listDepth
    paragraph.indent = -457_200
  }
  return paragraph
}

function toPptxRun(run: NeutralRun): TextRun {
  const out: TextRun = { text: run.text }
  if (run.bold) out.bold = true
  if (run.italic) out.italic = true
  if (run.underline) out.underline = true
  if (run.strike) out.strike = true
  if (run.sizePt !== undefined) out.fontSize = run.sizePt
  if (run.font) out.fontFamily = run.font
  if (run.color) out.color = `#${run.color}`
  // hyperlinks and vertical alignment are text-box rich-text territory the
  // flat textbox keeps out of the first cut
  return out
}

function frameBox(
  frame: XNode,
  deck: { widthEmu: number; heightEmu: number },
): { x: number; y: number; cx: number; cy: number } | null {
  const x = numberAttr(frame, 'svg:x')
  const y = numberAttr(frame, 'svg:y')
  const width = numberAttr(frame, 'svg:width')
  const height = numberAttr(frame, 'svg:height')
  if (x === null || y === null || width === null || height === null) {
    const fallback = DEFAULT_FRAME_BOX[attr(frame, 'presentation:class') ?? '']
    if (!fallback) return null
    return {
      x: Math.round(fallback.x * deck.widthEmu),
      y: Math.round(fallback.y * deck.heightEmu),
      cx: Math.round(fallback.w * deck.widthEmu),
      cy: Math.round(fallback.h * deck.heightEmu),
    }
  }
  return { x, y, cx: Math.max(1, width), cy: Math.max(1, height) }
}

/** svg:x/y/width/height are ODF lengths ("21.5cm") — EMU or null */
function numberAttr(node: XNode, name: string): number | null {
  return odfLengthToEmu(attr(node, name))
}

async function addImage(
  opened: OpenedPptx,
  slide: Slide,
  ref: InlineImageRef,
  box: { x: number; y: number; cx: number; cy: number },
  ctx: WalkContext,
  pack: OdfPackage,
): Promise<void> {
  let bytes: Uint8Array | null = null
  let path = ''
  if (ref.base64) {
    bytes = Uint8Array.from(atob(ref.base64), (character) => character.charCodeAt(0))
  } else if (ref.href) {
    path = ref.href
    bytes = await pack.read(ref.href)
  }
  if (!bytes || bytes.length === 0) {
    ctx.drops.add('images')
    return
  }
  const mime = sniffImageMime(bytes, path)
  if (!mime) {
    ctx.drops.add('images')
    return
  }
  addPicture(opened, slide, {
    bytes,
    ext: mime === 'image/jpeg' ? 'jpg' : mime === 'image/gif' ? 'gif' : 'png',
    offset: box,
  })
}
