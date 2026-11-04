/**
 * Stage 3c-1: header/footer stories from the PlcfHdd table (FIB pair 11).
 *
 * PlcfHdd is a CP-only PLC: n+2 CP entries for n stories. The first six
 * stories are the footnote/endnote separators (ignored here). Every section
 * then owns six stories in the order even header, odd header, even footer,
 * odd footer, first header, first footer. The CPs are offsets from the start
 * of the header document, which begins at CP ccpText + ccpFtn in the text
 * stream. At limit CP - 1 Word stores a paragraph-mark placeholder that is not
 * part of the story: strip exactly one trailing "\r" and treat the remaining
 * trailing paragraph mark as the story's final paragraph break.
 */
import { LegacyDocError, isLegacyDocError } from './errors'
import type { FibInfo } from './fib'
import type { PieceTable } from './clx'
import type { FormattedParagraph, NoteMark } from './formatting'
import { decodeWordDocumentRange, mapWordControlChars } from './text'
import type { SectionInfo } from './sections'
import type { FloatingShape } from './drawings'
import type { ImageBlock } from './images'

export interface HeaderFooterStory {
  /** mapped story text, paragraphs separated by newlines */
  text: string
  /** mapped paragraphs of the story (empty paragraphs included) */
  paragraphs: string[]
  /** the same paragraphs with resolved formatting, when a StoryFormatter was given */
  formatted?: FormattedParagraph[]
  /** floating pictures and text boxes anchored in the story, by paragraph index */
  pictures?: Map<number, FloatingShape[]>
  /** inline pictures of the story, by paragraph index */
  images?: Map<number, StoryInlineImage[]>
}

export interface StoryInlineImage {
  image: ImageBlock
  /** before any text of the paragraph, or after it */
  position: 'start' | 'end'
}

/** Resolves the formatted paragraphs of a CP range outside the body (header, footnote, endnote documents). */
export interface StoryFormatter {
  paragraphs(startCp: number, endCp: number, noteMarks?: NoteMark[]): FormattedParagraph[]
}

export interface SectionHeaderFooter {
  evenHeader?: HeaderFooterStory
  oddHeader?: HeaderFooterStory
  evenFooter?: HeaderFooterStory
  oddFooter?: HeaderFooterStory
  firstHeader?: HeaderFooterStory
  firstFooter?: HeaderFooterStory
}

export interface HeaderFooterResult {
  stories: SectionHeaderFooter[]
  warnings: string[]
}

export interface HeaderStoriesSource {
  fib: FibInfo
  wordDocument: Uint8Array
  table: Uint8Array
  pieces: PieceTable
}

/** story order inside one section (PlcfHdd index base 6 + section*6) */
const STORY_KEYS = [
  'evenHeader',
  'oddHeader',
  'evenFooter',
  'oddFooter',
  'firstHeader',
  'firstFooter',
] as const
const SEPARATOR_STORIES = 6

function badHdd(detail: string): never {
  throw new LegacyDocError('bad_hdd', detail)
}

function u32(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset, true)
}

/** map one raw story range to a story, stripping the guard paragraph mark */
export function mapStory(
  raw: string,
  startCp: number,
  formatter: StoryFormatter | undefined,
  drawings: FloatingShape[],
  inlineImages: ImageBlock[],
  keepBlank = false,
): HeaderFooterStory | undefined {
  let text = raw
  if (text.endsWith('\r')) text = text.slice(0, -1)
  const rawParagraphs = text.split('\r')
  // the final paragraph mark of the story leaves a trailing empty entry
  if (rawParagraphs.length > 1 && rawParagraphs[rawParagraphs.length - 1] === '')
    rawParagraphs.pop()
  const paragraphs = rawParagraphs.map(mapWordControlChars)
  const images = inlineImages.filter((image) => image.cp < text.length)
  if (paragraphs.length === 0) return undefined
  if (
    !keepBlank &&
    drawings.length === 0 &&
    images.length === 0 &&
    paragraphs.every((paragraph) => paragraph === '')
  )
    return undefined
  const story: HeaderFooterStory = { text: paragraphs.join('\n'), paragraphs }
  if (formatter) {
    const formatted = formatter.paragraphs(startCp, startCp + text.length)
    if (formatted.length > 1 && formatted[formatted.length - 1]!.text === '') formatted.pop()
    // the formatted read must split the story exactly like the plain mapping
    if (formatted.length === paragraphs.length) story.formatted = formatted
  }
  if (drawings.length > 0) {
    const pictures = new Map<number, FloatingShape[]>()
    for (const drawing of drawings) {
      let index = 0
      for (let cp = 0; cp < drawing.cp && cp < text.length; cp++) if (text[cp] === '\r') index++
      index = Math.min(index, paragraphs.length - 1)
      pictures.set(index, [...(pictures.get(index) ?? []), drawing])
    }
    story.pictures = pictures
  }
  if (images.length > 0) {
    const byParagraph = new Map<number, StoryInlineImage[]>()
    let paragraphStart = 0
    let index = 0
    for (const image of images) {
      while (
        index < rawParagraphs.length - 1 &&
        image.cp >= paragraphStart + rawParagraphs[index]!.length + 1
      ) {
        paragraphStart += rawParagraphs[index]!.length + 1
        index++
      }
      const before = mapWordControlChars(text.slice(paragraphStart, image.cp))
      byParagraph.set(index, [
        ...(byParagraph.get(index) ?? []),
        { image, position: before === '' ? 'start' : 'end' },
      ])
    }
    story.images = byParagraph
  }
  return story
}

/**
 * Parse the header/footer stories of every section. Broken structural data is
 * a 'bad_hdd' error; a document that declares header text but has no PlcfHdd
 * yields a warning and no stories (the writer then omits headers honestly).
 */
export function parseHeaderStories(
  source: HeaderStoriesSource,
  sections: SectionInfo[],
  formatter?: StoryFormatter,
  drawings: FloatingShape[] = [],
  inlineImages: ImageBlock[] = [],
): HeaderFooterResult {
  const { fib, wordDocument, table, pieces } = source
  const warnings: string[] = []
  const lcb = fib.lcbPlcfhdd
  if (lcb === 0) {
    if (fib.ccpHdd > 0)
      warnings.push(
        'the document declares header/footer text but has no header table; headers and footers are omitted',
      )
    return { stories: [], warnings }
  }
  if (lcb % 4 !== 0 || lcb < 12)
    badHdd(`the header/footer table has an invalid length (lcb ${lcb}, expected a multiple of 4)`)
  if (fib.fcPlcfhdd > table.length || lcb > table.length - fib.fcPlcfhdd)
    badHdd(
      `the header/footer table range (${fib.fcPlcfhdd}+${lcb}) lies outside the ${table.length}-byte table stream`,
    )
  const cpCount = lcb / 4
  const cps: number[] = []
  for (let index = 0; index < cpCount; index++) cps.push(u32(table, fib.fcPlcfhdd + index * 4))
  for (let index = 0; index + 1 < cpCount; index++) {
    if (cps[index + 1]! < cps[index]!)
      badHdd(`the header/footer CPs are not ordered (${cps[index]} > ${cps[index + 1]})`)
  }
  const storyCount = cpCount - 2
  const requiredStories = SEPARATOR_STORIES + sections.length * STORY_KEYS.length
  if (sections.length === 0) {
    warnings.push('the document has no section table; headers and footers are omitted')
    return { stories: [], warnings }
  }
  if (storyCount < requiredStories)
    badHdd(
      `the header/footer table has ${storyCount} stories, but ${sections.length} section(s) need ${requiredStories}`,
    )
  const headerStart = fib.ccpText + fib.ccpFtn
  const headerLength = fib.ccpHdd
  const totalCharacters = pieces.cps.at(-1) ?? 0
  if (headerStart + headerLength > totalCharacters)
    badHdd(
      `the header document (CPs ${headerStart}..${headerStart + headerLength}) lies outside the ${totalCharacters}-character text`,
    )
  const readStory = (storyIndex: number): HeaderFooterStory | undefined => {
    const relativeStart = Math.min(cps[storyIndex]!, headerLength)
    const relativeEnd = Math.min(cps[storyIndex + 1]!, headerLength)
    if (relativeEnd <= relativeStart) return undefined
    try {
      const raw = decodeWordDocumentRange(
        wordDocument,
        pieces,
        headerStart + relativeStart,
        headerStart + relativeEnd,
      )
      return mapStory(
        raw,
        headerStart + relativeStart,
        formatter,
        drawings
          .filter((drawing) => drawing.cp >= relativeStart && drawing.cp < relativeEnd)
          // header pictures never push body text: Word lays them out without wrapping
          .map((drawing) => ({
            ...drawing,
            cp: drawing.cp - relativeStart,
            wrap: 'none' as const,
          })),
        inlineImages
          .filter((image) => image.cp >= relativeStart && image.cp < relativeEnd)
          .map((image) => ({ ...image, cp: image.cp - relativeStart })),
        // a later section's blank story is an unlinked, empty header; a zero-length one inherits
        storyIndex >= SEPARATOR_STORIES + STORY_KEYS.length,
      )
    } catch (error) {
      if (isLegacyDocError(error))
        badHdd(`header/footer story ${storyIndex} cannot be read: ${error.message}`)
      throw error
    }
  }
  const stories: SectionHeaderFooter[] = []
  for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex++) {
    const sectionStories: SectionHeaderFooter = {}
    for (let variant = 0; variant < STORY_KEYS.length; variant++) {
      const story = readStory(SEPARATOR_STORIES + sectionIndex * STORY_KEYS.length + variant)
      if (story) sectionStories[STORY_KEYS[variant]!] = story
    }
    stories.push(sectionStories)
  }
  return { stories, warnings }
}
