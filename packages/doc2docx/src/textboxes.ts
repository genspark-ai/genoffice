/**
 * Text box stories (PlcftxbxTxt, FIB pair 56; PlcfHdrtxbxTxt, pair 58).
 *
 * The text box document follows the endnotes in the text stream; the header
 * text box document follows it. Each PLC entry is a 22-byte FTXBXS whose lid
 * is the OfficeArt shape id the story belongs to. Reusable (deleted) slots
 * and the terminating entry (lid 0) carry no story. Like header stories, a
 * story ends with a guard paragraph mark that mapStory strips.
 */
import type { FibInfo } from './fib'
import type { PieceTable } from './clx'
import type { ImageBlock } from './images'
import { mapStory, type HeaderFooterStory, type StoryFormatter } from './headers'
import { decodeWordDocumentRange } from './text'

export interface TextBoxStoriesSource {
  fib: FibInfo
  wordDocument: Uint8Array
  table: Uint8Array
  pieces: PieceTable
}

export interface TextBoxStory {
  /** index of the FTXBXS entry in its table */
  index: number
  /** OfficeArt shape id (FTXBXS.lid) */
  spid: number
  content: HeaderFooterStory
}

export interface TextBoxStories {
  body: TextBoxStory[]
  header: TextBoxStory[]
  warnings: string[]
}

const FTXBXS_BYTES = 22
const MAX_STORIES = 65536

function parseTable(
  source: TextBoxStoriesSource,
  fc: number,
  lcb: number,
  documentStart: number,
  documentLength: number,
  what: string,
  warnings: string[],
  formatter: StoryFormatter | undefined,
  inlineImages: ImageBlock[],
): TextBoxStory[] {
  if (lcb === 0) {
    if (documentLength > 0)
      warnings.push(
        `the document declares ${what} text box text but has no text box table; those text boxes are omitted`,
      )
    return []
  }
  const { table, wordDocument, pieces } = source
  const totalCharacters = pieces.cps.at(-1) ?? 0
  if (
    fc > table.length ||
    lcb > table.length - fc ||
    lcb < 4 ||
    (lcb - 4) % (4 + FTXBXS_BYTES) !== 0 ||
    (lcb - 4) / (4 + FTXBXS_BYTES) > MAX_STORIES ||
    documentStart + documentLength > totalCharacters
  ) {
    warnings.push(`the ${what} text box table is invalid; those text boxes are omitted`)
    return []
  }
  const count = (lcb - 4) / (4 + FTXBXS_BYTES)
  const dv = new DataView(table.buffer, table.byteOffset, table.byteLength)
  const out: TextBoxStory[] = []
  for (let index = 0; index < count; index++) {
    const base = fc + (count + 1) * 4 + index * FTXBXS_BYTES
    const reusable = dv.getUint16(base + 8, true) !== 0
    const spid = dv.getUint32(base + 14, true)
    if (reusable || spid === 0) continue
    const start = Math.min(dv.getUint32(fc + index * 4, true), documentLength)
    const end = Math.min(dv.getUint32(fc + (index + 1) * 4, true), documentLength)
    if (end <= start) continue
    const raw = decodeWordDocumentRange(
      wordDocument,
      pieces,
      documentStart + start,
      documentStart + end,
    )
    const content = mapStory(
      raw,
      documentStart + start,
      formatter,
      [],
      inlineImages
        .filter((image) => image.cp >= start && image.cp < end)
        .map((image) => ({ ...image, cp: image.cp - start })),
    ) ?? { text: '', paragraphs: [''] }
    out.push({ index, spid, content })
  }
  return out
}

/** CP where the text box document starts (header text boxes follow it) */
export function textBoxDocumentStart(fib: FibInfo): number {
  return fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpAtn + fib.ccpEdn
}

/**
 * Read the text box stories of the body and the header document. Missing or
 * malformed tables yield warnings, never errors. Inline image CPs are relative
 * to the start of their text box document.
 */
export function parseTextBoxStories(
  source: TextBoxStoriesSource,
  formatter?: StoryFormatter,
  inlineImages: { body?: ImageBlock[]; header?: ImageBlock[] } = {},
): TextBoxStories {
  const warnings: string[] = []
  const { fib } = source
  const bodyStart = textBoxDocumentStart(fib)
  const body = parseTable(
    source,
    fib.fcPlcftxbxTxt,
    fib.lcbPlcftxbxTxt,
    bodyStart,
    fib.ccpTxbx,
    'body',
    warnings,
    formatter,
    inlineImages.body ?? [],
  )
  const header = parseTable(
    source,
    fib.fcPlcfHdrtxbxTxt,
    fib.lcbPlcfHdrtxbxTxt,
    bodyStart + fib.ccpTxbx,
    fib.ccpHdrTxbx,
    'header',
    warnings,
    formatter,
    inlineImages.header ?? [],
  )
  return { body, header, warnings }
}
