/** The deck-wide text defaults that stage 3 inherits when a run does not
 *  override them: the FontCollection's typeface names and the master's
 *  TextMasterStyleAtom levels. Both live in the PowerPoint Document stream
 *  outside the slide containers, so they are scanned independently of the
 *  change directory. */

import { readTextMasterStyleAtom, type MasterStyleLevel } from './text-style'

export interface PptMasterStyles {
  /** FontCollection order; TextCFException font refs index into this */
  typefaces: string[]
  levels: MasterStyleLevel[]
}

const FONT_COLLECTION = 0x07d5
const FONT_ENTITY_ATOM = 0x0fb7
const TEXT_MASTER_STYLE_ATOM = 0x0fa3
const MAX_DEPTH = 24

interface RecordAt {
  version: number
  instance: number
  type: number
  start: number
  end: number
}

function* records(bytes: Uint8Array, from: number, to: number): Generator<RecordAt> {
  let at = from
  while (at + 8 <= to) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
    const head = dv.getUint16(at, true)
    const type = dv.getUint16(at + 2, true)
    const length = dv.getUint32(at + 4, true)
    const start = at + 8
    if (start + length > bytes.length) return
    yield {
      version: head & 0xf,
      instance: head >> 4,
      type,
      start,
      end: Math.min(start + length, to),
    }
    at = start + length
  }
}

/** lfFaceName is 64 bytes of UTF-16LE, NUL-terminated and capped at 32 chars. */
function faceNameOf(payload: Uint8Array): string {
  const chars: string[] = []
  for (let at = 0; at + 2 <= Math.min(payload.length, 64); at += 2) {
    const code = payload[at]! | (payload[at + 1]! << 8)
    if (code === 0) break
    chars.push(String.fromCharCode(code))
  }
  return chars.join('').trim()
}

export function masterStylesOf(document: Uint8Array): PptMasterStyles {
  const typefaces: string[] = []
  const levels: MasterStyleLevel[] = []
  const seenFonts = new Set<number>()
  const seenStyles = new Set<number>()
  const walk = (from: number, to: number, depth: number): void => {
    if (depth > MAX_DEPTH) return
    for (const record of records(document, from, to)) {
      if (record.type === FONT_COLLECTION) {
        for (const entry of records(document, record.start, record.end)) {
          if (entry.type === FONT_ENTITY_ATOM && !seenFonts.has(entry.start)) {
            seenFonts.add(entry.start)
            typefaces.push(faceNameOf(document.subarray(entry.start, entry.end)))
          }
        }
        continue
      }
      if (record.type === TEXT_MASTER_STYLE_ATOM && !seenStyles.has(record.start)) {
        seenStyles.add(record.start)
        levels.push(
          ...readTextMasterStyleAtom(document.subarray(record.start, record.end), record.instance),
        )
        continue
      }
      if (record.version === 0xf) walk(record.start, record.end, depth + 1)
    }
  }
  walk(0, document.length, 0)
  return { typefaces, levels }
}
