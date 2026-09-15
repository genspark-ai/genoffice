import type { TranslationUnit } from './types'

/**
 * Split a plain-text document body into translatable units.
 *
 * The rule is intentionally simple: blank-line-separated paragraphs become one
 * unit each, with a stable `unitId` derived from the scene + order. Long
 * paragraphs are split at sentence boundaries when they exceed `maxChars` so
 * one unit never blows past the model's input window. The renderer is free to
 * pre-chunk into headings/list items/table cells before calling this — the
 * helper exists for the "translate the whole document" path.
 *
 *   chunkDocument('Hello world.\n\nSecond paragraph here.', 'document')
 *     -> [{ unitId: 'document-0', kind: 'paragraph', order: 0, sourceText: 'Hello world.' }, ...]
 */
export function chunkDocument(
  text: string,
  scene: string = 'document',
  maxChars: number = 1200,
): TranslationUnit[] {
  if (!text || !text.trim()) return []
  const paragraphs = text
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter(Boolean)
  const out: TranslationUnit[] = []
  let order = 0
  for (const para of paragraphs) {
    if (para.length <= maxChars) {
      out.push({
        unitId: `${scene}-${order}`,
        kind: 'paragraph',
        order,
        sourceText: para,
      })
      order++
      continue
    }
    // long paragraph → split on sentence boundary, keep the delimiter attached
    const sentences = para
      .split(/(?<=[.!?。！？])\s+/)
      .map((s) => s.trim())
      .filter(Boolean)
    let buf: string[] = []
    let bufLen = 0
    const flush = () => {
      if (buf.length === 0) return
      out.push({
        unitId: `${scene}-${order}`,
        kind: 'paragraph',
        order,
        sourceText: buf.join(' '),
      })
      order++
      buf = []
      bufLen = 0
    }
    for (const sentence of sentences) {
      if (sentence.length >= maxChars) {
        // pathological single sentence longer than the cap — hard split
        flush()
        for (let i = 0; i < sentence.length; i += maxChars) {
          out.push({
            unitId: `${scene}-${order}`,
            kind: 'paragraph',
            order,
            sourceText: sentence.slice(i, i + maxChars),
          })
          order++
        }
        continue
      }
      if (bufLen + sentence.length + 1 > maxChars) flush()
      buf.push(sentence)
      bufLen += sentence.length + 1
    }
    flush()
  }
  return out
}

/** Stable, sortable id derived from a parent id and the unit's own offset. */
export function makeUnitId(parent: string, offset: number | string): string {
  return `${parent}-${offset}`
}
