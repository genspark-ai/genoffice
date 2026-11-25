import { escapeXmlAttr } from './xml-utils'
import { loadDocxZip } from './zip-load'

const STORY_PART = /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i

const REVISION_TAG =
  /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tblPrExChange|trPrChange|tcPrChange|cellIns|cellDel|cellMerge|numberingChange)\b[^>]*>/g

/** the story's tracked changes that name no author, named `author` */
export function labelUnattributedRevisionsXml(
  xml: string,
  author: string,
): { xml: string; count: number } {
  let count = 0
  const named = ` w:author="${escapeXmlAttr(author)}"`
  const out = xml.replace(REVISION_TAG, (tag) => {
    const own = /\sw:author="([^"]*)"/.exec(tag)
    if (own?.[1]) return tag
    count++
    if (own) return tag.replace(own[0], named)
    const close = tag.endsWith('/>') ? tag.length - 2 : tag.length - 1
    return tag.slice(0, close) + named + tag.slice(close)
  })
  return { xml: out, count }
}

/**
 * Combine Documents: a side's tracked changes that name no author take the
 * name given for that side. Returns the bytes as they came when there is none.
 */
export async function labelUnattributedRevisions(
  bytes: Uint8Array,
  author: string,
): Promise<{ bytes: Uint8Array; count: number }> {
  if (!author.trim()) return { bytes, count: 0 }
  const zip = await loadDocxZip(bytes)
  let total = 0
  for (const name of Object.keys(zip.files)) {
    if (!STORY_PART.test(name)) continue
    const xml = await zip.file(name)!.async('string')
    const { xml: out, count } = labelUnattributedRevisionsXml(xml, author.trim())
    if (count === 0) continue
    zip.file(name, out)
    total += count
  }
  if (total === 0) return { bytes, count: 0 }
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    count: total,
  }
}
