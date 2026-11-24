/**
 * Word's Unlink Fields on content kept as XML: the fields go, what they show
 * stays as ordinary runs.
 */
import { scanBody } from './scan'
import { loadDocxZip } from './zip-load'

const MAIN_PART = 'word/document.xml'
const EMPTY_RUN = /<w:r(?:\s[^>]*)?>\s*(?:<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>)?\s*<\/w:r>/g

/** block XML without its field characters, instructions and hyperlink wrappers: the results as plain runs */
export function unlinkFieldsXml(xml: string): string {
  return xml
    .replace(/<w:fldChar\b[^>]*?(?:\/>|>[\s\S]*?<\/w:fldChar>)/g, '')
    .replace(/<w:instrText\b[^>]*>[\s\S]*?<\/w:instrText>/g, '')
    .replace(/<w:fldSimple\b[^>]*\/>/g, '')
    .replace(/<\/?w:(?:fldSimple|hyperlink)\b[^>]*>/g, '')
    .replace(EMPTY_RUN, '')
}

/** the package with `bodyXml` as the whole body, in front of the document's last section properties */
export async function docxWithBody(bytes: Uint8Array, bodyXml: string): Promise<Uint8Array | null> {
  const zip = await loadDocxZip(bytes)
  const xml = await zip.file(MAIN_PART)?.async('string')
  if (!xml) return null
  const scan = scanBody(xml)
  if (scan.bodyContentStart < 0) return null
  const sectPr = scan.elements.filter((el) => el.name === 'w:sectPr').pop()
  zip.file(
    MAIN_PART,
    xml.slice(0, scan.bodyContentStart) +
      bodyXml +
      (sectPr ? xml.slice(sectPr.start, sectPr.end) : '') +
      xml.slice(scan.bodyContentEnd),
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
