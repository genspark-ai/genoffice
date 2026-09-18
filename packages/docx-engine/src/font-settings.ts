import { loadDocxZip } from './zip-load'
import { parseStyles } from './parse-styles'
import {
  mergeDefaultFontsXml,
  mergeStyleXml,
  type DefaultFonts,
  type StyleUpsert,
} from './style-upsert'
import type { ParsedDoc } from './types'

/** Resolve live style inheritance without reparsing the document body or its media. */
export async function previewFontSettings(
  parsed: ParsedDoc,
  upserts: StyleUpsert[],
  defaults?: DefaultFonts,
) {
  const zip = await loadDocxZip(parsed.internal.originalBytes)
  let xml =
    (await zip.file('word/styles.xml')?.async('string')) ??
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>'
  for (const up of upserts) {
    let found = false
    xml = xml.replace(/<w:style\b[^>]*(?:\/>|>[\s\S]*?<\/w:style>)/g, (style) => {
      if (!style.includes(`w:styleId="${up.styleId}"`)) return style
      found = true
      return mergeStyleXml(style, up)
    })
    if (!found) xml = xml.replace('</w:styles>', `${mergeStyleXml(null, up)}</w:styles>`)
  }
  if (defaults) xml = mergeDefaultFontsXml(xml, defaults)
  zip.file('word/styles.xml', xml)
  return parseStyles(zip, parsed.themeColors, parsed.themeFonts)
}
