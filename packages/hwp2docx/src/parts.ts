/**
 * Adding a part to the main document of a .docx: the part itself, its
 * relationship from word/document.xml and its content-type override.
 */

import type JSZip from 'jszip'

/**
 * Returns the relationship id the part was added under, or null when the
 * document already relates a part of that type (a singleton such as the font
 * table or settings) or has no relationships part. `repeatable` types (headers,
 * footers) always get a relationship of their own.
 */
export async function addDocumentPart(
  zip: JSZip,
  name: string,
  xml: string,
  relType: string,
  contentType: string,
  repeatable = false,
): Promise<string | null> {
  zip.file(`word/${name}`, xml)
  const relsPath = 'word/_rels/document.xml.rels'
  const rels = await zip.file(relsPath)?.async('string')
  let id: string | null = null
  if (rels && (repeatable || !rels.includes(`Type="${relType}"`))) {
    const used = new Set([...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]))
    let n = used.size + 1
    while (used.has(`rId${n}`)) n++
    id = `rId${n}`
    zip.file(
      relsPath,
      rels.replace(
        '</Relationships>',
        `<Relationship Id="${id}" Type="${relType}" Target="${name}"/></Relationships>`,
      ),
    )
  }
  const typesPath = '[Content_Types].xml'
  const types = await zip.file(typesPath)?.async('string')
  if (types && !types.includes(`PartName="/word/${name}"`)) {
    zip.file(
      typesPath,
      types.replace(
        '</Types>',
        `<Override PartName="/word/${name}" ContentType="${contentType}"/></Types>`,
      ),
    )
  }
  return id
}
