/**
 * AutoText a document carries: the building blocks of its glossary part
 * (word/glossary/document.xml) that sit in the AutoText gallery. A template
 * keeps its entries there, and so does a document made from one.
 */
import { decodeEntities } from './parse-xml-text'
import { scanBody } from './scan'
import { loadDocxZip } from './zip-load'

const MAIN_PART = 'word/document.xml'
const MAIN_RELS = 'word/_rels/document.xml.rels'
const GLOSSARY_REL = /<Relationship\b[^>]*\bType="[^"]*\/glossaryDocument"[^>]*>/
const DOC_PART = /<w:docPart>([\s\S]*?)<\/w:docPart>/g

export interface GlossaryAutoText {
  name: string
  /** the entry's text, paragraphs apart by a line break */
  text: string
}

const attr = (tag: string, name: string): string | null =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1] ?? null

async function glossaryPart(
  zip: Awaited<ReturnType<typeof loadDocxZip>>,
): Promise<{ path: string; xml: string } | null> {
  const rels = await zip.file(MAIN_RELS)?.async('string')
  const target = attr(GLOSSARY_REL.exec(rels ?? '')?.[0] ?? '', 'Target')
  if (!target) return null
  const path = target.startsWith('/') ? target.slice(1) : `word/${target}`
  const xml = await zip.file(path)?.async('string')
  return xml ? { path, xml } : null
}

function autoTextParts(xml: string): Array<{ name: string; body: string }> {
  const out: Array<{ name: string; body: string }> = []
  for (const m of xml.matchAll(DOC_PART)) {
    const part = m[1]!
    const pr = /<w:docPartPr>([\s\S]*?)<\/w:docPartPr>/.exec(part)?.[1] ?? ''
    if (attr(/<w:gallery\b[^>]*>/.exec(pr)?.[0] ?? '', 'w:val') !== 'autoTxt') continue
    const name = decodeEntities(attr(/<w:name\b[^>]*>/.exec(pr)?.[0] ?? '', 'w:val') ?? '')
    const body = /<w:docPartBody>([\s\S]*)<\/w:docPartBody>/.exec(part)?.[1]
    if (name && body !== undefined && !out.some((e) => e.name === name)) out.push({ name, body })
  }
  return out
}

const textOf = (body: string): string =>
  body
    .split(/<\/w:p>/)
    .map((p) =>
      [...p.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:(tab|br)\b[^>]*\/>/g)]
        .map((m) => (m[2] === 'tab' ? '\t' : m[2] ? '\n' : decodeEntities(m[1] ?? '')))
        .join(''),
    )
    .join('\n')
    .trimEnd()

/** the AutoText entries of the document's glossary part, in its order */
export async function docxAutoTextEntries(bytes: Uint8Array): Promise<GlossaryAutoText[]> {
  const part = await glossaryPart(await loadDocxZip(bytes))
  if (!part) return []
  return autoTextParts(part.xml).map((e) => ({ name: e.name, text: textOf(e.body) }))
}

/**
 * One entry as a package of its own that the ordinary parse reads: the
 * glossary document with the entry as its body, in the place of the main
 * document so its styles, numbering and pictures are the glossary's own.
 * `inline`: the entry is one paragraph, to go into a line of text.
 */
export async function docxAutoTextBody(
  bytes: Uint8Array,
  name: string,
): Promise<{ bytes: Uint8Array; inline: boolean } | null> {
  const zip = await loadDocxZip(bytes)
  const part = await glossaryPart(zip)
  const entry = part && autoTextParts(part.xml).find((e) => e.name === name)
  if (!part || !entry) return null
  const root = /<w:glossaryDocument\b([^>]*)>/.exec(part.xml)?.[1] ?? ''
  const main = await zip.file(MAIN_PART)?.async('string')
  const last = main
    ? scanBody(main)
        .elements.filter((el) => el.name === 'w:sectPr')
        .pop()
    : null
  zip.file(
    MAIN_PART,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document${root}><w:body>` +
      entry.body +
      (last ? main!.slice(last.start, last.end) : '') +
      '</w:body></w:document>',
  )
  // the glossary's own relationships, read from the main document's folder
  const dir = part.path.slice(0, part.path.lastIndexOf('/') + 1)
  const relsPath = `${dir}_rels/${part.path.slice(dir.length)}.rels`
  const rels = await zip.file(relsPath)?.async('string')
  const fromWord = dir.startsWith('word/') ? dir.slice('word/'.length) : `../${dir}`
  zip.file(
    MAIN_RELS,
    (
      rels ??
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    ).replace(/<Relationship\b[^>]*>/g, (tag) => {
      const target = attr(tag, 'Target')
      if (!target || /\bTargetMode="External"/.test(tag) || target.startsWith('/')) return tag
      return tag.replace(`Target="${target}"`, () => `Target="${fromWord}${target}"`)
    }),
  )
  const paragraphs = entry.body.match(/<w:p[ >]/g)?.length ?? 0
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    inline: paragraphs === 1 && !/<w:tbl[ >]/.test(entry.body),
  }
}
