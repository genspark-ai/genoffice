import JSZip from 'jszip'
import { resolveTarget } from './opc'
import { XMLParser } from 'fast-xml-parser'

// Text fidelity: no trim (xml:space="preserve" runs carry the spaces between words),
// no numeric coercion of tag values (otherwise <a:t>02139</a:t> becomes a number and loses characters).
// preserveOrder keeps <a:br> and <a:fld> in sequence with the <a:r> runs around them; grouped by
// tag name they lose that position, and a deck's soft breaks and field text land in the wrong place.
const parser = new XMLParser({
  ignoreAttributes: true,
  trimValues: false,
  parseTagValue: false,
  preserveOrder: true,
})

const manifestParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  attributeValueProcessor: (_name, value) => value.trim(),
})

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
}

function slideNumber(path: string): number {
  const m = /slide(\d+)\.xml$/.exec(path)
  return m ? Number(m[1]) : 0
}

async function presentationSlideEntries(zip: JSZip): Promise<(string | null)[] | null> {
  const presXml = await zipText(zip, 'ppt/presentation.xml')
  if (presXml === undefined) return null
  const pres = manifestParser.parse(presXml) as {
    'p:presentation'?: {
      'p:sldIdLst'?: { 'p:sldId'?: Record<string, string> | Record<string, string>[] }
    }
  }
  const slideIds = asArray(pres['p:presentation']?.['p:sldIdLst']?.['p:sldId'])

  const rels = new Map<string, { target: string; type: string; external: boolean }>()
  const relsXml = await zipText(zip, 'ppt/_rels/presentation.xml.rels')
  if (relsXml) {
    const doc = manifestParser.parse(relsXml) as {
      Relationships?: { Relationship?: Record<string, string> | Record<string, string>[] }
    }
    for (const rel of asArray(doc.Relationships?.Relationship)) {
      const id = String(rel['@_Id'] ?? '')
      if (!id) continue
      rels.set(id, {
        target: String(rel['@_Target'] ?? ''),
        type: String(rel['@_Type'] ?? ''),
        external: String(rel['@_TargetMode'] ?? '').toLowerCase() === 'external',
      })
    }
  }

  const entries: (string | null)[] = []
  for (const sldId of slideIds) {
    const rel = rels.get(sldId['@_r:id'] ?? '')
    entries.push(
      rel && !rel.external && rel.target && rel.type.endsWith('/slide')
        ? resolveTarget('ppt/presentation.xml', rel.target)
        : null,
    )
  }
  return entries
}

function legacySlidePaths(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => slideNumber(a) - slideNumber(b))
}

async function zipText(zip: JSZip, path: string): Promise<string | undefined> {
  const file = zip.files[path]
  return file ? file.async('text') : undefined
}

/**
 * One paragraph's text in document order. Only #text directly under a:t counts: untrimmed, the
 * whitespace laying out any other element is a value too. <a:br> is a soft line break, <a:tab>
 * is a tab stop between runs, and <a:fld> (slide number, date) contributes its own a:t where it sits.
 */
function collectText(nodes: readonly unknown[], out: string[], isText = false): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (key === '#text') {
        if (isText) out.push(String(value))
      } else if (key === 'a:br') {
        out.push('\n')
      } else if (key === 'a:tab') {
        out.push('\t')
      } else if (Array.isArray(value)) {
        collectText(value, out, key === 'a:t')
      }
    }
  }
}

/** walk the slide tree; each a:p paragraph becomes one output entry (a:br splits it further) */
function collectParagraphs(nodes: readonly unknown[], out: string[]): void {
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === 'a:p') {
        const texts: string[] = []
        collectText(value, texts)
        const line = texts.join('')
        if (line.trim()) out.push(line)
      } else {
        collectParagraphs(value, out)
      }
    }
  }
}

function countPictures(nodes: readonly unknown[]): number {
  let count = 0
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === 'p:pic') count += 1
      else count += countPictures(value)
    }
  }
  return count
}

/**
 * A slide whose only content is pictures yields no a:t text. Without a marker the model
 * (and the user reading the attachment chip) takes the bare "## Slide N" heading for a slide
 * that was read, when its figures never reached anyone.
 */
interface SlideSection {
  section: string
  hasText: boolean
  pictures: number
}

function slideSection(heading: string, xml: string): SlideSection {
  const tree = parser.parse(xml)
  const paras: string[] = []
  collectParagraphs(tree, paras)
  if (paras.length > 0)
    return { section: [heading, ...paras].join('\n'), hasText: true, pictures: 0 }
  const pictures = countPictures(tree)
  const note = `[picture-only slide: ${pictures} image${pictures === 1 ? '' : 's'}, no extractable text]`
  return { section: pictures > 0 ? `${heading}\n${note}` : heading, hasText: false, pictures }
}

function joinSections(sections: SlideSection[]): string {
  const body = sections.map((s) => s.section).join('\n\n')
  if (sections.some((s) => s.hasText) || !sections.some((s) => s.pictures > 0)) return body
  const n = sections.length
  return `[No extractable text: none of the ${n} slide${n === 1 ? '' : 's'} carries text; the content is in embedded images, which this extraction does not read.]\n\n${body}`
}

/** extract slide text from a pptx: one "## Slide N" section per slide, a line per paragraph */
export async function pptxToText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const slideEntries = await presentationSlideEntries(zip)
  const sections: SlideSection[] = []
  if (slideEntries) {
    for (const [index, path] of slideEntries.entries()) {
      if (path === null) continue
      const xml = await zipText(zip, path)
      if (!xml) continue
      sections.push(slideSection(`## Slide ${index + 1}`, xml))
    }
    return joinSections(sections)
  }
  for (const path of legacySlidePaths(zip)) {
    const xml = await zipText(zip, path)
    if (!xml) continue
    sections.push(slideSection(`## Slide ${slideNumber(path)}`, xml))
  }
  return joinSections(sections)
}
