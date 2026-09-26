import JSZip from 'jszip'
import { assertZipWithinLimits } from '@genoffice/docx-engine'
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
  removeNSPrefix: true,
})

function stripNamespacePrefix(name: string): string {
  const separator = name.indexOf(':')
  return separator < 0 ? name : name.slice(separator + 1)
}

const manifestParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  removeNSPrefix: false,
  transformTagName: stripNamespacePrefix,
  attributeValueProcessor: (_name, value) => value.trim(),
})

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
}

function relationshipId(node: Record<string, unknown>): string {
  const qualified = Object.entries(node).find(([key]) => /^@_[^:]+:id$/.test(key))
  return String(qualified?.[1] ?? node['@_id'] ?? '')
}

function slideNumber(path: string): number {
  const m = /slide(\d+)\.xml$/.exec(path)
  return m ? Number(m[1]) : 0
}

async function presentationSlideEntries(zip: JSZip): Promise<(string | null)[] | null> {
  const presXml = await zipText(zip, 'ppt/presentation.xml')
  if (presXml === undefined) return null
  const pres = manifestParser.parse(presXml) as {
    presentation?: {
      sldIdLst?: { sldId?: Record<string, string> | Record<string, string>[] }
    }
  }
  const slideIds = asArray(pres.presentation?.sldIdLst?.sldId)

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
    const rel = rels.get(relationshipId(sldId))
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

const NOTES_SLIDE_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'

async function notesXmlForSlide(zip: JSZip, slidePath: string): Promise<string | undefined> {
  const slash = slidePath.lastIndexOf('/')
  const relsPath = `${slidePath.slice(0, slash)}/_rels/${slidePath.slice(slash + 1)}.rels`
  const relsXml = await zipText(zip, relsPath)
  if (!relsXml) return undefined
  const doc = manifestParser.parse(relsXml) as {
    Relationships?: { Relationship?: Record<string, string> | Record<string, string>[] }
  }
  for (const rel of asArray(doc.Relationships?.Relationship)) {
    if (String(rel['@_Type'] ?? '') !== NOTES_SLIDE_REL) continue
    if (String(rel['@_TargetMode'] ?? '').toLowerCase() === 'external') continue
    const target = String(rel['@_Target'] ?? '')
    return target ? zipText(zip, resolveTarget(slidePath, target)) : undefined
  }
  return undefined
}

/** body placeholder only: the notes page's slide-number field would read as a stray digit */
function notesParagraphs(notesXml: string): string[] {
  const out: string[] = []
  for (const m of notesXml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
    if (!/<p:ph\b[^>]*type="body"/.test(m[0])) continue
    collectParagraphs(parser.parse(m[0]), out)
  }
  return out
}

/**
 * One compatibility branch of an mc:AlternateContent element: the Fallback when the
 * producer wrote one, else the first Choice. Reading both branches would duplicate
 * every run and every picture the element carries. The parser strips namespace
 * prefixes, so the branches arrive as AlternateContent / Choice / Fallback.
 */
function mcBranch(value: readonly unknown[]): readonly unknown[] {
  let choice: readonly unknown[] | undefined
  for (const entry of value) {
    if (entry == null || typeof entry !== 'object') continue
    for (const [key, branch] of Object.entries(entry)) {
      if (key === 'Fallback') return Array.isArray(branch) ? branch : []
      if (key === 'Choice' && !choice && Array.isArray(branch)) choice = branch
    }
  }
  return choice ?? []
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
      } else if (key === 'br') {
        out.push('\n')
      } else if (key === 'tab') {
        out.push('\t')
      } else if (Array.isArray(value)) {
        if (key === 'AlternateContent') collectText(mcBranch(value), out)
        else collectText(value, out, key === 't')
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
      if (key === 'p') {
        const texts: string[] = []
        collectText(value, texts)
        const line = texts.join('')
        if (line.trim()) out.push(line)
      } else {
        collectParagraphs(key === 'AlternateContent' ? mcBranch(value) : value, out)
      }
    }
  }
}

function countPictures(nodes: readonly unknown[]): number {
  let count = 0
  for (const node of nodes) {
    if (node == null || typeof node !== 'object') continue
    for (const [key, value] of Object.entries(node)) {
      if (key === 'pic') count += 1
      else if (Array.isArray(value))
        count += countPictures(key === 'AlternateContent' ? mcBranch(value) : value)
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

function slideSection(heading: string, xml: string, notesXml?: string): SlideSection {
  const tree = parser.parse(xml)
  const paras: string[] = []
  collectParagraphs(tree, paras)
  const notes = notesXml ? notesParagraphs(notesXml) : []
  const notesBlock = notes.length > 0 ? ['### Notes', ...notes] : []
  if (paras.length > 0)
    return {
      section: [heading, ...paras, ...notesBlock].join('\n'),
      hasText: true,
      pictures: 0,
    }
  const pictures = countPictures(tree)
  const note = `[picture-only slide: ${pictures} image${pictures === 1 ? '' : 's'}, no extractable text]`
  const lines = [heading, ...(pictures > 0 ? [note] : []), ...notesBlock]
  return { section: lines.join('\n'), hasText: notes.length > 0, pictures }
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
  assertZipWithinLimits(zip)
  const slideEntries = await presentationSlideEntries(zip)
  const sections: SlideSection[] = []
  if (slideEntries) {
    for (const [index, path] of slideEntries.entries()) {
      if (path === null) continue
      const xml = await zipText(zip, path)
      if (!xml) continue
      sections.push(slideSection(`## Slide ${index + 1}`, xml, await notesXmlForSlide(zip, path)))
    }
    return joinSections(sections)
  }
  for (const path of legacySlidePaths(zip)) {
    const xml = await zipText(zip, path)
    if (!xml) continue
    sections.push(
      slideSection(`## Slide ${slideNumber(path)}`, xml, await notesXmlForSlide(zip, path)),
    )
  }
  return joinSections(sections)
}
