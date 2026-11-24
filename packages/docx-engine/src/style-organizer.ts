import JSZip from 'jszip'
import { loadDocxZip } from './zip-load'
import { STYLE_SET_XML } from './style-sets-data'
import {
  dropStyleRefsXml,
  styleRemovalIds,
  upsertStyleXml,
  withDefaultPropsXml,
  type StyleUpsert,
} from './style-upsert'
import { bindUnboundPrefixes } from './xml-namespaces'

const STYLES_PATH = 'word/styles.xml'
const EMPTY_STYLES =
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>'
const STORY_PART_RE = /^word\/(?:document|header|footer|footnotes|endnotes|comments)\d*\.xml$/

/** one w:style of a styles part, as the Organizer lists it */
export interface OrganizerStyle {
  styleId: string
  name: string
  type: 'paragraph' | 'character' | 'table' | 'numbering'
  /** w:customStyle: only these can be deleted or renamed */
  custom: boolean
  isDefault: boolean
  basedOn?: string
  next?: string
  link?: string
  xml: string
}

const attrOf = (tag: string, name: string): string | undefined =>
  new RegExp(`\\b${name}=(["'])([^"']*)\\1`).exec(tag)?.[2]

const childVal = (xml: string, child: string): string | undefined =>
  new RegExp(`<${child}\\b[^>]*\\bw:val=(["'])([^"']*)\\1`).exec(xml)?.[2]

const unescapeXml = (text: string): string =>
  text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')

export function listStylesXml(xml: string): OrganizerStyle[] {
  const out: OrganizerStyle[] = []
  for (const m of xml.matchAll(/<w:style\b[^>]*(?:\/>|>[\s\S]*?<\/w:style>)/g)) {
    const head = m[0].slice(0, m[0].indexOf('>') + 1)
    const styleId = attrOf(head, 'w:styleId')
    if (styleId === undefined) continue
    const type = attrOf(head, 'w:type')
    const flag = (name: string) => /^(?:1|true|on)$/.test(attrOf(head, name) ?? '')
    const ref = (child: string) => {
      const v = childVal(m[0], child)
      return v === undefined ? undefined : unescapeXml(v)
    }
    out.push({
      styleId: unescapeXml(styleId),
      name: ref('w:name') ?? unescapeXml(styleId),
      type: type === 'character' || type === 'table' || type === 'numbering' ? type : 'paragraph',
      custom: flag('w:customStyle'),
      isDefault: flag('w:default'),
      ...(ref('w:basedOn') !== undefined ? { basedOn: ref('w:basedOn') } : {}),
      ...(ref('w:next') !== undefined ? { next: ref('w:next') } : {}),
      ...(ref('w:link') !== undefined ? { link: ref('w:link') } : {}),
      xml: m[0],
    })
  }
  return out
}

/** styles.xml of a docx / dotx, with pending definitions applied */
export async function readDocxStylesXml(
  bytes: Uint8Array,
  upserts: readonly StyleUpsert[] = [],
): Promise<string> {
  const zip = await loadDocxZip(bytes)
  let xml = (await zip.file(STYLES_PATH)?.async('string')) ?? EMPTY_STYLES
  for (const up of upserts) xml = upsertStyleXml(xml, up)
  return xml
}

/**
 * A style as it can live in another file: numbering links, revision ids and
 * anything outside the main namespace and its text effects (w14) belong to the
 * file it came from.
 */
function portableStyleXml(xml: string): string {
  return xml
    .replace(/<w:numPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:numPr>)/g, '')
    .replace(/<w:rsid\b[^>]*\/>/g, '')
    .replace(/<(?!w:|w14:)([A-Za-z0-9]+:[A-Za-z0-9]+)\b[^>]*(?:\/>|>[\s\S]*?<\/\1>)/g, '')
    .replace(/\s(?!w:|w14:)[A-Za-z0-9]+:[A-Za-z0-9]+=(?:"[^"]*"|'[^']*')/g, '')
    .replace(/\sw:default=(?:"[^"]*"|'[^']*')/, '')
}

const sameStyle = (a: OrganizerStyle, b: OrganizerStyle): boolean =>
  a.type === b.type && a.name.toLowerCase() === b.name.toLowerCase()

/**
 * The definitions that carry styles from one styles part into another. A
 * style the target already has under the same name is overwritten in place;
 * parents and the linked twin the target lacks come along, so the copy looks
 * as it did; a w:next the target cannot resolve is dropped. With `all`, every
 * source style is carried and the target's default paragraph style takes the
 * source's, whatever it is called.
 */
function copyUpserts(sourceXml: string, styleId: string | null, targetXml: string): StyleUpsert[] {
  const all = styleId === null
  const source = listStylesXml(sourceXml)
  const target = listStylesXml(targetXml)
  const byId = new Map(source.map((s) => [s.styleId, s]))
  const taken = new Set(target.map((s) => s.styleId))
  const idMap = new Map<string, string>()
  const order: OrganizerStyle[] = []
  const targetOf = (s: OrganizerStyle): OrganizerStyle | undefined =>
    target.find((d) => sameStyle(d, s)) ??
    (all && s.isDefault ? target.find((d) => d.isDefault && d.type === s.type) : undefined)
  const plan = (s: OrganizerStyle, overwrite: boolean): void => {
    if (idMap.has(s.styleId)) return
    const existing = targetOf(s)?.styleId
    if (existing !== undefined && !overwrite) {
      idMap.set(s.styleId, existing)
      return
    }
    let id = existing ?? s.styleId
    for (let n = 1; existing === undefined && taken.has(id); n++) id = `${s.styleId}${n}`
    taken.add(id)
    idMap.set(s.styleId, id)
    order.push(s)
    for (const dep of [s.basedOn, s.link]) {
      const d = dep === undefined ? undefined : byId.get(dep)
      if (d) plan(d, all)
    }
  }
  const roots = all ? source : [byId.get(styleId)].filter((s) => s !== undefined)
  for (const root of roots) plan(root, true)
  const attr = (v: string) => v.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
  // parents first: a style is written after what it is based on
  return order.reverse().map((s) => {
    const ref = (id: string | undefined): string | null => {
      const src = id === undefined ? undefined : byId.get(id)
      if (!src) return null
      return idMap.get(src.styleId) ?? targetOf(src)?.styleId ?? null
    }
    let xml = portableStyleXml(s.xml)
    const refs = { 'w:basedOn': s.basedOn, 'w:next': s.next, 'w:link': s.link }
    for (const [child, id] of Object.entries(refs)) {
      const to = ref(id)
      xml = xml.replace(new RegExp(`<${child}\\b[^>]*/>`), () =>
        to === null ? '' : `<${child} w:val="${attr(to)}"/>`,
      )
    }
    const over = targetOf(s)
    if (all && over) {
      xml = xml.replace(/<w:name\b[^>]*\/>/, () => `<w:name w:val="${attr(over.name)}"/>`)
      if (over.isDefault) xml = xml.replace(/^<w:style\b/, '<w:style w:default="1"')
    }
    return { styleId: idMap.get(s.styleId)!, sourceXml: xml }
  })
}

export const styleCopyUpserts = (
  sourceXml: string,
  styleId: string,
  targetXml: string,
): StyleUpsert[] => copyUpserts(sourceXml, styleId, targetXml)

export type StyleSetId = keyof typeof STYLE_SET_XML

/** Design > style sets, in Word's order */
export const STYLE_SET_IDS = Object.keys(STYLE_SET_XML) as StyleSetId[]

/** the definitions that give a document the look of one of Word's style sets; styles the set does not have keep theirs */
export const styleSetUpserts = (set: StyleSetId, targetXml: string): StyleUpsert[] =>
  copyUpserts(`<w:styles>${STYLE_SET_XML[set]}</w:styles>`, null, targetXml)

const elementIn = (xml: string, tag: string): string =>
  new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`).exec(xml)?.[1] ?? ''

/**
 * The styles of a document as a style set of its own: its built-in paragraph and
 * character styles, with the document defaults folded into the paragraph styles
 * that have no parent, so the set looks the same on a document with other defaults.
 */
export function styleSetXmlOf(stylesXml: string): string {
  const defaults = elementIn(stylesXml, 'w:docDefaults')
  const rPr = elementIn(elementIn(defaults, 'w:rPrDefault'), 'w:rPr')
  const pPr = elementIn(elementIn(defaults, 'w:pPrDefault'), 'w:pPr')
  const all = listStylesXml(stylesXml)
  const paragraphs = new Set(
    all.filter((s) => s.type === 'paragraph' && !s.custom).map((s) => s.styleId),
  )
  // a linked character twin is written as a custom style even for a built-in heading
  const kept = (s: OrganizerStyle): boolean =>
    s.type === 'paragraph'
      ? paragraphs.has(s.styleId)
      : s.type === 'character' && (!s.custom || (s.link !== undefined && paragraphs.has(s.link)))
  return all
    .filter(kept)
    .map((s) =>
      s.type === 'paragraph' && s.basedOn === undefined
        ? withDefaultPropsXml(portableSetStyle(s.xml), rPr, pPr)
        : portableSetStyle(s.xml),
    )
    .join('')
}

const portableSetStyle = (xml: string): string =>
  xml.replace(/<w:numPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:numPr>)/g, '').replace(/<w:rsid\b[^>]*\/>/g, '')

/** the definitions that give a document the look of a saved style set */
export const customStyleSetUpserts = (setXml: string, targetXml: string): StyleUpsert[] =>
  copyUpserts(`<w:styles>${setXml}</w:styles>`, null, targetXml)

/** every style of `stylesXml` as it stands there, for putting a document's styles back */
export const styleRestoreUpserts = (stylesXml: string): StyleUpsert[] =>
  listStylesXml(stylesXml).map((s) => ({ styleId: s.styleId, sourceXml: s.xml }))

/** a docx / dotx with the definitions applied; text using a removed style falls back to the default */
export async function writeDocxStyles(
  bytes: Uint8Array,
  upserts: readonly StyleUpsert[],
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(STYLES_PATH)
  if (!file) throw new Error('the file has no styles part')
  let xml = await file.async('string')
  const removed: string[] = []
  for (const up of upserts) {
    if (up.remove) removed.push(...styleRemovalIds(xml, up.styleId))
    xml = upsertStyleXml(xml, up)
  }
  zip.file(STYLES_PATH, bindUnboundPrefixes(xml), { date: file.date })
  if (removed.length > 0) {
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir || !STORY_PART_RE.test(name)) continue
      const part = await entry.async('string')
      const next = dropStyleRefsXml(part, removed)
      if (next !== part) zip.file(name, next, { date: entry.date })
    }
  }
  for (const [name, entry] of Object.entries(zip.files)) if (entry.dir) delete zip.files[name]
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}
