import type JSZip from 'jszip'
import type { SourceInfo } from './types'
import { escapeXmlText } from './xml-utils'

/**
 * Bibliography sources live in a customXml part using the Word b:Sources
 * schema, so sources added here appear in Word's own citation manager.
 */

export const SOURCES_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'

export const CUSTOM_XML_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'

/** find the zip path of the existing b:Sources custom xml part, if any */
export async function findSourcesPart(zip: JSZip): Promise<string | null> {
  for (const name of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(name)) continue
    const xml = await zip.file(name)!.async('string')
    if (xml.includes('Sources') && xml.includes(SOURCES_NS)) return name
  }
  return null
}

export function parseSourcesXml(xml: string): SourceInfo[] {
  const out: SourceInfo[] = []
  const re = /<b:Source>([\s\S]*?)<\/b:Source>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const body = m[1]
    const field = (tag: string) => {
      const f = new RegExp(`<b:${tag}>([\\s\\S]*?)</b:${tag}>`).exec(body)?.[1]
      return f ? decodeEntities(f.trim()) : undefined
    }
    const tag = field('Tag')
    if (!tag) continue
    // author: every person of the name list as "Last, First", or the corporate name
    let author = field('Corporate') ?? ''
    if (!author) {
      author = [...body.matchAll(/<b:Person>([\s\S]*?)<\/b:Person>/g)]
        .map((p) => {
          const part = (tag: string) =>
            decodeEntities(new RegExp(`<b:${tag}>([\\s\\S]*?)</b:${tag}>`).exec(p[1])?.[1] ?? '')
          return [part('Last').trim(), part('First').trim()].filter(Boolean).join(', ')
        })
        .filter(Boolean)
        .join('; ')
    }
    out.push({
      tag,
      type: field('SourceType') ?? 'Misc',
      author,
      title: field('Title') ?? '',
      year: field('Year') ?? '',
      publisher: field('Publisher') ?? field('JournalName') ?? field('InternetSiteTitle'),
      url: field('URL'),
      ...(field('City') ? { city: field('City') } : {}),
    })
  }
  return out
}

function decodeEntities(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function sourceEntryXml(source: SourceInfo): string {
  const t = (tag: string, value: string | undefined) =>
    value ? `<b:${tag}>${escapeXmlText(value)}</b:${tag}>` : ''
  const people = sourceAuthors(source)
    .map((p) => `<b:Person>${t('Last', p.last)}${t('First', p.first)}</b:Person>`)
    .join('')
  const authorXml = people
    ? `<b:Author><b:Author><b:NameList>${people}</b:NameList></b:Author></b:Author>`
    : ''
  const publisherTag =
    source.type === 'JournalArticle'
      ? 'JournalName'
      : source.type === 'InternetSite'
        ? 'InternetSiteTitle'
        : 'Publisher'
  return (
    '<b:Source>' +
    t('Tag', source.tag) +
    t('SourceType', source.type) +
    authorXml +
    t('Title', source.title) +
    t('Year', source.year) +
    t(publisherTag, source.publisher) +
    t('URL', source.url) +
    t('City', source.city) +
    '</b:Source>'
  )
}

/**
 * Surgical rebuild of b:Sources: unchanged b:Source entries keep their original bytes
 * (unmodeled fields like Editor/Volume/Pages/DOI/multi-author lists are preserved), the
 * root element attributes (SelectedStyle citation style) come from the original document,
 * and only new/edited entries are rebuilt from the modeled fields.
 */
export function buildSourcesXml(
  sources: SourceInfo[],
  originalXml: string | null = null,
  /** the style picked in this session; without one the part keeps the style it names */
  style?: CitationStyleId,
): string {
  const originals = new Map<string, { xml: string; parsed: SourceInfo }>()
  let rootOpen = `<b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" xmlns:b="${SOURCES_NS}" xmlns="${SOURCES_NS}">`
  if (originalXml) {
    const open = /<b:Sources(?:\s[^>]*)?>/.exec(originalXml)?.[0]
    if (open) rootOpen = open
    for (const m of originalXml.match(/<b:Source>[\s\S]*?<\/b:Source>/g) ?? []) {
      const parsed = parseSourcesXml(m)[0]
      if (parsed) originals.set(parsed.tag, { xml: m, parsed })
    }
  }
  const unchanged = (a: SourceInfo, b: SourceInfo) =>
    a.type === b.type &&
    a.author === b.author &&
    a.title === b.title &&
    a.year === b.year &&
    (a.publisher ?? '') === (b.publisher ?? '') &&
    (a.url ?? '') === (b.url ?? '')
  const body = sources
    .map((s) => {
      const orig = originals.get(s.tag)
      return orig && unchanged(orig.parsed, s) ? orig.xml : sourceEntryXml(s)
    })
    .join('')
  return (style ? withSelectedStyle(rootOpen, style) : rootOpen) + body + '</b:Sources>'
}

export function buildSourcesItemPropsXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<ds:datastoreItem ds:itemID="{4A8D2934-D5B1-4C15-A6F4-7A1B2C3D4E5F}"' +
    ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
    `<ds:schemaRefs><ds:schemaRef ds:uri="${SOURCES_NS}"/></ds:schemaRefs>` +
    '</ds:datastoreItem>'
  )
}

/** in-text citation in the list of sources, e.g. "(Smith, 2024)" */
export function citationText(source: SourceInfo): string {
  return `(${citeOne(STYLES.apa, source, {}, 0) || source.tag})`
}

export interface SourcePerson {
  last: string
  first: string
}

/** the people of a source: "Last, First; Last, First"; a name with no comma is a corporate author */
export function sourceAuthors(source: SourceInfo): SourcePerson[] {
  return (source.author ?? '')
    .split(';')
    .map((name) => name.trim())
    .filter(Boolean)
    .map((name) => {
      const comma = name.indexOf(',')
      return comma === -1
        ? { last: name, first: '' }
        : { last: name.slice(0, comma).trim(), first: name.slice(comma + 1).trim() }
    })
}

export type CitationStyleId = 'apa' | 'mla' | 'chicago' | 'ieee' | 'harvard'

/** a stretch of a bibliography line */
export interface CitationSegment {
  text: string
  italic?: boolean
}

interface CiteSwitches {
  /** `\p`: the pages cited */
  pages?: string
  /** `\n`, `\y`, `\t`: leave the author, the year or the title out */
  noAuthor?: boolean
  noYear?: boolean
  noTitle?: boolean
}

interface StyleDef {
  /** StyleName and SelectedStyle of the sources part, and the Version beside them */
  name: string
  xsl: string
  version: string
  /** citations are the source's number in brackets */
  numeric?: boolean
  /** the surnames of a citation */
  citeNames: (lasts: string[]) => string
  /** between the names and the year of a citation */
  citeYearSep: string | null
  /** before the pages of a citation */
  citePagesSep: string
  /** one author of a bibliography line; `lead` is the first of the line */
  bibName: (person: SourcePerson, lead: boolean) => string
  bibNames: (names: string[]) => string
  /** how the year joins the authors: in their chunk, as the next chunk, or closing the line */
  year: 'paren' | 'comma' | 'next' | 'end'
  /** between the chunks of a line */
  sep: string
  quoteArticle?: boolean
  urlLead: string
}

const initials = (first: string): string =>
  first
    .split(/[\s.]+/)
    .filter(Boolean)
    .map((word) => `${word[0]!.toUpperCase()}.`)
    .join(' ')

const etAl = (lasts: string[], and: string, max: number): string => {
  if (lasts.length <= 1) return lasts[0] ?? ''
  if (lasts.length > max) return `${lasts[0]} et al.`
  if (lasts.length === 2) return `${lasts[0]} ${and} ${lasts[1]}`
  return `${lasts.slice(0, -1).join(', ')}, ${and} ${lasts[lasts.length - 1]}`
}

const listed = (names: string[], and: string, comma: boolean): string => {
  if (names.length <= 1) return names[0] ?? ''
  if (names.length === 2) return `${names[0]}${comma ? ',' : ''} ${and} ${names[1]}`
  return `${names.slice(0, -1).join(', ')}, ${and} ${names[names.length - 1]}`
}

const lastInitials = (p: SourcePerson) => (p.first ? `${p.last}, ${initials(p.first)}` : p.last)
const lastFirst = (p: SourcePerson) => (p.first ? `${p.last}, ${p.first}` : p.last)
const firstLast = (p: SourcePerson) => (p.first ? `${p.first} ${p.last}` : p.last)

const STYLES: Record<CitationStyleId, StyleDef> = {
  apa: {
    name: 'APA',
    xsl: '\\APASixthEditionOfficeOnline.xsl',
    version: '6',
    citeNames: (l) => etAl(l, '&', 2),
    citeYearSep: ', ',
    citePagesSep: ', p. ',
    bibName: lastInitials,
    bibNames: (n) => listed(n, '&', true),
    year: 'paren',
    sep: '. ',
    urlLead: 'Retrieved from ',
  },
  mla: {
    name: 'MLA',
    xsl: '\\MLASeventhEditionOfficeOnline.xsl',
    version: '7',
    citeNames: (l) => etAl(l, 'and', 2),
    citeYearSep: null,
    citePagesSep: ' ',
    bibName: (p, lead) => (lead ? lastFirst(p) : firstLast(p)),
    bibNames: (n) => listed(n, 'and', false),
    year: 'end',
    sep: '. ',
    quoteArticle: true,
    urlLead: '',
  },
  chicago: {
    name: 'Chicago',
    xsl: '\\Chicago.xsl',
    version: '16',
    citeNames: (l) => etAl(l, 'and', 3),
    citeYearSep: ' ',
    citePagesSep: ', ',
    bibName: (p, lead) => (lead ? lastFirst(p) : firstLast(p)),
    bibNames: (n) => listed(n, 'and', true),
    year: 'next',
    sep: '. ',
    quoteArticle: true,
    urlLead: '',
  },
  ieee: {
    name: 'IEEE',
    xsl: '\\IEEE2006OfficeOnline.xsl',
    version: '2006',
    numeric: true,
    citeNames: () => '',
    citeYearSep: null,
    citePagesSep: ', p. ',
    bibName: (p) => (p.first ? `${initials(p.first)} ${p.last}` : p.last),
    bibNames: (n) => listed(n, 'and', false),
    year: 'end',
    sep: ', ',
    quoteArticle: true,
    urlLead: '',
  },
  harvard: {
    name: 'Harvard - Anglia',
    xsl: '\\HarvardAnglia2008OfficeOnline.xsl',
    version: '2008',
    citeNames: (l) => etAl(l, '&', 2),
    citeYearSep: ', ',
    citePagesSep: ', p. ',
    bibName: lastInitials,
    bibNames: (n) => listed(n, '&', false),
    year: 'comma',
    sep: '. ',
    urlLead: 'Available at: ',
  },
}

/** the styles the Style: picker offers, in its order */
export const CITATION_STYLES: ReadonlyArray<{ id: CitationStyleId; name: string }> = (
  ['apa', 'chicago', 'harvard', 'ieee', 'mla'] as const
).map((id) => ({ id, name: STYLES[id].name }))

/** the style a sources part names, when it is one of ours */
export function citationStyleOfXml(xml: string): CitationStyleId | null {
  const open = /<b:Sources(?:\s[^>]*)?>/.exec(xml)?.[0] ?? ''
  const named = `${/StyleName="([^"]*)"/.exec(open)?.[1] ?? ''} ${/SelectedStyle="([^"]*)"/.exec(open)?.[1] ?? ''}`
  const found = (Object.keys(STYLES) as CitationStyleId[]).find((id) =>
    new RegExp(`(?:^|[\\s\\\\/])${id}`, 'i').test(named),
  )
  return found ?? null
}

function withSelectedStyle(rootOpen: string, style: CitationStyleId): string {
  const def = STYLES[style]
  const set = (open: string, attr: string, value: string) =>
    new RegExp(`\\s${attr}="[^"]*"`).test(open)
      ? open.replace(new RegExp(`(\\s${attr}=")[^"]*"`), (_, lead: string) => `${lead}${value}"`)
      : open.replace(/^<b:Sources/, `<b:Sources ${attr}="${value}"`)
  return set(
    set(set(rootOpen, 'Version', def.version), 'StyleName', def.name),
    'SelectedStyle',
    def.xsl,
  )
}

/** the instruction of a citation of that source */
export const citationInstruction = (tag: string): string => `CITATION ${tag} \\l 1033`

/** what a CITATION instruction cites: its sources (`\m` adds one) and its switches */
export function citationFieldOf(instr: string): ({ tags: string[] } & CiteSwitches) | null {
  const tokens = instr.trim().match(/"[^"]*"|\S+/g) ?? []
  if (tokens[0]?.toUpperCase() !== 'CITATION' || !tokens[1]) return null
  const out: { tags: string[] } & CiteSwitches = { tags: [tokens[1].replace(/"/g, '')] }
  for (let i = 2; i < tokens.length; i++) {
    const arg = () => (tokens[++i] ?? '').replace(/"/g, '')
    const sw = tokens[i]!.toLowerCase()
    if (sw === '\\m') out.tags.push(arg())
    else if (sw === '\\p') out.pages = arg()
    else if (sw === '\\n') out.noAuthor = true
    else if (sw === '\\y') out.noYear = true
    else if (sw === '\\t') out.noTitle = true
    else if (/^\\[lsfv]$/.test(sw)) arg()
  }
  return out
}

/** one source inside a citation's brackets */
function citeOne(def: StyleDef, source: SourceInfo, sw: CiteSwitches, number: number): string {
  if (def.numeric) return `${number}${sw.pages ? def.citePagesSep + sw.pages : ''}`
  const names = sw.noAuthor ? '' : def.citeNames(sourceAuthors(source).map((p) => p.last))
  const lead = names || (sw.noTitle ? '' : source.title)
  const year = sw.noYear || def.citeYearSep === null ? '' : source.year || 'n.d.'
  return (
    [lead, year].filter(Boolean).join(def.citeYearSep ?? '') +
    (sw.pages ? def.citePagesSep + sw.pages : '')
  )
}

/**
 * The number each source takes in a numbered style: the order citations first
 * name them in the text, then the sources nothing cites yet.
 */
export function sourceNumbers(sources: SourceInfo[], citedOrder: string[]): Map<string, number> {
  const out = new Map<string, number>()
  for (const tag of [...citedOrder, ...sources.map((s) => s.tag)])
    if (!out.has(tag) && sources.some((s) => s.tag === tag)) out.set(tag, out.size + 1)
  return out
}

/** the text of a citation field in a style; null when it cites no source of the document */
export function formatCitation(
  instr: string,
  sources: SourceInfo[],
  style: CitationStyleId,
  citedOrder: string[] = [],
): string | null {
  const field = citationFieldOf(instr)
  if (!field) return null
  const def = STYLES[style]
  const numbers = sourceNumbers(sources, citedOrder)
  const parts = field.tags
    .map((tag) => sources.find((s) => s.tag === tag))
    .filter((s): s is SourceInfo => s !== undefined)
    .map((s) => citeOne(def, s, field, numbers.get(s.tag) ?? 0))
    .filter(Boolean)
  if (parts.length === 0) return null
  return def.numeric ? parts.map((p) => `[${p}]`).join(', ') : `(${parts.join('; ')})`
}

/** one bibliography line of a source */
function entrySegments(def: StyleDef, source: SourceInfo, number: number): CitationSegment[] {
  const chunks: CitationSegment[] = []
  const year = source.year || (def.year === 'end' ? '' : 'n.d.')
  const names = def.bibNames(sourceAuthors(source).map((p, i) => def.bibName(p, i === 0)))
  const article = source.type === 'JournalArticle'
  const title: CitationSegment = article
    ? { text: def.quoteArticle ? `“${source.title}”` : source.title }
    : { text: source.title, italic: true }
  const yearChunk = def.year === 'paren' ? `(${year})` : year
  if (names) {
    if (def.year === 'paren') chunks.push({ text: `${names} ${yearChunk}` })
    else if (def.year === 'comma') chunks.push({ text: `${names}, ${year}` })
    else chunks.push({ text: names })
    if (def.year === 'next') chunks.push({ text: year })
    if (source.title) chunks.push(title)
  } else {
    if (source.title) chunks.push(title)
    if (def.year !== 'end') chunks.push({ text: yearChunk })
  }
  if (source.publisher) {
    if (article) chunks.push({ text: source.publisher, italic: true })
    else if (source.type === 'InternetSite') chunks.push({ text: source.publisher })
    else chunks.push({ text: [source.city, source.publisher].filter(Boolean).join(': ') })
  }
  if (def.year === 'end' && year) chunks.push({ text: year })
  if (source.url) chunks.push({ text: def.urlLead + source.url })

  const out: CitationSegment[] = []
  const push = (seg: CitationSegment) => {
    const prev = out[out.length - 1]
    if (prev && !!prev.italic === !!seg.italic) prev.text += seg.text
    else out.push({ ...seg })
  }
  if (def.numeric) push({ text: `[${number}] ` })
  chunks.forEach((chunk, i) => {
    push(chunk)
    const last = i === chunks.length - 1
    // a chunk that closes on its own mark ("J.", "n.d.") takes no second one
    const closed = /[.?!]$/.test(chunk.text) && def.sep.startsWith('.')
    const sep = last ? '.' : def.sep
    push({ text: closed ? sep.slice(1) : sep })
  })
  return out.filter((seg) => seg.text !== '')
}

/**
 * The bibliography in a style, one line per source: by number where citations
 * are numbers, by author (then year) everywhere else.
 */
export function bibliographyEntries(
  sources: SourceInfo[],
  style: CitationStyleId,
  citedOrder: string[] = [],
): CitationSegment[][] {
  const def = STYLES[style]
  const numbers = sourceNumbers(sources, citedOrder)
  const key = (s: SourceInfo) => `${s.author || s.title}\u0000${s.year}`.toLowerCase()
  const ordered = [...sources].sort((a, b) =>
    def.numeric
      ? (numbers.get(a.tag) ?? 0) - (numbers.get(b.tag) ?? 0)
      : key(a).localeCompare(key(b)),
  )
  return ordered.map((s) => entrySegments(def, s, numbers.get(s.tag) ?? 0))
}

/** a BIBLIOGRAPHY field as one w:p per source, the field begun in the first and ended in the last */
export function bibliographyFieldXml(entries: CitationSegment[][]): string[] {
  const run = (inner: string, italic = false) =>
    `<w:r><w:rPr>${italic ? '<w:i/><w:iCs/>' : ''}<w:noProof/></w:rPr>${inner}</w:r>`
  const begin =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> BIBLIOGRAPHY </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  return entries.map(
    (segments, i) =>
      '<w:p><w:pPr><w:pStyle w:val="Bibliography"/><w:ind w:left="720" w:hanging="720"/>' +
      '<w:rPr><w:noProof/></w:rPr></w:pPr>' +
      (i === 0 ? begin : '') +
      segments
        .map((seg) => run(`<w:t xml:space="preserve">${escapeXmlText(seg.text)}</w:t>`, seg.italic))
        .join('') +
      (i === entries.length - 1 ? end : '') +
      '</w:p>',
  )
}
