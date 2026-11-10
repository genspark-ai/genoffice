import { decodeEntities } from './parse-xml-text'
import { applyRevisionResolutions } from './revision-xml'
import { patchParagraphTexts } from './text-patch'
import type { NoteInfo, NoteRun } from './types'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

/**
 * Footnotes / endnotes part handling (word/footnotes.xml, word/endnotes.xml).
 *
 * Both parts share one schema: a list of w:footnote / w:endnote elements.
 * Entries carrying a w:type attribute (separator, continuationSeparator, ...)
 * are structural and must be preserved; only typeless entries are real notes.
 */

export type NoteKind = 'footnote' | 'endnote'

const ROOT: Record<NoteKind, string> = { footnote: 'w:footnotes', endnote: 'w:endnotes' }
const ENTRY: Record<NoteKind, string> = { footnote: 'w:footnote', endnote: 'w:endnote' }

export const NOTE_PART_PATH: Record<NoteKind, string> = {
  footnote: 'word/footnotes.xml',
  endnote: 'word/endnotes.xml',
}

export const NOTE_REL_TYPE: Record<NoteKind, string> = {
  footnote: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnote: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
}

export const NOTE_CONTENT_TYPE: Record<NoteKind, string> = {
  footnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnote: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
}

const NOTE_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/**
 * Root attributes for a regenerated part. Entries are spliced back in as original bytes, so
 * the root has to keep declaring whatever prefixes those bytes use: Word puts w14:paraId on
 * every paragraph it writes, and a literal namespace list leaves that prefix unbound.
 *
 * `required` lists namespaces the rebuild itself emits (e.g. w14 on rebuilt comments): they
 * are appended when the reused original root does not declare them, since an original from
 * a non-Word producer may bind fewer prefixes than our generated markup uses.
 */
export function rootAttributes(
  originalXml: string | null,
  rootTag: string,
  fallback: string,
  required: Record<string, string> = {},
): string {
  const attrs = originalXml
    ? new RegExp(`<${rootTag}\\b([^>]*?)/?>`).exec(originalXml)?.[1]?.trim()
    : undefined
  const base = attrs && attrs.includes('xmlns:') ? attrs : fallback
  const missing = Object.entries(required)
    .filter(([prefix]) => !base.includes(`xmlns:${prefix}=`))
    .map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
    .join('')
  return base + missing
}

/**
 * The note's self-reference mark run (`<w:footnoteRef/>` / `<w:endnoteRef/>`).
 * Tolerates the spaced self-closing form (`<w:footnoteRef />`, .NET XmlWriter /
 * Open XML SDK output) and the open/close pair; the exact `<w:footnoteRef/>`
 * match hid every note number in such documents.
 */
const NOTE_REF_MARK_RE = /<w:(?:footnote|endnote)Ref\b\s*\/?>/

/** one note paragraph read through the document's run extraction and property resolver */
export interface NoteParagraph {
  /** display runs, the self-reference mark run excluded */
  runs: NoteRun[]
  /** w:pStyle of the paragraph */
  styleId?: string
  /** before / after / line resolved through the style chain and docDefaults */
  spacing: NonNullable<NoteInfo['spacing']>
}

/** supplied by the document parser; without it notes carry plain text only */
export type NoteParagraphReader = (pXml: string) => NoteParagraph

/** real notes (separator entries excluded), in file order */
export function parseNotesXml(xml: string, kind: NoteKind, read?: NoteParagraphReader): NoteInfo[] {
  return noteEntriesOf(xml, kind).map(({ id, text, xml: entryXml }) => {
    const paras = read ? noteParagraphsOf(entryXml).map(read) : []
    const first = paras[0]
    const richParas = paras.map((p) => p.runs)
    // the first paragraph opens with the self-reference mark and a spacer; a
    // spacer that was its own run empties out and is dropped entirely
    if (richParas[0]?.[0]) {
      richParas[0][0] = { ...richParas[0][0], text: richParas[0][0].text.replace(/^\s+/, '') }
      if (richParas[0][0].text === '') richParas[0].shift()
    }
    const styleId = first ? first.styleId : /<w:pStyle w:val="([^"]+)"/.exec(entryXml)?.[1]
    // Word draws the entry number only where a self-reference mark run exists
    // (empty-body notes keep their line but show no numeral — Word probe 2026-09-01)
    const noRefMark = !NOTE_REF_MARK_RE.test(entryXml)
    return {
      id,
      text,
      ...(first ? { richParas } : {}),
      ...(styleId ? { styleId } : {}),
      ...(/ADDIN\s+(?:ZOTERO_|CSL_)/.test(entryXml) ? { zoteroField: true as const } : {}),
      ...(first ? { spacing: first.spacing } : {}),
      ...(noRefMark ? { noRefMark: true as const } : {}),
    }
  })
}

/** the entry's paragraph slices in order (self-closing form first: the
 *  open-to-close alternative would swallow it) */
function noteParagraphsOf(entryXml: string): string[] {
  return entryXml.match(/<w:p(?:\s[^>]*)?\/>|<w:p[\s>][\s\S]*?<\/w:p>/g) ?? []
}

/** typeless (real) note entries with their exact XML slice + plain text */
function noteEntriesOf(
  xml: string,
  kind: NoteKind,
): Array<{ id: string; text: string; xml: string }> {
  const out: Array<{ id: string; text: string; xml: string }> = []
  const entry = ENTRY[kind]
  // the attribute run must not end in "/" so a self-closing empty entry
  // ("<w:footnote w:id="1"/>") is not read as an open tag over the next entry
  const re = new RegExp(`<${entry}(\\s[^>]*[^/>])?>([\\s\\S]*?)</${entry}>`, 'g')
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const attrs = m[1] ?? ''
    if (/w:type=(?:"[^"]*"|'[^']*')/.test(attrs)) continue // separator / continuation entries
    const id = /w:id=(?:"([^"]+)"|'([^']+)')/.exec(attrs)?.slice(1, 3).find(Boolean)
    if (!id) continue
    out.push({ id, text: noteEntryPlainText(m[2]), xml: m[0] })
  }
  return out
}

/** paragraph texts joined with newlines; the first paragraph starts with the self-reference mark + a spacer */
function noteEntryPlainText(innerXml: string): string {
  const paras = noteParagraphsOf(innerXml).map(notePlainText)
  if (paras.length > 0) paras[0] = paras[0].replace(/^\s+/, '')
  return paras.join('\n')
}

function notePlainText(xml: string): string {
  const texts: string[] = []
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) texts.push(m[1])
  return decodeEntities(texts.join(''))
}

/** default separator entries required by Word when the part is created fresh */
function separatorEntries(kind: NoteKind): string {
  const entry = ENTRY[kind]
  return (
    `<${entry} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></${entry}>` +
    `<${entry} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></${entry}>`
  )
}

/** one rich display run → run XML (size/font/bold…; save-side of NoteRun) */
function noteRunXml(run: NoteRun): string {
  const props: string[] = []
  const fonts: string[] = []
  if (run.fontAscii) {
    fonts.push(
      `w:ascii="${escapeXmlAttr(run.fontAscii)}" w:hAnsi="${escapeXmlAttr(run.fontAscii)}"`,
    )
  }
  if (run.font) fonts.push(`w:eastAsia="${escapeXmlAttr(run.font)}"`)
  if (fonts.length > 0) props.push(`<w:rFonts ${fonts.join(' ')}/>`)
  if (run.bold) props.push('<w:b/>')
  if (run.italic) props.push('<w:i/>')
  if (run.underline) props.push('<w:u w:val="single"/>')
  if (run.strike) props.push('<w:strike/>')
  if (run.color) props.push(`<w:color w:val="${escapeXmlAttr(run.color)}"/>`)
  if (run.sizeHalfPoints) {
    props.push(`<w:sz w:val="${run.sizeHalfPoints}"/><w:szCs w:val="${run.sizeHalfPoints}"/>`)
  }
  const rPr = props.length > 0 ? `<w:rPr>${props.join('')}</w:rPr>` : ''
  const textTag = run.del ? 'w:delText' : 'w:t'
  return `<w:r>${rPr}<${textTag} xml:space="preserve">${escapeXmlText(run.text)}</${textTag}></w:r>`
}

type NoteRev = NonNullable<NoteRun['ins']>
const noteRevKey = (r: NoteRun) =>
  r.ins || r.del ? JSON.stringify([r.ins ?? null, r.del ?? null]) : ''

/** w:ins / w:del attributes with ids unique across one notes part (kept
 *  original entries included, so a rebuilt entry never reuses their ids) */
type NoteRevAttrs = (info: NoteRev) => string
function noteRevAllocator(originalXml: string | null): NoteRevAttrs {
  let seq = 7001
  const emitted = new Set<string>()
  for (const m of (originalXml ?? '').matchAll(/<w:(?:ins|del)\s[^>]*?\bw:id="([^"]*)"/g)) {
    emitted.add(m[1])
  }
  return (info) => {
    let id = info.id && !emitted.has(info.id) ? info.id : String(seq++)
    while (emitted.has(id)) id = String(seq++)
    emitted.add(id)
    return (
      ` w:id="${escapeXmlAttr(id)}" w:author="${escapeXmlAttr(info.author)}"` +
      (info.date ? ` w:date="${escapeXmlAttr(info.date)}"` : '')
    )
  }
}

/** rich runs → run XML; consecutive runs of one tracked change share a w:ins / w:del wrapper */
function noteRunsXml(runs: NoteRun[], attrs: NoteRevAttrs): string {
  let out = ''
  for (let g = 0; g < runs.length;) {
    const key = noteRevKey(runs[g])
    let end = g
    while (end < runs.length && noteRevKey(runs[end]) === key) end++
    const { ins, del } = runs[g]
    const inner = runs.slice(g, end).map(noteRunXml).join('')
    const wrapped = del ? `<w:del${attrs(del)}>${inner}</w:del>` : inner
    out += ins ? `<w:ins${attrs(ins)}>${wrapped}</w:ins>` : wrapped
    g = end
  }
  return out
}

function noteEntryXml(kind: NoteKind, note: NoteInfo, revAttrs: NoteRevAttrs): string {
  const entry = ENTRY[kind]
  const refTag = kind === 'footnote' ? 'w:footnoteRef' : 'w:endnoteRef'
  if (note.richParas?.length) {
    // rich rebuild (P17): runs keep their measured size/font, the reference
    // mark + spacer shrink to the first run's size, and the paragraph pins
    // single spacing — a note rendered from measured source content must not
    // inflate past its source area via template docDefaults (after=120,
    // line=276) or a body-sized marker run
    const paras = note.richParas.map((runs, i) => {
      const sz = runs[0]?.sizeHalfPoints
      const szXml = sz ? `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` : ''
      const refRun =
        i === 0
          ? `<w:r><w:rPr><w:vertAlign w:val="superscript"/>${szXml}</w:rPr><${refTag}/></w:r>` +
            `<w:r>${sz ? `<w:rPr>${szXml}</w:rPr>` : ''}<w:t xml:space="preserve"> </w:t></w:r>`
          : ''
      const pPr = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
      return `<w:p>${pPr}${refRun}${noteRunsXml(runs, revAttrs)}</w:p>`
    })
    return `<${entry} w:id="${escapeXmlAttr(note.id)}">${paras.join('')}</${entry}>`
  }
  const paras = note.text.split('\n').map((line, i) => {
    // OOXML convention: the note body starts with the self-reference mark
    const refRun =
      i === 0
        ? `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><${refTag}/></w:r>` +
          '<w:r><w:t xml:space="preserve"> </w:t></w:r>'
        : ''
    const textRun =
      line === '' ? '' : `<w:r><w:t xml:space="preserve">${escapeXmlText(line)}</w:t></w:r>`
    return `<w:p>${refRun}${textRun}</w:p>`
  })
  return `<${entry} w:id="${escapeXmlAttr(note.id)}">${paras.join('')}</${entry}>`
}

/**
 * Regenerate the notes part from the full desired list. Structural entries
 * (separator/continuation) from the original part are kept byte-identical;
 * when there is no original part, standard separators are created.
 * Surgical: existing entries whose text is unchanged keep their original bytes (inline
 * formatting, images, and hyperlinks are preserved); entries with changed text first try
 * an in-paragraph w:t-level patch (formatting still preserved), and only fall back to a
 * plain-text rebuild when patching fails (paragraph count changed, etc.).
 */
export function buildNotesXml(
  kind: NoteKind,
  notes: NoteInfo[],
  originalXml: string | null,
): string {
  const entry = ENTRY[kind]
  let structural = ''
  const originals = new Map<string, { text: string; xml: string }>()
  if (originalXml) {
    const re = new RegExp(
      `<${entry}\\s[^>]*w:type=(?:"[^"]*"|'[^']*')[^>]*>[\\s\\S]*?</${entry}>`,
      'g',
    )
    structural = (originalXml.match(re) ?? []).join('')
    for (const e of noteEntriesOf(originalXml, kind)) originals.set(e.id, e)
  }
  if (!structural) structural = separatorEntries(kind)
  const revAttrs = noteRevAllocator(originalXml)
  const body = notes
    .map((n) => {
      let orig = originals.get(n.id)
      if (!orig) return noteEntryXml(kind, n, revAttrs)
      if (n.revisionsResolved?.length) {
        const xml = applyRevisionResolutions(orig.xml, n.revisionsResolved)
        const inner = new RegExp(`^<${entry}[^>]*>([\\s\\S]*)</${entry}>$`).exec(xml)?.[1] ?? ''
        orig = { text: noteEntryPlainText(inner), xml }
      }
      if (orig.text === n.text) return orig.xml
      const patched = patchParagraphTexts(orig.xml, n.text, {
        stripFirstParaLeadingSpace: true,
      })
      return patched ?? noteEntryXml(kind, n, revAttrs)
    })
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<${ROOT[kind]} ${rootAttributes(originalXml, ROOT[kind], NOTE_NS)}>${structural}${body}</${ROOT[kind]}>`
  )
}

/** next free numeric note id (separator ids -1/0 and existing notes considered) */
export function nextNoteId(notes: NoteInfo[]): string {
  let max = 0
  for (const note of notes) {
    const n = parseInt(note.id, 10)
    if (Number.isFinite(n)) max = Math.max(max, n)
  }
  return String(max + 1)
}
