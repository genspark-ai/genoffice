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

/** real notes (separator entries excluded), in file order */
export function parseNotesXml(xml: string, kind: NoteKind): NoteInfo[] {
  return noteEntriesOf(xml, kind).map(({ id, text, xml: entryXml }) => {
    const richParas = noteRichParas(entryXml)
    const hasFormat = richParas.some((paras) =>
      paras.some(
        (r) => r.bold || r.italic || r.underline || r.strike || r.color || r.sizeHalfPoints,
      ),
    )
    return { id, text, ...(hasFormat ? { richParas } : {}) }
  })
}

/** Display runs per paragraph (bold/italic/underline/strike, color, size); the footnote self-reference mark run is skipped */
function noteRichParas(entryXml: string): NoteRun[][] {
  const out: NoteRun[][] = []
  const pRe = /<w:p[\s>][\s\S]*?<\/w:p>|<w:p\/>/g
  let p: RegExpExecArray | null
  const flag = (rPr: string, tag: string) =>
    new RegExp(`<w:${tag}(?:\\s*/>|\\s(?![^>]*w:val="(?:0|false|none)")[^>]*/>)`).test(rPr)
  while ((p = pRe.exec(entryXml)) !== null) {
    const runs: NoteRun[] = []
    const rRe = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g
    let r: RegExpExecArray | null
    while ((r = rRe.exec(p[0])) !== null) {
      const inner = r[1]
      if (/<w:footnoteRef\/>|<w:endnoteRef\/>/.test(inner)) continue
      const text = notePlainText(inner)
      if (!text) continue
      const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(inner)?.[0] ?? ''
      const run: NoteRun = { text }
      if (flag(rPr, 'b')) run.bold = true
      if (flag(rPr, 'i')) run.italic = true
      // w:u is not a boolean prop: underline needs an explicit non-"none" w:val
      const uVal = /<w:u\s[^>]*w:val="([^"]*)"/.exec(rPr)?.[1]
      if (uVal && uVal !== 'none') run.underline = true
      if (flag(rPr, 'strike')) run.strike = true
      const color = /<w:color [^>]*w:val="([0-9A-Fa-f]{6})"/.exec(rPr)?.[1]
      if (color) run.color = color.toUpperCase()
      const sz = /<w:sz [^>]*w:val="(\d+)"/.exec(rPr)?.[1]
      if (sz) run.sizeHalfPoints = parseInt(sz, 10)
      runs.push(run)
    }
    out.push(runs)
  }
  // Strip the first paragraph's leading space (spacer after the self-reference mark)
  if (out[0]?.[0]) out[0][0].text = out[0][0].text.replace(/^\s+/, '')
  return out
}

/** typeless (real) note entries with their exact XML slice + plain text */
function noteEntriesOf(
  xml: string,
  kind: NoteKind,
): Array<{ id: string; text: string; xml: string }> {
  const out: Array<{ id: string; text: string; xml: string }> = []
  const entry = ENTRY[kind]
  const re = new RegExp(`<${entry}(\\s[^>]*)?>([\\s\\S]*?)</${entry}>`, 'g')
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const attrs = m[1] ?? ''
    if (/w:type="/.test(attrs)) continue // separator / continuation entries
    const id = /w:id="([^"]+)"/.exec(attrs)?.[1]
    if (!id) continue
    const paras: string[] = []
    const pRe = /<w:p[\s>][\s\S]*?<\/w:p>|<w:p\/>/g
    let p: RegExpExecArray | null
    while ((p = pRe.exec(m[2])) !== null) paras.push(notePlainText(p[0]))
    // the first paragraph starts with the self-reference mark + a spacer
    if (paras.length > 0) paras[0] = paras[0].replace(/^\s+/, '')
    out.push({ id, text: paras.join('\n'), xml: m[0] })
  }
  return out
}

function notePlainText(xml: string): string {
  const texts: string[] = []
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) texts.push(m[1])
  return texts
    .join('')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** default separator entries required by Word when the part is created fresh */
function separatorEntries(kind: NoteKind): string {
  const entry = ENTRY[kind]
  return (
    `<${entry} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></${entry}>` +
    `<${entry} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></${entry}>`
  )
}

function noteEntryXml(kind: NoteKind, note: NoteInfo): string {
  const entry = ENTRY[kind]
  const refTag = kind === 'footnote' ? 'w:footnoteRef' : 'w:endnoteRef'
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
    const re = new RegExp(`<${entry}\\s[^>]*w:type="[^"]*"[^>]*>[\\s\\S]*?</${entry}>`, 'g')
    structural = (originalXml.match(re) ?? []).join('')
    for (const e of noteEntriesOf(originalXml, kind)) originals.set(e.id, e)
  }
  if (!structural) structural = separatorEntries(kind)
  const body = notes
    .map((n) => {
      const orig = originals.get(n.id)
      if (!orig) return noteEntryXml(kind, n)
      if (orig.text === n.text) return orig.xml
      const patched = patchParagraphTexts(orig.xml, n.text, {
        stripFirstParaLeadingSpace: true,
      })
      return patched ?? noteEntryXml(kind, n)
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
