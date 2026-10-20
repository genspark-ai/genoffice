import { loadDocxZip } from './zip-load'

/**
 * Accept-all for tracked changes in the parts the editor never opens:
 * headers, footers, footnotes and endnotes. The body's revisions live in the
 * ProseMirror model, but these parts ride through the package as raw XML, so
 * `docs apply --ops '[{"op":"accept_changes","all":true}]'` rewrites them
 * here — otherwise a document whose header carries an insertion would come
 * back reporting that no tracked changes remain while one plainly does.
 *
 * Accepting means Word's semantics per element: inserted text stays as plain
 * runs, deleted text goes, moved text stays at its destination and
 * disappears from its source, and formatting-change records are dropped so
 * the current formatting stands. The tracked paragraph mark (`w:rPr/w:del`
 * inside `w:pPr`) is the one shape this pass does not accept, because it
 * merges two paragraphs: those are counted and reported instead of being
 * silently left behind.
 */

export interface PartRevisionAccept {
  xml: string
  /** run-level insertions, deletions, moves, and formatting records accepted */
  accepted: number
  /** tracked paragraph marks (self-closing w:del inside w:pPr/w:rPr) left in place */
  unhandledParaMarks: number
}

/** One accepted shape: paired deleted content, paired inserted content kept,
 *  or the attribute-only spelling of an inserted paragraph mark. */
const PAIRED = (tag: string) => new RegExp(`<${tag}(?:\\s[^>]*)?>[\\s\\S]*?</${tag}>`, 'g')
const SELF_CLOSING = (tag: string) => new RegExp(`<${tag}(?:\\s[^>]*)?/>`, 'g')

export function acceptRevisionsInPartXml(xml: string): PartRevisionAccept {
  let accepted = 0
  let out = xml

  // tracked paragraph-mark deletions cannot be accepted here (they merge
  // paragraphs); count them before anything touches the del elements
  const unhandledParaMarks = (xml.match(SELF_CLOSING('w:del')) ?? []).length

  // deleted content goes; self-closing w:del (the paragraph mark) was counted above
  out = out.replace(PAIRED('w:del'), () => {
    accepted++
    return ''
  })
  // inserted text stays: unwrap the element, keeping the runs
  out = out.replace(PAIRED('w:ins'), (match) => {
    accepted++
    return match.replace(/^<w:ins(?:\s[^>]*)?>/, '').replace(/<\/w:ins>$/, '')
  })
  // an inserted paragraph mark with no runs to hold: the mark simply stands
  out = out.replace(SELF_CLOSING('w:ins'), () => {
    accepted++
    return ''
  })
  // moves follow their ins/del twins: the source goes, the destination stays
  out = out.replace(PAIRED('w:moveFrom'), () => {
    accepted++
    return ''
  })
  out = out.replace(PAIRED('w:moveTo'), (match) => {
    accepted++
    return match.replace(/^<w:moveTo(?:\s[^>]*)?>/, '').replace(/<\/w:moveTo>$/, '')
  })
  // range markers name the move; with the moves accepted they mean nothing
  for (const marker of [
    'w:moveFromRangeStart',
    'w:moveFromRangeEnd',
    'w:moveToRangeStart',
    'w:moveToRangeEnd',
  ]) {
    out = out.replace(new RegExp(`<${marker}(?:\\s[^>]*)?/>`, 'g'), '')
  }
  // formatting changes: accepting keeps the current format, so the record goes
  for (const tag of ['w:rPrChange', 'w:pPrChange']) {
    out = out.replace(PAIRED(tag), () => {
      accepted++
      return ''
    })
  }

  return { xml: out, accepted, unhandledParaMarks }
}

const SIDE_PART_RE = /^word\/(header\d*|footer\d*)\.xml$|^word\/(footnotes|endnotes)\.xml$/

/**
 * Accepts every tracked change in the header, footer, footnote and endnote
 * parts of a package, as `accept_changes {all: true}` should. Returns the
 * rewritten parts by path (empty when nothing carried a revision), so the
 * caller can hand them to `saveDocx` as `partXml` overrides.
 */
export async function acceptRevisionsInSideParts(
  docxBytes: Uint8Array,
): Promise<{ parts: Record<string, string>; accepted: number; unhandledParaMarks: number }> {
  const zip = await loadDocxZip(docxBytes)
  const parts: Record<string, string> = {}
  let accepted = 0
  let unhandledParaMarks = 0
  for (const name of Object.keys(zip.files)) {
    if (!SIDE_PART_RE.test(name)) continue
    const file = zip.file(name)
    const xml = file ? await file.async('string') : null
    if (!xml || !/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange)[\s/>]/.test(xml)) continue
    const result = acceptRevisionsInPartXml(xml)
    if (result.accepted > 0) parts[name] = result.xml
    accepted += result.accepted
    unhandledParaMarks += result.unhandledParaMarks
  }
  return { parts, accepted, unhandledParaMarks }
}
