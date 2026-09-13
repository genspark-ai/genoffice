/**
 * Complex fields are begin/separate/end run triples, and Word treats an
 * unbalanced sequence as a damaged file (repair prompt, or the rest of the
 * document swallowed as one field result). Editing a field whose result spans
 * paragraphs (Zotero bibliographies) can drop the paragraph carrying the begin
 * or the end run, so the assembled body is balanced once more before it is
 * written: stray separate/end runs go, and a begin left open is closed at the
 * end of its own paragraph so the field keeps its code and first result line.
 */
const TOKEN_RE = /<w:fldChar\b[^>]*\bw:fldCharType="(begin|separate|end)"|<w:p\b[^>]*>|<\/w:p>/g

interface Edit {
  start: number
  end: number
  text: string
}

function runBounds(xml: string, at: number): [number, number] | null {
  let open = at
  do {
    open = xml.lastIndexOf('<w:r', open - 1)
    if (open < 0) return null
  } while (!/^<w:r[\s>]/.test(xml.slice(open, open + 5)))
  const close = xml.indexOf('</w:r>', at)
  if (close < 0) return null
  return [open, close + '</w:r>'.length]
}

export function balanceFieldChars(bodyXml: string): string {
  if (!bodyXml.includes('w:fldCharType')) return bodyXml
  const edits: Edit[] = []
  const open: Array<{ depth: number; paraEnd: number | null }> = []
  let depth = 0
  TOKEN_RE.lastIndex = 0
  for (let m = TOKEN_RE.exec(bodyXml); m; m = TOKEN_RE.exec(bodyXml)) {
    const token = m[0]
    if (token === '</w:p>') {
      for (const field of open) {
        if (field.depth === depth && field.paraEnd === null) field.paraEnd = m.index
      }
      depth = Math.max(0, depth - 1)
      continue
    }
    if (token.startsWith('<w:p')) {
      if (!token.endsWith('/>')) depth++
      continue
    }
    const type = m[1]
    if (type === 'begin') {
      open.push({ depth, paraEnd: null })
    } else if (type === 'end' && open.length > 0) {
      open.pop()
    } else {
      // separate outside a field, or end with nothing to close
      if (type === 'separate' && open.length > 0) continue
      const bounds = runBounds(bodyXml, m.index)
      if (bounds) edits.push({ start: bounds[0], end: bounds[1], text: '' })
    }
  }
  for (const field of open) {
    const at = field.paraEnd ?? bodyXml.length
    edits.push({ start: at, end: at, text: '<w:r><w:fldChar w:fldCharType="end"/></w:r>' })
  }
  if (edits.length === 0) return bodyXml
  edits.sort((a, b) => b.start - a.start || b.end - a.end)
  let out = bodyXml
  for (const edit of edits) out = out.slice(0, edit.start) + edit.text + out.slice(edit.end)
  return out
}
