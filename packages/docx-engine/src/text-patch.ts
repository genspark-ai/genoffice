import { decodeEntities } from './parse-xml-text'
import { escapeXmlText } from './xml-utils'

/**
 * Paragraph-level surgical text patch (saving edits to rich-text entries such as
 * comments/footnotes).
 *
 * Maps the edited plain text ('\n' between paragraphs, each paragraph = concatenated w:t)
 * back onto the original entry XML:
 * - Paragraphs whose text is unchanged keep their original bytes (inline formatting,
 *   hyperlinks, images, and fields are all preserved)
 * - Changed paragraphs get a minimal w:t replacement: only the w:t outside the common
 *   prefix/suffix are rewritten, other run bytes stay untouched
 * - If the paragraph count changed, or a changed paragraph has no w:t to anchor to,
 *   returns null and the caller falls back to rebuilding the whole entry
 *
 * stripFirstParaLeadingSpace: footnote/endnote first paragraphs start with a
 * self-reference mark + space run, which the plain-text model drops; when enabled,
 * that leading whitespace is treated as an immutable prefix.
 */
export function patchParagraphTexts(
  entryXml: string,
  newText: string,
  opts: { stripFirstParaLeadingSpace?: boolean } = {},
): string | null {
  const paras = paragraphSlices(entryXml)
  if (paras.length === 0) return null
  const oldParaTexts = paras.map((p, i) =>
    i === 0 && opts.stripFirstParaLeadingSpace
      ? decodedTextOf(p.xml).replace(/^\s+/, '')
      : decodedTextOf(p.xml),
  )
  if (oldParaTexts.join('\n') === newText) return entryXml
  const newParaTexts = newText.split('\n')
  if (newParaTexts.length !== paras.length) return null

  let out = ''
  let cursor = 0
  for (let i = 0; i < paras.length; i++) {
    const para = paras[i]
    out += entryXml.slice(cursor, para.start)
    if (oldParaTexts[i] === newParaTexts[i]) {
      out += para.xml
    } else {
      const skipLeading = i === 0 && opts.stripFirstParaLeadingSpace === true
      const patched = patchOneParagraph(para.xml, newParaTexts[i], skipLeading)
      if (patched === null) return null
      out += patched
    }
    cursor = para.end
  }
  out += entryXml.slice(cursor)

  // Self-check: re-extracted text must match the target, otherwise abandon the surgical patch
  const check = paragraphSlices(out).map((p, i) =>
    i === 0 && opts.stripFirstParaLeadingSpace
      ? decodedTextOf(p.xml).replace(/^\s+/, '')
      : decodedTextOf(p.xml),
  )
  return check.join('\n') === newText ? out : null
}

interface ParaSlice {
  start: number
  end: number
  xml: string
}

function paragraphSlices(xml: string): ParaSlice[] {
  // depth-aware scan: a paragraph may legally nest more <w:p> inside embedded
  // textbox content (w:txbxContent), so a lazy open-to-first-close match would
  // truncate the outer paragraph
  const out: ParaSlice[] = []
  const re = /<w:p(?:\s[^>]*)?\/>|<w:p(?:\s[^>]*)?>|<\/w:p>/g
  let m: RegExpExecArray | null
  let depth = 0
  let start = -1
  while ((m = re.exec(xml)) !== null) {
    const token = m[0]
    if (token === '</w:p>') {
      if (depth === 0) continue
      depth--
      if (depth === 0 && start !== -1) {
        const end = m.index + token.length
        out.push({ start, end, xml: xml.slice(start, end) })
        start = -1
      }
    } else if (token.endsWith('/>')) {
      if (depth === 0) out.push({ start: m.index, end: m.index + token.length, xml: token })
    } else {
      if (depth === 0) start = m.index
      depth++
    }
  }
  return out
}

interface TSlice {
  /** offsets of the inner content within the paragraph XML */
  innerStart: number
  innerEnd: number
  /** offset of the opening <w:t ...> tag */
  tagStart: number
  text: string
}

function tSlices(paraXml: string): TSlice[] {
  const out: TSlice[] = []
  const re = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(paraXml)) !== null) {
    const inner = m[2]
    const innerStart = m.index + m[0].length - inner.length - '</w:t>'.length
    out.push({
      tagStart: m.index,
      innerStart,
      innerEnd: innerStart + inner.length,
      text: decodeEntities(inner),
    })
  }
  return out
}

function decodedTextOf(paraXml: string): string {
  return tSlices(paraXml)
    .map((t) => t.text)
    .join('')
}

function patchOneParagraph(paraXml: string, newText: string, skipLeading: boolean): string | null {
  const slices = tSlices(paraXml)
  if (slices.length === 0) return null
  const raw = slices.map((t) => t.text).join('')
  // First paragraph's leading whitespace (spacer run after the self-reference mark) is
  // excluded from the diff and treated as an immutable prefix
  const lead = skipLeading ? raw.length - raw.replace(/^\s+/, '').length : 0
  const oldText = raw.slice(lead)
  if (oldText === newText) return paraXml

  let prefix = 0
  const maxCommon = Math.min(oldText.length, newText.length)
  while (prefix < maxCommon && oldText[prefix] === newText[prefix]) prefix++
  let suffix = 0
  while (
    suffix < maxCommon - prefix &&
    oldText[oldText.length - 1 - suffix] === newText[newText.length - 1 - suffix]
  ) {
    suffix++
  }
  // Changed range (relative to the full raw text)
  const changeStart = lead + prefix
  const changeEnd = lead + oldText.length - suffix
  const replacement = newText.slice(prefix, newText.length - suffix)

  // Find the w:t intersecting the changed range; for pure insertion (changeStart ===
  // changeEnd), take the slice covering that point
  let acc = 0
  let first = -1
  let last = -1
  const bounds: Array<{ start: number; end: number }> = []
  for (let i = 0; i < slices.length; i++) {
    const start = acc
    const end = acc + slices[i].text.length
    bounds.push({ start, end })
    const touches =
      changeStart === changeEnd
        ? start <= changeStart && changeStart <= end
        : start < changeEnd && end > changeStart
    if (touches) {
      if (first === -1) first = i
      last = i
    }
    acc = end
  }
  if (first === -1) {
    // Insertion point falls outside every w:t (in theory only an empty paragraph) —
    // no anchor, give up
    return null
  }

  // Rewrite [first..last]: first takes the prefix remainder + replacement text + last's
  // suffix remainder; the rest are emptied
  const newInner: string[] = []
  for (let i = first; i <= last; i++) {
    if (i === first) {
      const head = slices[i].text.slice(0, changeStart - bounds[i].start)
      const tail = slices[last].text.slice(changeEnd - bounds[last].start)
      newInner.push(escapeXmlText(head + replacement + tail))
    } else {
      newInner.push('')
    }
  }

  // Replace back-to-front so offsets stay valid; changed w:t all get xml:space="preserve"
  let out = paraXml
  for (let i = last; i >= first; i--) {
    const s = slices[i]
    const closeEnd = s.innerEnd + '</w:t>'.length
    out =
      out.slice(0, s.tagStart) +
      `<w:t xml:space="preserve">${newInner[i - first]}</w:t>` +
      out.slice(closeEnd)
  }
  return out
}
