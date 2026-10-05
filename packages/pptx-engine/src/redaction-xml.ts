/**
 * Reading and writing the "withheld from the model" marker.
 *
 * A span is a *mark over real words*, not a replacement for them: the deck and
 * the .pptx on disk keep the actual text, and only the model's view swaps it for
 * `{{label}}`. Losing the reader's own data to hide it from a model would be the
 * worse failure, so nothing here ever removes content.
 *
 * ## Why the marker has to be modelled
 *
 * The obvious idea is to leave the mark in the XML and let it ride along, the way
 * the engine already passes untouched `ByteAnchor.originalXml` through verbatim.
 * That works for a save that patches in place, and this engine was measured to
 * confirm it — but it breaks on the other path:
 *
 * - `patchTextElementXml` takes its **aligned** path when the model and the XML
 *   have the same number of runs, and patches each `<a:t>` in place. An
 *   `<a:extLst>` already inside `<a:rPr>` survives: the patcher rewrites the
 *   attributes it manages and keeps the rest of the node's inner XML.
 * - When the run count changes — splitting a run in half is exactly what marking
 *   half a word does — it falls through to `rebuildTxBody`, which regenerates
 *   `<a:txBody>` from the model. Anything the model does not carry is gone.
 *
 * So the mark is read into the model on parse and written back on save, in both
 * paths. The same measurement showed the visible half is not free either: the
 * patch path writes `u="none"` for a run whose model does not carry `underline`,
 * so the underline that shows the span has to be set by the caller as well.
 *
 * ## Where it lives
 *
 * `<a:extLst><a:ext uri="…"><go:redact w:label="…"/></a:ext></a:extLst>` inside
 * the run's `<a:rPr>`, or inside the element's `nvPr` for a picture. `a:extLst`
 * is the OOXML extension list — the schema's own slot for things it does not
 * know — so PowerPoint round-trips the file without noticing, and fast-xml-parser
 * hands the label back on the way in (verified: no element whitelist, so the
 * node survives parsing).
 *
 * `w:bdr` in docx and an inline `<span data-redaction>` in markdown are the same
 * idea in their own formats; the underline here is the counterpart of the
 * character border there.
 */

import { asXmlNode, escapeXmlAttr, xmlArray, type XmlNode } from './xml-utils'

/**
 * A fixed URI for our extension. Office's convention is that the `uri` is a
 * GUID identifying the extension's schema, so it must be stable: a deck that
 * round-trips through this engine has to land on the same value.
 */
export const REDACT_EXT_URI = '{9F1B2C3D-4E5F-4A6B-8C7D-9E0F1A2B3C4D}'

const REDACT_EL = 'go:redact'

/** The whole extension node, for splicing into an `<a:extLst>`. */
export function redactExtXml(label: string): string {
  return `<a:ext uri="${REDACT_EXT_URI}"><${REDACT_EL} w:label="${escapeXmlAttr(label)}"/></a:ext>`
}

/** True when the fragment already carries our mark. */
export function hasRedactExtIn(xml: string): boolean {
  // a /g regex keeps its lastIndex between .test() calls, so the answer would
  // alternate true/false on consecutive runs — this runs once per run per save
  OUR_EXT_RE.lastIndex = 0
  return OUR_EXT_RE.test(xml)
}

/**
 * Sync our mark on the element named `tag` inside `xml` — `a:rPr` for a run, the
 * picture's `nvPr` for media. A `label` of undefined removes it.
 *
 * Both save paths call this: the aligned patch path, which would otherwise leave
 * a mark the model no longer claims, and the rebuild path, which has just
 * regenerated the node from the model and needs the mark put back.
 */
export function syncRedactExt(xml: string, tag: string, label?: string): string {
  const open = new RegExp(`<${tag}\\b[^>]*?/>|<${tag}\\b[^>]*?>[\\s\\S]*?</${tag}>`).exec(xml)
  if (!open) return xml
  const next = label ? setRedactExt(open[0], label) : stripRedactExt(open[0])
  if (next === open[0]) return xml
  return xml.slice(0, open.index) + next + xml.slice(open.index + open[0].length)
}

/** One `<a:ext>` matching our URI, or null. */
function ourExt(extLst: unknown): XmlNode | null {
  for (const ext of xmlArray(asXmlNode(extLst)['a:ext'])) {
    const uri = ext['@_uri']
    if (uri === REDACT_EXT_URI) return ext
  }
  return null
}

/** The label, read out of an already-parsed `<a:rPr>` / `nvPr` node. */
export function readRedactLabel(node: unknown): string | undefined {
  const ext = ourExt(asXmlNode(node)['a:extLst'])
  if (!ext) return undefined
  const label = asXmlNode(ext[REDACT_EL])['@_w:label']
  return typeof label === 'string' && label !== '' ? label : undefined
}

/** Our `<a:ext>` as raw bytes, whether self-closing or paired. */
const OUR_EXT_RE = new RegExp(
  `<a:ext\\b[^>]*\\buri="${REDACT_EXT_URI}"[^>]*(?:/>|>[\\s\\S]*?</a:ext>)`,
  'g',
)

/** An `<a:extLst>` left holding nothing, either form. */
const EMPTY_EXTLST_RE = /<a:extLst\b[^>]*?\/>|<a:extLst\b[^>]*>([\s\s]*?)<\/a:extLst>/g

/** Remove our extension from a raw `<a:rPr>` / `nvPr` string, leaving everything else. */
export function stripRedactExt(xml: string): string {
  const withoutOurs = xml.replace(OUR_EXT_RE, '')
  // an apply/unapply cycle would otherwise leave an empty <a:extLst> behind,
  // and they would accumulate in a deck someone toggles marks in
  const pruned = withoutOurs.replace(EMPTY_EXTLST_RE, (_all, inner: string | undefined) =>
    inner?.trim() ? `<a:extLst>${inner}</a:extLst>` : '',
  )
  return collapseEmptyProps(pruned)
}

/**
 * Adding the mark gives a self-closing `<a:rPr/>` the paired form it needs;
 * take it back once the mark is gone, so toggling leaves no trace. Only an
 * empty node is collapsed — one that still holds a child keeps its form.
 */
function collapseEmptyProps(xml: string): string {
  return xml.replace(
    /<([\w:]+)\b([^>]*?)>([\s\s]*?)<\/\1>/g,
    (all, tag: string, attrs: string, inner: string) =>
      inner.trim() ? all : `<${tag}${attrs.replace(/\s+$/, '')}/>`,
  )
}

/**
 * Put the mark in a raw `<a:rPr>` (or `nvPr`) string, replacing any previous one.
 *
 * The `inner` of the node is preserved verbatim; only our own extension is added
 * or swapped. A self-closing `<a:rPr/>` gains a paired form because the mark
 * needs a child.
 */
export function setRedactExt(propsXml: string, label: string): string {
  const ext = redactExtXml(label)
  const open = /<([\w:]+)\b[^>]*?(\/?)>/.exec(propsXml)
  if (!open) return propsXml
  const [whole, tag, selfClose] = open
  const at = open.index
  if (selfClose === '/') {
    // self-closing: the mark needs a child, so it gains a paired form
    const head = whole.slice(0, -2) + '>'
    return (
      propsXml.slice(0, at) +
      `${head}<a:extLst>${ext}</a:extLst></${tag}>` +
      propsXml.slice(at + whole.length)
    )
  }
  const close = `</${tag}>`
  const closeAt = propsXml.lastIndexOf(close)
  if (closeAt === -1) return propsXml
  const inner = propsXml.slice(at + whole.length, closeAt)
  return (
    propsXml.slice(0, at) +
    whole +
    spliceExtLst(inner, ext) +
    close +
    propsXml.slice(closeAt + close.length)
  )
}

/** Replace our extension inside an existing `<a:extLst>`, or add one to the inner XML. */
function spliceExtLst(inner: string, ext: string): string {
  const list = /<a:extLst\b[^>]*?(?:\/>|>[\s\S]*?<\/a:extLst>)/.exec(inner)
  if (!list) return `<a:extLst>${ext}</a:extLst>${inner}`
  const kept = list[0].replace(OUR_EXT_RE, '')
  // our own extension goes last so a re-apply replaces rather than appends
  const next = kept.includes('</a:extLst>')
    ? kept.replace('</a:extLst>', `${ext}</a:extLst>`)
    : kept.replace(/\/>$/, `>${ext}</a:extLst>`)
  return inner.slice(0, list.index) + next + inner.slice(list.index + list[0].length)
}
