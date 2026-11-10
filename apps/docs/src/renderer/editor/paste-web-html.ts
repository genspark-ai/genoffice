/**
 * Web-HTML paste font handling (r181): text copied from a website pasted into
 * a document landed in the theme font instead of the text around it.
 *
 * Two gaps compound:
 *  1. Browsers put the copied site's styles on BLOCK elements when whole
 *     paragraphs are selected (`<p style="font-family:…">text</p>`), but the
 *     docTextStyle mark only parses `span` styles — a whole-paragraph web
 *     copy lost its font/size/color entirely.
 *  2. Sites styled with only generic families (`font-family: sans-serif` —
 *     e.g. Wikipedia) produce no concrete face after the family-chain filter,
 *     so the pasted run has no font at all and renders in the document
 *     default, visibly different from the surrounding text.
 *
 * The pushdown fixes 1 at the HTML level; the fill fixes 2 at the slice
 * level by giving font-less pasted text the insertion point's font slots —
 * the same "formats like typing" rule the plain-text lane already follows.
 * Both apply only to FOREIGN fragments: our own clipboard HTML carries
 * ProseMirror's data-pm-slice wrapper and must keep its exact semantics.
 */
import type { Fragment as PmFragment, Mark, MarkType, Schema } from '@tiptap/pm/model'
import { Fragment } from '@tiptap/pm/model'

/** Our own copies (and any other ProseMirror editor's) carry a data-pm-slice
 *  wrapper; explicit run payloads carry data-doc-style. Everything else is a
 *  foreign fragment: browsers, Word, spreadsheets. */
export function isForeignPasteHtml(html: string): boolean {
  return !html.includes('data-pm-slice') && !html.includes('data-doc-style')
}

const PUSHDOWN_PROPS = ['font-family', 'font-size', 'color'] as const

/** Table subtrees are left alone: spreadsheet payloads route through the
 *  cell-unwrap lanes, and web-table cell styling keeps today's behavior. */
const TABLE_SUBTREE = /^(TABLE|THEAD|TBODY|TFOOT|TR|TD|TH)$/

/**
 * Copy block-level inline font/size/color declarations down onto the spans
 * the mark parser reads. Every text node outside table subtrees gets its
 * nearest ancestor declaration per property, either on its existing direct
 * span parent (without overriding what that span declares itself) or via a
 * new wrapping span. Returns the input unchanged on any parse failure.
 */
export function pushDownWebInlineStyles(html: string): string {
  try {
    const doc = new window.DOMParser().parseFromString(html, 'text/html')
    const body = doc.body
    const textNodes: Text[] = []
    const collect = (el: Element): void => {
      for (const child of [...el.childNodes]) {
        if (child.nodeType === 3 /* TEXT_NODE */) {
          if ((child.textContent ?? '').trim()) textNodes.push(child as Text)
        } else if (child.nodeType === 1 /* ELEMENT_NODE */) {
          const tag = (child as Element).tagName
          if (TABLE_SUBTREE.test(tag) || tag === 'SCRIPT' || tag === 'STYLE') continue
          collect(child as Element)
        }
      }
    }
    collect(body)
    if (textNodes.length === 0) return html

    // nearest inline declaration per property, from the text node's element
    // chain up to (and including) body — blocks and spans alike
    const effective = (node: Text, prop: string): string => {
      for (let el = node.parentElement; el; el = el.parentElement) {
        const v = el.style.getPropertyValue(prop)
        if (v) return v
        if (el === body) break
      }
      return ''
    }

    let changed = false
    for (const node of textNodes) {
      const parent = node.parentElement
      if (!parent) continue
      const wanted: Array<[string, string]> = []
      for (const prop of PUSHDOWN_PROPS) {
        const v = effective(node, prop)
        if (v) wanted.push([prop, v])
      }
      if (wanted.length === 0) continue
      if (parent.tagName === 'SPAN') {
        // `wanted` skipped what the span declares itself, so this only adds
        // the inherited values; nested or shared spans included, since every
        // text child inherits the same chain
        for (const [prop, v] of wanted) {
          if (!parent.style.getPropertyValue(prop)) {
            parent.style.setProperty(prop, v)
            changed = true
          }
        }
      } else {
        const span = doc.createElement('span')
        for (const [prop, v] of wanted) span.style.setProperty(prop, v)
        parent.insertBefore(span, node)
        span.appendChild(node)
        changed = true
      }
    }
    return changed ? body.innerHTML : html
  } catch {
    return html
  }
}

const FONT_SLOT_ATTRS = ['font', 'fontAscii', 'csFont', 'eaSlotEmpty'] as const

/** The docTextStyle font-slot attrs of the insertion point, or null when the
 *  caret text carries no explicit font (pasting there stays font-less and
 *  follows the document default exactly like the surrounding text). */
export function caretFontSlots(
  marks: readonly Mark[],
  markType: MarkType,
): Record<string, unknown> | null {
  const style = marks.find((m) => m.type === markType)
  if (!style) return null
  const slots: Record<string, unknown> = {}
  for (const attr of FONT_SLOT_ATTRS) {
    if (style.attrs[attr] != null) slots[attr] = style.attrs[attr]
  }
  return 'font' in slots || 'fontAscii' in slots || 'csFont' in slots ? slots : null
}

/**
 * Give every pasted text run that ends up with no concrete font the
 * insertion point's font slots (size/color/etc. keep the source values —
 * only the unresolved font falls back, like Word mapping an unknown face).
 * Tables are skipped, matching the pushdown: cell text keeps inheriting.
 */
export function fillMissingRunFonts(
  fragment: PmFragment,
  slots: Record<string, unknown>,
  schema: Schema,
): PmFragment {
  const markType = schema.marks.docTextStyle
  if (!markType) return fragment
  const mapped: import('@tiptap/pm/model').Node[] = []
  fragment.forEach((node) => {
    if (node.type.spec.tableRole === 'table') {
      mapped.push(node)
      return
    }
    if (node.isText) {
      const style = node.marks.find((m) => m.type === markType)
      if (style && (style.attrs.font || style.attrs.fontAscii || style.attrs.csFont)) {
        mapped.push(node)
        return
      }
      const attrs = style ? { ...style.attrs, ...slots } : slots
      const mark = markType.create(attrs)
      mapped.push(node.mark(mark.addToSet(node.marks)))
      return
    }
    mapped.push(node.copy(fillMissingRunFonts(node.content, slots, schema)))
  })
  return Fragment.from(mapped)
}
