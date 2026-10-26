import { Extension, Node, mergeAttributes, nodeInputRule } from '@tiptap/core'
import type { Editor, MarkdownLexerConfiguration, MarkdownToken } from '@tiptap/core'
import { Plugin, TextSelection } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'
import { renderItem } from './markdownStyleRenderers'

/**
 * GFM footnotes, the two halves GitHub renders: a `[^label]` reference inline
 * and a `[^label]: text` definition at block level. The label holds no
 * whitespace or brackets. The definitions are ordinary top-level blocks —
 * draggable, movable, and paired with their source blocks by the block source
 * map — so a footnote file survives the lossless save untouched.
 */

/** `[^label]:` at a line start, up to three leading spaces; the body follows on the line */
const DEF_OPEN_RE = /^ {0,3}\[\^([^\][\s]+)\]:[^\S\n]*(.*)$/
/** a `[^label]` reference; anchored because the tokenizer answers for src[0] */
const REF_RE = /^\[\^([^\][\s]+)\]/
/** lines that end a lazy continuation: another definition, or a block of its own */
const INTERRUPT_RE =
  /^ {0,3}(?:#{1,6}(?:\s|$)|`{3,}|~{3,}|[-+*](?:\s|$)|\d{1,9}[.)](?:\s|$)|>|[-*_](?:[ \t]*[-*_]){2,}$|<[a-zA-Z][^>]*>)/

/** marked block tokens for a source fragment; supplied by the tokenizer helper */
type BlockLexer = (src: string) => MarkdownToken[]

export interface FootnoteDefinitionToken {
  type: 'footnoteDefinition'
  raw: string
  label: string
  tokens: MarkdownToken[]
}

/**
 * The lines of one definition: the `[^label]:` line, its indented body, and
 * lazy continuations directly below it. A blank line joins the body only when
 * the next line is indented (a second paragraph); a non-indented line after a
 * blank line ends the definition. The raw must tile the source exactly, so it
 * ends after the last kept line and leaves trailing blanks to the space
 * tokenizer.
 */
export function tokenizeFootnoteDefinition(
  src: string,
  blockTokens: BlockLexer,
): FootnoteDefinitionToken | undefined {
  const firstBreak = src.indexOf('\n')
  // the opening regex judges the first line alone: `$` cannot see past a
  // newline, so matching the whole source would miss every definition that
  // has blocks below it
  const open = DEF_OPEN_RE.exec(firstBreak < 0 ? src : src.slice(0, firstBreak))
  if (!open) return undefined
  const label = open[1]!
  const bodyLines: string[] = [open[2] ?? '']
  let keptEnd = firstBreak < 0 ? src.length : firstBreak
  let pos = keptEnd
  let pendingBlanks: string[] = []
  while (pos < src.length) {
    const lineStart = pos + 1
    if (lineStart > src.length) break
    const lineEnd = src.indexOf('\n', lineStart)
    const end = lineEnd < 0 ? src.length : lineEnd
    const line = src.slice(lineStart, end)
    if (/^\s*$/.test(line)) {
      pendingBlanks.push(line)
      pos = end
      continue
    }
    if (line[0] === '\t' || /^ {4}/.test(line)) {
      bodyLines.push(...pendingBlanks, line[0] === '\t' ? line.slice(1) : line.slice(4))
      pendingBlanks = []
      keptEnd = end
      pos = end
      continue
    }
    // a non-indented line continues the note only directly below its first line
    if (pendingBlanks.length === 0 && !DEF_OPEN_RE.test(line) && !INTERRUPT_RE.test(line)) {
      bodyLines.push(line)
      keptEnd = end
      pos = end
      continue
    }
    break
  }
  return {
    type: 'footnoteDefinition',
    raw: src.slice(0, keptEnd),
    label,
    tokens: blockTokens(`${bodyLines.join('\n')}\n`),
  }
}

/**
 * Offset of the first line-start `[^label]:` after a newline. marked consults
 * this to cut the paragraph that would otherwise swallow the definition line,
 * and it probes the source minus its first character, so position 0 — never a
 * true line start when a paragraph is being cut — is skipped.
 */
export function footnoteDefStart(src: string): number {
  let br = src.indexOf('\n')
  while (br >= 0) {
    const lineStart = br + 1
    const next = src.indexOf('\n', lineStart)
    const line = src.slice(lineStart, next < 0 ? undefined : next)
    if (DEF_OPEN_RE.test(line)) return lineStart
    br = next
  }
  return -1
}

/** offset of the first `[^label]` reference, for marked's inlineText probe */
export function footnoteRefStart(src: string): number {
  let at = src.indexOf('[^')
  while (at >= 0) {
    if (src[at - 1] !== '\\') {
      const match = REF_RE.exec(src.slice(at))
      if (match) return at
    }
    at = src.indexOf('[^', at + 1)
  }
  return -1
}

const tokenizeReference = (
  src: string,
): { type: 'footnoteReference'; raw: string; label: string } | undefined => {
  const match = REF_RE.exec(src)
  if (!match) return undefined
  return { type: 'footnoteReference', raw: match[0], label: match[1]! }
}

export const FootnoteReference = Node.create({
  name: 'footnoteReference',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,

  addAttributes() {
    return {
      label: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-footnote') ?? '',
        renderHTML: (attributes) => ({ 'data-footnote': String(attributes.label ?? '') }),
      },
    }
  },

  parseHTML() {
    return [{ tag: 'sup[data-footnote]' }]
  },

  renderHTML({ node }) {
    const label = String(node.attrs.label ?? '')
    return [
      'sup',
      mergeAttributes({ class: 'md-footnote-ref', 'data-footnote': label }),
      `[^${label}]`,
    ]
  },

  markdownTokenizer: {
    name: 'footnoteReference',
    level: 'inline',
    start: footnoteRefStart,
    tokenize: tokenizeReference,
  },

  parseMarkdown: (token, h) =>
    h.createNode('footnoteReference', { label: String(token.label ?? '') }),

  renderMarkdown: (node) => `[^${String(node.attrs?.label ?? '')}]`,

  addInputRules() {
    return [
      nodeInputRule({
        // a backtick or an open bracket before the caret means inline code or a
        // `[[wiki` under construction — not a reference
        find: /(?<![`[])\[\^([^\][\s]+)\]$/,
        type: this.type,
        getAttributes: (match) => ({ label: match[1] ?? '' }),
      }),
    ]
  },
})

export const FootnoteDefinition = Node.create({
  name: 'footnoteDefinition',
  group: 'block',
  content: 'block+',
  defining: true,

  addAttributes() {
    return {
      label: {
        default: '',
        parseHTML: (element) => element.getAttribute('data-footnote') ?? '',
        renderHTML: (attributes) => ({ 'data-footnote': String(attributes.label ?? '') }),
      },
    }
  },

  parseHTML() {
    return [{ tag: 'div[data-footnote]', contentElement: '.md-footnote-def-body' }]
  },

  renderHTML({ node }) {
    const label = String(node.attrs.label ?? '')
    return [
      'div',
      mergeAttributes({ class: 'md-footnote-def', 'data-footnote': label }),
      ['div', { class: 'md-footnote-def-label', contenteditable: 'false' }, `[^${label}]:`],
      ['div', { class: 'md-footnote-def-body' }, 0],
    ]
  },

  markdownTokenizer: {
    name: 'footnoteDefinition',
    level: 'block',
    start: footnoteDefStart,
    tokenize: (src: string, _tokens: MarkdownToken[], lexer: MarkdownLexerConfiguration) =>
      tokenizeFootnoteDefinition(src, (s) => lexer.blockTokens(s)),
  },

  parseMarkdown: (token, h) => {
    const children = h.parseBlockChildren?.(token.tokens ?? []) ?? []
    return h.createNode(
      'footnoteDefinition',
      { label: String(token.label ?? '') },
      children.length > 0 ? children : [h.createNode('paragraph')],
    )
  },

  /** body on the definition line, later blocks indented four and blank-line separated */
  renderMarkdown: (node, h) =>
    renderItem(node, h, `[^${String(node.attrs?.label ?? '')}]: `, true, 4),
})

const firstOfType = (doc: PmNode, type: string, label: string): number | null => {
  let found: number | null = null
  doc.descendants((node, pos) => {
    if (found !== null) return false
    if (node.type.name === type && String(node.attrs.label ?? '') === label) found = pos
  })
  return found
}

/** brief highlight on the jumped-to node; a class on the node's own DOM, no document state */
function flash(view: EditorView, pos: number): void {
  const el = view.nodeDOM(pos) as HTMLElement | null
  if (!el) return
  el.scrollIntoView({ behavior: 'smooth', block: 'center' })
  el.classList.add('md-footnote-flash')
  window.setTimeout(() => el.classList.remove('md-footnote-flash'), 1200)
}

/** click a reference to jump to its definition, and a definition back to its first reference */
export const FootnoteNav = Extension.create({
  name: 'footnoteNav',
  // above the core keymap (100): the Enter shortcut below must answer before
  // the core chain lifts the caret out through its own generic paths
  priority: 1000,

  addProseMirrorPlugins() {
    return [
      new Plugin({
        props: {
          handleClickOn: (view, _pos, node, _nodePos, event) => {
            if (node.type.name !== 'footnoteReference' && node.type.name !== 'footnoteDefinition') {
              return false
            }
            // clicks into the definition body are edits, only the label jumps back
            if (
              node.type.name === 'footnoteDefinition' &&
              !(event.target as Element | null)?.closest?.('.md-footnote-def-label')
            ) {
              return false
            }
            const target = firstOfType(
              view.state.doc,
              node.type.name === 'footnoteReference' ? 'footnoteDefinition' : 'footnoteReference',
              String(node.attrs.label ?? ''),
            )
            if (target === null) return false
            flash(view, target)
            return true
          },
        },
      }),
    ]
  },

  addKeyboardShortcuts() {
    return {
      /** Enter in the trailing empty paragraph leaves the definition instead of nesting */
      Enter: () => {
        const { state } = this.editor
        const { $from, empty } = state.selection
        if (!empty) return false
        if ($from.parent.type.name !== 'paragraph' || $from.parent.content.size > 0) return false
        if ($from.depth < 2) return false
        const def = $from.node($from.depth - 1)
        if (def.type.name !== 'footnoteDefinition') return false
        if ($from.index($from.depth - 1) !== def.childCount - 1) return false
        const defStart = $from.before($from.depth - 1)
        const paragraph = state.schema.nodes.paragraph.create()
        const tr = state.tr
        if (def.childCount === 1) {
          // the empty paragraph is the definition's only content: leaving would
          // strand an empty definition, so the definition becomes the paragraph
          tr.replaceWith(defStart, $from.after($from.depth - 1), paragraph)
          tr.setSelection(TextSelection.near(tr.doc.resolve(defStart + 1)))
        } else {
          // drop the empty paragraph; the definition shrinks by exactly its size
          const paraStart = $from.before()
          const paraSize = $from.parent.nodeSize
          const afterDef = defStart + def.nodeSize - paraSize
          tr.delete(paraStart, paraStart + paraSize)
          tr.insert(afterDef, paragraph)
          tr.setSelection(TextSelection.near(tr.doc.resolve(afterDef + 1)))
        }
        this.editor.view.dispatch(tr)
        return true
      },
    }
  },
})

/** next free numeric label: one past the highest `[^n]` in the document */
export function nextFootnoteLabel(doc: PmNode): string {
  let max = 0
  let any = false
  doc.descendants((node) => {
    if (node.type.name !== 'footnoteReference' && node.type.name !== 'footnoteDefinition') return
    const label = String(node.attrs.label ?? '')
    if (!/^\d+$/.test(label)) return
    any = true
    max = Math.max(max, parseInt(label, 10))
  })
  return String(any ? max + 1 : 1)
}

/**
 * The /footnote gesture: a reference at the caret and an empty definition at
 * the end of the document, caret placed in the definition to type the note.
 */
export function insertFootnote(editor: Editor): boolean {
  const refType = editor.schema.nodes.footnoteReference
  const defType = editor.schema.nodes.footnoteDefinition
  if (!refType || !defType) return false
  const label = nextFootnoteLabel(editor.state.doc)
  const caret = editor.state.selection.from
  const end = editor.state.doc.content.size
  return (
    editor
      .chain()
      .insertContentAt(caret, { type: 'footnoteReference', attrs: { label } })
      // the reference grew the document by one; the definition goes at its new end
      .insertContentAt(end + 1, {
        type: 'footnoteDefinition',
        attrs: { label },
        content: [{ type: 'paragraph' }],
      })
      // the definition's inner paragraph starts two opens in
      .setTextSelection(end + 3)
      .focus()
      .run()
  )
}
