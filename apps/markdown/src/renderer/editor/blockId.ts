import { Extension } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { Node as PmNode, ResolvedPos } from '@tiptap/pm/model'

/**
 * Obsidian block identifiers. A block is addressable once its last line ends
 * with `^id` (Latin letters, digits and dashes); a table, a quotation or a code
 * block, which has no line of its own to append to, takes the marker alone on
 * the line after it. The marker is plain text in the file — that is what
 * Obsidian writes and what `wikiFragmentPosition` already resolves — so no
 * schema carries it: this module writes the marker where Obsidian would, hands
 * back the `[[Note#^id]]` that addresses it, and dims the marker in the view.
 */

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    blockLink: {
      /** put a `[[Note#^id]]` for the block at the caret on the clipboard, writing the id if the block has none */
      copyBlockLink: () => ReturnType
    }
  }
}

/** the marker at a line's end: line start or whitespace before it, nothing but whitespace after */
const TRAILING_ID_RE = /(?:^|\s)\^([A-Za-z0-9-]+)\s*$/
/** a paragraph that is only a marker — the form that names the block before it */
const LONE_ID_RE = /^\s*\^([A-Za-z0-9-]+)\s*$/

const ID_ALPHABET = 'abcdefghijklmnopqrstuvwxyz0123456789'
/** six characters, as Obsidian generates them */
const ID_LENGTH = 6

/** every block id the document carries, lower-cased as the resolver compares them */
export function blockIds(doc: PmNode): Set<string> {
  const ids = new Set<string>()
  doc.descendants((node) => {
    if (!node.isTextblock) return
    const match = TRAILING_ID_RE.exec(node.textContent)
    if (match) ids.add(match[1]!.toLowerCase())
  })
  return ids
}

function randomId(): string {
  const bytes = new Uint8Array(ID_LENGTH)
  globalThis.crypto.getRandomValues(bytes)
  let out = ''
  for (const byte of bytes) out += ID_ALPHABET[byte % ID_ALPHABET.length]
  return out
}

function freshId(doc: PmNode): string {
  const taken = blockIds(doc)
  for (;;) {
    const id = randomId()
    if (!taken.has(id)) return id
  }
}

/** the file part of a link to this document: the stem, '' for an untitled one (`[[#^id]]` then addresses the note itself) */
export function noteNameOf(path: string): string {
  return path.replace(/^.*[/\\]/, '').replace(/\.[^.]+$/, '')
}

export interface BlockLinkPlan {
  /** the `[[...]]` that addresses the block */
  link: string
  /** writes the marker the document lacks; absent when the block already carries one */
  write?: (tr: Transaction) => void
}

/** a line with something to address: text that is not all whitespace, or an inline node (a picture, an embed) */
function hasContent(block: PmNode): boolean {
  let found = false
  block.forEach((child) => {
    if (!child.isText || /\S/.test(child.text ?? '')) found = true
  })
  return found
}

/** depth of the outermost table or quotation around `$pos`, or -1 */
function containerDepth($pos: ResolvedPos): number {
  for (let depth = 1; depth <= $pos.depth; depth++) {
    const name = $pos.node(depth).type.name
    if (name === 'table' || name === 'blockquote') return depth
  }
  return -1
}

/**
 * The link for the block at the selection, and what writing its id takes.
 *
 * A heading is addressed by its text, `[[Note#Heading]]`, the way Obsidian
 * links headings — no id is written for one (unless its text cannot sit in a
 * link). A table, a quotation or a code block is addressed as a whole and
 * takes a lone `^id` paragraph after it, as Obsidian documents: a marker on a
 * line inside would be part of the cell, the quote or the code. A list item
 * and every other textblock take ` ^id` at the end of their line. Null when
 * nothing at the caret can carry an id: an empty line.
 */
export function planBlockLink(state: EditorState, noteName: string): BlockLinkPlan | null {
  const { selection } = state
  const prefix = `[[${noteName}#`
  const { $from } = selection

  const container = containerDepth($from)
  if (container > 0) {
    return lonePlan(state, $from.before(container), $from.node(container), prefix)
  }
  // a selected block atom (a formula, a diagram) has no line of its own
  if (selection instanceof NodeSelection && selection.node.isBlock) {
    return lonePlan(state, selection.from, selection.node, prefix)
  }
  const block = $from.parent
  if (!block.isTextblock) return null

  if (block.type.name === 'codeBlock') {
    return lonePlan(state, $from.before(), block, prefix)
  }
  if (block.type.name === 'heading') {
    const title = block.textContent.replace(/\s+/g, ' ').trim()
    // a bracket or pipe in the title would end or alias the link
    if (title && !/[[\]|]/.test(title)) return { link: `${prefix}${title}]]` }
  }

  const text = block.textContent
  const existing = TRAILING_ID_RE.exec(text)
  if (existing) return { link: `${prefix}^${existing[1]}]]` }
  if (!hasContent(block)) return null

  const id = freshId(state.doc)
  const marker = `${text && !/\s$/.test(text) ? ' ' : ''}^${id}`
  const end = $from.end()
  return {
    link: `${prefix}^${id}]]`,
    // a bare text node: inheriting the marks at the line's end would put the
    // marker inside a code span or a link and change what the save writes
    write: (tr) => tr.insert(end, state.schema.text(marker)),
  }
}

/** a lone `^id` paragraph after the block at `pos`, reused when one is there */
function lonePlan(state: EditorState, pos: number, block: PmNode, prefix: string): BlockLinkPlan {
  const after = pos + block.nodeSize
  const next = state.doc.resolve(after).nodeAfter
  if (next?.type.name === 'paragraph') {
    const existing = LONE_ID_RE.exec(next.textContent)
    if (existing) return { link: `${prefix}^${existing[1]}]]` }
  }
  const id = freshId(state.doc)
  const paragraph = state.schema.nodes.paragraph!.create(null, state.schema.text(`^${id}`))
  return { link: `${prefix}^${id}]]`, write: (tr) => tr.insert(after, paragraph) }
}

/** the marker ranges to dim: a trailing or lone `^id` that sits whole in its line's last text node */
function markerDecorations(doc: PmNode): Decoration[] {
  const out: Decoration[] = []
  doc.descendants((node, pos) => {
    // code is code: a `^id` ending a code block is what the author typed
    if (!node.isTextblock || node.type.name === 'codeBlock') return
    const match = /(?:^|\s)(\^[A-Za-z0-9-]+)(\s*)$/.exec(node.textContent)
    const last = node.lastChild
    if (!match || !last?.isText) return
    const marker = match[1]!
    const tail = match[2]!
    // split by an inline mark, the marker is left as it reads
    if (!last.text!.endsWith(marker + tail)) return
    const end = pos + node.nodeSize - 1 - tail.length
    out.push(Decoration.inline(end - marker.length, end, { class: 'md-block-id' }))
  })
  return out
}

/** exported for tests: the marker pass owns this slice of plugin state */
export const blockIdMarkerKey = new PluginKey<DecorationSet>('blockIdMarkers')

export interface BlockLinkOptions {
  /** absolute path of the open document; '' while it is untitled */
  getDocumentPath(): string
  /** the link is on the clipboard */
  onCopied(link: string): void
  /** nothing at the caret can carry an id */
  onNothing(): void
  /** seam for tests; defaults to the clipboard */
  write(text: string): Promise<void> | void
}

export const BlockLink = Extension.create<BlockLinkOptions>({
  name: 'blockLink',

  addOptions() {
    return {
      getDocumentPath: () => '',
      onCopied: () => {},
      onNothing: () => {},
      write: (text) => navigator.clipboard.writeText(text),
    }
  },

  addCommands() {
    return {
      copyBlockLink:
        () =>
        ({ state, tr, dispatch }) => {
          const plan = planBlockLink(state, noteNameOf(this.options.getDocumentPath()))
          if (!plan) {
            if (dispatch) this.options.onNothing()
            return false
          }
          if (!dispatch) return true
          plan.write?.(tr)
          Promise.resolve(this.options.write(plan.link)).then(
            () => this.options.onCopied(plan.link),
            () => {},
          )
          return true
        },
    }
  },

  addKeyboardShortcuts() {
    return {
      'Mod-Shift-l': () => this.editor.commands.copyBlockLink(),
    }
  },

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: blockIdMarkerKey,
        state: {
          init: (_config, state) => DecorationSet.create(state.doc, markerDecorations(state.doc)),
          apply: (tr, old) =>
            tr.docChanged ? DecorationSet.create(tr.doc, markerDecorations(tr.doc)) : old,
        },
        props: {
          decorations: (state) => blockIdMarkerKey.getState(state),
        },
      }),
    ]
  },
})
