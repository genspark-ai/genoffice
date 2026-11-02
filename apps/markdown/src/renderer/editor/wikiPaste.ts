import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Transaction } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import { WIKI_LINK_RE } from './wikilink'
import { WIKI_EMBED_RE } from './wikiEmbed'

/**
 * Pasted wiki syntax goes live. The markdown tokenizers run on load, not on a
 * paste, so `[[Note]]` pasted from Obsidian — or the `[[Note#^id]]` the
 * copy-link button put on the clipboard — arrived as plain text that only
 * became a link on the next open. After a paste, the `[[...]]` literals in
 * the pasted text take the wikiLink mark and the `![[...]]` ones become embed
 * nodes, the shapes the tokenizers build, so the resolver styles and opens
 * them at once. Typing is left alone: the `[[` suggester is the live path
 * there, and a half-typed literal must not turn into an atom under the caret.
 */

interface Conversion {
  from: number
  to: number
  /** the embed literal to become a node; null for a link, which keeps its text */
  embed: string | null
}

/** the `[[...]]` and `![[...]]` literals touching [from, to]; typed ones elsewhere in the block stay text */
function conversions(doc: PmNode, from: number, to: number): Conversion[] {
  const out: Conversion[] = []
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true
    if (node.type.name === 'codeBlock') return false
    node.forEach((child, offset) => {
      if (!child.isText) return
      // already a link; or code, where brackets are brackets
      if (child.marks.some((mark) => mark.type.name === 'wikiLink' || mark.type.name === 'code'))
        return
      const text = child.text ?? ''
      const base = pos + 1 + offset
      let at = 0
      while (at < text.length) {
        const open = text.indexOf('[[', at)
        if (open < 0) break
        // `\[[` is an escaped literal, as the tokenizers read it
        if (open > 0 && text[open - 1] === '\\') {
          at = open + 2
          continue
        }
        const bang = open > 0 && text[open - 1] === '!' && !(open > 1 && text[open - 2] === '\\')
        const start = bang ? open - 1 : open
        const match = (bang ? WIKI_EMBED_RE : WIKI_LINK_RE).exec(text.slice(start))
        if (!match) {
          at = open + 2
          continue
        }
        const matchFrom = base + start
        const matchTo = matchFrom + match[0].length
        if (matchTo > from && matchFrom < to) {
          out.push({ from: matchFrom, to: matchTo, embed: bang ? match[0] : null })
        }
        at = start + match[0].length
      }
    })
    return false
  })
  return out
}

/** the range the paste transactions of this batch wrote, in the final document */
function pastedRange(transactions: readonly Transaction[]): { from: number; to: number } | null {
  let from = Infinity
  let to = -Infinity
  transactions.forEach((tr, index) => {
    if (!tr.docChanged || tr.getMeta('uiEvent') !== 'paste') return
    tr.mapping.maps.forEach((map, step) => {
      map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        // through the rest of this transaction, then through the later ones
        const rest = tr.mapping.slice(step + 1)
        let start = rest.map(newStart, -1)
        let end = rest.map(newEnd, 1)
        for (const later of transactions.slice(index + 1)) {
          start = later.mapping.map(start, -1)
          end = later.mapping.map(end, 1)
        }
        from = Math.min(from, start)
        to = Math.max(to, end)
      })
    })
  })
  return from === Infinity ? null : { from, to }
}

export const wikiPasteKey = new PluginKey('wikiPasteLive')

export const WikiPasteLive = Extension.create({
  name: 'wikiPasteLive',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: wikiPasteKey,
        appendTransaction: (transactions, _old, state) => {
          const range = pastedRange(transactions)
          if (!range) return null
          const found = conversions(state.doc, range.from, range.to)
          if (found.length === 0) return null
          const tr = state.tr
          const { schema } = state
          // from the end, so the positions before each replacement stay valid
          for (const c of found.sort((a, b) => b.from - a.from)) {
            if (c.embed !== null) {
              tr.replaceWith(c.from, c.to, schema.nodes.wikiEmbed!.create({ raw: c.embed }))
            } else {
              tr.addMark(c.from, c.to, schema.marks.wikiLink!.create())
            }
          }
          return tr
        },
      }),
    ]
  },
})
