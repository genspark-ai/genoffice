import { Extension } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { touchedTopLevelBlocks } from './touched-blocks'

// Word internal bookmarks (w:bookmarkStart — _Toc…, _Ref…) are kept on the
// paragraph as names only; the rendered link is a real <a href="#name">, but a
// fragment link is only clickable where the PDF (or the DOM) has an element
// with that id. Chromium's printToPDF turns href="#id" into a GoTo
// annotation, so each bookmark name gets a zero-size anchor widget at the
// start of its block: nothing in the document model moves (copying, undo and
// the DOCX save patch never see it), and no line box gains content.

function decorationsIn(doc: PmNode, from: number, to: number): Decoration[] {
  const widgets: Decoration[] = []
  doc.nodesBetween(from, to, (node, pos) => {
    const names = node.attrs.bookmarks as string[] | null
    if (!Array.isArray(names) || names.length === 0 || !node.isTextblock) return
    for (const name of names) {
      widgets.push(
        Decoration.widget(
          pos + 1,
          () => {
            const anchor = document.createElement('a')
            anchor.className = 'doc-bookmark-target'
            anchor.setAttribute('id', name)
            return anchor
          },
          // before the caret/anything else at the same position, so a link
          // landing at the block start resolves to this anchor first
          { side: -1, ignoreSelection: true },
        ),
      )
    }
    // the names ride the block; its children hold no bookmark attrs of their own
    return false
  })
  return widgets
}

const key = new PluginKey<DecorationSet>('bookmarkTargets')

export const BookmarkTargets = Extension.create({
  name: 'bookmarkTargets',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: (_, state) =>
            DecorationSet.create(state.doc, decorationsIn(state.doc, 0, state.doc.content.size)),
          // same economy as scriptFonts: only the top-level blocks a
          // transaction touched are rescanned, the rest maps through the change
          apply(tr, old) {
            if (!tr.docChanged) return old
            const touched = touchedTopLevelBlocks(tr)
            if (!touched)
              return DecorationSet.create(tr.doc, decorationsIn(tr.doc, 0, tr.doc.content.size))
            let set = old.map(tr.mapping, tr.doc)
            for (const pos of touched) {
              const block = tr.doc.nodeAt(pos)
              if (!block) continue
              const to = pos + block.nodeSize
              set = set.remove(set.find(pos, to))
              set = set.add(tr.doc, decorationsIn(tr.doc, pos, to))
            }
            return set
          },
        },
        props: { decorations: (state) => key.getState(state) },
      }),
    ]
  },
})
