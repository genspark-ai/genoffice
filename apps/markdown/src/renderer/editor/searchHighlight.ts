import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

export interface SearchRange {
  from: number
  to: number
}

interface SearchHighlight {
  ranges: SearchRange[]
  activeIndex: number
}

export const searchPluginKey = new PluginKey<DecorationSet>('mdSearch')

/** decorates find hits; the find target pushes ranges via setMeta(searchPluginKey) */
export const SearchHighlight = Extension.create({
  name: 'mdSearch',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: searchPluginKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            const meta = tr.getMeta(searchPluginKey) as SearchHighlight | undefined
            if (meta) {
              return DecorationSet.create(
                tr.doc,
                meta.ranges.map((r, i) =>
                  Decoration.inline(r.from, r.to, {
                    class: i === meta.activeIndex ? 'search-hit search-hit-active' : 'search-hit',
                  }),
                ),
              )
            }
            return old.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
      }),
    ]
  },
})
