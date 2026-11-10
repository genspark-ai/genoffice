import type { Editor } from '@tiptap/core'
import { findInText, type FindOptions, type FindTarget } from '@genoffice/ui'
import { searchPluginKey, type SearchRange } from './searchHighlight'

/** collect matches inside textblocks; inline content is flattened so matches spanning marks are found */
export function findMatches(editor: Editor, query: string, opts: FindOptions): SearchRange[] {
  const found: SearchRange[] = []
  if (!query) return found
  editor.state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    let text = ''
    const posAt: number[] = []
    node.forEach((child, offset) => {
      if (child.isText && child.text) {
        for (let k = 0; k < child.text.length; k++) posAt.push(pos + 1 + offset + k)
        text += child.text
      } else {
        posAt.push(pos + 1 + offset)
        text += '\0' // leaf placeholder (hard break, math) never matches
      }
    })
    for (const i of findInText(text, query, opts)) {
      found.push({ from: posAt[i]!, to: posAt[i + query.length - 1]! + 1 })
    }
    return false
  })
  return found
}

export function tiptapFindTarget(editor: Editor): FindTarget {
  let ranges: SearchRange[] = []
  const highlight = (activeIndex: number) => {
    editor.view.dispatch(
      editor.state.tr.setMeta(searchPluginKey, { ranges, activeIndex }).setMeta('uiOnly', true),
    )
  }
  return {
    get editable() {
      return editor.isEditable
    },
    search(query, opts, activeIndex) {
      ranges = findMatches(editor, query, opts)
      highlight(ranges.length === 0 ? 0 : Math.min(activeIndex, ranges.length - 1))
      return ranges.length
    },
    activate(index) {
      const range = ranges[index]
      if (!range) return
      highlight(index)
      const { node } = editor.view.domAtPos(range.from)
      const el = node instanceof HTMLElement ? node : node.parentElement
      el?.scrollIntoView({ block: 'center' })
    },
    replaceOne(index, replacement) {
      const m = ranges[index]
      if (!m) return
      editor.commands.command(({ tr }) => {
        tr.insertText(replacement, m.from, m.to)
        return true
      })
    },
    replaceAll(replacement) {
      if (ranges.length === 0) return
      const all = [...ranges].reverse()
      editor.commands.command(({ tr }) => {
        for (const m of all) tr.insertText(replacement, m.from, m.to)
        return true
      })
    },
    clear() {
      ranges = []
      highlight(0)
    },
    onDocChanged(listener) {
      const onUpdate = ({ transaction }: { transaction: { docChanged: boolean } }) => {
        if (transaction.docChanged) listener()
      }
      editor.on('update', onUpdate)
      return () => {
        editor.off('update', onUpdate)
      }
    },
  }
}
