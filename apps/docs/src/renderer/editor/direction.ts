import type { Editor } from '@tiptap/core'
import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Mapping } from '@tiptap/pm/transform'

/** paragraph-like blocks that carry the bidi attribute */
const DIR_BLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])

/** strong RTL (UAX#9 R/AL) scripts: Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan, Mandaic */
const RTL_CHAR =
  /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}]/u
const LETTER = /\p{L}/u

/**
 * Direction of the first strong directional character (dir="auto" semantics),
 * or null if none. Only letters decide: digits, combining marks, format
 * controls and punctuation (including script-Arabic marks like the Arabic
 * comma, which are bidi-neutral despite their script) are all weak.
 */
export function firstStrongDir(text: string): 'ltr' | 'rtl' | null {
  for (const ch of text) {
    if (!LETTER.test(ch)) continue
    return RTL_CHAR.test(ch) ? 'rtl' : 'ltr'
  }
  return null
}

/** bidi attr of the paragraph-like node at the cursor (false in textbox sub-editors, which have no bidi) */
export function activeBidi(editor: Editor): boolean {
  for (const name of ['docHeading', 'docListItem', 'docParagraph']) {
    if (editor.isActive(name)) return editor.getAttributes(name).bidi === true
  }
  return false
}

/**
 * The align attr the "align left/right" ribbon buttons, menu items and ⌘L/⌘R
 * should write. align stores the visual value and null means "start", so in an
 * RTL paragraph null already renders right: aligning to the paragraph's start
 * side clears the attr, aligning to the end side sets it explicitly.
 */
export function alignAttrFor(
  align: 'left' | 'center' | 'right' | 'justify',
  bidi: boolean,
): string | null {
  if (align === 'left') return bidi ? 'left' : null
  if (align === 'right') return bidi ? null : 'right'
  return align
}

/**
 * Attrs for flipping a paragraph's direction. Explicit left/right alignment
 * swaps: align stores the visual value, so the swap preserves the logical
 * (start/end) alignment, matching Word's behavior when toggling direction.
 */
function dirFlipAttrs(attrs: Record<string, unknown>, bidi: boolean): Record<string, unknown> {
  let align = attrs.align as string | null
  if (align === 'left') align = 'right'
  else if (align === 'right') align = 'left'
  return { ...attrs, bidi, align }
}

/**
 * Apply alignment to every paragraph-like block in the selection, resolving
 * the stored attr per paragraph against its own direction (a mixed LTR/RTL
 * selection must not inherit the cursor paragraph's resolution). Selected
 * images keep receiving the visual value via imageAlign, as before.
 */
export function setSelectionAlign(
  editor: Editor,
  align: 'left' | 'center' | 'right' | 'justify',
): boolean {
  return editor
    .chain()
    .focus()
    .command(({ state, tr, dispatch }) => {
      const { from, to } = state.selection
      let changed = false
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (DIR_BLOCKS.has(node.type.name)) {
          const value = alignAttrFor(align, node.attrs.bidi === true)
          tr.setNodeMarkup(pos, undefined, { ...node.attrs, align: value })
          changed = true
        } else if (node.type.name === 'docProtected') {
          // alignment also applies to selected images (w:jc on the image paragraph)
          tr.setNodeMarkup(pos, undefined, {
            ...node.attrs,
            imageAlign: alignAttrFor(align, false),
          })
          changed = true
        }
      })
      if (changed && dispatch) dispatch(tr)
      return changed
    })
    .run()
}

/** Set the writing direction of every paragraph-like block in the selection. */
export function setParagraphDirection(editor: Editor, dir: 'ltr' | 'rtl'): boolean {
  const bidi = dir === 'rtl'
  return editor
    .chain()
    .focus()
    .command(({ state, tr, dispatch }) => {
      const { from, to } = state.selection
      let changed = false
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!DIR_BLOCKS.has(node.type.name)) return
        if (node.attrs.bidi === undefined || node.attrs.bidi === bidi) return
        tr.setNodeMarkup(pos, undefined, dirFlipAttrs(node.attrs, bidi))
        changed = true
      })
      if (changed && dispatch) dispatch(tr)
      return changed
    })
    .run()
}

/**
 * Auto-detect paragraph direction while typing (Word's "detect language"):
 * when a text insertion gives a paragraph its first strong directional
 * character, the paragraph's bidi flag follows that character. Paragraphs
 * that already contained strong text are never touched, so existing content
 * and explicit direction choices survive edits.
 */
export const AutoDirectionExtension = Extension.create({
  name: 'autoDirection',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key: new PluginKey('autoDirection'),
        appendTransaction(transactions, oldState, newState) {
          if (!transactions.some((tr) => tr.docChanged)) return null
          // rewriting the paragraph DOM mid-composition would break the IME
          if (editor.view?.composing) return null

          // only plain text insertions (typing, text paste) qualify; block-level
          // replacements (document load, Enter splits, AI edits) are skipped
          let textInserted = false
          for (const tr of transactions) {
            for (const step of tr.steps) {
              const slice = (step as { slice?: { content?: import('@tiptap/pm/model').Fragment } })
                .slice
              if (!slice?.content || slice.content.childCount === 0) continue
              let onlyText = true
              slice.content.forEach((child) => {
                if (!child.isText) onlyText = false
              })
              if (onlyText) textInserted = true
              else return null
            }
          }
          if (!textInserted) return null

          const $head = newState.selection.$head
          let depth = 0
          for (let d = $head.depth; d > 0; d -= 1) {
            if (DIR_BLOCKS.has($head.node(d).type.name)) {
              depth = d
              break
            }
          }
          if (depth === 0) return null
          const para = $head.node(depth)
          if (para.attrs.bidi === undefined) return null
          const dir = firstStrongDir(para.textContent)
          if (!dir) return null
          const bidi = dir === 'rtl'
          if (para.attrs.bidi === bidi) return null

          // the flip only happens on the paragraph's first strong character:
          // if the pre-change paragraph already had one, leave it alone
          const mapping = new Mapping()
          for (const tr of transactions) mapping.appendMapping(tr.mapping)
          const oldPos = Math.min(
            mapping.invert().map(newState.selection.head),
            oldState.doc.content.size,
          )
          const $old = oldState.doc.resolve(oldPos)
          for (let d = $old.depth; d > 0; d -= 1) {
            if (!DIR_BLOCKS.has($old.node(d).type.name)) continue
            if (firstStrongDir($old.node(d).textContent)) return null
            break
          }

          // same attrs as the manual toggle, so auto-detect and ribbon
          // round-trip to the user's original alignment
          return newState.tr.setNodeMarkup(
            $head.before(depth),
            undefined,
            dirFlipAttrs(para.attrs, bidi),
          )
        },
      }),
    ]
  },
})
