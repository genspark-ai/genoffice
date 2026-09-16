/**
 * Word keeps an empty paragraph's pending format on its paragraph mark: type
 * in Calibri, press Enter, arrow up and back — typing still produces Calibri.
 * ProseMirror instead keeps the pending format in transient storedMarks,
 * which any selection move clears; the caret then re-derives marks from
 * adjacent text, and an empty paragraph has none — so navigation silently
 * reset new paragraphs to the document default.
 *
 * This extension gives empty blocks a pilcrow memory via the shared
 * `caretMarks` attr: while the caret sits in an empty block *with* stored
 * marks, the marks are stamped onto the block; when the caret re-enters the
 * block bare (storedMarks === null, i.e. cleared by navigation — an explicit
 * empty array from a user toggle is respected), the stamp is restored as
 * storedMarks. Session-only: the attr is not persisted to the file.
 */
import { Extension } from '@tiptap/core'
import type { Mark, Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'

/// Only pending FORMATTING survives; structural/annotation marks (comments,
/// links, revisions) must never re-materialize from a caret memory.
/** the caret carries FORMATTING only — annotation marks (comment, link,
 *  ins/del revisions) must never extend onto new typing */
export const FORMAT_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])

export const serializeMarks = (marks: readonly Mark[]): string | null => {
  const kept = marks
    .filter((mark) => FORMAT_MARKS.has(mark.type.name))
    .map((mark) => ({ type: mark.type.name, attrs: mark.attrs }))
  return kept.length > 0 ? JSON.stringify(kept) : null
}

/** the first text node's marks in [from, to] — Word's "formatting of the
 *  START of a range" rule, shared by Enter-replace and deletion carry */
export const firstTextMarksIn = (doc: PmNode, from: number, to: number): readonly Mark[] | null => {
  let marks: readonly Mark[] | null = null
  doc.nodesBetween(from, to, (node) => {
    if (marks) return false
    if (node.isText) {
      marks = node.marks
      return false
    }
    return true
  })
  return marks
}

/** Paragraph FORMATTING that must survive replacing whole paragraphs — the
 *  paragraph counterpart of FORMAT_MARKS. A selection that consumes whole
 *  blocks makes ProseMirror fill the hole with a DEFAULT paragraph, so
 *  style-derived fonts, alignment, indents and the paragraph-mark font all
 *  silently reset to the theme (r176 follow-up). Identity and annotation
 *  attrs (docxIndex, bookmarks, comments, revisions, sdtShell...) must never
 *  be cloned onto the fresh paragraph; pageBreakBefore stays off for the
 *  r157 reason (a newline must not clone a page break). */
export const PARA_FORMAT_ATTRS = [
  'styleId',
  'align',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'snapToGrid',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'spaceBeforeAuto',
  'spaceAfterAuto',
  'contextualSpacing',
  'bidi',
  'autoSpace',
  'wordWrap',
  'overflowPunct',
  'eaLang',
  'shadingFill',
  'shadingDisplay',
  'borders',
  'borderLines',
  'borderReset',
  'tabStops',
  'emptyRunSize',
  'emptyRunFont',
] as const

/** format attrs of the first textblock in [from, to] — the paragraph
 *  counterpart of firstTextMarksIn (Word's start-of-range rule) */
export const firstParaFormatIn = (
  doc: PmNode,
  from: number,
  to: number,
): Record<string, unknown> | null => {
  let found: Record<string, unknown> | null = null
  doc.nodesBetween(from, to, (node) => {
    if (found) return false
    if (node.isTextblock) {
      const picked: Record<string, unknown> = {}
      for (const key of PARA_FORMAT_ATTRS) if (key in node.attrs) picked[key] = node.attrs[key]
      found = picked
      return false
    }
    return true
  })
  return found
}

/** true when every format attr still sits at its type default — i.e. the
 *  block is the fresh filler paragraph a whole-block replace leaves behind,
 *  not a survivor that kept real formatting (which must not be stomped) */
export const paraFormatIsDefault = (node: PmNode): boolean =>
  PARA_FORMAT_ATTRS.every((key) => {
    const spec = node.type.spec.attrs?.[key]
    return !spec || node.attrs[key] === spec.default
  })

export const CaretMarksMemory = Extension.create({
  name: 'caretMarksMemory',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('caretMarksMemory'),
        appendTransaction: (transactions, oldState, newState) => {
          const { selection, storedMarks, schema } = newState
          if (!selection.empty) return null
          const $from = selection.$from
          const block = $from.parent
          if (!block.isTextblock || block.content.size > 0) return null
          if (!('caretMarks' in block.attrs)) return null
          const blockPos = $from.before($from.depth)
          if (storedMarks !== null) {
            // Caret holds pending marks (fresh Enter split, or a ribbon
            // format on the empty block): stamp them so they survive
            // navigation. Kept out of history — this is caret bookkeeping,
            // not a user edit to undo.
            const snapshot = serializeMarks(storedMarks)
            if (snapshot === block.attrs.caretMarks) return null
            return newState.tr
              .setNodeMarkup(blockPos, undefined, { ...block.attrs, caretMarks: snapshot })
              .setStoredMarks(storedMarks)
              .setMeta('addToHistory', false)
          }
          // Bare re-entry: navigation cleared the stored marks — restore the
          // block's pilcrow memory.
          const stamped = block.attrs.caretMarks as string | null
          if (!stamped) {
            // A deletion just emptied this block (Delete/Backspace/Cut over a
            // selection): Word keeps the formatting of the START of the
            // removed text on the pilcrow, so typing or pasting here must not
            // fall back to the theme font (r172; the typing-over and
            // Enter-over paths carry marks already — plain deletion did not)
            if (!transactions.some((tr) => tr.docChanged) || oldState.selection.empty) return null
            const { from, to } = oldState.selection
            const kept = (firstTextMarksIn(oldState.doc, from, to) ?? []).filter((mark) =>
              FORMAT_MARKS.has(mark.type.name),
            )
            // A selection that consumed whole paragraphs left a DEFAULT filler
            // block: its paragraph formatting (style, paragraph-mark font,
            // alignment...) must come back from the removed range's first
            // paragraph, or style-derived fonts reset to the theme (r176)
            const paraFormat = paraFormatIsDefault(block)
              ? firstParaFormatIn(oldState.doc, from, to)
              : null
            const restoresPara =
              paraFormat !== null &&
              PARA_FORMAT_ATTRS.some(
                (key) => key in block.attrs && paraFormat[key] !== block.attrs[key],
              )
            if (kept.length === 0 && !restoresPara) return null
            const tr = newState.tr
              .setNodeMarkup(blockPos, undefined, {
                ...block.attrs,
                ...(restoresPara ? paraFormat : null),
                caretMarks: serializeMarks(kept),
              })
              .setMeta('addToHistory', false)
            // an explicit empty stored-marks list means "user cleared" — only
            // set marks that actually carry something
            if (kept.length > 0) tr.setStoredMarks(kept)
            return tr
          }
          try {
            const parsed = JSON.parse(stamped) as { type: string; attrs: Record<string, unknown> }[]
            const marks = parsed.flatMap((entry) => {
              const type = schema.marks[entry.type]
              return type && FORMAT_MARKS.has(entry.type) ? [type.create(entry.attrs)] : []
            })
            if (marks.length === 0) return null
            return newState.tr.setStoredMarks(marks)
          } catch {
            return null
          }
        },
      }),
    ]
  },
})
