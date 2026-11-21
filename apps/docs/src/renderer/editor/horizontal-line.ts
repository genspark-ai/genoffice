import type { Editor } from '@tiptap/core'

/** the line Word's Borders and Shading > Horizontal Line writes: a full-width 1.5pt grey VML rule */
const HORIZONTAL_LINE_XML =
  '<w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0" stroked="f"/></w:pict>'

/** A horizontal line in a paragraph of its own: the caret's when it is empty, else a new one below. */
export function insertHorizontalLine(editor: Editor): boolean {
  const { state } = editor
  const { docInlineImage, docParagraph } = state.schema.nodes
  if (!docInlineImage || !docParagraph) return false
  const rule = docInlineImage.create({
    dataUrl: '',
    xml: HORIZONTAL_LINE_XML,
    rule: { colorHex: 'A0A0A0', thicknessPx: 2 },
  })
  const { $from } = state.selection
  const here = $from.parent
  const tr = state.tr
  if (here.type === docParagraph && here.content.size === 0) {
    tr.insert($from.pos, rule)
  } else {
    if ($from.depth < 1) return false
    const after = $from.after()
    const $after = state.doc.resolve(after)
    if (!$after.parent.canReplaceWith($after.index(), $after.index(), docParagraph)) return false
    tr.insert(after, docParagraph.create(null, rule))
  }
  editor.view.dispatch(tr.scrollIntoView())
  editor.commands.focus()
  return true
}
