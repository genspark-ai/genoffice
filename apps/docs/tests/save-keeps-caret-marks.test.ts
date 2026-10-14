import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { restoreMarksAfterDocRewrite } from '../src/renderer/file-actions'
import { serializeMarks } from '../src/renderer/editor/caret-marks'

/**
 * Autosave rewrites the whole document: save() re-parses the bytes it just
 * wrote and calls setContent() so the docxIndex anchors match the new file.
 * setContent() replaces the ProseMirror doc, which drops the caret's
 * storedMarks — the marks Tiptap applies to text typed *next*. A font the
 * user picked from the ribbon therefore survives only until the next autosave
 * tick (30s, or on window blur), after which new typing falls back to the
 * paragraph default. #2004.
 */

const para = (t: string) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: t }],
})

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('hello world')] },
  })
}

/** the mark set Tiptap would apply to text typed at the caret right now */
function marksAtCaret(editor: Editor): string[] {
  const { state } = editor
  const { $from } = state.selection
  return (state.storedMarks ?? $from.marks()).map((m) => m.type.name).sort()
}

/** pick Arial 12pt the way the ribbon does */
function pickFont(editor: Editor): void {
  editor
    .chain()
    .setTextSelection(1)
    .setMark('docTextStyle', { fontAscii: 'Arial', sizeHalfPoints: 24 })
    .run()
}

/** what saveOnce does around the reparse: setContent, then put the caret back */
function rewriteLikeSave(editor: Editor, caret = 1): void {
  editor.commands.setContent({ type: 'doc', content: [para('hello world')] } as never)
  editor.commands.setTextSelection(Math.min(caret, editor.state.doc.content.size))
}

describe('restoreMarksAfterDocRewrite', () => {
  it('puts back the formatting that a save would otherwise drop', () => {
    const editor = makeEditor()
    pickFont(editor)
    const snapshot = serializeMarks(editor.state.storedMarks!)
    expect(snapshot).toContain('docTextStyle')

    rewriteLikeSave(editor)
    // the marks are gone — this is the defect
    expect(marksAtCaret(editor)).not.toContain('docTextStyle')

    expect(restoreMarksAfterDocRewrite(editor, snapshot)).toBe(true)
    expect(marksAtCaret(editor)).toContain('docTextStyle')
  })

  it('restores the font the user picked, not just the mark name', () => {
    const editor = makeEditor()
    pickFont(editor)
    const snapshot = serializeMarks(editor.state.storedMarks!)

    rewriteLikeSave(editor)
    restoreMarksAfterDocRewrite(editor, snapshot)

    const style = editor.state.storedMarks!.find((m) => m.type.name === 'docTextStyle')
    expect(style!.attrs.fontAscii).toBe('Arial')
    expect(style!.attrs.sizeHalfPoints).toBe(24)
  })

  it('keeps the marks pending for the NEXT keystroke, not applied to the doc', () => {
    const editor = makeEditor()
    pickFont(editor)
    const snapshot = serializeMarks(editor.state.storedMarks!)

    rewriteLikeSave(editor)
    restoreMarksAfterDocRewrite(editor, snapshot)

    // storedMarks is the mechanism: the paragraph text itself stays unstyled,
    // so typing there still picks up Arial instead of the theme font
    const firstText = editor.state.doc.child(0).child(0)
    expect(firstText.marks.map((m) => m.type.name)).not.toContain('docTextStyle')
    expect(editor.state.storedMarks).not.toBeNull()
  })

  it('is a no-op when there was nothing to restore', () => {
    const editor = makeEditor()
    expect(restoreMarksAfterDocRewrite(editor, null)).toBe(false)
  })

  it('reports false for a snapshot that does not parse', () => {
    const editor = makeEditor()
    expect(restoreMarksAfterDocRewrite(editor, '{not json')).toBe(false)
  })

  it('reports false for a snapshot naming marks the editor does not have', () => {
    const editor = makeEditor()
    expect(
      restoreMarksAfterDocRewrite(
        editor,
        JSON.stringify([{ type: 'nonexistent-mark', attrs: {} }]),
      ),
    ).toBe(false)
  })

  it('never resurrects an annotation mark (a comment must not extend onto typing)', () => {
    const editor = makeEditor()
    // serializeMarks is the filter that keeps FORMAT_MARKS only
    const snapshot = serializeMarks(
      editor.schema.marks.docTextStyle!.create({ fontFamily: 'Arial' })
        ? [editor.schema.marks.bold!.create()]
        : [],
    )
    expect(snapshot).not.toContain('instrField')
    expect(restoreMarksAfterDocRewrite(editor, snapshot)).toBe(true)
    expect(marksAtCaret(editor)).toEqual(['bold'])
  })
})
