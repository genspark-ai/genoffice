/**
 * Protection and Read Mode switch whether the editor takes edits. That is view
 * state: it must not look like a content change to anything subscribed to
 * `update` (the dirty flag is set there), and it leaves the undo history alone.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { undoDepth } from '@tiptap/pm/history'
import { onEditableChange, setEditorEditable } from '../src/renderer/editor/editable'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('switching editability', () => {
  let editor: Editor
  let dirty: boolean

  beforeEach(() => {
    dirty = false
    editor = new Editor({
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Plain paragraph' }] }],
      },
      onUpdate: () => {
        dirty = true
      },
    })
  })
  afterEach(() => editor.destroy())

  it('emits no update, so the document stays clean', () => {
    const updates = vi.fn()
    const transactions = vi.fn()
    editor.on('update', updates)
    editor.on('transaction', transactions)
    setEditorEditable(editor, false)
    expect(editor.isEditable).toBe(false)
    setEditorEditable(editor, true)
    expect(editor.isEditable).toBe(true)
    expect(updates).not.toHaveBeenCalled()
    expect(transactions).not.toHaveBeenCalled()
    expect(dirty).toBe(false)
  })

  it('is what Tiptap itself would have reported as an update', () => {
    editor.setEditable(false)
    expect(dirty).toBe(true)
  })

  it('adds nothing to the undo history and keeps the document', () => {
    const before = editor.getJSON()
    const depth = undoDepth(editor.state)
    setEditorEditable(editor, false)
    setEditorEditable(editor, true)
    expect(undoDepth(editor.state)).toBe(depth)
    expect(editor.getJSON()).toEqual(before)
  })

  it('tells its own subscribers once per real change', () => {
    const changed = vi.fn()
    const off = onEditableChange(editor, changed)
    setEditorEditable(editor, true)
    expect(changed).not.toHaveBeenCalled()
    setEditorEditable(editor, false)
    expect(changed).toHaveBeenCalledTimes(1)
    setEditorEditable(editor, false)
    expect(changed).toHaveBeenCalledTimes(1)
    off()
    setEditorEditable(editor, true)
    expect(changed).toHaveBeenCalledTimes(1)
  })

  it('a real edit afterwards is still an update', () => {
    setEditorEditable(editor, false)
    setEditorEditable(editor, true)
    editor.commands.insertContentAt(1, 'x')
    expect(dirty).toBe(true)
  })
})
