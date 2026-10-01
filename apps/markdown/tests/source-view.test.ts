import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { parseDocText, serializeDocText } from '../src/renderer/markdown/docText'
import { spliceMarkdown } from '../src/renderer/markdown/sourceSplice'
import { applySourceText, type AppliedSource } from '../src/renderer/markdown/sourceView'

/**
 * Source view: the pane shows the exact file text, and a keystroke there is
 * re-parsed into the editor rather than written straight to disk. These cover
 * the two properties that makes safe — the view cannot misreport the file, and
 * the round-trip splice still protects markup the schema has no node for.
 */

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

/** load without a history step, the way a file open does */
function load(editor: Editor, markdown: string): void {
  editor
    .chain()
    .setMeta('addToHistory', false)
    .setContent(markdown, { contentType: 'markdown' })
    .run()
}

/** what a save would write for the current state, mirroring App.tsx's doSave */
function savedText(editor: Editor, applied: AppliedSource): string {
  expect(applied.sourceMap).not.toBeNull()
  return serializeDocText(
    applied.envelope,
    spliceMarkdown(editor, editor.state.doc, applied.sourceMap!),
  )
}

describe('source view', () => {
  it('re-parses the envelope so EOL, BOM and EOF newline survive a source edit', () => {
    const editor = createEditor()
    load(editor, '# Hi')
    const applied = applySourceText(editor, '﻿---\r\ntitle: t\r\n---\r\n\r\n# CRLF doc')
    expect(applied.envelope.bom).toBe(true)
    expect(applied.envelope.eol).toBe('\r\n')
    expect(applied.envelope.frontmatter).toBe('---\ntitle: t\n---\n\n')
    expect(applied.envelope.trailingNewline).toBe(false)
  })

  it('leaves the saved text byte-identical when the source is handed back unchanged', () => {
    const editor = createEditor()
    const raw = '---\ntitle: keep\n---\n\n# Head\n\nBody text.\n'
    load(editor, parseDocText(raw).body)
    const applied = applySourceText(editor, raw)
    expect(savedText(editor, applied).replace(/\r\n/g, '\n')).toBe(raw)
  })

  it('keeps markup the editor has no node for, via the round-trip splice', () => {
    const editor = createEditor()
    const raw = '# Head\n\nA footnote the editor cannot model:\n\n[^1]: the note body\n'
    load(editor, parseDocText(raw).body)
    const applied = applySourceText(editor, raw)
    expect(savedText(editor, applied)).toContain('[^1]: the note body')
  })

  it('puts an edited construct through the editor, not just the text', () => {
    const editor = createEditor()
    load(editor, '# Old')
    applySourceText(editor, '# New')
    expect(editor.getText().trim()).toBe('New')
  })

  it('strips legacy fenced divs the way the load path does', () => {
    const editor = createEditor()
    load(editor, 'plain')
    applySourceText(editor, ':::callout {type="note"}\ninside\n:::\n')
    expect(editor.getText()).not.toContain(':::')
    expect(editor.getText()).toContain('inside')
  })

  it('leaves no history step for a source keystroke', () => {
    const editor = createEditor()
    load(editor, '# A')
    applySourceText(editor, '# B')
    // the pane's own native undo covers the pane; one undo must not rewind it
    editor.commands.undo()
    expect(editor.getText().trim()).toBe('B')
  })
})
