import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { StyleInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { effectiveSizeHalfPoints } from '../src/renderer/editor/text-style-resolve'

const editors: Editor[] = []

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.length = 0
})

function createEditor(options?: {
  paragraphStyleId?: string
  characterStyleId?: string
  directSize?: number
}): Editor {
  const marks =
    options?.characterStyleId || options?.directSize
      ? [
          {
            type: 'docTextStyle',
            attrs: {
              styleId: options.characterStyleId ?? null,
              sizeHalfPoints: options.directSize ?? null,
            },
          },
        ]
      : undefined
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null, styleId: options?.paragraphStyleId ?? null },
          content: [{ type: 'text', text: 'Styled text', marks }],
        },
      ],
    },
  })
  editor.commands.setTextSelection(2)
  editors.push(editor)
  return editor
}

function styles(...entries: Array<[string, number]>): Map<string, StyleInfo> {
  return new Map(
    entries.map(([styleId, sizeHalfPoints]) => [
      styleId,
      {
        styleId,
        name: styleId,
        type: styleId.startsWith('Char') ? 'character' : 'paragraph',
        display: { sizeHalfPoints },
      },
    ]),
  )
}

describe('effectiveSizeHalfPoints', () => {
  it('resolves the paragraph style size at the caret', () => {
    const editor = createEditor({ paragraphStyleId: 'Title' })
    expect(effectiveSizeHalfPoints(editor, styles(['Title', 56]))).toBe(56)
  })

  it('resolves a character style before the paragraph style', () => {
    const editor = createEditor({
      paragraphStyleId: 'Title',
      characterStyleId: 'CharEmphasis',
    })
    expect(effectiveSizeHalfPoints(editor, styles(['Title', 56], ['CharEmphasis', 28]))).toBe(28)
  })

  it('prefers direct formatting over styles', () => {
    const editor = createEditor({ paragraphStyleId: 'Title', directSize: 24 })
    expect(effectiveSizeHalfPoints(editor, styles(['Title', 56]))).toBe(24)
  })

  it('falls back to document defaults', () => {
    const editor = createEditor()
    expect(effectiveSizeHalfPoints(editor, undefined, { sizeHalfPoints: 21 })).toBe(21)
  })
})
