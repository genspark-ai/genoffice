import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  droppedImageMime,
  imageFilesFromDataTransfer,
  insertImageFilesAt,
  isImageFileDrag,
  nearestTextSelection,
} from '../src/renderer/editor/image-drop'

const PNG_BYTES = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0])

function fakeTransfer(files: File[]): DataTransfer {
  return {
    files,
    items: files.map((f) => ({ kind: 'file', type: f.type })),
  } as unknown as DataTransfer
}

/** paragraph, TOC-like atom, paragraph */
function openDoc(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'Hello' }] },
        { type: 'docProtected', attrs: { blockType: 'toc', label: 'TOC' } },
        { type: 'docParagraph', content: [{ type: 'text', text: 'World' }] },
      ],
    },
  })
}

describe('image file drops', () => {
  it('accepts only the picture formats the insert dialog accepts', () => {
    const png = new File([PNG_BYTES], 'a.png', { type: 'image/png' })
    const jpg = new File([PNG_BYTES], 'b.jpg', { type: 'image/jpeg' })
    const webp = new File([PNG_BYTES], 'c.webp', { type: 'image/webp' })
    const pdf = new File([PNG_BYTES], 'd.pdf', { type: 'application/pdf' })
    const untyped = new File([PNG_BYTES], 'e.GIF', { type: '' })
    const untypedText = new File([PNG_BYTES], 'f.txt', { type: '' })
    expect(
      imageFilesFromDataTransfer(fakeTransfer([png, jpg, webp, pdf, untyped, untypedText])),
    ).toEqual([png, jpg, untyped])
    expect(droppedImageMime(untyped)).toBe('image/gif')
    expect(imageFilesFromDataTransfer(null)).toEqual([])
  })

  it('recognises an image file drag from the item types alone', () => {
    const dt = {
      items: [{ kind: 'file', type: 'image/png' }],
      files: [],
    } as unknown as DataTransfer
    expect(isImageFileDrag(dt)).toBe(true)
    expect(
      isImageFileDrag(fakeTransfer([new File([], 'x.pdf', { type: 'application/pdf' })])),
    ).toBe(false)
    expect(
      isImageFileDrag({
        items: [{ kind: 'string', type: 'image/png' }],
      } as unknown as DataTransfer),
    ).toBe(false)
  })

  it('resolves a drop inside a read-only block to the nearest text position', () => {
    const editor = openDoc()
    const doc = editor.state.doc
    const tocPos = doc.child(0).nodeSize
    const sel = nearestTextSelection(doc, tocPos + 1)
    expect(sel).not.toBeNull()
    expect(sel!.$from.parent.type.name).toBe('docParagraph')
    expect(nearestTextSelection(doc, 99999)?.$from.parent.textContent).toBe('World')
    editor.destroy()
  })

  describe('insertion', () => {
    class FakeImage {
      onload: (() => void) | null = null
      onerror: (() => void) | null = null
      naturalWidth = 120
      naturalHeight = 60
      set src(_v: string) {
        queueMicrotask(() => this.onload?.())
      }
    }
    let savedImage: typeof Image
    beforeEach(() => {
      savedImage = globalThis.Image
      globalThis.Image = FakeImage as unknown as typeof Image
    })
    afterEach(() => {
      globalThis.Image = savedImage
    })

    it('inserts pictures at the drop point and leaves the caret after them', async () => {
      const editor = openDoc()
      const files = [
        new File([PNG_BYTES], 'one.png', { type: 'image/png' }),
        new File([PNG_BYTES], 'two.png', { type: 'image/png' }),
      ]
      const inserted = await insertImageFilesAt(editor, files, 3)
      expect(inserted).toBe(2)
      const names = editor.state.doc.content.content.map((n) =>
        n.type.name === 'docProtected' ? `${n.type.name}:${n.attrs.blockType}` : n.type.name,
      )
      expect(names).toEqual([
        'docParagraph',
        'docProtected:image',
        'docProtected:image',
        'docParagraph',
        'docProtected:toc',
        'docParagraph',
      ])
      const image = editor.state.doc.child(1)
      expect(image.attrs.imageDataUrl).toMatch(/^data:image\/png;base64,/)
      expect(image.attrs.imageWidthPx).toBe(120)
      expect(image.attrs.label).toContain('one.png')
      expect(editor.state.doc.child(0).textContent).toBe('He')
      expect(editor.state.selection.$from.parent.textContent).toBe('llo')
      editor.destroy()
    })

    it('skips files whose type is not a supported picture', async () => {
      const editor = openDoc()
      const inserted = await insertImageFilesAt(
        editor,
        [new File([PNG_BYTES], 'x.webp', { type: 'image/webp' })],
        1,
      )
      expect(inserted).toBe(0)
      expect(editor.state.doc.childCount).toBe(3)
      editor.destroy()
    })
  })
})
