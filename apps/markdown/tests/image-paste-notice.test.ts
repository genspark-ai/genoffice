import { afterEach, describe, expect, it, vi } from 'vitest'
import { Slice } from '@tiptap/pm/model'
import { setToastEmitter, type ToastData } from '../src/renderer/components/toast-bus'
import { strings } from '../src/renderer/i18n/strings'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error) — destroy every editor we create.
const editors: import('@tiptap/core').Editor[] = []
let toasts: ToastData[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
  setToastEmitter(null)
  toasts = []
  vi.restoreAllMocks()
})

function stubSaveImage(result: string | null) {
  Object.defineProperty(window, 'markdownApi', {
    configurable: true,
    value: { saveImage: vi.fn(async () => result) },
  })
}

async function newEditor() {
  const { Editor } = await import('@tiptap/core')
  const { buildExtensions } = await import('../src/renderer/editor/extensions')
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

/** Invoke the editor's handlePaste props with a synthetic clipboard payload.
 * Mirrors ProseMirror semantics: first handler returning truthy wins.
 * getData is stubbed because CodeBlock's paste handler reads text/html. */
function paste(editor: import('@tiptap/core').Editor, files: File[]): boolean {
  const clipboardData = {
    files,
    getData: () => '',
  } as unknown as DataTransfer
  const event = { clipboardData } as unknown as ClipboardEvent
  let handled = false
  editor.view.someProp('handlePaste', (fn) => {
    if (handled) return
    handled = fn(editor.view, event, Slice.empty) === true
  })
  return handled
}

function pngFile(): File {
  // 4-byte payload is enough — persistAndInsert only reads bytes + MIME type
  return new File([new Uint8Array([137, 80, 78, 71])], 'shot.png', { type: 'image/png' })
}

describe('image paste into an untitled document', () => {
  it('toasts instead of silently dropping when the image cannot be persisted', async () => {
    stubSaveImage(null)
    setToastEmitter((toast) => toasts.push(toast))
    const editor = await newEditor()

    const handled = paste(editor, [pngFile()])
    expect(handled).toBe(true)

    await vi.waitFor(() => expect(toasts).toHaveLength(1))
    expect(toasts[0]).toEqual({ text: strings.zh.imageNeedsSavedDocument, kind: 'error' })
    expect(editor.state.doc.toString()).not.toContain('image')
    expect(window.markdownApi.saveImage).toHaveBeenCalledOnce()
  })

  it('inserts the image node when persistence succeeds, without a toast', async () => {
    stubSaveImage('assets/shot.png')
    setToastEmitter((toast) => toasts.push(toast))
    const editor = await newEditor()

    paste(editor, [pngFile()])

    await vi.waitFor(() => {
      let found: string | null = null
      editor.state.doc.descendants((node) => {
        if (node.type.name === 'image') found = String(node.attrs.src)
        return true
      })
      expect(found).toBe('assets/shot.png')
    })
    expect(toasts).toHaveLength(0)
  })

  it('ignores clipboards without image files', async () => {
    stubSaveImage('assets/shot.png')
    const editor = await newEditor()
    const text = new File(['plain'], 'notes.txt', { type: 'text/plain' })
    expect(paste(editor, [text])).toBe(false)
    expect(window.markdownApi.saveImage).not.toHaveBeenCalled()
  })
})
