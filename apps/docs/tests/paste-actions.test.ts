/**
 * UI-started pastes (ribbon split button, Paste Special, context menu) go
 * through the same editor paste pipeline as Ctrl+V and arm the same
 * paste-mode handshake the post-paste chip uses.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  consumeForeignPaste,
  setDefaultPasteMode,
  takePastePayload,
} from '../src/renderer/editor/paste-options'

const insertImageFromDataUrl = vi.fn(async (..._args: unknown[]) => true)
vi.mock('../src/renderer/components/ribbon-tabs', () => ({
  insertImageFromDataUrl: (...args: unknown[]) => insertImageFromDataUrl(...args),
}))

class FakeDataTransfer {
  private data = new Map<string, string>()
  setData(type: string, value: string) {
    this.data.set(type, value)
  }
  getData(type: string) {
    return this.data.get(type) ?? ''
  }
  get types() {
    return [...this.data.keys()]
  }
}
class FakeClipboardEvent extends Event {
  clipboardData: FakeDataTransfer | null
  constructor(type: string, init?: { clipboardData?: FakeDataTransfer }) {
    super(type)
    this.clipboardData = init?.clipboardData ?? null
  }
}
const g = globalThis as Record<string, unknown>
g.DataTransfer = FakeDataTransfer
g.ClipboardEvent = FakeClipboardEvent

const { pastePayload, pasteFromClipboard, readClipboardPayload } =
  await import('../src/renderer/editor/paste-actions')

const FOREIGN = `<meta charset='utf-8'><p style="font-family: Arial;">web <b>text</b></p>`
const INTERNAL = `<div data-pm-slice="1 1 []"><p>ours</p></div>`

// the document editor installs transformPasted (App); sub-editors do not
function makeEditor(consumesModes = true) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    editorProps: consumesModes ? { transformPasted: (slice) => slice } : {},
  })
  const pasteHTML = vi.spyOn(editor.view, 'pasteHTML').mockReturnValue(true)
  const pasteText = vi.spyOn(editor.view, 'pasteText').mockReturnValue(true)
  return { editor, pasteHTML, pasteText }
}

function clipboardItem(flavors: Record<string, Blob>) {
  return {
    types: Object.keys(flavors),
    getType: async (type: string) => flavors[type]!,
  }
}

function mockClipboard(clipboard: {
  read?: () => Promise<unknown>
  readText?: () => Promise<string>
}) {
  Object.defineProperty(navigator, 'clipboard', { value: clipboard, configurable: true })
}

beforeEach(() => {
  localStorage.clear()
  consumeForeignPaste()
  takePastePayload()
  insertImageFromDataUrl.mockClear()
})

describe('pastePayload', () => {
  it('default paste of foreign HTML arms the default mode and stashes the chip payload', () => {
    setDefaultPasteMode('merge')
    const { editor, pasteHTML, pasteText } = makeEditor()
    expect(pastePayload(editor, { html: FOREIGN, text: 'web text', image: null })).toBe(true)
    expect(pasteHTML).toHaveBeenCalledTimes(1)
    expect(pasteText).not.toHaveBeenCalled()
    const event = pasteHTML.mock.calls[0]![1] as unknown as FakeClipboardEvent
    expect(event.clipboardData?.getData('text/plain')).toBe('web text')
    expect(consumeForeignPaste()).toBe('merge')
    expect(takePastePayload()).toMatchObject({ html: FOREIGN, text: 'web text', mode: 'merge' })
    editor.destroy()
  })

  it('an explicit mode overrides the default for that paste only', () => {
    const { editor, pasteHTML } = makeEditor()
    expect(pastePayload(editor, { html: FOREIGN, text: 'web text', image: null }, 'text')).toBe(
      true,
    )
    expect(pasteHTML).toHaveBeenCalledTimes(1)
    expect(consumeForeignPaste()).toBe('text')
    expect(takePastePayload()?.mode).toBe('text')
    // the override is spent
    pastePayload(editor, { html: FOREIGN, text: 'web text', image: null })
    expect(consumeForeignPaste()).toBe('source')
    editor.destroy()
  })

  it('our own HTML keeps its runs, except Keep Text Only which pastes the text', () => {
    const { editor, pasteHTML, pasteText } = makeEditor()
    expect(pastePayload(editor, { html: INTERNAL, text: 'ours', image: null }, 'merge')).toBe(true)
    expect(pasteHTML).toHaveBeenCalledTimes(1)
    expect(consumeForeignPaste()).toBeNull()
    expect(takePastePayload()).toBeNull()
    expect(pastePayload(editor, { html: INTERNAL, text: '', image: null }, 'text')).toBe(true)
    expect(pasteText).toHaveBeenCalledWith('ours', expect.anything())
    editor.destroy()
  })

  it('sub-editors without the paste lanes never arm the handshake', () => {
    setDefaultPasteMode('text')
    const { editor, pasteHTML, pasteText } = makeEditor(false)
    expect(pastePayload(editor, { html: FOREIGN, text: 'web text', image: null })).toBe(true)
    expect(pasteText).toHaveBeenCalledWith('web text', expect.anything())
    expect(pastePayload(editor, { html: FOREIGN, text: 'web text', image: null }, 'merge')).toBe(
      true,
    )
    expect(pasteHTML).toHaveBeenCalledTimes(1)
    // the default mode is for other programs: an in-document copy keeps its runs
    expect(pastePayload(editor, { html: INTERNAL, text: 'ours', image: null })).toBe(true)
    expect(pasteHTML).toHaveBeenCalledTimes(2)
    expect(consumeForeignPaste()).toBeNull()
    expect(takePastePayload()).toBeNull()
    const main = makeEditor()
    pastePayload(main.editor, { html: FOREIGN, text: 'web text', image: null })
    expect(consumeForeignPaste()).toBe('text')
    main.editor.destroy()
    editor.destroy()
  })

  it('plain text goes down the text lane; a lone bitmap becomes a picture', async () => {
    const { editor, pasteHTML, pasteText } = makeEditor()
    expect(pastePayload(editor, { html: '', text: 'plain', image: null })).toBe(true)
    expect(pasteText).toHaveBeenCalledWith('plain', expect.anything())
    const image = new Blob(['x'], { type: 'image/png' })
    expect(pastePayload(editor, { html: '', text: ' ', image })).toBe(true)
    await vi.waitFor(() => expect(insertImageFromDataUrl).toHaveBeenCalledTimes(1))
    expect(pasteHTML).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('Paste Special choices: picture needs a bitmap, text needs text', () => {
    const { editor, pasteText } = makeEditor()
    expect(pastePayload(editor, { html: '', text: 'plain', image: null }, 'picture')).toBe(false)
    expect(pastePayload(editor, { html: '', text: '', image: null }, 'text')).toBe(false)
    const image = new Blob(['x'], { type: 'image/png' })
    expect(pastePayload(editor, { html: '', text: 'alt', image }, 'text')).toBe(true)
    expect(pasteText).toHaveBeenCalledWith('alt', expect.anything())
    expect(insertImageFromDataUrl).not.toHaveBeenCalled()
    editor.destroy()
  })
})

describe('readClipboardPayload', () => {
  it('collects html, text and the first bitmap across items', async () => {
    const image = new Blob(['x'], { type: 'image/png' })
    mockClipboard({
      read: async () => [
        clipboardItem({ 'text/plain': new Blob(['t']), 'text/html': new Blob([FOREIGN]) }),
        clipboardItem({ 'image/png': image }),
      ],
    })
    expect(await readClipboardPayload()).toEqual({ html: FOREIGN, text: 't', image })
  })

  it('falls back to readText when read() is unavailable', async () => {
    mockClipboard({
      read: async () => {
        throw new Error('denied')
      },
      readText: async () => 'fallback',
    })
    expect(await readClipboardPayload()).toEqual({ html: '', text: 'fallback', image: null })
    const { editor, pasteText } = makeEditor()
    expect(await pasteFromClipboard(editor, 'text')).toBe(true)
    expect(pasteText).toHaveBeenCalledWith('fallback', expect.anything())
    editor.destroy()
  })
})
