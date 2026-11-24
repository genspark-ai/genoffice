/** The editor's focus command: the page is focused before any further input, and nothing done since is undone. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions, textboxSubExtensions } from '../src/renderer/editor/extensions'
import { hfSubExtensions } from '../src/renderer/editor/hf-editor'
import { PromptFocusExtension } from '../src/renderer/editor/prompt-focus'

const microtasks = () => new Promise((done) => setTimeout(done, 0))

describe('focus command', () => {
  let editor: Editor
  let host: HTMLElement
  let other: HTMLInputElement
  let frames: FrameRequestCallback[]
  const runFrames = () => {
    for (const frame of frames.splice(0)) frame(0)
  }

  beforeEach(() => {
    frames = []
    vi.spyOn(window, 'requestAnimationFrame').mockImplementation((cb) => frames.push(cb))
    host = document.body.appendChild(document.createElement('div'))
    other = document.body.appendChild(document.createElement('input'))
    editor = new Editor({
      element: host,
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'first' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'second' }] },
        ],
      },
    })
  })
  afterEach(() => {
    editor.destroy()
    host.remove()
    other.remove()
    vi.restoreAllMocks()
  })

  it('is the one every Docs editor runs', () => {
    for (const list of [editorExtensions, textboxSubExtensions, hfSubExtensions])
      expect(list).toContain(PromptFocusExtension)
  })

  it('focuses the page before a frame has run', async () => {
    expect(editor.view.hasFocus()).toBe(false)
    editor.commands.focus()
    await microtasks()
    expect(editor.view.hasFocus()).toBe(true)
  })

  it('sets the caret where asked', async () => {
    editor.chain().focus('end').run()
    await microtasks()
    expect(editor.view.hasFocus()).toBe(true)
    expect(editor.state.selection.$from.parent.textContent).toBe('second')
  })

  it('leaves the focus with a control focused since the command ran', async () => {
    editor.commands.focus()
    other.focus()
    await microtasks()
    runFrames()
    expect(document.activeElement).toBe(other)
  })

  it('takes the focus from the control the command was run from', async () => {
    other.focus()
    editor.commands.focus()
    await microtasks()
    expect(editor.view.hasFocus()).toBe(true)
  })

  it.each(['keydown', 'pointerdown', 'wheel'])(
    'writes nothing to the page in the frame after a %s',
    async (type) => {
      editor.commands.focus()
      await microtasks()
      ;(editor.view.dom as HTMLElement).blur()
      const focus = vi.spyOn(editor.view, 'focus')
      const dispatch = vi.spyOn(editor.view, 'dispatch')
      window.dispatchEvent(new Event(type))
      runFrames()
      expect(focus).not.toHaveBeenCalled()
      expect(dispatch).not.toHaveBeenCalled()
    },
  )

  it('waits for a key that is going down to run its course: Enter in a dialog is not typed into the page', async () => {
    let focusedInTime = true
    other.addEventListener('keydown', () => {
      editor.commands.focus()
      queueMicrotask(() => (focusedInTime = editor.view.hasFocus()))
    })
    other.focus()
    other.dispatchEvent(new Event('keydown', { bubbles: true }))
    await microtasks()
    expect(focusedInTime).toBe(false)
    expect(editor.view.hasFocus()).toBe(true)
  })

  it('takes the focus from a button that took it as it was pressed', async () => {
    const button = document.body.appendChild(document.createElement('button'))
    button.addEventListener('mousedown', () => {
      editor.commands.focus()
      button.focus()
    })
    button.dispatchEvent(new Event('mousedown', { bubbles: true }))
    await microtasks()
    button.remove()
    expect(editor.view.hasFocus()).toBe(true)
  })

  it('run as the pointer goes down, is not given up for the mousedown of the same press', async () => {
    const button = document.body.appendChild(document.createElement('button'))
    button.addEventListener('pointerdown', () => editor.commands.focus())
    button.dispatchEvent(new Event('pointerdown', { bubbles: true }))
    button.dispatchEvent(new Event('mousedown', { bubbles: true }))
    await microtasks()
    button.remove()
    expect(editor.view.hasFocus()).toBe(true)
  })

  it('takes the focus back in the frame when a control that held it has gone', async () => {
    editor.commands.focus(null, { scrollIntoView: false })
    await microtasks()
    other.focus()
    other.blur()
    expect(editor.view.hasFocus()).toBe(false)
    runFrames()
    expect(editor.view.hasFocus()).toBe(true)
  })

  it('does not focus when the command is only tested', async () => {
    expect(editor.can().focus()).toBe(true)
    await microtasks()
    runFrames()
    expect(editor.view.hasFocus()).toBe(false)
  })
})
