/** The frame after a focus command scrolls the selection into view: it must not throw where there is no layout. */
import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

let editor: Editor | undefined
afterEach(() => editor?.destroy())

it('the scroll of the frame after focus runs on a live editor', async () => {
  const host = document.body.appendChild(document.createElement('div'))
  editor = new Editor({
    element: host,
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'first' }] }],
    },
  })
  const errors: unknown[] = []
  const onError = (event: ErrorEvent) => errors.push(event.error)
  window.addEventListener('error', onError)
  editor.commands.focus('end')
  await new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done)))
  window.removeEventListener('error', onError)
  expect(errors).toEqual([])
  expect(editor.state.selection.from).toBe(6)
})
