import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { installMcpBridge } from '../src/renderer/mcp-bridge'
import type { FileActionContext } from '../src/renderer/file-actions'
import type { McpCommandMessage, McpCommandResult } from '../src/shared/ipc'

/**
 * Renderer half of the MCP bridge against a real Tiptap editor: the shell
 * forwards the same commands an external agent sends, and the bridge runs them
 * through the built-in agent's executors.
 */

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (value: string): JsonNode => ({ type: 'text', text: value })
const heading = (value: string, level = 1): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(value)],
})
const para = (value: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [text(value)],
})

const liveEditors: Editor[] = []

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

function install(editor: Editor) {
  let handler: ((message: McpCommandMessage) => void) | undefined
  let notify: ((result: McpCommandResult) => void) | undefined
  window.desktop = {
    onMcpCommand: (callback: (message: McpCommandMessage) => void) => {
      handler = callback
      return () => {
        handler = undefined
      }
    },
    reportMcpResult: (result: McpCommandResult) => notify?.(result),
    signalMcpReady: () => {},
  } as unknown as typeof window.desktop
  const ctx = {
    editor,
    doc: { parsed: { blocks: [] }, isBlank: false },
  } as unknown as FileActionContext
  const uninstall = installMcpBridge({ getCtx: () => ctx })
  let seq = 0
  const send = (command: McpCommandMessage['command'], payload: unknown) => {
    const result = new Promise<McpCommandResult>((resolve) => {
      notify = resolve
    })
    handler!({ requestId: `t-${++seq}`, command, payload })
    return result
  }
  return { send, uninstall }
}

describe('MCP bridge read_blocks', () => {
  it('returns the full block HTML instead of the clipped preview', async () => {
    const long = 'A sentence that is far longer than a block-list preview. '.repeat(12).trim()
    const editor = createEditor([heading('Title'), para(long)])
    const { send, uninstall } = install(editor)
    try {
      const result = await send('read_blocks', { startBlockIndex: 1, endBlockIndex: 1 })
      expect(result.ok).toBe(true)
      expect((result.result as { text: string }).text).toBe(`<p>${long}</p>`)
    } finally {
      uninstall()
    }
  })

  it('passes the paging offset through and reassembles the range', async () => {
    const long = 'A'.repeat(30_000)
    const editor = createEditor([para(long)])
    const { send, uninstall } = install(editor)
    try {
      const first = await send('read_blocks', { startBlockIndex: 0, endBlockIndex: 0 })
      expect(first.ok).toBe(true)
      const firstText = (first.result as { text: string }).text
      expect(firstText).toContain('truncated')
      const offset = Number(/offset=(\d+)/.exec(firstText)?.[1])
      expect(offset).toBeGreaterThan(0)

      const second = await send('read_blocks', { startBlockIndex: 0, endBlockIndex: 0, offset })
      expect(second.ok).toBe(true)
      expect((second.result as { text: string }).text).toContain('end of range')
    } finally {
      uninstall()
    }
  })

  it('clears the stale-index guard after a successful read', async () => {
    const editor = createEditor([para('Original')])
    const { send, uninstall } = install(editor)
    try {
      const read = await send('read_blocks', { startBlockIndex: 0, endBlockIndex: 0 })
      expect(read.ok).toBe(true)

      // a user edit outside the agent pipeline makes the block indexes stale
      editor.view.dispatch(
        editor.state.tr.insertText(' user edit', editor.state.doc.child(0).content.size + 1),
      )
      const refused = await send('replace_blocks', {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>Rewritten</p>',
      })
      expect(refused.ok).toBe(false)
      expect(refused.error).toContain('edited by the user')

      const fresh = await send('read_blocks', { startBlockIndex: 0, endBlockIndex: 0 })
      expect(fresh.ok).toBe(true)
      expect((fresh.result as { text: string }).text).toContain('Original user edit')

      const retried = await send('replace_blocks', {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>Rewritten</p>',
      })
      expect(retried.ok).toBe(true)
      expect(editor.state.doc.child(0).textContent).toBe('Rewritten')
    } finally {
      uninstall()
    }
  })

  it('reports an invalid range as an error', async () => {
    const editor = createEditor([para('Only block')])
    const { send, uninstall } = install(editor)
    try {
      const result = await send('read_blocks', { startBlockIndex: 0, endBlockIndex: 5 })
      expect(result.ok).toBe(false)
      expect(result.error).toContain('block index invalid or out of range')
    } finally {
      uninstall()
    }
  })
})
