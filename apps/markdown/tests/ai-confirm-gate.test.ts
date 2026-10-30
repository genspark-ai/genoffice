// The markdown panel's approval gate: risk:'high' write tools (apply_ops,
// write_document, insert_image, generate_image) park at an inline allow/deny
// card before they execute. Denying must leave the document untouched and hand
// the model a do-not-retry result; allowing runs the tool. Driven end-to-end
// through the stubbed window.markdownApi stream, like ai-panel-prefs.test.ts.
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import type { AiStreamChunk, AiStreamRequest, AiSettings } from '@genoffice/ai-provider'
import { TOOL_DENIED_OUTPUT } from '@genoffice/agent-core'
import { AiPanel, type MarkdownAiDeps } from '../src/renderer/ai/AiPanel'
import { buildExtensions } from '../src/renderer/editor/extensions'
import type { MarkdownApi } from '../src/shared/ipc'

type StreamHandler = (chunk: AiStreamChunk) => void

const roots: Root[] = []
const editors: Editor[] = []

beforeAll(() => {
  ;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  Element.prototype.scrollTo = () => {}
})

afterAll(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

afterEach(() => {
  for (const root of roots.splice(0)) act(() => root.unmount())
  document.body.innerHTML = ''
  delete (window as unknown as { markdownApi?: MarkdownApi }).markdownApi
})

function createEditor(md: string): Editor {
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
  editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

/**
 * The transport answers turn 1 with an apply_ops insert call, turn 2 with
 * text. Records every request so the assertions can see what the model
 * received.
 */
function installStreamStub() {
  const requests: AiStreamRequest[] = []
  let listener: StreamHandler | null = null
  const api = {
    aiGskStatus: async () => ({ loggedIn: false }),
    getAiSettings: async () => ({ providers: [] }) as unknown as AiSettings,
    setAiSettings: async () => {},
    onAiSettingsChanged: () => () => {},
    openAiModelSettings: async () => {},
    skillsList: async () => [],
    skillText: async () => '',
    setAiPanelPrefs: async (prefs: unknown) => prefs,
    openSourcePath: async () => {},
    aiStream: async (request: AiStreamRequest) => {
      requests.push(request)
    },
    aiStreamCancel: async () => {},
    onAiStream: (handler: StreamHandler) => {
      listener = handler
      return () => {
        listener = null
      }
    },
  }
  ;(window as unknown as { markdownApi: MarkdownApi }).markdownApi = api as unknown as MarkdownApi

  const send = (chunk: AiStreamChunk) => {
    act(() => {
      listener?.(chunk)
    })
  }
  /** pump the task queue until the loop has started a request past `after` */
  const waitForTurn = async (after = 0): Promise<AiStreamRequest> => {
    for (let i = 0; i < 200 && requests.length <= after; i++) {
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 1))
      })
    }
    if (requests.length <= after) throw new Error('the agent loop never started the next turn')
    return requests[requests.length - 1]!
  }
  /** answer the current turn with the apply_ops call, then close it */
  const answerWithToolCall = (turn: AiStreamRequest): void => {
    send({
      requestId: turn.requestId,
      type: 'tool-call',
      toolCall: {
        id: 't1',
        name: 'apply_ops',
        input: { ops: [{ op: 'insertContent', after: -1, markdown: 'AI inserted' }] },
      },
    })
    send({ requestId: turn.requestId, type: 'done' })
  }
  return { requests, send, waitForTurn, answerWithToolCall }
}

function renderPanel(deps: MarkdownAiDeps): HTMLDivElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  roots.push(root)
  // the preset auto-sends once on mount, the way a ribbon chip does
  act(() => {
    root.render(
      createElement(AiPanel, {
        deps,
        filePath: null,
        preset: { text: 'add a paragraph at the top', nonce: 1 },
        onCollapse: () => {},
      }),
    )
  })
  return container
}

const flush = async (): Promise<void> => {
  for (let i = 0; i < 6; i++) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 1))
    })
  }
}

describe('AI panel tool confirmation', () => {
  it('parks a write tool at the allow/deny card; denying leaves the document untouched', async () => {
    const editor = createEditor('original body')
    const deps: MarkdownAiDeps = {
      getEditor: () => editor,
      getFrontmatter: () => '',
      setFrontmatter: () => {},
      getSnapshot: () => ({ body: 'original body', frontmatter: '' }),
      restoreSnapshot: () => {},
      onRunDone: () => {},
    }
    const stream = installStreamStub()
    const container = renderPanel(deps)

    const turn = await stream.waitForTurn()
    stream.answerWithToolCall(turn)
    // the card is in the conversation and the loop holds at the gate
    expect(container.querySelector('.ai-confirm-card')).not.toBeNull()
    expect(stream.requests).toHaveLength(1)

    const deny = container.querySelector<HTMLButtonElement>('.ai-confirm-deny')!
    await act(async () => {
      deny.click()
    })
    // the settled decision can land the second turn inside the same act flush
    const turn2 = await stream.waitForTurn(1)

    // the tool never ran: the document still says what it said
    expect(editor.getMarkdown()).toContain('original body')
    expect(editor.getMarkdown()).not.toContain('AI inserted')
    // the model received the denial as an error tool result it can act on
    const toolMessage = turn2.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
    stream.send({ requestId: turn2.requestId, type: 'delta', text: 'ok' })
    stream.send({ requestId: turn2.requestId, type: 'done' })
    await flush()
    // the settled card stays in the flow as the record of the decision
    expect(container.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('denied')
    expect(container.querySelector('.ai-confirm-deny')).toBeNull()
  })

  it('allowing executes the write tool', async () => {
    const editor = createEditor('original body')
    const deps: MarkdownAiDeps = {
      getEditor: () => editor,
      getFrontmatter: () => '',
      setFrontmatter: () => {},
      getSnapshot: () => ({ body: 'original body', frontmatter: '' }),
      restoreSnapshot: () => {},
      onRunDone: () => {},
    }
    const stream = installStreamStub()
    const container = renderPanel(deps)

    const turn = await stream.waitForTurn()
    stream.answerWithToolCall(turn)
    const allow = container.querySelector<HTMLButtonElement>('.ai-confirm-allow')!
    await act(async () => {
      allow.click()
    })
    const turn2 = await stream.waitForTurn(1)

    // the write happened through the real editor path
    expect(editor.getMarkdown()).toContain('AI inserted')
    const toolMessage = turn2.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).not.toBe(TOOL_DENIED_OUTPUT)
    stream.send({ requestId: turn2.requestId, type: 'delta', text: 'ok' })
    stream.send({ requestId: turn2.requestId, type: 'done' })
    await flush()
    expect(container.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe(
      'allowed',
    )
  })

  it('a read tool never asks: it runs straight through', async () => {
    const editor = createEditor('original body')
    const deps: MarkdownAiDeps = {
      getEditor: () => editor,
      getFrontmatter: () => '',
      setFrontmatter: () => {},
      getSnapshot: () => ({ body: 'original body', frontmatter: '' }),
      restoreSnapshot: () => {},
      onRunDone: () => {},
    }
    const stream = installStreamStub()
    const container = renderPanel(deps)

    const turn = await stream.waitForTurn()
    stream.send({
      requestId: turn.requestId,
      type: 'tool-call',
      toolCall: { id: 't1', name: 'read_blocks', input: { startIndex: 0, endIndex: 0 } },
    })
    stream.send({ requestId: turn.requestId, type: 'done' })
    await stream.waitForTurn(1)
    await flush()
    expect(container.querySelector('.ai-confirm-card')).toBeNull()
  })

  it('unmounting the panel denies the pending request', async () => {
    const editor = createEditor('original body')
    const deps: MarkdownAiDeps = {
      getEditor: () => editor,
      getFrontmatter: () => '',
      setFrontmatter: () => {},
      getSnapshot: () => ({ body: 'original body', frontmatter: '' }),
      restoreSnapshot: () => {},
      onRunDone: () => {},
    }
    const stream = installStreamStub()
    const container = renderPanel(deps)

    const turn = await stream.waitForTurn()
    stream.answerWithToolCall(turn)
    expect(container.querySelector('.ai-confirm-card')).not.toBeNull()
    const root = roots.pop()!
    await act(async () => {
      root.unmount()
      await new Promise((resolve) => setTimeout(resolve, 0))
    })
    // the tool never ran, and the panel's unmount cancels the run: no second
    // turn is opened (the pending decision itself settles as denied)
    expect(editor.getMarkdown()).not.toContain('AI inserted')
    expect(stream.requests).toHaveLength(1)
  })
})
