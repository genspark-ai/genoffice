// #543 P1: a rollback point outlives the session it was taken in. The transcript
// is restored from the JSONL, the documents come back from the store on demand,
// and a point whose snapshot is gone says so instead of turning into a dead button.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiPanel } from '../src/renderer/ai/AiPanel'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'
import { t } from '../src/renderer/i18n/locale'

vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

/** The document as it was before the stored turn — what a roll back must restore */
const SNAPSHOT_DOC = {
  type: 'doc',
  content: [
    {
      type: 'docParagraph',
      attrs: { docxIndex: 0 },
      content: [{ type: 'text', text: 'before the AI rewrote this' }],
    },
  ],
}

const SNAPSHOT_ID = 'snap-11111111'

/** What a reopened file's transcript looks like: the ref survives, the bytes do not */
const TRANSCRIPT = [
  { role: 'user' as const, text: 'rewrite the intro' },
  {
    role: 'assistant' as const,
    text: 'Rewritten.',
    version: { id: 4, label: 'rewrite the intro', time: '10:05', snapshotId: SNAPSHOT_ID },
  },
]

function mockProjectApi(options: {
  transcript?: unknown[]
  available?: string[]
  snapshot?: string | null
}) {
  const win = window as unknown as { projectApi?: unknown }
  const previous = win.projectApi
  const api = {
    resolveChat: vi.fn(async () => ({ projectId: 'p', chatId: 'c' })),
    loadChat: vi.fn(async () => options.transcript ?? []),
    listChatSnapshots: vi.fn(async () => options.available ?? []),
    loadChatSnapshot: vi.fn(async () => options.snapshot ?? null),
    appendChat: vi.fn(async () => {}),
    rebindChat: vi.fn(async () => ({ projectId: 'p', chatId: 'c' })),
    saveChatSnapshot: vi.fn(async () => ({ snapshotId: null, bytes: 0 })),
  }
  win.projectApi = api
  return {
    api,
    restore: () => {
      win.projectApi = previous
    },
  }
}

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'the AI rewrote this' }],
        },
      ],
    },
  })
}

let mounted: Array<{ root: Root; container: HTMLElement }> = []
let editors: Editor[] = []

function mountPanel(editor: Editor): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(AiPanel, {
        editor,
        blocks: [],
        settings,
        open: true,
        onExpand: () => {},
        onCollapse: () => {},
      }),
    ),
  )
  mounted.push({ root, container })
  return container
}

/** Let resolveChat → loadChat → listChatSnapshots settle */
async function flush(times = 6): Promise<void> {
  for (let i = 0; i < times; i++) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0))
    })
  }
}

const docText = (editor: Editor) =>
  editor.state.doc.textBetween(0, editor.state.doc.content.size, '\n')

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  for (const { root, container } of mounted) {
    act(() => root.unmount())
    container.remove()
  }
  mounted = []
  for (const editor of editors) editor.destroy()
  editors = []
  vi.unstubAllGlobals()
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  vi.restoreAllMocks()
})

describe('rollback points after reopening a file', () => {
  it('lists the restored versions from the transcript', async () => {
    const mock = mockProjectApi({ transcript: TRANSCRIPT, available: [SNAPSHOT_ID] })
    const editor = createEditor()
    editors.push(editor)
    const container = mountPanel(editor)
    await flush()
    mock.restore()

    const rows = container.querySelectorAll('.ai-version-row')
    expect(rows.length).toBe(1)
    expect(rows[0].textContent).toContain('rewrite the intro')
    expect(rows[0].textContent).toContain('10:05')
    // its snapshot is still on disk, so it is offered as a real action
    expect(container.querySelector('.ai-version-rollback')?.textContent).toBe(t('aiRollback'))
    expect(docText(editor)).toBe('the AI rewrote this')
  })

  it('rolls the document back to a snapshot read from the store', async () => {
    const mock = mockProjectApi({
      transcript: TRANSCRIPT,
      available: [SNAPSHOT_ID],
      snapshot: JSON.stringify(SNAPSHOT_DOC),
    })
    const editor = createEditor()
    editors.push(editor)
    const container = mountPanel(editor)
    await flush()

    act(() => container.querySelector<HTMLButtonElement>('.ai-version-rollback')!.click())
    await flush()

    // the document came from the store, not from this session's memory
    expect(docText(editor)).toBe('before the AI rewrote this')
    expect(mock.api.loadChatSnapshot).toHaveBeenCalledWith({
      projectId: 'p',
      chatId: 'c',
      snapshotId: SNAPSHOT_ID,
    })
    // and the action flipped to its undo, exactly like an in-session roll back
    expect(container.querySelector('.ai-version-rollback')?.textContent).toBe(t('aiRollbackUndo'))
    mock.restore()
  })

  it('marks a version expired when its snapshot is gone', async () => {
    // retention dropped it (or it was never small enough to store): the row stays
    // in the record, and says why it cannot be used instead of hiding the action
    const mock = mockProjectApi({ transcript: TRANSCRIPT, available: [] })
    const editor = createEditor()
    editors.push(editor)
    const container = mountPanel(editor)
    await flush()
    mock.restore()

    expect(container.querySelectorAll('.ai-version-row').length).toBe(1)
    expect(container.querySelector('.ai-version-rollback')).toBeNull()
    expect(container.querySelector('.ai-version-note')?.textContent).toBe(t('aiVersionExpired'))
    // clicking nothing happened: the document is untouched
    expect(docText(editor)).toBe('the AI rewrote this')
  })

  it('marks a version expired when the store can no longer read it', async () => {
    const mock = mockProjectApi({
      transcript: TRANSCRIPT,
      available: [SNAPSHOT_ID],
      snapshot: null,
    })
    const editor = createEditor()
    editors.push(editor)
    const container = mountPanel(editor)
    await flush()

    act(() => container.querySelector<HTMLButtonElement>('.ai-version-rollback')!.click())
    await flush()
    mock.restore()

    // the file was listed a moment ago but is unreadable now
    expect(docText(editor)).toBe('the AI rewrote this')
    expect(container.querySelector('.ai-version-note')?.textContent).toBe(t('aiVersionExpired'))
  })

  it('leaves the panel without versions when the transcript has none', async () => {
    const mock = mockProjectApi({
      transcript: [{ role: 'user', text: 'just a question' }],
      available: [],
    })
    const editor = createEditor()
    editors.push(editor)
    const container = mountPanel(editor)
    await flush()
    mock.restore()

    expect(container.querySelector('.ai-versions')).toBeNull()
  })
})
