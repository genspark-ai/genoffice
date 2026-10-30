// @vitest-environment jsdom
// Sheets: the approval gate for risk:'high' workbook-write tools
// (propose_operations, create_document, merge_attached_workbooks). The App's
// AgentLoop parks them at the shared useToolConfirm card (rendered by
// AiChatPanel). These tests pin the gate on the real loop with the real
// workbook skill (deps stubbed at the apply surface), plus the panel's card
// wiring — App.tsx contributes only the three-line hook plumbing between them.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import {
  AgentLoop,
  composeSkills,
  TOOL_DENIED_OUTPUT,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentTransport,
} from '@genoffice/agent-core'
import { createWorkbookSkill } from '../src/renderer/ai/workbook-skill'
import type { ActiveSheetInfo, SheetsSkillDeps } from '../src/renderer/ai/tools'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { ChangePlan } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { ToolConfirmCard } from '@genoffice/ui'
import { AiChatPanel } from '../src/renderer/ai/AiChatPanel'

// ── the loop-level gate: real AgentLoop + real workbook skill ──

const EMPTY_PLAN: ChangePlan = {
  transactionId: 'agent-1',
  baseRevision: 0,
  cellChanges: [
    { sheetId: 'sheet-1', address: 'A1', before: { value: 'old' }, after: { value: 'new' } },
  ],
  sheetRenames: [],
  structuralChanges: [],
  formatChanges: [],
  warnings: [],
}

/** deps stubbed at the one surface a write lands on */
function fakeDeps(proposeOperations: SheetsSkillDeps['proposeOperations']): SheetsSkillDeps {
  const info: ActiveSheetInfo = {
    mode: 'demo',
    sheetId: 'sheet-1',
    sheetName: 'Sheet1',
    revision: 0,
    knownAddresses: ['A1', 'B1'],
    sheets: [{ id: 'sheet-1', name: 'Sheet1' }],
    selection: 'B2:C4',
  }
  return {
    getActiveSheetInfo: () => info,
    readCells: () => ({}),
    readFormats: () => ({}),
    readSheetFeatures: () => 'Feature state of sheet Sheet1 (id=sheet-1):\nAutoFilter: none',
    findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
    selectRange: () => ({ ok: true, sheetName: 'Sheet1' }),
    tracePrecedents: () => ({ refs: [] }),
    traceDependents: () => ({ dependents: [], truncated: false, incompleteSheets: [] }),
    proposeOperations,
  }
}

/** transport stub: records requests, lets the test script each turn */
function stubTransport() {
  const requests: AgentStreamRequest[] = []
  let callbacks: AgentStreamCallbacks | null = null
  const transport: AgentTransport = {
    stream(request, cb) {
      requests.push(request)
      callbacks = cb
      return { cancel: () => {} }
    },
  }
  const toolCall = (name: string, input: Record<string, unknown>) =>
    callbacks?.onToolCall({ id: 't1', name, input })
  const done = () => callbacks?.onDone()
  const delta = (text: string) => callbacks?.onDelta(text)
  return { requests, transport, toolCall, done, delta }
}

/** the panel hook's stand-in: one pending decision at a time, settled by hand */
function manualConfirm() {
  let settle: ((allow: boolean) => void) | null = null
  return {
    confirmTool: () =>
      new Promise<boolean>((resolve) => {
        settle = resolve
      }),
    pending: () => settle !== null,
    settle: (allow: boolean) => {
      const resolve = settle
      settle = null
      resolve?.(allow)
    },
  }
}

const tick = async (): Promise<void> => {
  for (let i = 0; i < 6; i++) await new Promise((resolve) => setTimeout(resolve, 0))
}

describe('agent loop tool confirmation', () => {
  it('denying a write batch never touches the workbook apply surface', async () => {
    const proposeOperations = vi.fn((operations: readonly WorkbookOperation[], summary: string) => {
      void operations
      void summary
      return { ok: true as const, plan: EMPTY_PLAN }
    })
    const { requests, transport, toolCall, done } = stubTransport()
    const confirm = manualConfirm()
    const loop = new AgentLoop({
      transport,
      confirmTool: confirm.confirmTool,
      skill: composeSkills('sheets', '', [createWorkbookSkill(fakeDeps(proposeOperations))]),
    })
    void loop.run('set A1 to AI edited')

    await tick()
    expect(requests).toHaveLength(1)
    toolCall('propose_operations', {
      operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'AI edited' }],
      summary: 'set A1',
    })
    done()

    // the loop holds at the gate: the write surface was not touched and no
    // second turn opened
    await tick()
    expect(confirm.pending()).toBe(true)
    expect(proposeOperations).not.toHaveBeenCalled()
    expect(requests).toHaveLength(1)

    confirm.settle(false)
    await tick()
    // the model received the denial as an error tool result it can act on
    expect(proposeOperations).not.toHaveBeenCalled()
    expect(requests).toHaveLength(2)
    const toolMessage = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).toBe(TOOL_DENIED_OUTPUT)
  })

  it('allowing a write batch runs it through the apply surface', async () => {
    const proposeOperations = vi.fn((operations: readonly WorkbookOperation[], summary: string) => {
      void operations
      void summary
      return { ok: true as const, plan: EMPTY_PLAN }
    })
    const { requests, transport, toolCall, done, delta } = stubTransport()
    const confirm = manualConfirm()
    const loop = new AgentLoop({
      transport,
      confirmTool: confirm.confirmTool,
      skill: composeSkills('sheets', '', [createWorkbookSkill(fakeDeps(proposeOperations))]),
    })
    void loop.run('set A1 to AI edited')

    await tick()
    toolCall('propose_operations', {
      operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'AI edited' }],
      summary: 'set A1',
    })
    done()
    await tick()
    expect(confirm.pending()).toBe(true)

    confirm.settle(true)
    await tick()
    // the write happened through the real skill path
    expect(proposeOperations).toHaveBeenCalledTimes(1)
    expect(proposeOperations.mock.calls[0]![0]).toEqual([
      { op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'AI edited' },
    ])
    // the tool result went back to the model (its round-trip continued)
    expect(requests).toHaveLength(2)
    const toolMessage = requests[1]!.messages.find((m) => m.role === 'tool')
    expect(toolMessage?.results?.[0]?.output).not.toBe(TOOL_DENIED_OUTPUT)
    delta('ok')
    done()
  })

  it('a read tool never asks: it runs straight through', async () => {
    const proposeOperations = vi.fn((operations: readonly WorkbookOperation[], summary: string) => {
      void operations
      void summary
      return { ok: true as const, plan: EMPTY_PLAN }
    })
    const { requests, transport, toolCall, done } = stubTransport()
    const confirm = manualConfirm()
    const loop = new AgentLoop({
      transport,
      confirmTool: confirm.confirmTool,
      skill: composeSkills('sheets', '', [createWorkbookSkill(fakeDeps(proposeOperations))]),
    })
    void loop.run('what is in A1')

    await tick()
    toolCall('read_cells', { addresses: ['A1'] })
    done()
    await tick()
    expect(confirm.pending()).toBe(false)
    expect(proposeOperations).not.toHaveBeenCalled()
    expect(requests).toHaveLength(2)
  })
})

// ── the panel-level card: AiChatPanel renders what the App hook owns ──

const roots: Root[] = []

beforeAll(() => {
  ;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
})

afterEach(() => {
  for (const root of roots.splice(0)) act(() => root.unmount())
  document.body.innerHTML = ''
})

function mountPanel(confirmCard: ToolConfirmCard | null): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  roots.push(root)
  act(() =>
    root.render(
      createElement(AiChatPanel, {
        isOpen: true,
        hasContent: true,
        chat: [
          { role: 'user', text: 'set A1', tools: [] },
          { role: 'assistant', text: '', tools: [], streaming: true },
        ],
        attachments: [],
        attachedSkills: [],
        onToggleSkill: () => {},
        attachNotice: null,
        onPickAttachments: () => {},
        onAddAttachmentPaths: () => {},
        onAddPastedImage: () => {},
        onRemoveAttachment: () => {},
        prompt: '',
        preview: null,
        aiBusy: true,
        confirmCard,
        onConfirmAllow: () => {},
        onConfirmDeny: () => {},
        onPromptChange: () => {},
        onSend: () => {},
        onStop: () => {},
        onNewChat: () => {},
        onUndo: () => {},
        scopeRange: null,
        scopeColumns: null,
        scopeLocked: false,
        onScopeDismiss: () => {},
        onCitation: () => {},
        onExpand: () => {},
        onCollapse: () => {},
      }),
    ),
  )
  return container
}

describe('AiChatPanel confirm card', () => {
  it('renders the pending card with live allow/deny buttons', () => {
    const container = mountPanel({
      toolName: 'propose_operations',
      args: '{"operations":[{"op":"set_cell"}]}',
      outcome: 'awaiting',
      timedOut: false,
    })
    expect(container.querySelector('.ai-confirm-card')).not.toBeNull()
    expect(container.querySelector('.ai-confirm-allow')).not.toBeNull()
    expect(container.querySelector('.ai-confirm-deny')).not.toBeNull()
  })

  it('renders nothing while the gate is idle and settles the card visually', () => {
    const idle = mountPanel(null)
    expect(idle.querySelector('.ai-confirm-card')).toBeNull()

    const settled = mountPanel({
      toolName: 'propose_operations',
      args: '{}',
      outcome: 'denied',
      timedOut: false,
    })
    // a settled card is a record: the buttons are gone, the outcome stays
    expect(settled.querySelector('.ai-confirm-deny')).toBeNull()
    expect(settled.querySelector('.ai-confirm-card')?.getAttribute('data-outcome')).toBe('denied')
  })
})
