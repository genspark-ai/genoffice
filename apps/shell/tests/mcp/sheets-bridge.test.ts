import { beforeEach, describe, expect, it, vi } from 'vitest'
import { EventEmitter } from 'node:events'

/**
 * The sheets MCP bridge's reply plumbing.
 *
 * A workbook lives in the renderer, so this bridge passes commands over IPC and
 * matches replies by requestId. requestIds are guessable, so a reply must be
 * accepted only from the tab the command targeted, and a tab that goes away must
 * fail its in-flight commands instead of leaving them to time out.
 */

interface FakeSender {
  id: number
  isDestroyed: () => boolean
  send: (channel: string, payload: unknown) => void
  once: (event: string, handler: () => void) => void
}

const handlers = new Map<string, (event: { sender: FakeSender }, payload?: unknown) => void>()
const senders = new Map<number, FakeSender>()
const liveSenders = new Map<number, FakeSender>()

vi.mock('electron', () => ({
  ipcMain: {
    on: (channel: string, handler: (event: { sender: FakeSender }, payload?: unknown) => void) => {
      handlers.set(channel, handler)
    },
  },
  webContents: {
    fromId: (id: number) => liveSenders.get(id),
  },
}))

function senderFor(id: number): FakeSender {
  const existing = senders.get(id)
  if (existing) return existing
  const emitter = new EventEmitter()
  const sender: FakeSender = {
    id,
    isDestroyed: () => !liveSenders.has(id),
    send: vi.fn(),
    once: (event, handler) => {
      emitter.once(event, handler)
    },
  }
  senders.set(id, sender)
  liveSenders.set(id, sender)
  ;(sender as unknown as { destroy: () => void }).destroy = () => {
    liveSenders.delete(id)
    emitter.emit('destroyed')
  }
  return sender
}

/** the command payload the bridge pushed to a tab */
function lastCommand(wcId: number): { requestId: string; command: string } {
  const sender = senders.get(wcId)!
  const call = (sender.send as ReturnType<typeof vi.fn>).mock.calls.at(-1)
  expect(call?.[0]).toBe('sheets:mcp-command')
  return call?.[1] as { requestId: string; command: string }
}

/**
 * runCommand awaits readiness before it registers the pending entry, so a test
 * must let that turn complete before it can observe or pre-empt the command.
 */
async function nextCommand(wcId: number): Promise<{ requestId: string; command: string }> {
  for (let i = 0; i < 50; i++) {
    const sender = senders.get(wcId)
    if ((sender?.send as ReturnType<typeof vi.fn> | undefined)?.mock.calls.length) {
      return lastCommand(wcId)
    }
    await new Promise((resolve) => setTimeout(resolve, 1))
  }
  throw new Error('the bridge never sent a command')
}

async function loadBridge() {
  vi.resetModules()
  handlers.clear()
  ;(senders as Map<number, FakeSender>).clear?.()
  const mod = await import('../../src/main/mcp/sheets-bridge')
  return mod
}

beforeEach(() => {
  senders.clear()
  liveSenders.clear()
  handlers.clear()
})

describe('sheets bridge reply routing', () => {
  it('ignores a reply whose requestId belongs to another tab', async () => {
    const { installSheetsBridge, createSheetsControl } = await loadBridge()
    installSheetsBridge()

    const target = senderFor(1)
    const attacker = senderFor(2)
    handlers.get('sheets:mcp-ready')!({ sender: target })

    const control = createSheetsControl({ openBlankTab: async () => 1 })
    const pending = control.runCommand(1, 'read_sheet', {})
    const { requestId } = await nextCommand(1)

    // a different tab answers with the guessable id: it must not resolve the command
    handlers.get('sheets:mcp-result')!(
      { sender: attacker },
      { requestId, ok: true, result: { spoofed: true } },
    )

    let settled = false
    void pending.then(
      () => (settled = true),
      () => (settled = true),
    )
    await new Promise((resolve) => setTimeout(resolve, 20))
    expect(settled).toBe(false)

    // the real target can still answer it
    handlers.get('sheets:mcp-result')!(
      { sender: target },
      { requestId, ok: true, result: { real: true } },
    )
    await expect(pending).resolves.toEqual({ real: true })
  })

  it('fails in-flight commands when the target tab goes away', async () => {
    const { installSheetsBridge, createSheetsControl } = await loadBridge()
    installSheetsBridge()

    const target = senderFor(1)
    handlers.get('sheets:mcp-ready')!({ sender: target })

    const control = createSheetsControl({ openBlankTab: async () => 1 })
    const pending = control.runCommand(1, 'read_sheet', {})
    await nextCommand(1)
    ;(target as unknown as { destroy: () => void }).destroy()

    await expect(pending).rejects.toThrow(/tab was closed/)
  })

  it('fails a read that is waiting for a tab that never becomes ready', async () => {
    const { installSheetsBridge, createSheetsControl } = await loadBridge()
    installSheetsBridge()

    const target = senderFor(1)
    const control = createSheetsControl({ openBlankTab: async () => 1 })
    // no sheets:mcp-ready for this tab — the wait must not hang forever
    const pending = control.runCommand(1, 'read_sheet', {})
    ;(target as unknown as { destroy: () => void }).destroy()
    await expect(pending).rejects.toThrow(/tab was closed/)
  })
})

describe('create_session cleanup', () => {
  it('abandons the blank tab when readiness fails, so no orphan is left behind', async () => {
    const { installSheetsBridge, createSheetsControl } = await loadBridge()
    installSheetsBridge()

    // the tab exists but never announces readiness
    const target = senderFor(7)
    const abandon = vi.fn()
    const control = createSheetsControl({
      openBlankTab: async () => 7,
      abandonBlankTab: abandon,
    })

    const pending = control.openBlankTab()
    ;(target as unknown as { destroy: () => void }).destroy()
    await expect(pending).rejects.toThrow(/tab was closed/)
    expect(abandon).toHaveBeenCalledWith(7)
  })

  it('leaves a ready tab alone', async () => {
    const { installSheetsBridge, createSheetsControl } = await loadBridge()
    installSheetsBridge()

    const target = senderFor(8)
    const abandon = vi.fn()
    const control = createSheetsControl({
      openBlankTab: async () => 8,
      abandonBlankTab: abandon,
    })
    const pending = control.openBlankTab()
    handlers.get('sheets:mcp-ready')!({ sender: target })

    await expect(pending).resolves.toBe(8)
    expect(abandon).not.toHaveBeenCalled()
  })
})
