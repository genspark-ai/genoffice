import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * A stand-in for ipcMain that keeps real Electron's most important property:
 * registering a second handler for the same channel throws. That is what makes
 * the placement of registerAgentToolIpc relative to each app's
 * already-registered guard load-bearing.
 */
const handlers = new Map<string, (...args: unknown[]) => unknown>()

vi.mock('electron', () => ({
  ipcMain: {
    handle(channel: string, fn: (...args: unknown[]) => unknown) {
      if (handlers.has(channel)) {
        throw new Error(`Attempted to register a second handler for '${channel}'`)
      }
      handlers.set(channel, fn)
    },
    removeHandler(channel: string) {
      handlers.delete(channel)
    },
  },
}))

// pulls in @genspark/cli otherwise; none of these tests exercise extraction
vi.mock('@genoffice/ai-search', () => ({ tavilyExtract: vi.fn() }))

const { registerAgentToolIpc } = await import('../src/agent-ipc')
const { AgentInstructionsStore } = await import('../src/agent-instructions-store')

let dir: string
let store: InstanceType<typeof AgentInstructionsStore>

beforeEach(() => {
  handlers.clear()
  dir = mkdtempSync(join(tmpdir(), 'agent-ipc-'))
  store = new AgentInstructionsStore(dir)
  registerAgentToolIpc(() => store)
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

const invoke = (channel: string, ...args: unknown[]): unknown =>
  handlers.get(channel)!({ sender: {} }, ...args)

describe('registerAgentToolIpc', () => {
  // The point of the module: these are exactly the channels createWebSkill's
  // bridge calls, and standalone slides/sheets had none of them.
  it('registers every channel the web skill needs', () => {
    expect([...handlers.keys()].sort()).toEqual([
      'ai:browse-page',
      'ai:capture-page',
      'ai:extract-pages',
      'ai:forget',
      'ai:instructions-prompt',
      'ai:remember',
      'ai:skill-body',
    ])
  })

  // Guards the mistake this fix nearly shipped: the call sitting before an
  // app's own "already registered" early-return, so a second window throws.
  it('throws when called twice, so it must sit behind each app-level guard', () => {
    expect(() => registerAgentToolIpc(() => store)).toThrow(/second handler/)
  })

  it('resolves the store lazily, not at registration time', () => {
    handlers.clear()
    const get = vi.fn(() => store)
    registerAgentToolIpc(get)
    expect(get).not.toHaveBeenCalled()
    invoke('ai:instructions-prompt', 'docx')
    expect(get).toHaveBeenCalled()
  })

  it('records a memory through the remember channel', () => {
    expect(invoke('ai:remember', 'prefers metric units')).toBe(true)
    expect(store.readMemories().map((m) => m.text)).toEqual(['prefers metric units'])
  })

  it('forgets by wording, case-insensitively, and reports a miss', () => {
    invoke('ai:remember', 'prefers metric units')
    expect(invoke('ai:forget', 'PREFERS METRIC UNITS')).toBe(true)
    expect(store.readMemories()).toEqual([])
    expect(invoke('ai:forget', 'never said this')).toBe(false)
  })

  it('puts recorded memories into the prompt it builds', () => {
    invoke('ai:remember', 'prefers metric units')
    expect(String(invoke('ai:instructions-prompt', 'docx'))).toContain('prefers metric units')
  })
})

describe('ai:capture-page', () => {
  const capture = (image: unknown, rect?: unknown) =>
    handlers.get('ai:capture-page')!(
      { sender: { capturePage: vi.fn().mockResolvedValue(image) } },
      rect,
    )

  it('returns base64 for a real capture', async () => {
    const image = { isEmpty: () => false, toPNG: () => Buffer.from('png-bytes') }
    await expect(capture(image)).resolves.toBe(Buffer.from('png-bytes').toString('base64'))
  })

  // an occluded or zero-sized view answers with an empty image; '' lets the
  // tool say so instead of handing the model a blank picture
  it('returns empty string for an empty capture', async () => {
    await expect(capture({ isEmpty: () => true, toPNG: () => Buffer.alloc(0) })).resolves.toBe('')
  })

  it('survives capturePage rejecting', async () => {
    const sender = { capturePage: vi.fn().mockRejectedValue(new Error('gone')) }
    await expect(handlers.get('ai:capture-page')!({ sender })).resolves.toBe('')
  })

  it('rounds a fractional rect and drops a degenerate one', async () => {
    const capturePage = vi
      .fn()
      .mockResolvedValue({ isEmpty: () => false, toPNG: () => Buffer.of(1) })
    await handlers.get('ai:capture-page')!(
      { sender: { capturePage } },
      { x: 10.4, y: 20.6, width: 100.5, height: 50.5 },
    )
    expect(capturePage).toHaveBeenCalledWith({ x: 10, y: 21, width: 101, height: 51 })

    capturePage.mockClear()
    // a rect the renderer could not resolve must capture the whole view, not a
    // zero-sized region that comes back empty
    await handlers.get('ai:capture-page')!(
      { sender: { capturePage } },
      { x: 0, y: 0, width: 0, height: 0 },
    )
    expect(capturePage).toHaveBeenCalledWith()
  })
})
