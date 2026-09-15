/**
 * Tests for the frozen-selection extension (W9 deliverable).
 *
 * Verifies:
 *   - session_start fires the extension's handler
 *   - The handler calls ctx.ui.setCustomData with the frozen snapshot
 *   - Stale editor (no selection) is a no-op (no customData write)
 *   - Custom key + custom fingerprint are honoured
 *   - Multiple session_start events refresh the capturedAt timestamp
 */

import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { ReactUIAdapter } from '@genoffice/agent-runtime'
import type { ExtensionAPI, ExtensionContext } from '@earendil-works/pi-coding-agent'
import type {
  BeforeAgentStartEvent,
  BeforeAgentStartEventResult,
  SessionStartEvent,
  ToolCallEvent,
  ToolCallEventResult,
} from '@earendil-works/pi-coding-agent'
import {
  createFrozenSelectionExtension,
  type FrozenSelectionEditor,
} from '../src/extensions/frozen-selection'

// ---------------------------------------------------------------------------
// Mock editor
// ---------------------------------------------------------------------------

class MockEditor implements FrozenSelectionEditor<{ from: number; to: number }> {
  scope: { from: number; to: number } | null
  units: string[]

  constructor(units: string[], scope: { from: number; to: number } | null = { from: 0, to: 2 }) {
    this.units = units
    this.scope = scope
  }

  getSelectionScope() {
    return this.scope
  }

  getUnitCount() {
    return this.units.length
  }

  getUnitText(index: number) {
    return this.units[index] ?? ''
  }
}

// ---------------------------------------------------------------------------
// Mini pi harness — captures handlers registered by extensions
// ---------------------------------------------------------------------------

type Handler<E, R> = (event: E, ctx: ExtensionContext) => Promise<R | undefined> | R | undefined

interface CapturedHandlers {
  session_start: Array<Handler<SessionStartEvent, void>>
  before_agent_start: Array<Handler<BeforeAgentStartEvent, BeforeAgentStartEventResult>>
  tool_call: Array<Handler<ToolCallEvent, ToolCallEventResult>>
}

function makeHarness(ui: unknown): {
  api: ExtensionAPI
  captured: CapturedHandlers
  fire: (k: keyof CapturedHandlers, e: unknown) => Promise<unknown[]>
} {
  const captured: CapturedHandlers = {
    session_start: [],
    before_agent_start: [],
    tool_call: [],
  }

  const api: ExtensionAPI = {
    // Cast — pi overloads make TS inference painful here; tests use a loose API surface.
    on(event: string, handler: unknown) {
      const key = event as keyof CapturedHandlers
      if (!captured[key]) throw new Error(`Unexpected event ${event}`)
      ;(captured[key] as unknown[]).push(handler as never)
    },
  } as unknown as ExtensionAPI

  const ctx = { ui } as unknown as ExtensionContext

  const fire = async (k: keyof CapturedHandlers, e: unknown) => {
    const out: unknown[] = []
    for (const h of captured[k]) {
      out.push(await (h as (event: unknown, ctx: ExtensionContext) => Promise<unknown>)(e, ctx))
    }
    return out
  }

  return { api, captured, fire }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('frozen-selection extension', () => {
  let adapter: ReactUIAdapter

  beforeEach(() => {
    adapter = new ReactUIAdapter()
  })

  afterEach(() => {
    // No state to clean up — adapter is per-test.
  })

  it('captures the selection scope at session_start', async () => {
    const editor = new MockEditor(['para 1', 'para 2', 'para 3'])
    const ext = createFrozenSelectionExtension({ getEditor: () => editor })
    const { api, captured, fire } = makeHarness(adapter)
    ext(api)

    expect(captured.session_start.length).toBe(1)

    await fire('session_start', { type: 'session_start', reason: 'startup' })

    const frozen = adapter.getCustomData<{
      scope: { from: number; to: number }
      docFingerprint: string
      capturedAt: number
    }>('frozenSelection')
    expect(frozen).toBeDefined()
    expect(frozen!.scope).toEqual({ from: 0, to: 2 })
    expect(frozen!.docFingerprint).toBe('3|para 1|para 3')
    expect(frozen!.capturedAt).toBeGreaterThan(0)
  })

  it('is a no-op when the editor has no selection', async () => {
    const editor = new MockEditor(['para 1'], null)
    const ext = createFrozenSelectionExtension({ getEditor: () => editor })
    const { api, captured, fire } = makeHarness(adapter)
    ext(api)

    await fire('session_start', { type: 'session_start', reason: 'startup' })
    expect(adapter.getCustomData('frozenSelection')).toBeUndefined()
  })

  it('is a no-op when the editor is not available', async () => {
    const ext = createFrozenSelectionExtension({ getEditor: () => null })
    const { api, captured, fire } = makeHarness(adapter)
    ext(api)

    await fire('session_start', { type: 'session_start', reason: 'startup' })
    expect(adapter.getCustomData('frozenSelection')).toBeUndefined()
  })

  it('honours customDataKey and fingerprint overrides', async () => {
    const editor = new MockEditor(['a', 'b', 'c'], { from: 5, to: 7 })
    const ext = createFrozenSelectionExtension({
      getEditor: () => editor,
      customDataKey: 'docs:frozen',
      fingerprint: (e) => `units=${e.getUnitCount()}`,
    })
    const { api, captured, fire } = makeHarness(adapter)
    ext(api)

    await fire('session_start', { type: 'session_start', reason: 'startup' })

    const frozen = adapter.getCustomData<{ docFingerprint: string; scope: { from: number; to: number } }>('docs:frozen')
    expect(frozen).toBeDefined()
    expect(frozen!.docFingerprint).toBe('units=3')
    expect(frozen!.scope).toEqual({ from: 5, to: 7 })

    // Default key is not set when override is used
    expect(adapter.getCustomData('frozenSelection')).toBeUndefined()
  })

  it('refreshes the snapshot on subsequent session_start events', async () => {
    const editor = new MockEditor(['x', 'y'], { from: 0, to: 1 })
    const ext = createFrozenSelectionExtension({ getEditor: () => editor })
    const { api, captured, fire } = makeHarness(adapter)
    ext(api)

    await fire('session_start', { type: 'session_start', reason: 'new' })
    const first = adapter.getCustomData<{ capturedAt: number }>('frozenSelection')

    // Wait at least 1ms to ensure timestamp moves
    await new Promise((r) => setTimeout(r, 5))

    editor.scope = { from: 1, to: 1 }
    await fire('session_start', { type: 'session_start', reason: 'reload' })
    const second = adapter.getCustomData<{ capturedAt: number; scope: { from: number; to: number } }>('frozenSelection')

    expect(second).toBeDefined()
    expect(second!.capturedAt).toBeGreaterThan(first!.capturedAt)
    expect(second!.scope).toEqual({ from: 1, to: 1 })
  })
})

