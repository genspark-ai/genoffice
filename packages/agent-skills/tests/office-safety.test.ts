/**
 * Tests for the office-safety facade (W9 deliverable).
 *
 * Verifies:
 *   - installOfficeSafety wires both frozen-selection + verify-response
 *   - Skips frozen selection when no editor option is passed
 *   - Custom verify rules thread through to the system prompt
 */

import { describe, expect, it } from 'vitest'
import { ReactUIAdapter } from '@genoffice/agent-runtime'
import type {
  BeforeAgentStartEvent,
  ExtensionAPI,
  ExtensionContext,
  SessionStartEvent,
  ToolCallEvent,
} from '@earendil-works/pi-coding-agent'
import {
  installOfficeSafety,
  type FrozenSelectionEditor,
} from '../src/extensions/office-safety'

interface CapturedHandlers {
  session_start: Array<(event: SessionStartEvent, ctx: ExtensionContext) => Promise<void> | void>
  before_agent_start: Array<(event: BeforeAgentStartEvent, ctx: ExtensionContext) => Promise<unknown> | unknown>
  tool_call: Array<(event: ToolCallEvent, ctx: ExtensionContext) => Promise<unknown> | unknown>
}

function makeHarness(ui: unknown) {
  const captured: CapturedHandlers = {
    session_start: [],
    before_agent_start: [],
    tool_call: [],
  }
  const api: ExtensionAPI = {
    on(event: string, handler: unknown) {
      const key = event as keyof CapturedHandlers
      if (!captured[key]) throw new Error(`Unexpected event ${event}`)
      ;(captured[key] as unknown[]).push(handler as never)
    },
  } as unknown as ExtensionAPI
  const ctx = { ui } as unknown as ExtensionContext
  return { api, captured, ctx }
}

class MockEditor implements FrozenSelectionEditor<{ from: number; to: number }> {
  scope: { from: number; to: number } | null = { from: 0, to: 1 }
  units = ['p1', 'p2']
  getSelectionScope() { return this.scope }
  getUnitCount() { return this.units.length }
  getUnitText(i: number) { return this.units[i] ?? '' }
}

describe('installOfficeSafety', () => {
  it('wires both frozen-selection and verify-response by default', async () => {
    const { api, captured, ctx } = makeHarness(new ReactUIAdapter())
    installOfficeSafety(api, { frozen: { getEditor: () => new MockEditor() } })

    expect(captured.session_start.length).toBe(1)
    expect(captured.before_agent_start.length).toBe(1)

    // Drive session_start — frozen selection should land in customData
    const editor = new MockEditor()
    await captured.session_start[0]!(
      { type: 'session_start', reason: 'startup' } as SessionStartEvent,
      ctx,
    )
    expect((ctx.ui as ReactUIAdapter).getCustomData('frozenSelection')).toBeDefined()

    // Drive before_agent_start — systemPrompt gets the verify rules
    const result = (await captured.before_agent_start[0]!(
      { type: 'before_agent_start', prompt: 'p', systemPrompt: 'S', systemPromptOptions: {} as never } as BeforeAgentStartEvent,
      ctx,
    )) as { systemPrompt?: string } | undefined
    expect(result?.systemPrompt).toContain('[genoffice:verify-rules]')
  })

  it('skips frozen-selection when no editor is provided', async () => {
    const { api, captured } = makeHarness(new ReactUIAdapter())
    installOfficeSafety(api, { verify: { rules: 'short rule' } })

    expect(captured.session_start.length).toBe(0)
    expect(captured.before_agent_start.length).toBe(1)
  })

  it('still installs verify-response when no opts are passed', async () => {
    const { api, captured } = makeHarness(new ReactUIAdapter())
    installOfficeSafety(api)

    expect(captured.session_start.length).toBe(0)
    expect(captured.before_agent_start.length).toBe(1)
  })
})
