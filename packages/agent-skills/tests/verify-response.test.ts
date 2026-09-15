/**
 * Tests for the verify-response extension (W9 deliverable).
 *
 * Verifies:
 *   - before_agent_start returns a systemPrompt that contains the rules
 *   - Multiple chained handlers compose correctly (second sees first's prompt)
 *   - Custom rules and custom marker are honoured
 *   - Default marker is `[genoffice:verify-rules]`
 */

import { describe, expect, it } from 'vitest'
import type {
  BeforeAgentStartEvent,
  ExtensionAPI,
  ExtensionContext,
  SessionStartEvent,
  ToolCallEvent,
} from '@earendil-works/pi-coding-agent'
import {
  createVerifyResponseExtension,
  DEFAULT_VERIFY_RULES,
  VERIFY_BLOCK_MARKER,
} from '../src/extensions/verify-response'

interface CapturedHandlers {
  session_start: Array<(event: SessionStartEvent, ctx: ExtensionContext) => Promise<void> | void>
  before_agent_start: Array<(event: BeforeAgentStartEvent, ctx: ExtensionContext) => Promise<unknown> | unknown>
  tool_call: Array<(event: ToolCallEvent, ctx: ExtensionContext) => Promise<unknown> | unknown>
}

function makeHarness(): { api: ExtensionAPI; captured: CapturedHandlers; fireBefore: (prompt: string, sysPrompt: string) => Promise<string | undefined> } {
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

  const ctx = {} as unknown as ExtensionContext

  const fireBefore = async (userPrompt: string, sysPrompt: string): Promise<string | undefined> => {
    let last: string | undefined
    for (const h of captured.before_agent_start) {
      const ev: BeforeAgentStartEvent = {
        type: 'before_agent_start',
        prompt: userPrompt,
        systemPrompt: sysPrompt,
        systemPromptOptions: {} as never,
      }
      const result = await h(ev, ctx)
      if (result && typeof result === 'object' && 'systemPrompt' in result) {
        last = (result as { systemPrompt: string }).systemPrompt
        sysPrompt = last
      }
    }
    return last
  }

  return { api, captured, fireBefore }
}

describe('verify-response extension', () => {
  it('appends default rules + marker to the system prompt', async () => {
    const { api, fireBefore } = makeHarness()
    createVerifyResponseExtension()(api)

    const out = await fireBefore('user prompt', 'ORIGINAL SYSTEM PROMPT')
    expect(out).toBeDefined()
    expect(out).toContain('ORIGINAL SYSTEM PROMPT')
    expect(out).toContain(VERIFY_BLOCK_MARKER)
    expect(out).toContain(DEFAULT_VERIFY_RULES.split('\n')[1]!) // the first numbered rule
  })

  it('preserves user prompt + original system prompt order', async () => {
    const { api, fireBefore } = makeHarness()
    createVerifyResponseExtension()(api)

    const out = await fireBefore('user prompt', 'SYSTEM-A')
    expect(out).toBeDefined()
    // System prompt comes first, then rules block
    expect(out!.indexOf('SYSTEM-A')).toBeLessThan(out!.indexOf(VERIFY_BLOCK_MARKER))
    expect(out!.indexOf(VERIFY_BLOCK_MARKER)).toBeLessThan(out!.lastIndexOf(VERIFY_BLOCK_MARKER))
  })

  it('honours custom rules and marker', async () => {
    const { api, fireBefore } = makeHarness()
    createVerifyResponseExtension({
      rules: 'Never claim "done" without evidence',
      marker: '[my-rules]',
    })(api)

    const out = await fireBefore('p', 'S')
    expect(out).toContain('[my-rules]')
    expect(out).toContain('Never claim "done" without evidence')
    // Default rules are NOT present
    expect(out).not.toContain('Only claim an action')
    expect(out).not.toContain(VERIFY_BLOCK_MARKER)
  })

  it('chains cleanly with another extension that also rewrites systemPrompt', async () => {
    const { api, captured, fireBefore } = makeHarness()
    // Simulate another extension running first — add a prepend handler
    captured.before_agent_start.push(async (event) => {
      return { systemPrompt: `[OTHER]\n${event.systemPrompt}` }
    })
    createVerifyResponseExtension()(api)

    const out = await fireBefore('p', 'S0')
    expect(out).toBeDefined()
    expect(out).toContain('[OTHER]')
    expect(out).toContain('S0')
    expect(out).toContain(VERIFY_BLOCK_MARKER)
    expect(out!.indexOf('[OTHER]')).toBeLessThan(out!.indexOf('S0'))
    expect(out!.indexOf('S0')).toBeLessThan(out!.indexOf(VERIFY_BLOCK_MARKER))
  })

  it('does not register handlers for events it does not care about', async () => {
    const { api, captured } = makeHarness()
    createVerifyResponseExtension()(api)
    expect(captured.before_agent_start.length).toBe(1)
    expect(captured.session_start.length).toBe(0)
    expect(captured.tool_call.length).toBe(0)
  })
})
