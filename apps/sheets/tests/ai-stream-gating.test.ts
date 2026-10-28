import { describe, expect, it } from 'vitest'
import type { AiChatMessage } from '../src/renderer/ai/AiChatPanel'
import { streamedTextPatch } from '../src/renderer/ai/stream-patch'

const pending: AiChatMessage = {
  role: 'assistant',
  text: '',
  tools: [],
  streaming: true,
}

describe('streamedTextPatch', () => {
  it('streams the live cumulative text when the pref is on', () => {
    const next = streamedTextPatch(pending, 'partial answer', true)
    expect(next.text).toBe('partial answer')
    expect(next.streaming).toBe(true)
  })

  it('clears a stale error flag on a successful retry even with streaming off', () => {
    const failed: AiChatMessage = { ...pending, text: 'apply failed', isError: true }
    const next = streamedTextPatch(failed, 'retrying', false)
    expect(next.isError).toBe(false)
  })

  it('holds the text back while streaming is off — the bubble keeps its pending state', () => {
    const next = streamedTextPatch(pending, 'partial answer', false)
    expect(next.text).toBe('')
    expect(next.streaming).toBe(true)
  })
})
