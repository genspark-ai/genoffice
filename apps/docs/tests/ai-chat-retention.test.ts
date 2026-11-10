// Every completed AI turn pins a full-document JSON snapshot on its chat entry
// and renders an inline roll-back action from it. Nothing released them, so a
// 122-page document over 20 turns of conversation held hundreds of MB in React
// state forever. The transcript itself must stay bounded too.
import { describe, expect, it } from 'vitest'
import {
  MAX_CHAT_ENTRIES,
  MAX_ROLLBACK_SNAPSHOTS,
  boundChatHistory,
} from '../src/renderer/ai/chat-retention'

interface Entry {
  role: 'user' | 'assistant'
  text: string
  streaming?: boolean
  snapshot?: { type: string; payload: string }
  tools?: Array<{ name: string; summary: string }>
}

/** one turn: a user entry plus the assistant entry that carries the snapshot */
function turn(i: number, docBytes = 8): Entry[] {
  return [
    { role: 'user', text: `question ${i}` },
    {
      role: 'assistant',
      text: `answer ${i}`,
      // stands in for the ProseMirror JSON of a long document
      snapshot: { type: 'doc', payload: 'x'.repeat(docBytes) },
    },
  ]
}

function transcript(turns: number, docBytes = 8): Entry[] {
  return Array.from({ length: turns }, (_, i) => turn(i, docBytes)).flat()
}

const withSnapshots = (entries: Entry[]) => entries.filter((e) => e.snapshot).length

describe('AI chat retention', () => {
  it('releases every snapshot past the rollback cap', () => {
    const bounded = boundChatHistory(transcript(20))
    expect(withSnapshots(bounded)).toBeLessThanOrEqual(MAX_ROLLBACK_SNAPSHOTS)
  })

  it('keeps the rollback points of the most recent runs', () => {
    const bounded = boundChatHistory(transcript(20))
    const kept = bounded.filter((e) => e.snapshot).map((e) => e.text)
    expect(kept).toEqual(
      Array.from(
        { length: MAX_ROLLBACK_SNAPSHOTS },
        (_, i) => `answer ${20 - MAX_ROLLBACK_SNAPSHOTS + i}`,
      ),
    )
  })

  it('drops a snapshot as undefined, so the roll-back action cannot be offered', () => {
    // the panel renders RollbackButton under `entry.snapshot &&`; a dropped
    // snapshot must therefore be undefined, never a stale placeholder object
    const bounded = boundChatHistory(transcript(20))
    const dropped = bounded.find((e) => e.text === 'answer 0')
    expect(dropped).toBeDefined()
    expect(dropped!.snapshot).toBeUndefined()
    expect(dropped!.text).toBe('answer 0') // the transcript itself survives
  })

  it('keeps the text of every turn whose snapshot was released', () => {
    const bounded = boundChatHistory(transcript(20))
    expect(bounded.filter((e) => e.role === 'user').map((e) => e.text)).toHaveLength(20)
    expect(bounded.filter((e) => e.role === 'assistant').map((e) => e.text)).toContain('answer 0')
  })

  it('caps the number of retained entries', () => {
    const bounded = boundChatHistory(transcript(400))
    expect(bounded.length).toBeLessThanOrEqual(MAX_CHAT_ENTRIES)
    // the newest entries are the ones kept
    expect(bounded.at(-1)!.text).toBe('answer 399')
  })

  it('never releases the newest snapshot or drops the live turn', () => {
    const live: Entry[] = [
      ...transcript(3),
      { role: 'user', text: 'live question' },
      { role: 'assistant', text: '', streaming: true },
    ]
    const bounded = boundChatHistory(live)
    expect(withSnapshots(bounded)).toBeLessThanOrEqual(MAX_ROLLBACK_SNAPSHOTS)
    expect(bounded.at(-1)).toEqual({ role: 'assistant', text: '', streaming: true })
    expect(bounded.some((e) => e.text === 'live question')).toBe(true)
  })

  it('bounds the retained snapshot payload, not just the entry count', () => {
    // the reported leak: 20 turns of a 122-page document's JSON
    const before = withSnapshots(transcript(20, 10_000_000))
    const after = withSnapshots(boundChatHistory(transcript(20, 10_000_000)))
    expect(before).toBe(20)
    expect(after).toBeLessThanOrEqual(MAX_ROLLBACK_SNAPSHOTS)
  })

  it('returns the same array when nothing needs trimming', () => {
    const entries = transcript(2)
    // identity matters: a fresh array on every delta would re-render the log
    expect(boundChatHistory(entries)).toBe(entries)
  })

  it('is idempotent, because several append paths bound the same array in turn', () => {
    const once = boundChatHistory(transcript(20))
    const twice = boundChatHistory(once)
    expect(twice).toBe(once)
    expect(withSnapshots(twice)).toBe(MAX_ROLLBACK_SNAPSHOTS)
  })

  it('releases only the snapshot, keeping the rest of the turn', () => {
    const entries: Entry[] = [
      {
        role: 'assistant',
        text: 'early',
        streaming: false,
        snapshot: { type: 'doc', payload: 'x'.repeat(64) },
        tools: [{ name: 'replace', summary: 'edited 3 spots' }],
      },
      // enough later runs to push this entry past the rollback cap
      ...transcript(6),
    ]
    const bounded = boundChatHistory(entries)
    const early = bounded[0]!
    expect(early.snapshot).toBeUndefined()
    expect(early.text).toBe('early')
    expect(early.tools).toEqual([{ name: 'replace', summary: 'edited 3 spots' }])
  })
})
