import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { dirname, join } from 'node:path'
import { tmpdir } from 'node:os'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Chat-transcript compaction (issue #27 item 8): the JSONL grows without end
 * while loadChat only ever sees the newest MAX_CHAT_FILE_READ_BYTES
 * (readChatTail). Past twice that, an append rewrites the file to exactly its
 * read window — with the record the thread label comes from kept ahead,
 * because firstQuestionInFile labels a chat from the file's head.
 */

// the compaction's failure path needs a rename that fails on demand: the
// mock throws only for the one chat wired to fail; everything else passes
// through to the real fs
const renameFailures = vi.fn()
vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    renameSync: (from: unknown, to: unknown) => {
      if (String(to).endsWith('chat-err.jsonl')) {
        renameFailures()
        throw new Error('EIO: rename failed (simulated)')
      }
      return (actual.renameSync as (f: unknown, t: unknown) => void)(from, to)
    },
  }
})

import { ProjectStore } from '../src/store.js'

const FILLER_KB = 28
const OVER_THRESHOLD_RECORDS = 620
/** ~10 MB: past the 8 MB read window, under the 16 MB compaction threshold */
const MID_BAND_RECORDS = 360
const UNDER_THRESHOLD_RECORDS = 40

let tmpDir: string
let store: ProjectStore
let chatPath: string

function chatLine(seq: number, text: string): string {
  const ts = new Date(2026, 0, 1, 0, Math.floor(seq / 60), seq % 60).toISOString()
  return `${JSON.stringify({ seq, ts, role: 'assistant', text })}\n`
}

/** first user record + N assistant records with big texts */
function buildChat(records: number): string {
  const first = `${JSON.stringify({
    seq: 0,
    ts: new Date(2026, 0, 1).toISOString(),
    role: 'user',
    text: 'the first question ever',
  })}\n`
  const filler = 'x'.repeat(FILLER_KB * 1024)
  let out = first
  for (let i = 1; i <= records; i++) out += chatLine(i, `${filler} ${i}`)
  return out
}

beforeEach(() => {
  tmpDir = mkdtempSync(join(tmpdir(), 'project-store-compaction-'))
  store = new ProjectStore(tmpDir)
  store.ensureDefaultProject()
  chatPath = join(tmpDir, 'projects', 'default', 'chats', 'chat-1.jsonl')
  mkdirSync(dirname(chatPath), { recursive: true })
  renameFailures.mockReset()
})

afterEach(() => {
  rmSync(tmpDir, { recursive: true, force: true })
})

describe('chat transcript compaction (issue 27 item 8)', () => {
  it('leaves a small transcript byte-for-byte alone', () => {
    const original = buildChat(UNDER_THRESHOLD_RECORDS)
    writeFileSync(chatPath, original)
    store.appendChatMessage('default', 'chat-1', { role: 'assistant', text: 'added' })
    const after = readFileSync(chatPath, 'utf8')
    // nothing rewritten: the original bytes plus exactly one new record
    expect(after.startsWith(original)).toBe(true)
    const added = after.slice(original.length)
    expect(added).toContain('"text":"added"')
    expect(added).toMatch(/^\{.*\}\n$/)
  })

  it('a transcript between the read window and the threshold is left alone too', () => {
    // ~10 MB: past the 8 MB read window but under the 16 MB threshold. This
    // is the size band where an ungated compaction would show — the rewrite
    // would keep the head line and the newest 8 MB and drop the middle — so
    // the byte-for-byte check here is what pins the size gate itself
    const original = buildChat(MID_BAND_RECORDS)
    writeFileSync(chatPath, original)
    store.appendChatMessage('default', 'chat-1', { role: 'assistant', text: 'added' })
    const after = readFileSync(chatPath, 'utf8')
    expect(after.startsWith(original)).toBe(true)
    expect(statSync(chatPath).size).toBeGreaterThan(original.length)
    expect(
      readdirSync(join(tmpDir, 'projects', 'default', 'chats')).filter((f) => f.endsWith('.tmp')),
    ).toEqual([])
  })

  it('a rebind merges the compacted window, not the pre-compaction history', () => {
    // the rebind merge path reads whole files (mergeChatFiles), so a
    // compacted source carries its visible window into the target — the
    // documented compaction behavior change
    const original = buildChat(OVER_THRESHOLD_RECORDS)
    writeFileSync(chatPath, original)
    store.appendChatMessage('default', 'chat-1', { role: 'assistant', text: 'final marker' })
    // compacted to its read window, not the original ~17 MB
    expect(statSync(chatPath).size).toBeLessThan(9 * 1024 * 1024)

    const dstPath = join(tmpDir, 'projects', 'default', 'chats', 'chat-dst.jsonl')
    writeFileSync(dstPath, buildChat(UNDER_THRESHOLD_RECORDS))
    store.rebindChat('default', 'chat-1', 'chat-dst')

    // the source file is consumed; the target carries the compacted window
    // on top of its own records — far short of the 660 an uncompacted
    // source would have moved
    expect(existsSync(chatPath)).toBe(false)
    const merged = store.loadChat('default', 'chat-dst', 10_000)
    expect(merged.length).toBeGreaterThan(UNDER_THRESHOLD_RECORDS)
    expect(merged.length).toBeLessThan(UNDER_THRESHOLD_RECORDS + 400)
    expect(merged.map((m) => m.text)).toContain('the first question ever')
    expect(merged[merged.length - 1]).toMatchObject({ text: 'final marker' })
  })

  it('compacts past the threshold to the read window, keeping the head record', () => {
    // ~17 MB: past the 16 MB threshold, with a tail larger than the 8 MB
    // read window — the head record can never be inside the tail
    const original = buildChat(OVER_THRESHOLD_RECORDS)
    writeFileSync(chatPath, original)
    const firstLineBefore = original.slice(0, original.indexOf('\n'))
    // what the reader could see before compaction: the newest 200 records
    const before = store.loadChat('default', 'chat-1', 200)

    store.appendChatMessage('default', 'chat-1', { role: 'assistant', text: 'final marker' })

    const size = statSync(chatPath).size
    expect(size).toBeLessThan(9 * 1024 * 1024)
    // the head record survives at the top of the rewritten file
    const firstLineAfter = readFileSync(chatPath, 'utf8').slice(0, firstLineBefore.length + 1)
    expect(firstLineAfter).toBe(firstLineBefore + '\n')
    // nothing the reader could see is gone: the pre-append newest 199 are
    // still the newest 199, and the marker is the new tail
    const after = store.loadChat('default', 'chat-1', 200)
    expect(after).toHaveLength(200)
    expect(after[199]).toMatchObject({ text: 'final marker' })
    expect(after.slice(0, 199).map((m) => ({ seq: m.seq, text: m.text }))).toEqual(
      before.slice(1).map((m) => ({ seq: m.seq, text: m.text })),
    )
    // the atomic rewrite left no temp file behind
    expect(
      readdirSync(join(tmpDir, 'projects', 'default', 'chats')).filter((f) => f.endsWith('.tmp')),
    ).toEqual([])
  })

  it('keeps the record the label comes from when it is not the first line', () => {
    // firstQuestionInFile labels from the first line that parses to a user
    // turn — a torn or non-record first line means the label lives on a
    // later line, and keeping the literal first line would not save it
    const tornLine = 'this line is not a record\n'
    const original = buildChat(OVER_THRESHOLD_RECORDS)
    writeFileSync(chatPath, tornLine + original)
    const userLine = original.slice(0, original.indexOf('\n'))
    // the labeler reads the user record from line 2
    expect(store.listChats('default')[0]?.firstQuestion).toBe('the first question ever')

    store.appendChatMessage('default', 'chat-1', { role: 'assistant', text: 'final marker' })

    // the rewritten head is the user record, so the label survives
    const firstLineAfter = readFileSync(chatPath, 'utf8').slice(0, userLine.length + 1)
    expect(firstLineAfter).toBe(userLine + '\n')
    expect(store.listChats('default')[0]?.firstQuestion).toBe('the first question ever')
    expect(statSync(chatPath).size).toBeLessThan(9 * 1024 * 1024)
  })

  it('a failed compaction leaves the transcript appended but uncompacted', () => {
    chatPath = join(tmpDir, 'projects', 'default', 'chats', 'chat-err.jsonl')
    const original = buildChat(OVER_THRESHOLD_RECORDS)
    writeFileSync(chatPath, original)

    // the append must succeed despite the compaction failure — compaction
    // is best effort and must never lose the message that triggered it
    store.appendChatMessage('default', 'chat-err', { role: 'assistant', text: 'final marker' })

    expect(renameFailures).toHaveBeenCalled()
    const after = readFileSync(chatPath, 'utf8')
    // uncompacted: the original bytes plus the appended record, no rewrite
    expect(after.startsWith(original)).toBe(true)
    expect(after).toContain('"text":"final marker"')
    // the temp file of the failed rename is cleaned up
    expect(readdirSync(dirname(chatPath)).filter((f) => f.endsWith('.tmp'))).toEqual([])
  })

  it('seq continues from the compacted file after a fresh store instance', () => {
    const original = buildChat(OVER_THRESHOLD_RECORDS)
    writeFileSync(chatPath, original)
    store.appendChatMessage('default', 'chat-1', { role: 'assistant', text: 'final marker' })
    expect(statSync(chatPath).size).toBeLessThan(9 * 1024 * 1024)

    // a fresh instance has no seq counter cache: nextSeq initializes from
    // the file, so the compacted tail's max seq must survive the rewrite
    const reopened = new ProjectStore(tmpDir)
    reopened.appendChatMessage('default', 'chat-1', { role: 'assistant', text: 'after restart' })
    const newest = reopened.loadChat('default', 'chat-1', 2)
    expect(newest).toHaveLength(2)
    expect(newest[1]).toMatchObject({ text: 'after restart' })
    expect(newest[1]!.seq).toBe(newest[0]!.seq + 1)
    expect(newest[0]).toMatchObject({ text: 'final marker' })
    // the file stayed in the compacted size class, and the record just
    // appended is the last line on disk — a rewrite that dropped either
    // of the two would show up here
    expect(existsSync(chatPath)).toBe(true)
    const raw = readFileSync(chatPath, 'utf8')
    expect(statSync(chatPath).size).toBeLessThan(9 * 1024 * 1024)
    const lines = raw.trimEnd().split('\n')
    expect(lines[lines.length - 1]).toContain('"after restart"')
  })
})
