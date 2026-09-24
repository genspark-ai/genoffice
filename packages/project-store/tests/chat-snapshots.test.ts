import { existsSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { randomBytes } from 'node:crypto'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ProjectStore } from '../src/store'
import { MAX_CHAT_SNAPSHOTS, MAX_CHAT_SNAPSHOT_BYTES } from '../src/store'

/** A stand-in for the AI panel's document JSON: any serializable document */
const docJson = (marker: string) => JSON.stringify({ type: 'doc', marker })

describe('chat rollback-point snapshots', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-snapshots-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
    vi.restoreAllMocks()
  })

  const snapDir = (chatId: string) =>
    join(tmpDir, 'projects', 'default', 'chats', `${chatId}.snapshots`)

  it('round-trips a document through save and load', () => {
    const saved = store.saveChatSnapshot('default', 'chat1', docJson('before turn 1'))
    expect(saved.snapshotId).toBeTruthy()
    expect(saved.bytes).toBeGreaterThan(0)
    expect(store.loadChatSnapshot('default', 'chat1', saved.snapshotId!)).toBe(
      docJson('before turn 1'),
    )
  })

  it('stores the snapshot gzipped, beside the chat file', () => {
    const saved = store.saveChatSnapshot('default', 'chat1', docJson('compress me'))
    const file = join(snapDir('chat1'), `${saved.snapshotId}.json.gz`)
    expect(existsSync(file)).toBe(true)
    // the file on disk is gzip: reading it back needs gunzip, not JSON.parse
    const gz = readdirSync(snapDir('chat1'))
    expect(gz).toEqual([`${saved.snapshotId}.json.gz`])
  })

  it('keeps only the newest snapshots (ring of 20)', () => {
    const ids: string[] = []
    for (let i = 0; i < MAX_CHAT_SNAPSHOTS + 5; i++) {
      ids.push(store.saveChatSnapshot('default', 'chat1', docJson(`turn ${i}`))!.snapshotId ?? '')
    }
    const stored = store.listChatSnapshots('default', 'chat1')
    expect(stored).toHaveLength(MAX_CHAT_SNAPSHOTS)
    // the first five were evicted, the newest survived
    expect(stored).not.toContain(ids[0])
    expect(stored).toContain(ids.at(-1))
    expect(store.loadChatSnapshot('default', 'chat1', ids[0])).toBeNull()
  })

  it('refuses a snapshot past the size cap instead of writing it', () => {
    // random bytes do not compress, so this really does land over the cap
    const huge = JSON.stringify({
      type: 'doc',
      blob: randomBytes(5 * 1024 * 1024).toString('base64'),
    })
    expect(Buffer.byteLength(huge)).toBeGreaterThan(MAX_CHAT_SNAPSHOT_BYTES)
    const saved = store.saveChatSnapshot('default', 'chat1', huge)
    expect(saved.snapshotId).toBeNull()
    expect(saved.bytes).toBeGreaterThan(MAX_CHAT_SNAPSHOT_BYTES)
    expect(store.listChatSnapshots('default', 'chat1')).toEqual([])
    expect(existsSync(snapDir('chat1'))).toBe(false)
  })

  it('returns null for a snapshot that is gone, without throwing', () => {
    const saved = store.saveChatSnapshot('default', 'chat1', docJson('before turn 1'))!
    rmSync(join(snapDir('chat1'), `${saved.snapshotId}.json.gz`))
    expect(store.loadChatSnapshot('default', 'chat1', saved.snapshotId!)).toBeNull()
    expect(store.loadChatSnapshot('default', 'chat1', 'never-existed')).toBeNull()
  })

  it('lists nothing for a chat that never stored one', () => {
    expect(store.listChatSnapshots('default', 'no-such-chat')).toEqual([])
  })

  it('refuses traversal ids', () => {
    expect(() => store.saveChatSnapshot('default', '../evil', '{}')).toThrow(/Invalid chatId/)
    expect(() => store.loadChatSnapshot('default', 'chat1', '../evil')).toThrow(
      /Invalid snapshotId/,
    )
    expect(() => store.listChatSnapshots('default', '../evil')).toThrow(/Invalid chatId/)
  })

  it('survives a file that is not our gzip payload', () => {
    const saved = store.saveChatSnapshot('default', 'chat1', docJson('before turn 1'))!
    writeFileSync(join(snapDir('chat1'), `${saved.snapshotId}.json.gz`), 'not gzip at all', 'utf8')
    expect(store.loadChatSnapshot('default', 'chat1', saved.snapshotId!)).toBeNull()
  })

  it('reads back what a fresh store instance wrote (no in-memory state needed)', () => {
    const saved = store.saveChatSnapshot('default', 'chat1', docJson('durable'))!
    const reopened = new ProjectStore(tmpDir)
    expect(reopened.loadChatSnapshot('default', 'chat1', saved.snapshotId!)).toBe(
      docJson('durable'),
    )
    expect(reopened.listChatSnapshots('default', 'chat1')).toEqual([saved.snapshotId])
  })
})

describe('snapshots follow their chat', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-snapshots-move-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  const snapDir = (projectId: string, chatId: string) =>
    join(tmpDir, 'projects', projectId, 'chats', `${chatId}.snapshots`)

  it('moves with an unsaved session that first hits disk', () => {
    // the AI panel writes the snapshot while the chat is still "unsaved-…",
    // and the app rebinds the chat when the document is first saved
    store.appendChatMessage('default', 'unsaved-1', { role: 'user', text: 'rewrite this' })
    const saved = store.saveChatSnapshot('default', 'unsaved-1', docJson('before turn 1'))!
    expect(existsSync(join(snapDir('default', 'unsaved-1'), `${saved.snapshotId}.json.gz`))).toBe(
      true,
    )

    const rebound = store.rebindChatToFile('default', 'unsaved-1', '/tmp/report.docx')
    expect(store.loadChatSnapshot(rebound.projectId, rebound.chatId, saved.snapshotId!)).toBe(
      docJson('before turn 1'),
    )
    expect(existsSync(snapDir('default', 'unsaved-1'))).toBe(false)
  })

  it('moves with a chat when its file changes project', () => {
    const project = store.createProject('Second')
    const ids = store.resolveChatForFile('/tmp/notes.docx')
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'assistant', text: 'done' })
    const saved = store.saveChatSnapshot(ids.projectId, ids.chatId, docJson('before turn 1'))!

    store.moveFileToProject('/tmp/notes.docx', project.id)

    expect(store.loadChatSnapshot(project.id, ids.chatId, saved.snapshotId!)).toBe(
      docJson('before turn 1'),
    )
    expect(store.loadChatSnapshot(ids.projectId, ids.chatId, saved.snapshotId!)).toBeNull()
  })
})

describe('chat records carry the version ref', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-version-ref-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('persists the ref and hands it back on load', () => {
    const saved = store.saveChatSnapshot('default', 'chat1', docJson('before turn 1'))!
    store.appendChatMessage('default', 'chat1', {
      role: 'assistant',
      text: 'rewrote the intro',
      version: { id: 3, label: 'rewrite the intro', time: '10:04', snapshotId: saved.snapshotId! },
    })
    const [msg] = store.loadChat('default', 'chat1')
    expect(msg.version).toEqual({
      id: 3,
      label: 'rewrite the intro',
      time: '10:04',
      snapshotId: saved.snapshotId,
    })
  })

  it('omits the ref when the snapshot was too large to keep', () => {
    store.appendChatMessage('default', 'chat1', {
      role: 'assistant',
      text: 'rewrote the intro',
      version: { id: 1, label: 'a very large document', time: '10:04' },
    })
    const [msg] = store.loadChat('default', 'chat1')
    expect(msg.version).toEqual({ id: 1, label: 'a very large document', time: '10:04' })
    expect(msg.version?.snapshotId).toBeUndefined()
  })

  it('caps the label so one line cannot grow without bound', () => {
    store.appendChatMessage('default', 'chat1', {
      role: 'assistant',
      text: 'done',
      version: { id: 1, label: 'x'.repeat(5000), time: '10:04' },
    })
    expect(store.loadChat('default', 'chat1')[0].version?.label.length).toBe(200)
  })

  it('ignores a ref on messages written by an older build (no version field)', () => {
    store.appendChatMessage('default', 'chat1', { role: 'user', text: 'no version here' })
    expect(store.loadChat('default', 'chat1')[0].version).toBeUndefined()
  })

  it('reads a transcript whose version ref survived a merge', () => {
    const saved = store.saveChatSnapshot('default', 'unsaved-2', docJson('before turn 1'))!
    store.appendChatMessage('default', 'unsaved-2', {
      role: 'assistant',
      text: 'done',
      version: { id: 1, label: 'turn one', time: '09:00', snapshotId: saved.snapshotId! },
    })
    const rebound = store.rebindChatToFile('default', 'unsaved-2', '/tmp/merged.docx')
    const [msg] = store.loadChat(rebound.projectId, rebound.chatId)
    const snapshotId = msg.version?.snapshotId
    expect(snapshotId).toBeTruthy()
    // the ref still resolves to bytes after the move
    expect(store.loadChatSnapshot(rebound.projectId, rebound.chatId, snapshotId!)).toBe(
      docJson('before turn 1'),
    )
  })
})
