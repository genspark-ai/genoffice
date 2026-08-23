import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

// ────────────────────────────────────────────────────────────
// Test helpers
// ────────────────────────────────────────────────────────────

function makeTempDir(): string {
  return mkdtempSync(join(tmpdir(), 'project-store-test-'))
}

// ────────────────────────────────────────────────────────────
// 1. Default project creation
// ────────────────────────────────────────────────────────────

describe('ensureDefaultProject', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('creates the default project on first call', () => {
    const proj = store.ensureDefaultProject()
    expect(proj.id).toBe('default')
    expect(proj.name).toBe('Default Project')
    expect(typeof proj.createdAt).toBe('string')
  })

  it('is idempotent: repeated calls return the same project without errors', () => {
    const first = store.ensureDefaultProject()
    const second = store.ensureDefaultProject()
    expect(second.id).toBe(first.id)
    expect(second.createdAt).toBe(first.createdAt)
  })

  it('default project is discoverable via index.json', () => {
    store.ensureDefaultProject()
    const projects = store.listProjects()
    expect(projects.find((p) => p.id === 'default')).toBeTruthy()
  })
})

// ────────────────────────────────────────────────────────────
// 2. resolveProjectForFile
// ────────────────────────────────────────────────────────────

describe('resolveProjectForFile', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('unknown files are assigned to default', () => {
    const projectId = store.resolveProjectForFile('/fake/path/doc.docx')
    expect(projectId).toBe('default')
  })

  it('querying the same path twice returns the same projectId (cached in fileMap)', () => {
    const p1 = store.resolveProjectForFile('/fake/path/doc.docx')
    const p2 = store.resolveProjectForFile('/fake/path/doc.docx')
    expect(p1).toBe(p2)
  })

  it('records the file in the default project files list', () => {
    store.resolveProjectForFile('/fake/test.xlsx')
    const proj = store.getProject('default')
    expect(proj?.files).toContain('/fake/test.xlsx')
  })
})

// ────────────────────────────────────────────────────────────
// 3. appendChatMessage + loadChat
// ────────────────────────────────────────────────────────────

describe('appendChatMessage + loadChat', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('appended messages can be read back', () => {
    store.appendChatMessage('default', 'chat1', { role: 'user', text: 'hello' })
    const msgs = store.loadChat('default', 'chat1')
    expect(msgs).toHaveLength(1)
    expect(msgs[0].role).toBe('user')
    expect(msgs[0].text).toBe('hello')
  })

  it('seq increases monotonically', () => {
    store.appendChatMessage('default', 'chat1', { role: 'user', text: 'a' })
    store.appendChatMessage('default', 'chat1', { role: 'assistant', text: 'b' })
    store.appendChatMessage('default', 'chat1', { role: 'user', text: 'c' })
    const msgs = store.loadChat('default', 'chat1')
    expect(msgs[0].seq).toBe(0)
    expect(msgs[1].seq).toBe(1)
    expect(msgs[2].seq).toBe(2)
  })

  it('ts is a valid UTC ISO string', () => {
    store.appendChatMessage('default', 'chat1', { role: 'user', text: 'hi' })
    const msgs = store.loadChat('default', 'chat1')
    expect(() => new Date(msgs[0].ts).toISOString()).not.toThrow()
  })

  it('tools field is persisted correctly', () => {
    store.appendChatMessage('default', 'chat1', {
      role: 'assistant',
      text: 'done',
      tools: [{ name: 'write_file', summary: 'wrote 3 lines', isError: false }],
    })
    const msgs = store.loadChat('default', 'chat1')
    expect(msgs[0].tools?.[0].name).toBe('write_file')
  })

  it('limit parameter: returns only the most recent N messages', () => {
    for (let i = 0; i < 10; i++) {
      store.appendChatMessage('default', 'chatLimit', { role: 'user', text: `msg${i}` })
    }
    const msgs = store.loadChat('default', 'chatLimit', 3)
    expect(msgs).toHaveLength(3)
    expect(msgs[0].text).toBe('msg7')
    expect(msgs[2].text).toBe('msg9')
  })

  it('nonexistent chat returns an empty array', () => {
    const msgs = store.loadChat('default', 'nonexistent')
    expect(msgs).toEqual([])
  })

  it('empty text can be written and read back', () => {
    store.appendChatMessage('default', 'chat1', { role: 'assistant', text: '' })
    const msgs = store.loadChat('default', 'chat1')
    expect(msgs[0].text).toBe('')
  })

  it('a runaway reply is truncated instead of stored whole', () => {
    // A model stuck in a repetition loop; stored whole it would also be replayed
    // into the model context when the file reopens
    store.appendChatMessage('default', 'chatHuge', {
      role: 'assistant',
      text: 'shame '.repeat(20_000),
    })
    const msgs = store.loadChat('default', 'chatHuge')
    expect(msgs[0].text.length).toBeLessThan(33_000)
    expect(msgs[0].text.endsWith('[truncated]')).toBe(true)
  })

  it('text just under the cap is stored verbatim', () => {
    const text = 'x'.repeat(31_999)
    store.appendChatMessage('default', 'chatUnder', { role: 'assistant', text })
    const msgs = store.loadChat('default', 'chatUnder')
    expect(msgs[0].text).toBe(text)
  })
})

// ────────────────────────────────────────────────────────────
// 4. Tolerance for corrupted JSONL lines
// ────────────────────────────────────────────────────────────

describe('JSONL corrupted-line tolerance', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('skips bad lines and reads the rest normally', () => {
    // Write normal messages first
    store.appendChatMessage('default', 'chatCorrupt', { role: 'user', text: 'good1' })
    store.appendChatMessage('default', 'chatCorrupt', { role: 'assistant', text: 'good2' })

    // Manually insert a corrupted line into the JSONL
    const chatPath = join(tmpDir, 'projects', 'default', 'chats', 'chatCorrupt.jsonl')
    writeFileSync(chatPath, readFileSyncHelper(chatPath) + '{ invalid json line \n', 'utf8')

    // Append another normal message (new store instance; the seq counter rebuilds from the file)
    const store2 = new ProjectStore(tmpDir)
    store2.appendChatMessage('default', 'chatCorrupt', { role: 'user', text: 'good3' })

    const msgs = store2.loadChat('default', 'chatCorrupt')
    expect(msgs.length).toBe(3) // the bad line was skipped
    expect(msgs.map((m) => m.text)).toEqual(['good1', 'good2', 'good3'])
  })

  it('all bad lines: returns an empty array without crashing', () => {
    const chatPath = join(tmpDir, 'projects', 'default', 'chats', 'badChat.jsonl')
    const dir = join(tmpDir, 'projects', 'default', 'chats')
    mkdirSyncHelper(dir)
    writeFileSync(chatPath, 'not json\nalso bad\n{incomplete\n', 'utf8')

    const msgs = store.loadChat('default', 'badChat')
    expect(msgs).toEqual([])
  })
})

// ────────────────────────────────────────────────────────────
// 5. rebindChat
// ────────────────────────────────────────────────────────────

describe('rebindChat', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('after rebind the old chatId has no history and the new chatId does', () => {
    store.appendChatMessage('default', 'unsaved-1234', { role: 'user', text: 'hello' })
    store.appendChatMessage('default', 'unsaved-1234', { role: 'assistant', text: 'world' })

    const newChatId = ProjectStore.chatIdForFile('/home/user/doc.docx')
    store.rebindChat('default', 'unsaved-1234', newChatId)

    const oldMsgs = store.loadChat('default', 'unsaved-1234')
    expect(oldMsgs).toHaveLength(0)

    const newMsgs = store.loadChat('default', newChatId)
    expect(newMsgs).toHaveLength(2)
  })

  it('appending after rebind continues seq from the correct position', () => {
    store.appendChatMessage('default', 'unsaved-5678', { role: 'user', text: 'msg0' })
    store.appendChatMessage('default', 'unsaved-5678', { role: 'assistant', text: 'msg1' })

    const newId = 'aabbccdd11223344'
    store.rebindChat('default', 'unsaved-5678', newId)
    store.appendChatMessage('default', newId, { role: 'user', text: 'msg2' })

    const msgs = store.loadChat('default', newId)
    expect(msgs).toHaveLength(3)
    expect(msgs[2].seq).toBe(2) // seq starts at 0 and continues after rebind
  })

  it('tempId === newChatId: no-op without errors', () => {
    store.appendChatMessage('default', 'sameId', { role: 'user', text: 'hi' })
    expect(() => store.rebindChat('default', 'sameId', 'sameId')).not.toThrow()
    expect(store.loadChat('default', 'sameId')).toHaveLength(1)
  })

  it('does not throw when the source file does not exist', () => {
    expect(() => store.rebindChat('default', 'ghost-id', 'new-id')).not.toThrow()
  })

  it('merges when the target file exists: old history kept, temp messages re-sequenced at the end', () => {
    const newId = ProjectStore.chatIdForFile('/home/user/existing.docx')
    store.appendChatMessage('default', newId, { role: 'user', text: 'old-0' })
    store.appendChatMessage('default', newId, { role: 'assistant', text: 'old-1' })
    store.appendChatMessage('default', 'unsaved-merge', { role: 'user', text: 'new-0' })

    store.rebindChat('default', 'unsaved-merge', newId)

    const msgs = store.loadChat('default', newId)
    expect(msgs.map((m) => m.text)).toEqual(['old-0', 'old-1', 'new-0'])
    expect(msgs.map((m) => m.seq)).toEqual([0, 1, 2])
    expect(store.loadChat('default', 'unsaved-merge')).toHaveLength(0)

    store.appendChatMessage('default', newId, { role: 'assistant', text: 'after' })
    expect(store.loadChat('default', newId).map((m) => m.seq)).toEqual([0, 1, 2, 3])
  })
})

// ────────────────────────────────────────────────────────────
// 5b. rebindChatToFile
// ────────────────────────────────────────────────────────────

describe('rebindChatToFile', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('computes chatId from path, registers fileMap, returns new ids, and history is recoverable by path', () => {
    store.appendChatMessage('default', 'unsaved-9999', { role: 'user', text: 'generate a deck' })
    store.appendChatMessage('default', 'unsaved-9999', { role: 'assistant', text: 'generated' })

    const draftPath = '/Users/x/Documents/GenOffice/AI Strategy.pptx'
    const ids = store.rebindChatToFile('default', 'unsaved-9999', draftPath)

    expect(ids.chatId).toBe(ProjectStore.chatIdForFile(draftPath))
    // Simulate reopening: resolving by path yields the same projectId/chatId with full history
    expect(store.resolveProjectForFile(draftPath)).toBe(ids.projectId)
    const msgs = store.loadChat(ids.projectId, ids.chatId)
    expect(msgs.map((m) => m.text)).toEqual(['generate a deck', 'generated'])
  })

  it('moves history into the owning project when the file already belongs to another project', () => {
    const proj = store.createProject('Research Project')
    const filePath = '/Users/x/Documents/GenOffice/Beijing Research.pptx'
    store.moveFileToProject(filePath, proj.id)

    store.appendChatMessage('default', 'unsaved-7777', { role: 'user', text: 'hi' })
    store.appendChatMessage('default', 'unsaved-7777', { role: 'assistant', text: 'ok' })
    const ids = store.rebindChatToFile('default', 'unsaved-7777', filePath)

    expect(ids.projectId).toBe(proj.id)
    expect(store.loadChat(proj.id, ids.chatId)).toHaveLength(2)
    expect(store.loadChat('default', ids.chatId)).toHaveLength(0)
  })
})

// ────────────────────────────────────────────────────────────
// 5c. Stable chatId mapping and rename tracking
// ────────────────────────────────────────────────────────────

describe('resolveChatForFile / fileRenamed', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('chatId stays the same after rename and history follows the file', () => {
    const oldPath = '/Users/x/Documents/GenOffice/old-name.pptx'
    const ids = store.resolveChatForFile(oldPath)
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'assistant', text: 'a' })

    const newPath = '/Users/x/Documents/GenOffice/new-name.pptx'
    store.fileRenamed(oldPath, newPath)

    const after = store.resolveChatForFile(newPath)
    expect(after.chatId).toBe(ids.chatId)
    expect(store.loadChat(after.projectId, after.chatId)).toHaveLength(2)
  })

  it('legacy data without mapping (path-hash chatId) can still recover history after rename', () => {
    const oldPath = '/Users/x/Documents/GenOffice/legacy.docx'
    const legacyChatId = ProjectStore.chatIdForFile(oldPath)
    store.resolveProjectForFile(oldPath) // registers only fileMap, not chatIdByPath (simulates old data)
    store.appendChatMessage('default', legacyChatId, { role: 'user', text: 'q' })
    store.appendChatMessage('default', legacyChatId, { role: 'assistant', text: 'a' })

    const newPath = '/Users/x/Documents/GenOffice/legacy-renamed.docx'
    store.fileRenamed(oldPath, newPath)

    const after = store.resolveChatForFile(newPath)
    expect(after.chatId).toBe(legacyChatId)
    expect(store.loadChat(after.projectId, after.chatId)).toHaveLength(2)
  })

  it('rename after rebindChatToFile still resolves to the same history', () => {
    store.appendChatMessage('default', 'unsaved-1', { role: 'user', text: 'q' })
    store.appendChatMessage('default', 'unsaved-1', { role: 'assistant', text: 'a' })
    const draft = '/Users/x/Documents/GenOffice/draft.pptx'
    const ids = store.rebindChatToFile('default', 'unsaved-1', draft)

    const renamed = '/Users/x/Documents/GenOffice/final-name.pptx'
    store.fileRenamed(draft, renamed)
    const after = store.resolveChatForFile(renamed)
    expect(after.chatId).toBe(ids.chatId)
    expect(store.loadChat(after.projectId, after.chatId).map((m) => m.text)).toEqual(['q', 'a'])
  })
})

// ────────────────────────────────────────────────────────────
// 5d. Buffering before the first assistant message (empty sessions create no file)
// ────────────────────────────────────────────────────────────

describe('appendChatMessage opening buffer', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  const chatFile = (chatId: string) =>
    join(tmpDir, 'projects', 'default', 'chats', `${chatId}.jsonl`)

  it('no file is created with only user messages, but loadChat still reads the buffer', () => {
    store.appendChatMessage('default', 'unsaved-buf', {
      role: 'user',
      text: 'question without answer',
    })
    expect(existsSync(chatFile('unsaved-buf'))).toBe(false)
    expect(store.loadChat('default', 'unsaved-buf').map((m) => m.text)).toEqual([
      'question without answer',
    ])
  })

  it('creates the file when the first assistant message arrives, flushing the buffer with continuous seq', () => {
    store.appendChatMessage('default', 'unsaved-buf2', { role: 'user', text: 'q' })
    store.appendChatMessage('default', 'unsaved-buf2', { role: 'assistant', text: 'a' })
    expect(existsSync(chatFile('unsaved-buf2'))).toBe(true)
    const msgs = store.loadChat('default', 'unsaved-buf2')
    expect(msgs.map((m) => m.text)).toEqual(['q', 'a'])
    expect(msgs.map((m) => m.seq)).toEqual([0, 1])
  })

  it('buffer is materialized on rebind: user-only messages are kept with the file', () => {
    store.appendChatMessage('default', 'unsaved-buf3', {
      role: 'user',
      text: 'generation interrupted',
    })
    const ids = store.rebindChatToFile(
      'default',
      'unsaved-buf3',
      '/Users/x/Documents/GenOffice/interrupted.pptx',
    )
    expect(store.loadChat(ids.projectId, ids.chatId).map((m) => m.text)).toEqual([
      'generation interrupted',
    ])
  })

  it('oversized tool input/output is truncated to 16k on disk; attachment metadata is preserved', () => {
    const big = 'x'.repeat(20_000)
    store.appendChatMessage('default', 'tools-chat', {
      role: 'user',
      text: 'q',
      attachments: [{ name: 'asset.pdf', path: '/tmp/asset.pdf', ext: 'pdf', sizeBytes: 123 }],
    })
    store.appendChatMessage('default', 'tools-chat', {
      role: 'assistant',
      text: 'a',
      tools: [{ name: 'read_slide', summary: 'read page 1', input: big, output: big }],
    })
    const msgs = store.loadChat('default', 'tools-chat')
    expect(msgs[0].attachments).toEqual([
      { name: 'asset.pdf', path: '/tmp/asset.pdf', ext: 'pdf', sizeBytes: 123 },
    ])
    expect(msgs[1].tools?.[0].input).toHaveLength(16_000)
    expect(msgs[1].tools?.[0].output).toHaveLength(16_000)
    expect(msgs[1].tools?.[0].summary).toBe('read page 1')
  })

  it('user messages appended to a chat with an existing file are written directly, not buffered', () => {
    store.appendChatMessage('default', 'has-file', { role: 'user', text: 'q1' })
    store.appendChatMessage('default', 'has-file', { role: 'assistant', text: 'a1' })
    store.appendChatMessage('default', 'has-file', { role: 'user', text: 'q2' })
    const raw = store.loadChat('default', 'has-file')
    expect(raw.map((m) => m.text)).toEqual(['q1', 'a1', 'q2'])
    expect(existsSync(chatFile('has-file'))).toBe(true)
  })
})

// ────────────────────────────────────────────────────────────
// 6. listChats
// ────────────────────────────────────────────────────────────

describe('listChats', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('returns the list of existing chats', () => {
    store.appendChatMessage('default', 'chat-a', { role: 'user', text: 'a' })
    store.appendChatMessage('default', 'chat-a', { role: 'assistant', text: 'ra' })
    store.appendChatMessage('default', 'chat-b', { role: 'user', text: 'b' })
    store.appendChatMessage('default', 'chat-b', { role: 'assistant', text: 'rb' })
    const list = store.listChats('default')
    const ids = list.map((c) => c.chatId).sort()
    expect(ids).toEqual(['chat-a', 'chat-b'])
  })

  it('returns an empty array when the chats directory does not exist', () => {
    expect(store.listChats('nonexistent-project')).toEqual([])
  })
})

// ────────────────────────────────────────────────────────────
// 7. chatIdForFile
// ────────────────────────────────────────────────────────────

describe('chatIdForFile', () => {
  it('same path returns the same id', () => {
    const a = ProjectStore.chatIdForFile('/home/user/test.pptx')
    const b = ProjectStore.chatIdForFile('/home/user/test.pptx')
    expect(a).toBe(b)
  })

  it('different paths return different ids', () => {
    const a = ProjectStore.chatIdForFile('/home/user/a.pptx')
    const b = ProjectStore.chatIdForFile('/home/user/b.pptx')
    expect(a).not.toBe(b)
  })

  it('returns a 16-character hex string', () => {
    const id = ProjectStore.chatIdForFile('/any/path.xlsx')
    expect(id).toHaveLength(16)
    expect(/^[0-9a-f]{16}$/.test(id)).toBe(true)
  })
})

// ────────────────────────────────────────────────────────────
// Helpers (Node built-in wrappers, avoiding import cycles)
// ────────────────────────────────────────────────────────────

import { readFileSync as _readFileSync, mkdirSync as _mkdirSync } from 'node:fs'

function readFileSyncHelper(p: string): string {
  try {
    return _readFileSync(p, 'utf8')
  } catch {
    return ''
  }
}

// ────────────────────────────────────────────────────────────
// 8. P1: createProject / renameProject / deleteProject
// ────────────────────────────────────────────────────────────

describe('createProject', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('creates a project and shows it in listProjects', () => {
    const proj = store.createProject('Research Project')
    expect(proj.name).toBe('Research Project')
    expect(proj.id).toMatch(/^proj-/)
    const list = store.listProjects()
    expect(list.find((p) => p.id === proj.id)).toBeTruthy()
  })

  it('throws on an empty name', () => {
    expect(() => store.createProject('   ')).toThrow()
  })

  it('multiple creations produce different ids', () => {
    const a = store.createProject('Project A')
    const b = store.createProject('Project B')
    expect(a.id).not.toBe(b.id)
  })
})

describe('renameProject', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('renames successfully', () => {
    const proj = store.createProject('Old Name')
    store.renameProject(proj.id, 'New Name')
    const list = store.listProjects()
    const found = list.find((p) => p.id === proj.id)
    expect(found?.name).toBe('New Name')
  })

  it('cannot rename the default project', () => {
    expect(() => store.renameProject('default', 'Renamed')).toThrow()
  })

  it('throws on an empty name', () => {
    const proj = store.createProject('Test')
    expect(() => store.renameProject(proj.id, '')).toThrow()
  })
})

describe('deleteProject', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('deleted project disappears from listProjects', () => {
    const proj = store.createProject('Temp Project')
    store.deleteProject(proj.id)
    const list = store.listProjects()
    expect(list.find((p) => p.id === proj.id)).toBeUndefined()
  })

  it('soft delete: directory is moved into .trash', () => {
    const proj = store.createProject('Soft Delete Test')
    const projDir = join(tmpDir, 'projects', proj.id)
    store.deleteProject(proj.id)
    // The original directory no longer exists
    expect(existsSync(projDir)).toBe(false)
    // The .trash directory exists
    const trashDir = join(tmpDir, 'projects', '.trash')
    expect(existsSync(trashDir)).toBe(true)
  })

  it('files of a deleted project revert to default', () => {
    const proj = store.createProject('Migration Project')
    // Assign the file to that project first
    store.resolveProjectForFile('/test/file.docx') // goes to default
    store.moveFileToProject('/test/file.docx', proj.id)
    store.deleteProject(proj.id)
    // fileMap should point to default
    const _resolved = store.resolveProjectForFile('/test/file.docx')
    // resolveProjectForFile doesn't rewrite existing mappings, so check the index first
    const index = (store as any).readIndex()
    expect(index.fileMap['/test/file.docx']).toBe('default')
  })

  it('cannot delete the default project', () => {
    expect(() => store.deleteProject('default')).toThrow()
  })
})

// ────────────────────────────────────────────────────────────
// 9. P1: moveFileToProject
// ────────────────────────────────────────────────────────────

describe('moveFileToProject', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('file belongs to the target project after moving', () => {
    const proj = store.createProject('Target Project')
    store.resolveProjectForFile('/docs/a.docx') // assigned to default
    store.moveFileToProject('/docs/a.docx', proj.id)
    const index = (store as any).readIndex()
    expect(index.fileMap['/docs/a.docx']).toBe(proj.id)
    const targetProj = store.getProject(proj.id)
    expect(targetProj?.files).toContain('/docs/a.docx')
  })

  it('original project no longer contains the file after moving', () => {
    const proj = store.createProject('Target Project 2')
    store.resolveProjectForFile('/docs/b.docx')
    store.moveFileToProject('/docs/b.docx', proj.id)
    const defaultProj = store.getProject('default')
    expect(defaultProj?.files).not.toContain('/docs/b.docx')
  })

  it('moves the chat jsonl file when history exists', () => {
    const proj = store.createProject('Target Project 3')
    store.resolveProjectForFile('/docs/c.docx')
    const chatId = ProjectStore.chatIdForFile('/docs/c.docx')
    store.appendChatMessage('default', chatId, { role: 'user', text: 'test message' })
    store.moveFileToProject('/docs/c.docx', proj.id)
    // The new project can read the messages
    const msgs = store.loadChat(proj.id, chatId)
    expect(msgs).toHaveLength(1)
    expect(msgs[0].text).toBe('test message')
    // The old project cannot
    const oldMsgs = store.loadChat('default', chatId)
    expect(oldMsgs).toHaveLength(0)
  })

  it('moving to the same project does not throw', () => {
    store.resolveProjectForFile('/docs/d.docx')
    expect(() => store.moveFileToProject('/docs/d.docx', 'default')).not.toThrow()
  })

  it('throws when the target project does not exist', () => {
    store.resolveProjectForFile('/docs/e.docx')
    expect(() => store.moveFileToProject('/docs/e.docx', 'nonexistent')).toThrow()
  })
})

// ────────────────────────────────────────────────────────────
// 10. P1: listProjectsSummary
// ────────────────────────────────────────────────────────────

describe('listProjectsSummary', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('default project is flagged isDefault=true', () => {
    const summaries = store.listProjectsSummary()
    const def = summaries.find((s) => s.id === 'default')
    expect(def?.isDefault).toBe(true)
  })

  it('fileCount reflects the number of owned files', () => {
    const docPath = join(tmpDir, 'x.docx')
    const sheetPath = join(tmpDir, 'y.xlsx')
    writeFileSync(docPath, 'doc')
    writeFileSync(sheetPath, 'sheet')
    store.resolveProjectForFile(docPath)
    store.resolveProjectForFile(sheetPath)
    const summaries = store.listProjectsSummary()
    const def = summaries.find((s) => s.id === 'default')
    expect(def?.fileCount).toBe(2)
  })

  it('fileCount and project files exclude stale and duplicate paths', () => {
    const existingPath = join(tmpDir, 'existing.docx')
    const stalePath = join(tmpDir, 'deleted.docx')
    writeFileSync(existingPath, 'doc')
    writeFileSync(stalePath, 'deleted')
    store.resolveProjectForFile(existingPath)
    store.resolveProjectForFile(stalePath)
    rmSync(stalePath)

    const project = store.getProject('default')
    project!.files.push(existingPath)
    writeFileSync(
      join(tmpDir, 'projects', 'default', 'project.json'),
      JSON.stringify(project, null, 2),
      'utf8',
    )

    expect(store.listProjectFiles('default')).toEqual([existingPath])
    const summary = store.listProjectsSummary().find((item) => item.id === 'default')
    expect(summary?.fileCount).toBe(1)
  })

  it('newly created project appears in the summary', () => {
    const proj = store.createProject('Summary Test')
    const summaries = store.listProjectsSummary()
    expect(summaries.find((s) => s.id === proj.id)).toBeTruthy()
  })
})

// ────────────────────────────────────────────────────────────
// 11. P1: getProjectTimeline
// ────────────────────────────────────────────────────────────

describe('getProjectTimeline', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('empty project returns an empty array', () => {
    const proj = store.createProject('Empty Project')
    expect(store.getProjectTimeline(proj.id)).toEqual([])
  })

  it('returns timeline entries when messages exist', () => {
    store.resolveProjectForFile('/docs/tl.docx')
    const chatId = ProjectStore.chatIdForFile('/docs/tl.docx')
    store.appendChatMessage('default', chatId, { role: 'user', text: 'hello world' })
    store.appendChatMessage('default', chatId, { role: 'assistant', text: 'hello!' })
    const timeline = store.getProjectTimeline('default')
    expect(timeline.length).toBeGreaterThanOrEqual(2)
    // preview is capped at 120 chars
    timeline.forEach((e) => expect(e.preview.length).toBeLessThanOrEqual(120))
  })

  it('limit parameter is honored', () => {
    store.resolveProjectForFile('/docs/limit.docx')
    const chatId = ProjectStore.chatIdForFile('/docs/limit.docx')
    for (let i = 0; i < 15; i++) {
      store.appendChatMessage('default', chatId, { role: 'user', text: `message${i}` })
      store.appendChatMessage('default', chatId, { role: 'assistant', text: `reply${i}` })
    }
    const timeline = store.getProjectTimeline('default', 10)
    expect(timeline).toHaveLength(10)
  })

  it('timeline is sorted by ts in descending order', () => {
    store.resolveProjectForFile('/docs/order.docx')
    const chatId = ProjectStore.chatIdForFile('/docs/order.docx')
    store.appendChatMessage('default', chatId, { role: 'user', text: 'early' })
    store.appendChatMessage('default', chatId, { role: 'user', text: 'late' })
    const timeline = store.getProjectTimeline('default')
    if (timeline.length >= 2) {
      expect(timeline[0].ts >= timeline[1].ts).toBe(true)
    }
  })

  it('fileName is the basename of the file', () => {
    store.resolveProjectForFile('/some/path/myFile.docx')
    const chatId = ProjectStore.chatIdForFile('/some/path/myFile.docx')
    store.appendChatMessage('default', chatId, { role: 'user', text: 'content' })
    store.appendChatMessage('default', chatId, { role: 'assistant', text: 'ok' })
    const timeline = store.getProjectTimeline('default')
    const entry = timeline.find((e) => e.chatId === chatId)
    expect(entry?.fileName).toBe('myFile.docx')
  })
})

function mkdirSyncHelper(p: string): void {
  try {
    _mkdirSync(p, { recursive: true })
  } catch {
    /* ignore */
  }
}
