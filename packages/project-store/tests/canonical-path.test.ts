import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { join, sep } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore, canonicalPathKey } from '../src/store.js'

function withDotSegment(filePath: string): string {
  const cut = filePath.lastIndexOf(sep)
  return `${filePath.slice(0, cut)}${sep}.${sep}${filePath.slice(cut + 1)}`
}

function readIndex(tmpDir: string): {
  fileMap: Record<string, string>
  chatIdByPath?: Record<string, string>
} {
  return JSON.parse(readFileSync(join(tmpDir, 'projects', 'index.json'), 'utf8'))
}

describe('canonical path identity', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-canonical-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  describe('canonicalPathKey', () => {
    it('resolves an existing file to its real path', () => {
      const filePath = join(tmpDir, 'exists.docx')
      writeFileSync(filePath, 'doc', 'utf8')
      expect(canonicalPathKey(filePath)).toBe(canonicalPathKey(withDotSegment(filePath)))
      expect(canonicalPathKey(filePath)).not.toBe(withDotSegment(filePath))
    })

    it('keys a missing file under its resolved directory', () => {
      const missing = join(tmpDir, 'missing.docx')
      expect(canonicalPathKey(missing)).toBe(join(canonicalPathKey(tmpDir), 'missing.docx'))
    })

    it('uses the given path when neither it nor its directory resolve', () => {
      const missing = join(tmpDir, 'no-such-dir', 'missing.docx')
      expect(canonicalPathKey(missing)).toBe(missing)
    })

    it('maps Unicode-normalization variants of one name to one key', () => {
      const nfc = join(tmpDir, 'caf\u00e9.docx')
      const nfd = join(tmpDir, 'cafe\u0301.docx')
      expect(nfc).not.toBe(nfd)
      expect(canonicalPathKey(nfc)).toBe(canonicalPathKey(nfd))
    })
  })

  it('resolves equivalent spellings of one file to a single membership and chat', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const dotted = withDotSegment(filePath)

    const first = store.resolveChatForFile(filePath)
    store.appendChatMessage(first.projectId, first.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(first.projectId, first.chatId, { role: 'assistant', text: 'a' })

    const second = store.resolveChatForFile(dotted)
    expect(second.projectId).toBe(first.projectId)
    expect(second.chatId).toBe(first.chatId)

    const index = readIndex(tmpDir)
    expect(Object.keys(index.fileMap)).toHaveLength(1)
    expect(Object.keys(index.chatIdByPath ?? {})).toHaveLength(1)
    expect(store.getProject('default')?.files).toHaveLength(1)
  })

  it('appends through either spelling land in the same transcript', () => {
    const filePath = join(tmpDir, 'plan.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const dotted = withDotSegment(filePath)

    const first = store.resolveChatForFile(filePath)
    store.appendChatMessage(first.projectId, first.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(first.projectId, first.chatId, { role: 'assistant', text: 'a' })

    const viaDotted = store.chatIdForPath(dotted, first.projectId)
    store.appendChatMessage(first.projectId, viaDotted, { role: 'user', text: 'q2' })

    expect(store.loadChat(first.projectId, first.chatId).map((m) => m.text)).toEqual([
      'q',
      'a',
      'q2',
    ])
  })

  it('finds a chat an older version wrote under the raw path hash', () => {
    const filePath = join(tmpDir, 'legacy.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const dotted = withDotSegment(filePath)

    // How the previous version derived the id: sha256 of the raw spelling
    const legacyId = createHash('sha256').update(dotted).digest('hex').slice(0, 16)
    expect(legacyId).not.toBe(ProjectStore.chatIdForFile(dotted))
    const chatsDir = join(tmpDir, 'projects', 'default', 'chats')
    mkdirSync(chatsDir, { recursive: true })
    writeFileSync(
      join(chatsDir, `${legacyId}.jsonl`),
      `${JSON.stringify({ seq: 0, ts: new Date().toISOString(), role: 'user', text: 'old' })}\n`,
      'utf8',
    )

    expect(store.chatIdForPath(dotted, 'default')).toBe(legacyId)
    const resolved = store.resolveChatForFile(dotted)
    expect(resolved.chatId).toBe(legacyId)
    expect(store.loadChat(resolved.projectId, resolved.chatId).map((m) => m.text)).toEqual(['old'])
  })

  it('rename through an equivalent spelling keeps the same chat', () => {
    const filePath = join(tmpDir, 'before.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const ids = store.resolveChatForFile(filePath)
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'assistant', text: 'a' })

    const renamed = join(tmpDir, 'after.docx')
    renameFile(filePath, renamed)
    store.fileRenamed(withDotSegment(filePath), renamed)

    const after = store.resolveChatForFile(renamed)
    expect(after.chatId).toBe(ids.chatId)
    expect(store.loadChat(after.projectId, after.chatId).map((m) => m.text)).toEqual(['q', 'a'])
  })

  it('moves a file into another project through an equivalent spelling', () => {
    const filePath = join(tmpDir, 'movable.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const dotted = withDotSegment(filePath)
    const ids = store.resolveChatForFile(filePath)
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(ids.projectId, ids.chatId, { role: 'assistant', text: 'a' })

    const project = store.createProject('Target')
    store.moveFileToProject(dotted, project.id)

    expect(store.resolveProjectForFile(filePath)).toBe(project.id)
    expect(store.loadChat(project.id, ids.chatId)).toHaveLength(2)
    expect(store.getProject('default')?.files).toHaveLength(0)
    expect(store.getProject(project.id)?.files).toHaveLength(1)
  })

  it('derives a stable 16 hex chat id from the canonical key', () => {
    const filePath = join(tmpDir, 'stable.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    expect(ProjectStore.chatIdForFile(filePath)).toMatch(/^[0-9a-f]{16}$/)
    expect(ProjectStore.chatIdForFile(filePath)).toBe(
      ProjectStore.chatIdForFile(withDotSegment(filePath)),
    )
  })
})

describe('canonical path identity for symlinks', () => {
  let tmpDir: string
  let store: ProjectStore
  let canSymlink = false

  beforeAll(() => {
    const probe = mkdtempSync(join(tmpdir(), 'project-store-symlink-probe-'))
    try {
      const target = join(probe, 'target.docx')
      writeFileSync(target, 'doc', 'utf8')
      symlinkSync(target, join(probe, 'link.docx'), 'file')
      canSymlink = true
    } catch {
      canSymlink = false
    } finally {
      rmSync(probe, { recursive: true, force: true })
    }
  })

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-symlink-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  afterAll(() => {
    canSymlink = false
  })

  it.skipIf(!canSymlink)('resolves a symlink to the same identity as its target', () => {
    const target = join(tmpDir, 'target.docx')
    writeFileSync(target, 'doc', 'utf8')
    const link = join(tmpDir, 'link.docx')
    symlinkSync(target, link, 'file')

    const real = store.resolveChatForFile(target)
    store.appendChatMessage(real.projectId, real.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(real.projectId, real.chatId, { role: 'assistant', text: 'a' })

    const viaLink = store.resolveChatForFile(link)
    expect(viaLink.projectId).toBe(real.projectId)
    expect(viaLink.chatId).toBe(real.chatId)
    expect(Object.keys(readIndex(tmpDir).fileMap)).toHaveLength(1)
    expect(store.getProject('default')?.files).toHaveLength(1)
  })
})

describe.skipIf(process.platform !== 'win32')('canonical path identity for case variants', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-case-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('resolves Windows case variants of one file to a single identity', () => {
    const filePath = join(tmpDir, 'MixedCase.docx')
    writeFileSync(filePath, 'doc', 'utf8')
    const shouted = filePath.toUpperCase()
    expect(shouted).not.toBe(filePath)
    expect(existsSync(shouted)).toBe(true)

    const real = store.resolveChatForFile(filePath)
    store.appendChatMessage(real.projectId, real.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(real.projectId, real.chatId, { role: 'assistant', text: 'a' })

    const viaShout = store.resolveChatForFile(shouted)
    expect(viaShout.projectId).toBe(real.projectId)
    expect(viaShout.chatId).toBe(real.chatId)
    expect(Object.keys(readIndex(tmpDir).fileMap)).toHaveLength(1)
    expect(store.getProject('default')?.files).toHaveLength(1)
  })
})

function renameFile(from: string, to: string): void {
  writeFileSync(to, readFileSync(from))
  rmSyncFile(from)
}

function rmSyncFile(path: string): void {
  rmSync(path, { force: true })
}
