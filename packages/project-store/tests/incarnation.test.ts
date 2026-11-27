import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { canonicalPathKey, fileIncarnation, ProjectStore } from '../src/store.js'
import type { ProjectIndex } from '../src/types.js'

let tmpDir: string
let store: ProjectStore
beforeEach(() => {
  tmpDir = mkdtempSync(join(tmpdir(), 'project-store-incarnation-'))
  store = new ProjectStore(tmpDir)
})
afterEach(() => {
  rmSync(tmpDir, { recursive: true, force: true })
})

const writeDoc = (path: string, text: string) => writeFileSync(path, text)

describe('file incarnation', () => {
  it('changes when a file at the same path is deleted and recreated', () => {
    const path = join(tmpDir, 'plan.docx')
    writeDoc(path, 'old')
    const before = fileIncarnation(path)
    expect(before).toBeDefined()
    rmSync(path)
    writeDoc(path, 'new')
    expect(fileIncarnation(path)).toBeDefined()
    expect(fileIncarnation(path)).not.toBe(before)
  })

  it('is undefined for a path with no file', () => {
    expect(fileIncarnation(join(tmpDir, 'missing.docx'))).toBeUndefined()
  })

  it('mints a fresh chat with an empty transcript for a replacement file', () => {
    const path = join(tmpDir, 'plan.docx')
    writeDoc(path, 'old')
    const first = store.resolveChatForFile(path)
    store.appendChatMessage(first.projectId, first.chatId, {
      role: 'user',
      text: 'old document question',
    })
    expect(store.loadChat(first.projectId, first.chatId)).toHaveLength(1)

    // deleted and replaced outside GenOffice: the path mapping survives
    rmSync(path)
    writeDoc(path, 'unrelated')
    const second = store.resolveChatForFile(path)
    expect(second.chatId).not.toBe(first.chatId)
    expect(second.projectId).toBe(first.projectId)
    expect(store.loadChat(second.projectId, second.chatId)).toEqual([])
    // the old conversation is still there, untouched
    expect(store.loadChat(first.projectId, first.chatId)[0]!.text).toBe('old document question')
  })

  it('repeated resolves of the same file keep one stable chat id', () => {
    const path = join(tmpDir, 'stable.docx')
    writeDoc(path, 'same')
    const first = store.resolveChatForFile(path)
    const second = store.resolveChatForFile(path)
    const third = store.resolveChatForFile(path)
    expect(second.chatId).toBe(first.chatId)
    expect(third.chatId).toBe(first.chatId)
    expect(existsSync(join(tmpDir, 'projects', 'index.json'))).toBe(true)
  })

  it('a rename keeps the chat because the physical file is the same', () => {
    const oldPath = join(tmpDir, 'old.docx')
    const newPath = join(tmpDir, 'new.docx')
    writeDoc(oldPath, 'same bytes')
    const first = store.resolveChatForFile(oldPath)
    store.appendChatMessage(first.projectId, first.chatId, {
      role: 'user',
      text: 'keep me',
    })
    renameSync(oldPath, newPath)
    store.fileRenamed(oldPath, newPath)
    const after = store.resolveChatForFile(newPath)
    expect(after.chatId).toBe(first.chatId)
    expect(store.loadChat(after.projectId, after.chatId)[0]!.text).toBe('keep me')
  })

  it('adopts an existing mapping written by an older version', () => {
    const path = join(tmpDir, 'legacy.docx')
    writeDoc(path, 'legacy')
    // an index with the chat mapping but no incarnation record
    const legacy: ProjectIndex = {
      projects: [],
      fileMap: { [path]: 'default' },
      chatIdByPath: { [path]: 'legacyChat' },
    }
    mkdirSync(join(tmpDir, 'projects'))
    writeFileSync(join(tmpDir, 'projects', 'index.json'), JSON.stringify(legacy))
    const resolved = store.resolveChatForFile(path)
    expect(resolved.chatId).toBe('legacyChat')
    const index = JSON.parse(
      readFileSync(join(tmpDir, 'projects', 'index.json'), 'utf8'),
    ) as ProjectIndex
    expect(index.incarnationByPath?.[canonicalPathKey(path)]).toBe(fileIncarnation(path))
  })

  it('deleting a chat drops its incarnation so the path can be rebound', () => {
    const path = join(tmpDir, 'again.docx')
    writeDoc(path, 'one')
    const first = store.resolveChatForFile(path)
    store.deleteChat(first.projectId, first.chatId)
    const index = JSON.parse(
      readFileSync(join(tmpDir, 'projects', 'index.json'), 'utf8'),
    ) as ProjectIndex
    expect(index.chatIdByPath?.[canonicalPathKey(path)]).toBeUndefined()
    expect(index.incarnationByPath?.[canonicalPathKey(path)]).toBeUndefined()
    // a new file at the same path can reuse the original id since the old chat is gone
    rmSync(path)
    writeDoc(path, 'two')
    expect(store.resolveChatForFile(path).chatId).toBe(first.chatId)
  })
})
