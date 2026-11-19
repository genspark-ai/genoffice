import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { join, sep } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

let tmpDir: string
let store: ProjectStore

beforeEach(() => {
  tmpDir = mkdtempSync(join(tmpdir(), 'project-store-tar-'))
  store = new ProjectStore(tmpDir)
  store.ensureDefaultProject()
})

afterEach(() => {
  rmSync(tmpDir, { recursive: true, force: true })
  vi.restoreAllMocks()
})

const chatFile = (chatId: string) => join(tmpDir, 'projects', 'default', 'chats', `${chatId}.jsonl`)

describe('loadChat with a record larger than the tail window', () => {
  it('reads further back instead of returning an empty chat', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    mkdirSync(join(tmpDir, 'projects', 'default', 'chats'), { recursive: true })
    const huge = { seq: 1, ts: 't', role: 'assistant', text: 'x'.repeat(9 * 1024 * 1024) }
    writeFileSync(
      chatFile('big'),
      JSON.stringify({ seq: 0, ts: 't', role: 'user', text: 'q' }) +
        '\n' +
        JSON.stringify(huge) +
        '\n',
    )
    const messages = store.loadChat('default', 'big')
    expect(messages.map((m) => m.seq)).toEqual([0, 1])
    expect(warn).toHaveBeenCalled()
  })
})

describe('appendChatMessage write failure', () => {
  it('keeps the buffered opening messages and throws', () => {
    vi.spyOn(console, 'warn').mockImplementation(() => {})
    store.appendChatMessage('default', 'c1', { role: 'user', text: 'q' })
    const chatsDir = join(tmpDir, 'projects', 'default', 'chats')
    mkdirSync(chatsDir, { recursive: true })
    // A directory at the transcript path makes the append fail
    mkdirSync(chatFile('c1'))
    expect(() =>
      store.appendChatMessage('default', 'c1', { role: 'assistant', text: 'a' }),
    ).toThrow()
    expect(store.loadChat('default', 'c1').map((m) => m.text)).toEqual(['q'])
    rmSync(chatFile('c1'), { recursive: true })
    store.appendChatMessage('default', 'c1', { role: 'assistant', text: 'a2' })
    expect(store.loadChat('default', 'c1').map((m) => m.text)).toEqual(['q', 'a2'])
  })
})

describe('fileRenamed onto an already-mapped path', () => {
  it('source mappings win and the displaced entries are cleaned up', () => {
    const src = join(tmpDir, 'src.docx')
    const dst = join(tmpDir, 'dst.docx')
    writeFileSync(src, '')
    writeFileSync(dst, '')
    const other = store.createProject('other')
    store.moveFileToProject(dst, other.id)
    const dstIds = store.resolveChatForFile(dst)
    store.appendChatMessage(dstIds.projectId, dstIds.chatId, { role: 'user', text: 'old q' })
    store.appendChatMessage(dstIds.projectId, dstIds.chatId, { role: 'assistant', text: 'old a' })
    const srcIds = store.resolveChatForFile(src)
    store.appendChatMessage(srcIds.projectId, srcIds.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(srcIds.projectId, srcIds.chatId, { role: 'assistant', text: 'a' })

    store.fileRenamed(src, dst)

    const after = store.resolveChatForFile(dst)
    expect(after).toEqual(srcIds)
    expect(store.loadChat(after.projectId, after.chatId).map((m) => m.text)).toEqual(['q', 'a'])
    expect(store.getProject(other.id)?.files).toEqual([])
    expect(store.getProject('default')?.files.filter((f) => f === dst)).toHaveLength(1)
    expect(existsSync(join(tmpDir, 'projects', other.id, 'chats', `${dstIds.chatId}.jsonl`))).toBe(
      false,
    )
    expect(readdirSync(join(tmpDir, 'projects', '.trash'))).toHaveLength(1)
    expect(store.knownFilePaths().filter((p) => p.endsWith('src.docx'))).toEqual([])
  })

  it('also displaces a destination stored under a legacy non-canonical key', () => {
    const src = join(tmpDir, 'src.docx')
    const dst = join(tmpDir, 'dst.docx')
    writeFileSync(src, '')
    writeFileSync(dst, '')
    mkdirSync(join(tmpDir, 'sub'))
    // Spelled with a ".." segment: an older version keyed the raw string
    const rawDst = `${tmpDir}${sep}sub${sep}..${sep}dst.docx`
    const dstIds = store.resolveChatForFile(dst)
    store.appendChatMessage(dstIds.projectId, dstIds.chatId, { role: 'user', text: 'old q' })
    store.appendChatMessage(dstIds.projectId, dstIds.chatId, { role: 'assistant', text: 'old a' })
    const srcIds = store.resolveChatForFile(src)
    store.appendChatMessage(srcIds.projectId, srcIds.chatId, { role: 'user', text: 'q' })
    store.appendChatMessage(srcIds.projectId, srcIds.chatId, { role: 'assistant', text: 'a' })
    const indexPath = join(tmpDir, 'projects', 'index.json')
    const index = JSON.parse(readFileSync(indexPath, 'utf8'))
    const canonicalDst = Object.keys(index.chatIdByPath).find(
      (k) => index.chatIdByPath[k] === dstIds.chatId,
    )!
    index.fileMap[rawDst] = index.fileMap[canonicalDst]
    delete index.fileMap[canonicalDst]
    index.chatIdByPath[rawDst] = dstIds.chatId
    delete index.chatIdByPath[canonicalDst]
    writeFileSync(indexPath, JSON.stringify(index))

    store.fileRenamed(src, rawDst)

    expect(store.resolveChatForFile(dst)).toEqual(srcIds)
    const after = JSON.parse(readFileSync(indexPath, 'utf8'))
    expect(after.fileMap[rawDst]).toBeUndefined()
    expect(after.chatIdByPath[rawDst]).toBeUndefined()
    expect(existsSync(chatFile(dstIds.chatId))).toBe(false)
    expect(readdirSync(join(tmpDir, 'projects', '.trash'))).toHaveLength(1)
  })
})
