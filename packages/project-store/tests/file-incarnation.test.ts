import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtempSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

/**
 * A path is not an identity. When the file behind a path is deleted and an
 * unrelated one takes its place, the new file must start with an empty
 * history instead of inheriting the old document's transcript
 * (genoffice#1021).
 *
 * The identity is the device and inode the file lives on, plus the inode's
 * creation time. GenOffice's own save replaces the file (a temp file is
 * renamed over it), so the recorded identity is refreshed after a save —
 * without that, every save would read as "a different file".
 */
describe('a reused path does not inherit the previous file history', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-incarnation-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  const filePath = (): string => join(tmpDir, 'plan.docx')

  it('keeps the same chat for the same file', () => {
    const path = filePath()
    writeFileSync(path, 'the first document', 'utf8')
    const first = store.resolveChatForFile(path)
    store.appendChatMessage('default', first.chatId, { role: 'user', text: 'hello' })
    const second = store.resolveChatForFile(path)
    expect(second.chatId).toBe(first.chatId)
    expect(store.loadChat('default', second.chatId)).toHaveLength(1)
  })

  it('starts a new chat when an unrelated file takes the path', () => {
    const path = filePath()
    writeFileSync(path, 'the first document', 'utf8')
    const first = store.resolveChatForFile(path)
    store.appendChatMessage('default', first.chatId, { role: 'user', text: 'hello' })

    // the old file goes away, an unrelated one takes its place
    rmSync(path)
    writeFileSync(path, 'a completely different document', 'utf8')

    const second = store.resolveChatForFile(path)
    expect(second.chatId).not.toBe(first.chatId)
    expect(store.loadChat('default', second.chatId)).toEqual([])
    // the old transcript is still there, under the old chat
    expect(store.loadChat('default', first.chatId)).toHaveLength(1)
  })

  it('keeps the chat across a save that replaces the file', () => {
    const path = filePath()
    writeFileSync(path, 'the first document', 'utf8')
    const first = store.resolveChatForFile(path)
    store.appendChatMessage('default', first.chatId, { role: 'user', text: 'hello' })

    // what GenOffice's atomic write does: a temp file renamed over the path,
    // which gives the path a new inode
    const tmp = join(tmpDir, 'plan.docx.tmp')
    writeFileSync(tmp, 'the edited document', 'utf8')
    renameSync(tmp, path)
    // the store learns about the new inode, the way the save path tells it
    store.refreshFileIdentity(path)

    const second = store.resolveChatForFile(path)
    expect(second.chatId).toBe(first.chatId)
    expect(store.loadChat('default', second.chatId)).toHaveLength(1)
  })

  it('keeps the chat of an index written before identities were recorded', () => {
    const path = filePath()
    writeFileSync(path, 'the first document', 'utf8')
    const first = store.resolveChatForFile(path)

    // an index from an older version: the chat is mapped, no identity yet
    const indexPath = join(tmpDir, 'projects', 'index.json')
    const index = JSON.parse(readFileSync(indexPath, 'utf8'))
    delete index.fileIdentityByPath
    writeFileSync(indexPath, JSON.stringify(index))

    const second = store.resolveChatForFile(path)
    expect(second.chatId).toBe(first.chatId)
    // and the identity is recorded from here on
    const updated = JSON.parse(readFileSync(indexPath, 'utf8'))
    expect(Object.keys(updated.fileIdentityByPath ?? {})).toHaveLength(1)
  })

  it('leaves a file that is not there yet alone', () => {
    const path = filePath()
    const first = store.resolveChatForFile(path)
    const second = store.resolveChatForFile(path)
    expect(second.chatId).toBe(first.chatId)
  })
})
