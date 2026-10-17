import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { ProjectStore } from '../src/store'

function makeTempDir(): string {
  return mkdtempSync(join(tmpdir(), 'project-store-delete-chat-'))
}

describe('deleteChat', () => {
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

  it('removes the transcript it names and leaves the others alone', () => {
    // The assistant turn is what materializes the file; a lone user turn is
    // still only in the pending buffer.
    store.appendChatMessage('default', 'keep', { role: 'user', text: 'kept' })
    store.appendChatMessage('default', 'keep', { role: 'assistant', text: 'kept answer' })
    store.appendChatMessage('default', 'drop', { role: 'user', text: 'dropped' })
    store.appendChatMessage('default', 'drop', { role: 'assistant', text: 'dropped answer' })
    store.deleteChat('default', 'drop')

    expect(store.listChats('default').map((chat) => chat.chatId)).toEqual(['keep'])
    expect(store.loadChat('default', 'drop')).toEqual([])
    expect(store.loadChat('default', 'keep')).toHaveLength(2)
  })

  it('deletes a conversation whose file was never written', () => {
    // The normal case for a question just asked and then dismissed: it is
    // still in the pending buffer, and loadChat can still hand it back until
    // the queue is dropped.
    store.appendChatMessage('default', 'pending', { role: 'user', text: 'asked' })
    expect(store.loadChat('default', 'pending')).toHaveLength(1)
    store.deleteChat('default', 'pending')
    expect(store.loadChat('default', 'pending')).toEqual([])
    expect(store.listChats('default')).toEqual([])
  })

  it('is a no-op for a conversation that is not there', () => {
    // Dismissing a thread a previous run already removed must not throw: the
    // list and the reader can disagree by one entry, and the reader is the one
    // that was asked to do something.
    store.appendChatMessage('default', 'here', { role: 'user', text: 'text' })
    expect(() => store.deleteChat('default', 'never-existed')).not.toThrow()
  })

  it('refuses an id that would reach outside the chats directory', () => {
    // The id arrives over IPC from a renderer; a path separator in it must not
    // turn a delete into an arbitrary unlink.
    expect(() => store.deleteChat('default', '../project.json')).toThrow(/Invalid chatId/)
    expect(() => store.deleteChat('default', 'a/b')).toThrow(/Invalid chatId/)
  })

  it('keeps a project that holds only file-less conversations', () => {
    store.appendChatMessage('default', 'kb-ask', { role: 'user', text: 'q' })
    store.appendChatMessage('default', 'kb-ask', { role: 'assistant', text: 'a' })
    expect(store.listChats('default')).toHaveLength(1)
    store.deleteChat('default', 'kb-ask')
    expect(store.listChats('default')).toEqual([])
    // The default project is still there to be written to again.
    store.appendChatMessage('default', 'kb-ask', { role: 'user', text: 'again' })
    expect(store.loadChat('default', 'kb-ask')).toHaveLength(1)
  })
})
