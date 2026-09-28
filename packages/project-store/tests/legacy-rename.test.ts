import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { join, sep } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

/**
 * The chat id an older version derived: sha256 of the path exactly as it was
 * handed to it, first 16 hex chars — no filesystem canonicalization.
 */
const legacyChatId = (rawPath: string) =>
  createHash('sha256').update(rawPath.normalize('NFC')).digest('hex').slice(0, 16)

describe('fileRenamed with a transcript under the legacy path hash', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-legacy-rename-'))
    store = new ProjectStore(tmpDir)
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('carries the legacy chat id over to the new path', () => {
    // Spelled with a ".." segment (path.join would normalize it away): the
    // parent still resolves, so the canonical key (realpath of the directory)
    // differs from the raw string the older version hashed.
    const subDir = join(tmpDir, 'sub')
    mkdirSync(subDir)
    const oldPath = `${subDir}${sep}..${sep}doc.txt`
    const newPath = join(tmpDir, 'renamed.txt')
    writeFileSync(newPath, 'content')

    // The old version registered the file in the default project and wrote the
    // transcript under the hash of the raw path, with no chatIdByPath entry.
    store.resolveProjectForFile(oldPath)
    const chatId = legacyChatId(oldPath)
    const chatsDir = join(tmpDir, 'projects', 'default', 'chats')
    mkdirSync(chatsDir, { recursive: true })
    writeFileSync(
      join(chatsDir, `${chatId}.jsonl`),
      JSON.stringify({
        seq: 0,
        ts: '2026-01-01T00:00:00.000Z',
        role: 'user',
        text: 'legacy hello',
      }) + '\n',
    )

    store.fileRenamed(oldPath, newPath)

    const resolved = store.resolveChatForFile(newPath)
    expect(resolved.chatId).toBe(chatId)
    expect(store.loadChat('default', resolved.chatId).map((m) => m.text)).toEqual(['legacy hello'])
  })
})
