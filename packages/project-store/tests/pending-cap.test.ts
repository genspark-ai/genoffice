import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { MAX_PENDING_OPENING_MESSAGES, ProjectStore } from '../src/store.js'

describe('pending opening message buffer cap', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-pending-cap-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('materializes the file instead of growing memory without bound', () => {
    for (let i = 0; i < MAX_PENDING_OPENING_MESSAGES + 50; i++) {
      store.appendChatMessage('default', 'chat1', { role: 'user', text: `msg${i}` })
    }
    const msgs = store.loadChat('default', 'chat1', MAX_PENDING_OPENING_MESSAGES + 50)
    expect(msgs).toHaveLength(MAX_PENDING_OPENING_MESSAGES + 50)
    expect(msgs[0]!.text).toBe('msg0')
  })
})
