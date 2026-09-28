import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'
import type { ToolActivity } from '../src/types.js'

describe('appendChatMessage with a non-string tool field', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-tool-field-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('keeps the assistant turn instead of dropping it on a thrown truncate', () => {
    // A provider can hand the input over as an object, not as the string the
    // type promises: .slice() on it threw and the record was never written.
    const tools = [
      { name: 'read_file', summary: 'read a.txt', input: { path: 'a.txt' } },
    ] as unknown as ToolActivity[]

    store.appendChatMessage('default', 'chat1', { role: 'assistant', text: 'here it is', tools })

    const messages = store.loadChat('default', 'chat1')
    expect(messages.map((m) => m.text)).toEqual(['here it is'])
    expect(messages[0]?.tools?.[0]?.input).toBe('{"path":"a.txt"}')
  })
})
