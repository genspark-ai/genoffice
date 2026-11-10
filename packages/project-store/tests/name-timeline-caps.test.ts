import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { MAX_PROJECT_NAME_CHARS, ProjectStore } from '../src/store.js'

describe('project name length and timeline limit caps', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-caps-'))
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('rejects overlong names on create and rename', () => {
    const big = 'x'.repeat(MAX_PROJECT_NAME_CHARS + 1)
    expect(() => store.createProject(big)).toThrow(/too long/)
    const p = store.createProject('ok')
    expect(() => store.renameProject(p.id, big)).toThrow(/too long/)
  })

  it('bounds timeline limit for NaN/huge/negative', () => {
    expect(store.getProjectTimeline('default', NaN)).toEqual([])
    expect(store.getProjectTimeline('default', 1e9)).toEqual([])
    expect(store.getProjectTimeline('default', -5)).toEqual([])
  })

  it('loads more than 200 messages per chat when the limit allows', () => {
    const proj = store.createProject('Timeline Cap')
    const chatId = ProjectStore.chatIdForFile(join(tmpDir, 'deck.pptx'))
    for (let i = 0; i < 250; i++) {
      store.appendChatMessage(proj.id, chatId, { role: 'user', text: `m${i}` })
    }
    expect(store.getProjectTimeline(proj.id, 300)).toHaveLength(250)
  })
})
