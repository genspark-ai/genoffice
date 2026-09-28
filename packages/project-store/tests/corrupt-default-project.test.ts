import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

describe('ensureDefaultProject with an unparseable project.json', () => {
  let tmpDir: string

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-corrupt-default-'))
    new ProjectStore(tmpDir).resolveProjectForFile(join(tmpDir, 'registered.txt'))
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('leaves the corrupt file alone instead of writing an empty project over it', () => {
    const projectJson = join(tmpDir, 'projects', 'default', 'project.json')
    // Truncated mid-document, the way an interrupted write leaves it
    const corrupt = readFileSync(projectJson, 'utf8').slice(0, -3)
    writeFileSync(projectJson, corrupt)

    // Every file resolve calls this
    new ProjectStore(tmpDir).resolveProjectForFile(join(tmpDir, 'other.txt'))

    expect(readFileSync(projectJson, 'utf8')).toBe(corrupt)
  })
})
