import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
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

  it('moves the corrupt file aside and starts a fresh project next to it', () => {
    const dir = join(tmpDir, 'projects', 'default')
    const projectJson = join(dir, 'project.json')
    // Truncated mid-document, the way an interrupted write leaves it
    const corrupt = readFileSync(projectJson, 'utf8').slice(0, -3)
    writeFileSync(projectJson, corrupt)

    // Every file resolve calls this
    const store = new ProjectStore(tmpDir)
    store.resolveProjectForFile(join(tmpDir, 'other.txt'))

    const aside = readdirSync(dir).filter((f) => f.startsWith('project.json.corrupt-'))
    expect(aside).toHaveLength(1)
    expect(readFileSync(join(dir, aside[0]!), 'utf8')).toBe(corrupt)
    // the store is usable again: the fresh project parses and records the file
    const fresh = JSON.parse(readFileSync(projectJson, 'utf8')) as { files: unknown[] }
    expect(fresh.files).toHaveLength(1)
  })
})
