import { existsSync, mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { clearOpenDocuments, publishOpenDocuments } from '../src/main/open-documents'

describe('open-documents registry', () => {
  it('writes the deduplicated sorted path list with this pid and removes it on clear', () => {
    const path = join(mkdtempSync(join(tmpdir(), 'open-docs-')), 'open-documents.json')
    publishOpenDocuments(path, ['/b.docx', '/a.xlsx', '/b.docx'])
    const registry = JSON.parse(readFileSync(path, 'utf8'))
    expect(registry).toMatchObject({ pid: process.pid, paths: ['/a.xlsx', '/b.docx'] })
    expect(typeof registry.updatedAt).toBe('string')
    publishOpenDocuments(path, [])
    expect(JSON.parse(readFileSync(path, 'utf8')).paths).toEqual([])
    clearOpenDocuments(path)
    expect(existsSync(path)).toBe(false)
    clearOpenDocuments(path)
  })

  it('reports a failed write once per distinct error instead of swallowing it', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    const missing = '/nonexistent-dir/x/open-documents.json'
    expect(publishOpenDocuments(missing, ['/a'])).toBe(false)
    expect(publishOpenDocuments(missing, ['/b'])).toBe(false)
    expect(warn).toHaveBeenCalledTimes(1)
    expect(warn.mock.calls[0]?.[1]).toMatch(/ENOENT/)
    const ok = join(mkdtempSync(join(tmpdir(), 'open-docs-')), 'open-documents.json')
    expect(publishOpenDocuments(ok, ['/a'])).toBe(true)
  })

  afterEach(() => vi.restoreAllMocks())
})
