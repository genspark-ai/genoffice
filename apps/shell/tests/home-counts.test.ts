import { afterEach, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createI18n } from '@genoffice/i18n'
import { pageRecentPaths } from '../src/main/recent-files'
import { fileCountKey, timelineCountKey, visiblePageCount } from '../src/renderer/src/counts'
import { strings } from '../src/renderer/src/strings'

const tempDirs: string[] = []

afterEach(() => {
  for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

describe('home visible counts', () => {
  it('uses the filtered total for the sidebar count', () => {
    const dir = mkdtempSync(join(tmpdir(), 'shell-counts-'))
    tempDirs.push(dir)
    const docPath = join(dir, 'notes.docx')
    const slidePath = join(dir, 'deck.pptx')
    writeFileSync(docPath, 'doc')
    writeFileSync(slidePath, 'slide')

    const page = pageRecentPaths(
      [docPath, slidePath],
      { ext: 'docx', offset: 0, limit: 50 },
      new Set(),
    )

    expect(page.totalAll).toBe(2)
    expect(page.total).toBe(1)
    expect(page.entries.map((entry) => entry.path)).toEqual([docPath])
    expect(visiblePageCount(page)).toBe(1)
  })

  it('counts .xlsm under the sheets (xlsx) filter', () => {
    const dir = mkdtempSync(join(tmpdir(), 'shell-counts-'))
    tempDirs.push(dir)
    const bookPath = join(dir, 'book.xlsx')
    const macroPath = join(dir, 'macro.xlsm')
    const docPath = join(dir, 'notes.docx')
    writeFileSync(bookPath, 'sheet')
    writeFileSync(macroPath, 'sheet')
    writeFileSync(docPath, 'doc')

    const page = pageRecentPaths(
      [bookPath, macroPath, docPath],
      { ext: 'xlsx', offset: 0, limit: 50 },
      new Set(),
    )

    expect(page.total).toBe(2)
    expect(page.entries.map((entry) => entry.path)).toEqual([bookPath, macroPath])
  })

  it('excludes stale paths from totals and rows', () => {
    const dir = mkdtempSync(join(tmpdir(), 'shell-counts-'))
    tempDirs.push(dir)
    const existingPath = join(dir, 'existing.xlsx')
    const missingPath = join(dir, 'missing.xlsx')
    writeFileSync(existingPath, 'sheet')

    const page = pageRecentPaths([missingPath, existingPath], {}, new Set())

    expect(page.total).toBe(1)
    expect(page.totalAll).toBe(1)
    expect(page.entries.map((entry) => entry.path)).toEqual([existingPath])
  })
})

describe('count labels', () => {
  const translate = createI18n(strings)

  it('uses singular and plural file labels', () => {
    expect(translate('en', fileCountKey(1), { n: 1 })).toBe('1 file')
    expect(translate('en', fileCountKey(2), { n: 2 })).toBe('2 files')
  })

  it('uses singular and plural activity item labels', () => {
    expect(translate('en', timelineCountKey(1), { n: 1 })).toBe('1 item')
    expect(translate('en', timelineCountKey(2), { n: 2 })).toBe('2 items')
  })

  it('picks the singular form in every locale with plural inflection', () => {
    expect(translate('fr', fileCountKey(1), { n: 1 })).toBe('1 fichier')
    expect(translate('de', fileCountKey(1), { n: 1 })).toBe('1 Datei')
    expect(translate('zh', fileCountKey(1), { n: 1 })).toBe('1 个文件')
  })
})
