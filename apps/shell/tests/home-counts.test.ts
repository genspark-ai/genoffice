import { afterEach, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createI18n, LANGS, type Lang } from '@genoffice/i18n'
import { normalizeRecentQuery, pageRecentPaths } from '../src/main/recent-files'
import { fileCountLabel, visiblePageCount } from '../src/renderer/src/counts'
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

  it('counts legacy .xls under the sheets (xlsx) filter', () => {
    const dir = mkdtempSync(join(tmpdir(), 'shell-counts-'))
    tempDirs.push(dir)
    const bookPath = join(dir, 'book.xlsx')
    const legacyPath = join(dir, 'legacy.xls')
    const docPath = join(dir, 'notes.docx')
    writeFileSync(bookPath, 'sheet')
    writeFileSync(legacyPath, 'sheet')
    writeFileSync(docPath, 'doc')

    const page = pageRecentPaths(
      [bookPath, legacyPath, docPath],
      { ext: 'xlsx', offset: 0, limit: 50 },
      new Set(),
    )

    expect(page.total).toBe(2)
    expect(page.entries.map((entry) => entry.path)).toEqual([bookPath, legacyPath])
  })

  it('keeps unavailable paths listed at their position, flagged missing (r158)', () => {
    const dir = mkdtempSync(join(tmpdir(), 'shell-counts-'))
    tempDirs.push(dir)
    const existingPath = join(dir, 'existing.xlsx')
    const missingPath = join(dir, 'missing.xlsx')
    writeFileSync(existingPath, 'sheet')

    const page = pageRecentPaths([missingPath, existingPath], {}, new Set())

    // a transiently unstat-able file (disconnected drive, pending mount) must
    // not vanish from the list — it renders dimmed with an unavailable state
    expect(page.total).toBe(2)
    expect(page.totalAll).toBe(2)
    expect(page.entries.map((entry) => [entry.path, entry.missing === true])).toEqual([
      [missingPath, true],
      [existingPath, false],
    ])
    expect(page.entries[0].mtimeMs).toBe(0)
    expect(page.entries[0].ext).toBe('xlsx')
  })
})

describe('recent query ext normalization', () => {
  it('trims whitespace, strips leading dots, and lowercases the filter', () => {
    expect(normalizeRecentQuery({ ext: '.XLSX' }).ext).toBe('xlsx')
    expect(normalizeRecentQuery({ ext: ' xlsx ' }).ext).toBe('xlsx')
    expect(normalizeRecentQuery({ ext: '...md' }).ext).toBe('md')
    expect(normalizeRecentQuery({ ext: '...' }).ext).toBeUndefined()
    expect(normalizeRecentQuery({ ext: '' }).ext).toBeUndefined()
    expect(normalizeRecentQuery({}).ext).toBeUndefined()
  })

  it('applies the normalized filter to the page', () => {
    const dir = mkdtempSync(join(tmpdir(), 'shell-counts-'))
    tempDirs.push(dir)
    const bookPath = join(dir, 'book.xlsx')
    writeFileSync(bookPath, 'sheet')

    const page = pageRecentPaths([bookPath], { ext: '.XLSX', limit: 50 }, new Set())
    expect(page.total).toBe(1)
    expect(page.entries.map((entry) => entry.path)).toEqual([bookPath])
  })

  it('shares the sheets/html families with the starred view (same helper)', () => {
    const dir = mkdtempSync(join(tmpdir(), 'shell-counts-'))
    tempDirs.push(dir)
    const htmPath = join(dir, 'page.htm')
    const htmlPath = join(dir, 'page.html')
    const legacyPath = join(dir, 'legacy.xls')
    writeFileSync(htmPath, 'html')
    writeFileSync(htmlPath, 'html')
    writeFileSync(legacyPath, 'sheet')
    // pageRecentPaths is the recents helper; starred now calls the same
    // matchesExtFamily, so assert the family includes both spellings
    expect(
      pageRecentPaths([htmPath, legacyPath], { ext: 'html', limit: 50 }, new Set()).total,
    ).toBe(1)
    expect(
      pageRecentPaths([htmPath, legacyPath], { ext: 'xlsx', limit: 50 }, new Set()).total,
    ).toBe(1)
    expect(pageRecentPaths([htmlPath], { ext: 'htm', limit: 50 }, new Set()).total).toBe(0)
  })
})

describe('count labels', () => {
  const translate = createI18n(strings)
  const label = (lang: Lang, n: number) =>
    fileCountLabel(n, lang, (key, params) => translate(lang, key, params))

  it('uses singular and plural file labels', () => {
    expect(label('en', 1)).toBe('1 file')
    expect(label('en', 2)).toBe('2 files')
    expect(label('en', 0)).toBe('0 files')
  })

  it('picks the singular form in every locale with plural inflection', () => {
    expect(label('fr', 1)).toBe('1 fichier')
    expect(label('de', 1)).toBe('1 Datei')
    expect(label('zh', 1)).toBe('1 \u4e2a\u6587\u4ef6')
  })

  it('follows CLDR categories: fr zero, cs few, ru one at 21, ar dual/many', () => {
    expect(label('fr', 0)).toBe('0 fichier')
    expect(label('cs', 1)).toBe('1 soubor')
    expect(label('cs', 3)).toBe('3 soubory')
    expect(label('cs', 5)).toBe('5 soubor\u016f')
    expect(label('ru', 21)).toBe('21 \u0444\u0430\u0439\u043b')
    expect(label('ru', 5)).toBe('\u0424\u0430\u0439\u043b\u043e\u0432: 5')
    expect(label('pl', 2)).toBe('Pliki: 2')
    expect(label('ar', 0)).toBe('لا توجد ملفات')
    expect(label('ar', 2)).toBe('ملفان')
    expect(label('ar', 3)).toBe('3 ملفات')
    expect(label('ar', 11)).toBe('11 ملفًا')
    expect(label('ar', 100)).toBe(strings.ar.fileCount.replace('{n}', '100'))
  })

  it('falls back to the one/other pair in every locale', () => {
    for (const lang of LANGS) {
      for (const n of [0, 1, 2, 5, 11, 21, 100]) expect(label(lang, n)).not.toContain('{n}')
    }
  })
})
