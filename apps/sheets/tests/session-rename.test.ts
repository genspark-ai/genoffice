import { describe, expect, it } from 'vitest'

import { sessionAfterRename, xlsxSiblingOf } from '../src/main/session-rename'

const base = { path: '/tmp/genoffice/import/report.xlsx', snapshotPath: '/tmp/snap', sha256: 'abc' }

describe('sessionAfterRename', () => {
  it('moves a plain workbook session with its file', () => {
    const session = { ...base, path: '/docs/report.xlsx' }
    expect(sessionAfterRename(session, '/docs/report.xlsx', '/docs/q3.xlsx')).toEqual({
      ...session,
      path: '/docs/q3.xlsx',
    })
  })

  it('ignores sessions unrelated to the renamed file', () => {
    expect(sessionAfterRename({ ...base }, '/docs/other.xlsx', '/docs/q3.xlsx')).toBeNull()
    expect(
      sessionAfterRename(
        { ...base, csvSourcePath: '/docs/data.csv', suggestSaveAs: '/docs/data.xlsx' },
        '/docs/other.csv',
        '/docs/renamed.csv',
      ),
    ).toBeNull()
  })

  it('follows a renamed .csv through csvSourcePath so Save writes to the new name', () => {
    // the converted copy in the temp dir stays where it is; only the source moves
    const session = { ...base, csvImport: true, csvSourcePath: '/docs/data.csv' }
    expect(sessionAfterRename(session, '/docs/data.csv', '/docs/sales-2026.csv')).toEqual({
      ...session,
      csvSourcePath: '/docs/sales-2026.csv',
    })
  })

  it('re-derives the Save As default of a converted .xls / .tsv copy', () => {
    const xls = { ...base, suggestSaveAs: '/docs/legacy.xlsx' }
    expect(sessionAfterRename(xls, '/docs/legacy.xls', '/docs/archive.xls')).toEqual({
      ...xls,
      suggestSaveAs: '/docs/archive.xlsx',
    })
    const tsv = { ...base, suggestSaveAs: '/docs/variants.xlsx' }
    expect(sessionAfterRename(tsv, '/docs/variants.tsv', '/docs/uniprot.tsv')).toEqual({
      ...tsv,
      suggestSaveAs: '/docs/uniprot.xlsx',
    })
  })

  it('does not touch a Save As default that merely shares the stem with an unrelated rename', () => {
    // a real .xlsx named like the converted copy's default: its own path matches, so
    // only `path` moves and the default is left alone
    const session = { ...base, path: '/docs/legacy.xlsx', suggestSaveAs: '/docs/legacy.xlsx' }
    expect(sessionAfterRename(session, '/docs/legacy.xlsx', '/docs/new.xlsx')).toEqual({
      ...session,
      path: '/docs/new.xlsx',
    })
  })

  it('follows a restored recovery copy through restoreTarget', () => {
    const session = { ...base, path: '/tmp/recovery/x.xlsx', restoreTarget: '/docs/budget.xlsx' }
    expect(sessionAfterRename(session, '/docs/budget.xlsx', '/docs/budget-final.xlsx')).toEqual({
      ...session,
      restoreTarget: '/docs/budget-final.xlsx',
    })
  })
})

describe('xlsxSiblingOf', () => {
  it('swaps the extension for .xlsx', () => {
    expect(xlsxSiblingOf('/a/b.tsv')).toBe('/a/b.xlsx')
    expect(xlsxSiblingOf('/a/b.tar.csv')).toBe('/a/b.tar.xlsx')
  })
})
