import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { FileIndexStore } from '../src/main/file-index/store'

let dir: string
let store: FileIndexStore

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-index-'))
  store = new FileIndexStore(join(dir, 'index.db'))
})
afterEach(() => {
  store.close()
  rmSync(dir, { recursive: true, force: true })
})

const meta = (path: string) => ({ path, mtimeMs: 1, sizeBytes: 10 })

describe('FileIndexStore', () => {
  it('finds CJK content inside a longer run and ranks name hits first', () => {
    store.upsert(
      meta('/docs/plan.docx'),
      '\u672c\u5e74\u5ea6\u65b0\u80fd\u6e90\u6c7d\u8f66\u5e02\u573a\u8c03\u7814\u62a5\u544a',
      'ok',
    )
    store.upsert(meta('/docs/\u62a5\u544a\u6c47\u603b.xlsx'), 'nothing relevant', 'ok')
    store.upsert(meta('/docs/other.pdf'), 'quarterly numbers', 'ok')
    const r = store.search('\u62a5\u544a')
    expect(r.total).toBe(2)
    expect(r.hits.map((h) => h.name)).toEqual(['\u62a5\u544a\u6c47\u603b.xlsx', 'plan.docx'])
    expect(r.hits[0]!.needles).toEqual(['\u62a5\u544a'])
    expect(r.hits[1]!.snippet!.some((p) => p.hit && p.text === '\u62a5\u544a')).toBe(true)
  })

  it('does not match bigram collisions across a substring boundary', () => {
    store.upsert(meta('/a.md'), '\u65b0\u80fd\u6e90', 'ok')
    expect(store.search('\u65b0\u6e90').total).toBe(0)
  })

  it('answers single-character and mixed lone-character queries', () => {
    store.upsert(meta('/q.md'), '\u62a5\u7b2c1\u5b63\u5ea6\u603b\u7ed3', 'ok')
    expect(store.search('\u544a').total).toBe(0)
    expect(store.search('\u7ed3').total).toBe(1)
    expect(store.search('\u7b2c1\u5b63\u5ea6').total).toBe(1)
  })

  it('prefix-matches Latin words and honours the extension filter', () => {
    store.upsert(meta('/r/annual-report.docx'), 'the fiscal year', 'ok')
    store.upsert(meta('/r/notes.md'), 'reporting lines', 'ok')
    expect(store.search('rep').total).toBe(2)
    expect(store.search('rep', { exts: ['md'] }).hits.map((h) => h.ext)).toEqual(['md'])
    expect(store.search('report fiscal').total).toBe(1)
  })

  it('falls back to any-term when no file has every term', () => {
    store.upsert(meta('/x.md'), 'alpha only', 'ok')
    store.upsert(meta('/y.md'), 'beta only', 'ok')
    expect(store.search('alpha beta').total).toBe(2)
  })

  it('reaches a phrase typed one character too far through token coverage', () => {
    store.upsert(
      meta('/t/\u5927\u6a21\u578b\u6280\u672f\u6f14\u8fdb.pdf'),
      'TECH BRIEFING \u5927\u6a21\u578b\u6280\u672f\u6f14\u8fdb\u8d8b\u52bf',
      'ok',
    )
    store.upsert(meta('/t/other.md'), '\u6280\u672f\u76d8\u70b9', 'ok')
    const r = store.search('\u5927\u6a21\u578b\u6280\u672f\u6211')
    expect(r.hits.map((h) => h.name)).toEqual(['\u5927\u6a21\u578b\u6280\u672f\u6f14\u8fdb.pdf'])
    expect(r.hits[0]!.needles).toEqual([
      '\u5927\u6a21',
      '\u6a21\u578b',
      '\u578b\u6280',
      '\u6280\u672f',
    ])
    expect(r.hits[0]!.snippet!.filter((p) => p.hit).map((p) => p.text)).toEqual([
      '\u5927\u6a21\u578b\u6280\u672f',
    ])
  })

  it('ranks the file covering more of the query first and stays out of short queries', () => {
    store.upsert(meta('/c/full.md'), '\u5927\u6a21\u578b\u7684\u6280\u672f', 'ok')
    store.upsert(meta('/c/half.md'), '\u6280\u672f\u8def\u7ebf', 'ok')
    const r = store.search('\u5927\u6a21\u578b\u6280\u672f')
    expect(r.hits.map((h) => h.name)).toEqual(['full.md'])
    expect(store.search('\u6a21\u578b\u5927').total).toBe(0)
  })

  it('reaches a sentence-shaped query through the words it shares with the file', () => {
    store.upsert(
      meta('/c/tencent.docx'),
      '\u817e\u8baf\u662f\u4e2d\u56fd\u7684\u6e38\u620f\u516c\u53f8',
      'ok',
    )
    store.upsert(meta('/c/beijing.docx'), '\u4e2d\u56fd\u7684\u516c\u53f8', 'ok')
    store.upsert(meta('/c/google.docx'), '\u5e02\u503c\u6700\u9ad8\u7684\u516c\u53f8', 'ok')
    const r = store.search('\u4e2d\u56fd\u505a\u6e38\u620f\u7684\u516c\u53f8')
    expect(r.hits.map((h) => h.name)).toEqual(['tencent.docx', 'beijing.docx'])
    expect(r.hits[0]!.needles).toEqual(['\u4e2d\u56fd', '\u6e38\u620f', '\u516c\u53f8'])
  })

  it('finds a folder name run together with unrelated characters', () => {
    store.upsert(
      meta('/docs/\u5927\u6570\u636e\u5e73\u53f0\u8fc1\u79fb\u65b9\u6848/report.docx'),
      null,
      'name-only',
    )
    store.upsert(meta('/docs/\u6570\u636e/a.docx'), null, 'name-only')
    expect(store.search('\u5927\u6570\u636e\u662f\u7684').hits.map((h) => h.name)).toEqual([
      'report.docx',
    ])
    expect(store.search('\u6570\u636e\u662f\u7684').total).toBe(0)
  })

  it('relaxes each term on its own and puts name or folder hits before body hits', () => {
    store.upsert(
      meta('/docs/\u5927\u6570\u636e\u5e73\u53f0\u8fc1\u79fb/report.docx'),
      null,
      'name-only',
    )
    store.upsert(
      meta('/docs/\u5927\u6570\u636e\u5e73\u53f0\u8fc1\u79fb/slides.pptx'),
      null,
      'name-only',
    )
    store.upsert(meta('/docs/gen/brief.docx'), '\u5317\u6781\u718a\u6816\u606f\u5730', 'ok')
    store.upsert(meta('/docs/gen/limits.docx'), '\u6700\u5927\u6570\u91cf\u9650\u5236', 'ok')
    const r = store.search('\u5927\u6570\u636e\u662f\u7684\u8fc1\u79fb \u5317\u6781\u718a')
    expect(r.total).toBe(3)
    expect(r.hits.map((h) => h.name)).toEqual(['report.docx', 'slides.pptx', 'brief.docx'])
    expect(r.hits[2]!.needles).toEqual(['\u5317\u6781\u718a'])
    const byNeedle = store.search('\u5927\u6570')
    expect(byNeedle.hits.map((h) => h.name)).toEqual(['report.docx', 'slides.pptx', 'limits.docx'])
  })

  it('indexes name-only entries and removes files', () => {
    store.upsert(meta('/big/huge-deck.pptx'), null, 'name-only')
    expect(store.search('huge').hits[0]!.snippet).toBeNull()
    expect(store.listAll().get('/big/huge-deck.pptx')?.status).toBe('name-only')
    store.remove(['/big/huge-deck.pptx'])
    expect(store.search('huge').total).toBe(0)
    expect(store.count()).toBe(0)
  })

  it('replaces an existing path on upsert', () => {
    store.upsert(meta('/v.md'), 'first version', 'ok')
    store.upsert({ path: '/v.md', mtimeMs: 2, sizeBytes: 12 }, 'second version', 'ok')
    expect(store.search('first').total).toBe(0)
    expect(store.search('second').total).toBe(1)
    expect(store.count()).toBe(1)
  })

  it('searches folder names, also with lone-character queries', () => {
    store.upsert(meta('/Users/me/\u9879\u76ee\u7532/a.docx'), 'x', 'ok')
    expect(store.search('\u9879\u76ee\u7532').total).toBe(1)
    expect(store.search('\u7532').total).toBe(1)
    store.upsert(meta('/Users/me/\u7b2c1\u5b63\u5ea6/b.docx'), 'x', 'ok')
    expect(store.search('\u7b2c1\u5b63').hits.map((h) => h.name)).toEqual(['b.docx'])
  })

  it('applies the type filter before falling back to any-term matches', () => {
    store.upsert(meta('/f/both.pdf'), 'alpha beta', 'ok')
    store.upsert(meta('/f/one.md'), 'alpha only', 'ok')
    expect(store.search('alpha beta').hits.map((h) => h.name)).toEqual(['both.pdf'])
    expect(store.search('alpha beta', { exts: ['md'] }).hits.map((h) => h.name)).toEqual(['one.md'])
  })
})
