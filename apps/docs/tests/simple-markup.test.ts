import { describe, expect, it } from 'vitest'
import type { Block } from '@genoffice/docx-engine'
import { mergeChangeLines } from '../src/renderer/editor/margin-annotations'
import {
  blocksHaveRevisions,
  initialMarkupMode,
  revisionHiddenIn,
} from '../src/renderer/editor/revision-view'

const rev = { author: 'Ann', date: '2026-09-01T00:00:00Z' }
const para = (extra: Partial<Block> = {}): Block =>
  ({ type: 'paragraph', runs: [{ text: 'plain' }], ...extra }) as unknown as Block

describe('blocksHaveRevisions', () => {
  it('is false for plain paragraphs and tables', () => {
    const table = { rows: [[{ paras: ['a'], richParas: [{ runs: [{ text: 'a' }] }] }]] }
    expect(blocksHaveRevisions([para(), para({ table } as Partial<Block>)])).toBe(false)
  })

  it('detects run-level w:ins / w:del / w:rPrChange', () => {
    expect(blocksHaveRevisions([para({ runs: [{ text: 'x', ins: rev }] } as Partial<Block>)])).toBe(
      true,
    )
    expect(blocksHaveRevisions([para({ runs: [{ text: 'x', del: rev }] } as Partial<Block>)])).toBe(
      true,
    )
    expect(
      blocksHaveRevisions([para({ runs: [{ text: 'x', rPrChange: rev }] } as Partial<Block>)]),
    ).toBe(true)
  })

  it('detects paragraph-level revisions', () => {
    expect(blocksHaveRevisions([para({ pPrChangeInfo: rev })])).toBe(true)
    expect(blocksHaveRevisions([para({ paraMarkDel: rev })])).toBe(true)
    expect(blocksHaveRevisions([para({ moveRevision: 'to' })])).toBe(true)
    expect(blocksHaveRevisions([para({ blockRevision: { kind: 'del', ...rev } })])).toBe(true)
  })

  it('detects revisions inside tables, rows, cells and nested tables', () => {
    const cellRun = {
      rows: [[{ paras: ['x'], richParas: [{ runs: [{ text: 'x', del: rev }] }] }]],
    }
    const rowRev = { rows: [[{ paras: ['x'] }]], rowRevisions: [{ kind: 'ins', ...rev }] }
    const cellRev = { rows: [[{ paras: ['x'], cellRevision: { kind: 'del', ...rev } }]] }
    const nested = { rows: [[{ paras: [], nestedTables: [rowRev] }]] }
    for (const table of [cellRun, rowRev, cellRev, nested]) {
      expect(blocksHaveRevisions([para({ table } as Partial<Block>)])).toBe(true)
    }
  })
})

describe('initialMarkupMode', () => {
  it('opens a document with tracked changes in Simple Markup by default', () => {
    expect(initialMarkupMode('all', false, true)).toBe('simple')
  })

  it('leaves a document without revisions in the current view', () => {
    expect(initialMarkupMode('all', false, false)).toBe('all')
    expect(initialMarkupMode('simple', false, false)).toBe('simple')
  })

  it('keeps a view the user picked this session', () => {
    expect(initialMarkupMode('all', true, true)).toBe('all')
    expect(initialMarkupMode('none', true, true)).toBe('none')
    expect(initialMarkupMode('original', false, true)).toBe('original')
  })
})

describe('revisionHiddenIn', () => {
  it('hides deletions in Simple / No Markup and insertions in Original only', () => {
    expect(revisionHiddenIn('simple', 'del')).toBe(true)
    expect(revisionHiddenIn('none', 'rowDel')).toBe(true)
    expect(revisionHiddenIn('simple', 'ins')).toBe(false)
    expect(revisionHiddenIn('simple', 'pPrChange')).toBe(false)
    expect(revisionHiddenIn('original', 'ins')).toBe(true)
    expect(revisionHiddenIn('original', 'del')).toBe(false)
    expect(revisionHiddenIn('all', 'del')).toBe(false)
    expect(revisionHiddenIn('all', 'both')).toBe(false)
    expect(revisionHiddenIn('simple', 'both')).toBe(true)
  })
})

describe('mergeChangeLines', () => {
  it('merges touching lines into one bar and keeps the earliest position', () => {
    const bars = mergeChangeLines([
      { top: 40, bottom: 56, pos: 120 },
      { top: 10, bottom: 26, pos: 30 },
      { top: 24, bottom: 40, pos: 12 },
      { top: 100, bottom: 116 },
    ])
    expect(bars).toEqual([
      { top: 10, bottom: 56, pos: 12 },
      { top: 100, bottom: 116 },
    ])
  })

  it('keeps lines further than the join tolerance apart', () => {
    expect(
      mergeChangeLines([
        { top: 0, bottom: 10, pos: 1 },
        { top: 14, bottom: 24, pos: 2 },
      ]),
    ).toHaveLength(2)
  })
})
