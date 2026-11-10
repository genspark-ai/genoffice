import { describe, expect, it } from 'vitest'
import type { Block } from '@genoffice/docx-engine'
import { mergeChangeLines } from '../src/renderer/editor/margin-annotations'
import { hfTabLines, hfTabSegments, hfValueInView } from '../src/renderer/editor/hf-dom'
import {
  blocksHaveRevisions,
  documentHasRevisions,
  noteRunsInView,
  runHiddenIn,
  markupPrintsIn,
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

describe('revisionHiddenIn', () => {
  it('hides deletions in Simple / No Markup and insertions in Original only', () => {
    expect(revisionHiddenIn('simple', { kind: 'del' })).toBe(true)
    expect(revisionHiddenIn('none', { kind: 'rowDel' })).toBe(true)
    expect(revisionHiddenIn('simple', { kind: 'ins' })).toBe(false)
    expect(revisionHiddenIn('simple', { kind: 'pPrChange' })).toBe(false)
    expect(revisionHiddenIn('original', { kind: 'ins' })).toBe(true)
    expect(revisionHiddenIn('original', { kind: 'del' })).toBe(false)
    expect(revisionHiddenIn('all', { kind: 'del' })).toBe(false)
  })

  it('hides insert-then-delete text in every view but All Markup', () => {
    expect(revisionHiddenIn('simple', { kind: 'ins', overlapped: true })).toBe(true)
    expect(revisionHiddenIn('original', { kind: 'del', overlapped: true })).toBe(true)
    expect(revisionHiddenIn('none', { kind: 'ins', overlapped: true })).toBe(true)
    expect(revisionHiddenIn('all', { kind: 'ins', overlapped: true })).toBe(false)
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

describe('markupPrintsIn', () => {
  it('prints balloons and revisions only from All Markup in Print Layout', () => {
    expect(markupPrintsIn('all', 'print')).toBe(true)
    for (const mode of ['simple', 'none', 'original'] as const) {
      expect(markupPrintsIn(mode, 'print')).toBe(false)
    }
    expect(markupPrintsIn('all', 'web')).toBe(false)
  })
})

describe('documentHasRevisions', () => {
  const base = { blocks: [para()], footnotes: [], endnotes: [] }
  it('is false for a clean document', () => {
    expect(documentHasRevisions(base)).toBe(false)
  })
  it('sees footnote, header/footer and textbox revisions the body does not carry', () => {
    expect(
      documentHasRevisions({
        ...base,
        footnotes: [{ id: '1', text: 'n', richParas: [[{ text: 'n', del: rev }]] }],
      }),
    ).toBe(true)
    expect(
      documentHasRevisions({ ...base, footerParas: [{ runs: [{ text: 'f', ins: rev }] }] }),
    ).toBe(true)
    expect(
      documentHasRevisions({
        ...base,
        hfParts: {
          rId9: { text: '', hasPageNumber: false, paras: [{ runs: [{ text: 'h', del: rev }] }] },
        },
      }),
    ).toBe(true)
    expect(
      documentHasRevisions({
        ...base,
        blocks: [
          para({
            textboxes: [{ paras: [{ runs: [{ text: 't', del: rev }] }] }],
          } as unknown as Partial<Block>),
        ],
      }),
    ).toBe(true)
  })
})

describe('out-of-body runs per Markup view', () => {
  const paras = [[{ text: 'a' }, { text: 'b', del: rev }, { text: 'c', ins: rev }]]
  it('No Markup / Simple drop deletions, Original drops insertions, All keeps both', () => {
    expect(runHiddenIn('none', { del: rev })).toBe(true)
    expect(runHiddenIn('simple', { ins: rev })).toBe(false)
    expect(runHiddenIn('original', { ins: rev })).toBe(true)
    expect(runHiddenIn('original', { ins: rev, del: rev })).toBe(false)
    expect(noteRunsInView('none', paras)![0].map((r) => r.text)).toEqual(['a', 'c'])
    expect(noteRunsInView('original', paras)![0].map((r) => r.text)).toEqual(['a', 'b'])
    expect(noteRunsInView('all', paras)).toBe(paras)
  })
})

describe('header/footer strip model per Markup view', () => {
  const paras = [{ runs: [{ text: 'a' }, { text: '\tb', del: rev }, { text: 'c', ins: rev }] }]
  it('drops view-hidden runs so they neither reserve height nor place tab segments', () => {
    const root = document.createElement('div')
    root.className = 'app rev-display-none'
    document.body.append(root)
    try {
      expect(hfValueInView({ text: '', paras })!.paras![0].runs.map((r) => r.text)).toEqual([
        'a',
        'c',
      ])
      expect(hfTabSegments(paras[0])).toBeNull()
      expect(hfTabLines(paras[0])).toBeNull()
      // ptabAligns index every authored tab: a hidden tab must not hand its
      // alignment to the next visible one
      const ptab = {
        runs: [{ text: 'a\t', del: rev }, { text: 'b\tc' }],
        tabStops: [{ pos: 1440, val: 'left' as const }],
        ptabAligns: ['right' as const, undefined],
      }
      expect(hfTabSegments(ptab)?.segments[0].left).toHaveProperty('px')
      expect(hfTabLines(ptab)?.[0].segments[0].left).toHaveProperty('px')
      // the view-filtered part keeps the same re-indexed alignments
      const inView = hfValueInView({ text: '', paras: [ptab] })!.paras![0]
      expect(inView.ptabAligns).toEqual([undefined])
      expect(hfTabSegments(inView)?.segments[0].left).toHaveProperty('px')
      root.className = 'app'
      expect(hfTabSegments(paras[0])?.segments).toHaveLength(1)
      expect(hfTabSegments(ptab)?.segments[0].left).toHaveProperty('pct')
    } finally {
      root.remove()
    }
  })

  it('lays multi-line paragraphs out from the shown runs only', () => {
    const multi = {
      runs: [{ text: 'a\tb' }, { text: '\nx\ty', del: rev }, { text: '\nc', ins: rev }],
    }
    const root = document.createElement('div')
    root.className = 'app rev-display-none'
    document.body.append(root)
    try {
      const lines = hfTabLines(multi)!
      expect(lines).toHaveLength(2)
      expect(lines[1].lead.map((r) => r.text)).toEqual(['c'])
      root.className = 'app rev-display-original'
      expect(hfTabLines(multi)!.map((l) => l.segments.length)).toEqual([1, 1])
      root.className = 'app'
      expect(hfTabLines(multi)).toHaveLength(3)
    } finally {
      root.remove()
    }
  })
})
