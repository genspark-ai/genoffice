import { describe, expect, it } from 'vitest'

import {
  cellFontTextStyle,
  joinManualBreaks,
  toRichTextDocument,
} from '../src/renderer/univer-sync'
import { createEditJournal, recordSetRangeValues, toSaveEdits } from '../src/renderer/edit-journal'
import type { WorkbookRichRun } from '../src/shared/desktop-api'

const plainRun = (text: string): WorkbookRichRun => ({
  text,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
})

describe('joinManualBreaks', () => {
  it('joins Alt+Enter lines with no separator, matching Excel with wrap off', () => {
    // prod ref: '010-5007-5707\n/010-8604-4376' shows as one joined line.
    expect(joinManualBreaks('010-5007-5707\n/010-8604-4376')).toBe('010-5007-5707/010-8604-4376')
    expect(joinManualBreaks('A\r\nB\rC\nD')).toBe('ABCD')
  })
})

describe('toRichTextDocument with line breaks', () => {
  it('converts \\n to Univer paragraph breaks with per-paragraph markers', () => {
    const p = toRichTextDocument('Line1\nLine2\nLine3')
    expect(p?.body?.dataStream).toBe('Line1\rLine2\rLine3\r\n')
    expect(p?.body?.paragraphs).toEqual([{ startIndex: 5 }, { startIndex: 11 }, { startIndex: 17 }])
    expect(p?.body?.sectionBreaks).toEqual([{ startIndex: 18 }])
    expect(p?.body?.textRuns).toEqual([])
  })

  it('keeps textRun offsets aligned across the 1:1 replacement', () => {
    const runs: WorkbookRichRun[] = [
      plainRun('Line1\nLine2\n'),
      { text: 'Line3', bold: false, italic: true, underline: true, strikethrough: false },
    ]
    const p = toRichTextDocument('Line1\nLine2\nLine3', runs)
    const styled = p?.body?.textRuns?.[1]
    expect(styled).toMatchObject({ st: 12, ed: 17, ts: { it: 1, ul: { s: 1 } } })
    expect(p?.body?.dataStream?.slice(styled!.st, styled!.ed)).toBe('Line3')
  })
})

describe('toRichTextDocument cell-font inheritance', () => {
  const cellStyle = {
    fontFamily: 'Meiryo UI',
    fontSize: 48,
    bold: true,
    italic: false,
    underline: false,
    strikethrough: false,
    wrapText: true,
    fontColor: '#112233',
    diagonalUp: false,
    diagonalDown: false,
  }

  it('covers a plain multiline cell with the cell font', () => {
    const p = toRichTextDocument('Line1\nLine2', [], cellFontTextStyle(cellStyle))
    expect(p?.body?.textRuns).toEqual([
      {
        st: 0,
        ed: 11,
        ts: { ff: 'Meiryo UI', fs: 48, bl: 1, cl: { rgb: '#112233' } },
      },
    ])
  })

  it('applies the cell font to runs without rPr and under formatted runs', () => {
    const runs: WorkbookRichRun[] = [
      plainRun('Title '),
      {
        text: 'small',
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        size: 36,
      },
    ]
    const p = toRichTextDocument('Title small', runs, cellFontTextStyle(cellStyle))
    expect(p?.body?.textRuns?.[0]).toEqual({
      st: 0,
      ed: 6,
      ts: { ff: 'Meiryo UI', fs: 48, bl: 1, cl: { rgb: '#112233' } },
    })
    // rPr present: its boolean flags are authoritative (bold off), while
    // family/size/color fall back to the cell font unless the run sets them.
    expect(p?.body?.textRuns?.[1]).toMatchObject({
      st: 6,
      ed: 11,
      ts: { ff: 'Meiryo UI', fs: 36, bl: 0, cl: { rgb: '#112233' } },
    })
  })

  it('does not lose a point to float underflow when scaling run sizes', () => {
    const runs: WorkbookRichRun[] = [
      { text: 'x', bold: false, italic: false, underline: false, strikethrough: false, size: 9 },
    ]
    // 9 * (6/9) is 5.999…; the shrink scale must still yield 6.
    const p = toRichTextDocument('x', runs, {}, 6 / 9)
    expect(p?.body?.textRuns?.[0]?.ts?.fs).toBe(6)
  })

  it('collapses CRLF to a single paragraph break and keeps offsets aligned', () => {
    const runs: WorkbookRichRun[] = [
      plainRun('A\r\nB\r\n'),
      { text: 'C', bold: true, italic: false, underline: false, strikethrough: false },
    ]
    const p = toRichTextDocument('A\r\nB\r\nC', runs)
    expect(p?.body?.dataStream).toBe('A\rB\rC\r\n')
    expect(p?.body?.paragraphs).toEqual([{ startIndex: 1 }, { startIndex: 3 }, { startIndex: 5 }])
    expect(p?.body?.textRuns?.[1]).toMatchObject({ st: 4, ed: 5 })
    expect(p?.body?.dataStream?.slice(4, 5)).toBe('C')
  })
})

describe('digit-leading font families', () => {
  it('escapes the family for the canvas and unescapes on journal extraction', () => {
    const p = toRichTextDocument(
      'Line1\nLine2',
      [],
      cellFontTextStyle({
        fontFamily: '12ABC',
        fontSize: 48,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        wrapText: false,
        diagonalUp: false,
        diagonalDown: false,
      }),
    )
    expect(p?.body?.textRuns?.[0]?.ts?.ff).toBe('\\31 2ABC')

    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { p } } })
    const entry = journal.cells.get('sheet-1')?.get('0:0')
    expect(entry?.rich?.[0]?.family).toBe('12ABC')
  })
})

describe('multiline round trip through the edit journal', () => {
  it('restores \\n in rich multi-run cells and keeps run boundaries', () => {
    const runs: WorkbookRichRun[] = [
      plainRun('Line1\nLine2\n'),
      { text: 'Line3', bold: false, italic: true, underline: true, strikethrough: false },
    ]
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: { 0: { p: toRichTextDocument('Line1\nLine2\nLine3', runs) } },
    })
    const entry = journal.cells.get('sheet-1')?.get('0:0')
    expect(entry?.value).toBe('Line1\nLine2\nLine3')
    expect(entry?.rich).toEqual(runs)
  })

  it('restores \\n in single-format cells (Alt+Enter, no styled runs)', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: { 0: { p: toRichTextDocument('foo\nbar\nbaz') } },
    })
    const entry = journal.cells.get('sheet-1')?.get('0:0')
    expect(entry?.value).toBe('foo\nbar\nbaz')
    expect(entry?.rich).toBeUndefined()
  })

  it('writes \\n to the save payload for editor-produced streams', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: { 0: { p: { body: { dataStream: 'edited\rin place\r\n' } } } },
    })
    const edits = toSaveEdits(journal)
    expect(edits).toHaveLength(1)
    expect(edits[0]?.value).toBe('edited\nin place')
    expect(edits[0]?.rich).toBeUndefined()
  })
})
