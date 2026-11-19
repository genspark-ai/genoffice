/**
 * Pure halves of the F7 / Review ▸ Spelling engine: string-cell collection
 * over the whole workbook (skip formulas / letterless / withheld cells,
 * sheet order, row-major), tokenising, the white-list, replacement casing
 * and word boundaries, rich-text run preservation, and the synchronous run
 * against a scripted frame engine, plus the localStorage
 * preference/dictionary round-trip.
 */
import { describe, expect, it } from 'vitest'
import {
  collectWorkbookTextCells,
  indexOfWordOf,
  isWhiteListed,
  replaceInRichDoc,
  replaceSuggestion,
  runSheetSpellCheck,
  tokenizeCellText,
  type SpellCellRef,
  type SpellEngine,
  type SpellSnapshot,
} from '../src/renderer/spell-check'
import {
  loadLearnedWords,
  saveLearnedWords,
  spellcheckEnabled,
  storeSpellcheckEnabled,
} from '../src/renderer/spellcheck-pref'

const cell = (row: number, column: number, v: unknown, f?: unknown) => ({ v, f })
const sheet = (cellData: Record<string, Record<string, ReturnType<typeof cell>>>) => ({
  sheets: { 'sheet-1': { name: 'Data', cellData } },
})

describe('collectWorkbookTextCells', () => {
  it('keeps plain string cells only: no formulas, numbers, empties or letterless text', () => {
    const { cells } = collectWorkbookTextCells(
      sheet({
        0: { 0: cell(0, 0, 'Helo world'), 1: cell(0, 1, 'jumsp', '=SUM(1)'), 2: cell(0, 2, 42) },
        1: { 0: cell(1, 0, '   '), 1: cell(1, 1, '123-456'), 2: cell(1, 2, 'recieved') },
      }),
      [{ sheetId: 'sheet-1', sheetName: 'Data' }],
    )
    expect(cells.map((entry) => entry.address)).toEqual(['A1', 'C2'])
    expect(cells[0]?.text).toBe('Helo world')
  })

  it('walks sheets in the given order, rows before columns in numeric order', () => {
    const snapshot: SpellSnapshot = {
      sheets: {
        'sheet-1': {
          name: 'One',
          cellData: { 10: { 2: cell(10, 2, 'ten') }, 2: { 1: cell(2, 1, 'two') } },
        },
        'sheet-2': { name: 'Two', cellData: { 0: { 0: cell(0, 0, 'also here') } } },
      },
    }
    const { cells } = collectWorkbookTextCells(snapshot, [
      { sheetId: 'sheet-2', sheetName: 'Two' },
      { sheetId: 'sheet-1', sheetName: 'One' },
    ])
    expect(cells.map((entry) => `${entry.sheetName}!${entry.address}`)).toEqual([
      'Two!A1',
      'One!B3',
      'One!C11',
    ])
  })

  it('never collects a withheld cell', () => {
    const { cells } = collectWorkbookTextCells(
      sheet({ 0: { 0: cell(0, 0, 'secret name'), 1: cell(0, 1, 'visible') } }),
      [{ sheetId: 'sheet-1', sheetName: 'Data' }],
      (sheetId, row, column) => sheetId === 'sheet-1' && row === 0 && column === 0,
    )
    expect(cells.map((entry) => entry.address)).toEqual(['B1'])
  })

  it('caps cells and reports truncation', () => {
    const many: SpellSnapshot = {
      sheets: {
        'sheet-1': {
          name: 'Data',
          cellData: Object.fromEntries(
            Array.from({ length: 5 }, (_, row) => [
              String(row),
              { 0: cell(row, 0, `text ${row}`) },
            ]),
          ),
        },
      },
    }
    const { cells, truncated } = collectWorkbookTextCells(
      many,
      [{ sheetId: 'sheet-1', sheetName: 'Data' }],
      undefined,
      3,
    )
    expect(cells).toHaveLength(3)
    expect(truncated).toBe(true)
  })

  it('returns nothing for a sheet without cell data', () => {
    expect(
      collectWorkbookTextCells({ sheets: {} }, [{ sheetId: 'missing', sheetName: 'X' }]).cells,
    ).toEqual([])
  })
})

describe('tokenizeCellText', () => {
  it('splits on non-letters, keeps the first-seen casing and order without repeats', () => {
    expect(tokenizeCellText('The the THE qick, qick! brown')).toEqual(['The', 'qick', 'brown'])
  })

  it('keeps apostrophes and dashes inside a word, drops single letters and digits', () => {
    expect(tokenizeCellText("don't A1 42 co-op x")).toEqual(["don't", 'co-op'])
  })

  it('caps word length: a 65-letter run is a code fragment, not a word', () => {
    expect(tokenizeCellText('a'.repeat(65))).toEqual([])
  })
})

describe('white-list and replacement', () => {
  it('matches white-list entries case-insensitively', () => {
    expect(isWhiteListed('Qick', new Set(['qick']), new Set())).toBe(true)
    expect(isWhiteListed('qick', new Set(['QICK']), new Set())).toBe(true)
    expect(isWhiteListed('qick', new Set(), new Set(['Qick']))).toBe(true)
    expect(isWhiteListed('qick', new Set(), new Set())).toBe(false)
  })

  it('replaces the flagged occurrence and follows obvious casing', () => {
    expect(replaceSuggestion('the qick fox', 'qick', 'quick')).toBe('the quick fox')
    expect(replaceSuggestion('the Qick fox', 'Qick', 'quick')).toBe('the Quick fox')
    expect(replaceSuggestion('the QICK fox', 'QICK', 'quick')).toBe('the QUICK fox')
  })

  it('replaces only a standalone occurrence, never a substring of a longer word', () => {
    expect(indexOfWordOf('the cat toy', 'cat')).toBe(4)
    expect(indexOfWordOf('concatenate cats', 'cat')).toBe(-1)
    expect(replaceSuggestion('the cat toy', 'cat', 'dog')).toBe('the dog toy')
    expect(replaceSuggestion('concatenate cats', 'cat', 'dog')).toBe('concatenate cats')
  })
})

describe('replaceInRichDoc', () => {
  it('rewrites the dataStream and shifts the runs after the replacement', () => {
    const doc = replaceInRichDoc(
      {
        body: {
          dataStream: 'the qick fox',
          textRuns: [
            { st: 0, ed: 8, ts: { bold: true } },
            { st: 8, ed: 12 },
          ],
        },
      },
      'qick',
      'quick',
    )
    expect(doc?.body?.dataStream).toBe('the quick fox')
    expect(doc?.body?.textRuns).toEqual([
      { st: 0, ed: 9, ts: { bold: true } },
      { st: 9, ed: 13 },
    ])
  })

  it('re-splits runs that cross the replaced span so coverage stays contiguous', () => {
    // the word sits at 4..8; run A covers up to 6, run B starts at 6: the
    // replacement joins run B, run A is trimmed to the span's start
    const doc = replaceInRichDoc(
      {
        body: {
          dataStream: 'the qick fox',
          textRuns: [
            { st: 0, ed: 6 },
            { st: 6, ed: 12 },
          ],
        },
      },
      'qick',
      'quick',
    )
    expect(doc?.body?.dataStream).toBe('the quick fox')
    expect(doc?.body?.textRuns).toEqual([
      { st: 0, ed: 4 },
      { st: 4, ed: 13 },
    ])
  })

  it('drops a run that lived wholly inside the replaced word', () => {
    const doc = replaceInRichDoc(
      {
        body: {
          dataStream: 'the qick fox',
          textRuns: [
            { st: 0, ed: 4 },
            { st: 5, ed: 7 },
            { st: 8, ed: 12 },
          ],
        },
      },
      'qick',
      'q',
    )
    expect(doc?.body?.dataStream).toBe('the q fox')
    expect(doc?.body?.textRuns).toEqual([
      { st: 0, ed: 4 },
      { st: 5, ed: 9 },
    ])
  })

  it('shifts later runs left through a longer replacement', () => {
    const doc = replaceInRichDoc(
      {
        body: {
          dataStream: 'aa qick bb',
          textRuns: [
            { st: 0, ed: 3 },
            { st: 3, ed: 7 },
            { st: 7, ed: 10 },
          ],
        },
      },
      'qick',
      'quicker',
    )
    expect(doc?.body?.dataStream).toBe('aa quicker bb')
    expect(doc?.body?.textRuns).toEqual([
      { st: 0, ed: 3 },
      { st: 3, ed: 10 },
      { st: 10, ed: 13 },
    ])
  })

  it('returns null when the word does not occur standalone', () => {
    expect(replaceInRichDoc({ body: { dataStream: 'scatter' } }, 'cat', 'dog')).toBeNull()
    expect(replaceInRichDoc({ body: {} }, 'cat', 'dog')).toBeNull()
  })
})

describe('runSheetSpellCheck', () => {
  const ref = (address: string, text: string, sheetId = 'sheet-1'): SpellCellRef => ({
    sheetId,
    sheetName: 'Data',
    address,
    row: 0,
    column: 0,
    text,
  })
  /** an engine whose dictionary flags exactly the words in `wrong` */
  const engineOf = (wrong: readonly string[], calls: string[] = []): SpellEngine => ({
    isWordMisspelled(word) {
      calls.push(word)
      return wrong.includes(word)
    },
    getWordSuggestions(word) {
      return [`fix-${word}`, `fix-${word}`, word, 'other']
    },
  })

  it('flags misspelled words per cell with as-written casing and capped suggestions', () => {
    const findings = runSheetSpellCheck(
      [ref('A1', 'the Qick fox'), ref('B1', 'all good')],
      engineOf(['Qick']),
      new Set(),
      new Set(),
    )
    expect(findings).toHaveLength(1)
    expect(findings[0]?.word).toBe('Qick')
    expect(findings[0]?.suggestions).toEqual(['fix-Qick', 'other'])
  })

  it('asks the engine with the spelling as written: a proper noun is not lower-cased', () => {
    const calls: string[] = []
    runSheetSpellCheck([ref('A1', 'to London')], engineOf(['london'], calls), new Set(), new Set())
    expect(calls).toEqual(['to', 'London'])
  })

  it('asks the engine once per distinct word, not once per cell', () => {
    const calls: string[] = []
    runSheetSpellCheck(
      [ref('A1', 'qick qick'), ref('B1', 'qick again')],
      engineOf(['qick'], calls),
      new Set(),
      new Set(),
    )
    expect(calls.filter((word) => word === 'qick')).toHaveLength(1)
  })

  it('drops white-listed words before the engine sees them', () => {
    const calls: string[] = []
    const findings = runSheetSpellCheck(
      [ref('A1', 'qick name')],
      engineOf(['qick', 'name'], calls),
      new Set(['name']),
      new Set(),
    )
    // 'name' is white-listed: only 'qick' reaches the engine
    expect(findings.map((finding) => finding.word)).toEqual(['qick'])
    expect(calls).toEqual(['qick'])
  })
})

describe('spellcheck preference and dictionary storage', () => {
  const storage = () => {
    const map = new Map<string, string>()
    return {
      getItem: (key: string) => map.get(key) ?? null,
      setItem: (key: string, value: string) => void map.set(key, value),
    }
  }

  it('defaults on and stores the off choice (docs spellcheck-pref shape)', () => {
    const s = storage()
    expect(spellcheckEnabled(s)).toBe(true)
    storeSpellcheckEnabled(false, s)
    expect(spellcheckEnabled(s)).toBe(false)
    storeSpellcheckEnabled(true, s)
    expect(spellcheckEnabled(s)).toBe(true)
  })

  it('round-trips learned words, dedupes, caps and survives corrupt JSON', () => {
    const s = storage()
    saveLearnedWords(['acme', 'acme'], s)
    expect(loadLearnedWords(s)).toEqual(['acme'])
    const many = Array.from({ length: 600 }, (_, index) => `w${index}`)
    saveLearnedWords(many, s)
    expect(loadLearnedWords(s)).toHaveLength(500)
    expect(loadLearnedWords(s)[0]).toBe('w100')
    s.setItem('aisheets.spell-dictionary', '{not json')
    expect(loadLearnedWords(s)).toEqual([])
    s.setItem('aisheets.spell-dictionary', '[1, null, "ok"]')
    expect(loadLearnedWords(s)).toEqual(['ok'])
  })
})
