/**
 * Pure halves of the F7 / Review ▸ Spelling engine: string-cell collection
 * (skip formulas / letterless / withheld cells, row-major order), batch
 * dedupe + truncation, checker-reply parsing (hallucination guards), the
 * white-list, replacement casing, and the batch runner against a scripted
 * chat, plus the localStorage preference/dictionary round-trip.
 */
import { describe, expect, it } from 'vitest'
import {
  buildTextBatches,
  collectSheetTextCells,
  isWhiteListed,
  parseSpellReply,
  replaceSuggestion,
  runSheetSpellCheck,
  type SpellCellRef,
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

describe('collectSheetTextCells', () => {
  it('keeps plain string cells only: no formulas, numbers, empties or letterless text', () => {
    const cells = collectSheetTextCells(
      sheet({
        0: { 0: cell(0, 0, 'Helo world'), 1: cell(0, 1, 'jumsp', '=SUM(1)'), 2: cell(0, 2, 42) },
        1: { 0: cell(1, 0, '   '), 1: cell(1, 1, '123-456'), 2: cell(1, 2, 'recieved') },
      }),
      'sheet-1',
      'Data',
    )
    expect(cells.map((entry) => entry.address)).toEqual(['A1', 'C2'])
    expect(cells[0]?.text).toBe('Helo world')
  })

  it('walks rows before columns in numeric order, not string-key order', () => {
    const cells = collectSheetTextCells(
      sheet({
        10: { 2: cell(10, 2, 'ten') },
        2: { 1: cell(2, 1, 'two') },
      }),
      'sheet-1',
      'Data',
    )
    expect(cells.map((entry) => entry.address)).toEqual(['B3', 'C11'])
  })

  it('never collects a withheld cell', () => {
    const cells = collectSheetTextCells(
      sheet({ 0: { 0: cell(0, 0, 'secret name'), 1: cell(0, 1, 'visible') } }),
      'sheet-1',
      'Data',
      (sheetId, row, column) => sheetId === 'sheet-1' && row === 0 && column === 0,
    )
    expect(cells.map((entry) => entry.address)).toEqual(['B1'])
  })

  it('returns nothing for a sheet without cell data', () => {
    expect(collectSheetTextCells({ sheets: {} }, 'missing', 'X')).toEqual([])
  })
})

describe('buildTextBatches', () => {
  const ref = (address: string, text: string): SpellCellRef => ({
    sheetId: 's',
    sheetName: 'S',
    address,
    row: 0,
    column: 0,
    text,
  })

  it('dedupes texts and maps every duplicate cell to the one text', () => {
    const { batches, cellsByText } = buildTextBatches([
      ref('A1', 'same typo'),
      ref('B1', 'same typo'),
      ref('C1', 'other'),
    ])
    expect(batches).toEqual([['same typo', 'other']])
    expect(cellsByText.get('same typo')?.map((entry) => entry.address)).toEqual(['A1', 'B1'])
  })

  it('caps texts and reports truncation, still mapping overflow cells', () => {
    const many = Array.from({ length: 5 }, (_, index) => ref(`A${index + 1}`, `text ${index}`))
    const { batches, truncated } = buildTextBatches(many, 2, 3)
    expect(batches).toEqual([['text 0', 'text 1'], ['text 2']])
    expect(truncated).toBe(true)
  })
})

describe('parseSpellReply', () => {
  const texts = ['the qick brown fox', 'no problems here', 'a seperate issue']

  it('reads a clean minified reply', () => {
    const parsed = parseSpellReply(
      '{"findings":[{"i":1,"w":"qick","s":["quick"]},{"i":3,"w":"seperate","s":["separate"]}]}',
      texts,
    )
    expect(parsed?.findings).toEqual([
      { index: 1, word: 'qick', suggestions: ['quick'] },
      { index: 3, word: 'seperate', suggestions: ['separate'] },
    ])
  })

  it('tolerates code fences and stray prose around the JSON', () => {
    const parsed = parseSpellReply(
      'Sure!\n```json\n{"findings":[{"i":1,"w":"qick","s":["quick"]}]}\n```',
      texts,
    )
    expect(parsed?.findings).toEqual([{ index: 1, word: 'qick', suggestions: ['quick'] }])
  })

  it('drops findings whose word does not occur in the flagged text', () => {
    const parsed = parseSpellReply(
      '{"findings":[{"i":1,"w":"helo","s":["hello"]},{"i":2,"w":"problemz","s":["problem"]},{"i":9,"w":"qick","s":["quick"]},{"i":0,"w":"x","s":["y"]}]}',
      texts,
    )
    expect(parsed?.findings).toEqual([])
  })

  it('keeps suggestion lists sane: drops empties, self-repeats, caps length', () => {
    const parsed = parseSpellReply(
      '{"findings":[{"i":1,"w":"qick","s":["quick","quick","","qick","fast","swift","zap","boom"]}]}',
      texts,
    )
    expect(parsed?.findings[0]?.suggestions).toEqual(['quick', 'fast', 'swift', 'zap', 'boom'])
  })

  it('returns null for non-JSON replies and replies without a findings array', () => {
    expect(parseSpellReply('all good', texts)).toBeNull()
    expect(parseSpellReply('{"result":[]}', texts)).toBeNull()
    expect(parseSpellReply('{"findings": "nope"}', texts)).toBeNull()
  })
})

describe('white-list and replacement', () => {
  it('matches white-list entries case-insensitively', () => {
    const learned = new Set(['acme'])
    const ignored = new Set(['IgnoredWord'])
    expect(isWhiteListed('ACME', learned, ignored)).toBe(true)
    expect(isWhiteListed('ignoredword', learned, ignored)).toBe(true)
    expect(isWhiteListed('typo', learned, ignored)).toBe(false)
  })

  it('replaces the flagged occurrence and follows obvious casing', () => {
    expect(replaceSuggestion('a qick fix', 'qick', 'quick')).toBe('a quick fix')
    expect(replaceSuggestion('Qick start', 'Qick', 'quick')).toBe('Quick start')
    expect(replaceSuggestion('QICK START', 'QICK', 'quick')).toBe('QUICK START')
    expect(replaceSuggestion('gone', 'missing', 'x')).toBe('gone')
    // only the flagged occurrence: the second one is the user's to decide
    expect(replaceSuggestion('one qick, two qick', 'qick', 'quick')).toBe('one quick, two qick')
  })
})

describe('runSheetSpellCheck', () => {
  const ref = (address: string, text: string): SpellCellRef => ({
    sheetId: 's',
    sheetName: 'S',
    address,
    row: 0,
    column: 0,
    text,
  })
  const okChat = (body: string) => async () => ({ ok: true, content: body })

  it('expands findings to every cell sharing the text and reports progress', async () => {
    const progress: Array<[number, number]> = []
    const seen: string[] = []
    const result = await runSheetSpellCheck({
      cells: [ref('A1', 'a qick fox'), ref('B9', 'a qick fox'), ref('C1', 'fine')],
      learnedWords: new Set(),
      ignoredWords: new Set(),
      requestChat: async (_system, user) => {
        seen.push(user)
        return okChat('{"findings":[{"i":1,"w":"qick","s":["quick"]}]}')()
      },
      onProgress: (done, total) => progress.push([done, total]),
    })
    expect(result).toEqual({ status: 'done', truncated: false })
    // one batch: the duplicate text was sent once
    expect(seen).toHaveLength(1)
    expect(JSON.parse(seen[0]!).texts).toEqual(['a qick fox', 'fine'])
    // the finding resolved to both cells that carry the text
    const findings: SpellCellRef[] = []
    await runSheetSpellCheck({
      cells: [ref('A1', 'a qick fox'), ref('B9', 'a qick fox')],
      learnedWords: new Set(),
      ignoredWords: new Set(),
      requestChat: okChat('{"findings":[{"i":1,"w":"qick","s":["quick"]}]}'),
      onFindings: (batch) => findings.push(...batch.map((finding) => finding.cell)),
    })
    expect(findings.map((entry) => entry.address)).toEqual(['A1', 'B9'])
    expect(progress.at(-1)).toEqual([2, 2])
  })

  it('filters white-listed words out of model findings', async () => {
    const findings: string[] = []
    await runSheetSpellCheck({
      cells: [ref('A1', 'acme qick'), ref('B1', 'whitel listed')],
      learnedWords: new Set(['acme']),
      ignoredWords: new Set(['whitel']),
      requestChat: okChat(
        '{"findings":[{"i":1,"w":"acme","s":["a cme"]},{"i":2,"w":"whitel","s":["white"]},{"i":1,"w":"qick","s":["quick"]}]}',
      ),
      onFindings: (batch) => findings.push(...batch.map((finding) => finding.word)),
    })
    expect(findings).toEqual(['qick'])
  })

  it('sends the known words with each batch so the model can skip them', async () => {
    const seen: string[] = []
    await runSheetSpellCheck({
      cells: [ref('A1', 'anything')],
      learnedWords: new Set(['acme']),
      ignoredWords: new Set(['zgok']),
      requestChat: async (_system, user) => {
        seen.push(user)
        return { ok: true, content: '{"findings":[]}' }
      },
    })
    expect(JSON.parse(seen[0]!).knownCorrectWords).toEqual(['acme', 'zgok'])
  })

  it('stops between batches once shouldContinue turns false', async () => {
    let calls = 0
    const result = await runSheetSpellCheck({
      cells: [ref('A1', 'one'), ref('B1', 'two'), ref('C1', 'three')],
      learnedWords: new Set(),
      ignoredWords: new Set(),
      requestChat: async () => {
        calls += 1
        return { ok: true, content: '{"findings":[]}' }
      },
      batchSize: 1,
      // the check runs before each batch: two batches go out, the third never starts
      shouldContinue: () => calls < 2,
    })
    expect(result).toEqual({ status: 'canceled' })
    expect(calls).toBe(2)
  })

  it('surfaces a failed chat request as an error result', async () => {
    const result = await runSheetSpellCheck({
      cells: [ref('A1', 'text')],
      learnedWords: new Set(),
      ignoredWords: new Set(),
      requestChat: async () => ({ ok: false, error: 'no provider' }),
    })
    expect(result).toEqual({ status: 'error', error: 'no provider' })
  })

  it('carries truncation when more unique texts exist than the cap', async () => {
    const cells = Array.from({ length: 4 }, (_, index) => ref(`A${index + 1}`, `t${index}`))
    const result = await runSheetSpellCheck({
      cells,
      learnedWords: new Set(),
      ignoredWords: new Set(),
      requestChat: okChat('{"findings":[]}'),
      maxTexts: 3,
    })
    expect(result).toEqual({ status: 'done', truncated: true })
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
