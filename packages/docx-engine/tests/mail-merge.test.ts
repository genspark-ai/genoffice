import { describe, expect, it } from 'vitest'
import type { Block } from '../src/types'
import {
  compareOperands,
  mergeBlockTree,
  mergeDataFromRows,
  mergeFieldName,
  mergeFieldText,
  mergeRecords,
  type MergeData,
} from '../src/mail-merge'

/**
 * Mail merge, engine half: substitute MERGEFIELD runs with a recipient row's
 * values and honour Word's control fields (NEXT / SKIPIF / IF). Requested in
 * genoffice#1984; the field itself survives a round-trip via the parse whitelist, this
 * is what gives it a value.
 */

/** the fields every parsed Block carries; the merge reads neither, but the
 *  signature does, so the fixtures carry them */
const blk = (runs: unknown[]): Block =>
  ({ id: '', docxIndex: 0, originalXml: '', type: 'paragraph', runs }) as unknown as Block

const p = (text: string) => blk([{ text }])

/** "Dear «Name»," as Word stores it once the parse whitelist keeps the field */
const mergeName = (name = 'Name') => blk([{ text: `«${name}»`, instrField: `MERGEFIELD ${name}` }])

const control = (instr: string) => blk([{ text: '', instrField: instr }])

/** the text of each block, so an assertion reads as a list of lines */
const lines = (blocks: readonly { runs?: readonly { text?: string }[] }[]): string[] =>
  blocks.map((b) => (b.runs ?? []).map((r) => r.text ?? '').join(''))

describe('mergeFieldName', () => {
  it('reads a bare name and stops at the first switch', () => {
    expect(mergeFieldName('MERGEFIELD CustomerName \\* MERGEFORMAT')).toBe('CustomerName')
  })

  it('reads a quoted name whole, so a header with a space survives', () => {
    expect(mergeFieldName('MERGEFIELD "Postal Address" \\* MERGEFORMAT')).toBe('Postal Address')
  })

  it('returns null for another keyword and for a name-less MERGEFIELD', () => {
    expect(mergeFieldName('PAGE')).toBeNull()
    expect(mergeFieldName('MERGEFIELD')).toBeNull()
  })
})

describe('mergeFieldText', () => {
  const rec = { First: 'ada', Last: 'LOVELACE', Empty: '' }

  it('applies \\b and \\f only around a non-empty value', () => {
    expect(mergeFieldText('MERGEFIELD First \\f " "', rec)).toBe('ada ')
    expect(mergeFieldText('MERGEFIELD Empty \\b "(" \\f ")"', rec)).toBe('')
    expect(mergeFieldText('MERGEFIELD Last \\b "Ms. "', rec)).toBe('Ms. LOVELACE')
  })

  it('applies the case switches and ignores MERGEFORMAT', () => {
    expect(mergeFieldText('MERGEFIELD First \\* Upper', rec)).toBe('ADA')
    expect(mergeFieldText('MERGEFIELD Last \\* Lower', rec)).toBe('lovelace')
    expect(mergeFieldText('MERGEFIELD Last \\* FirstCap \\* MERGEFORMAT', rec)).toBe('Lovelace')
    expect(mergeFieldText('MERGEFIELD Last \\* Caps', { Last: 'van der berg' })).toBe(
      'Van Der Berg',
    )
  })
})

describe('mergeDataFromRows', () => {
  it('reads a header row and builds one record per data row', () => {
    const data = mergeDataFromRows([
      ['Name', 'City'],
      ['Ada', 'London'],
      ['Grace', 'NY'],
    ])
    expect(data.fields).toEqual(['Name', 'City'])
    expect(data.records).toEqual([
      { Name: 'Ada', City: 'London' },
      { Name: 'Grace', City: 'NY' },
    ])
  })

  it('keeps a value that contains a comma (the CSV parser kept it whole)', () => {
    const data = mergeDataFromRows([['Name'], ['Doe, Jane']])
    expect(data.records[0]).toEqual({ Name: 'Doe, Jane' })
  })

  it('trims header whitespace but not the values', () => {
    const data = mergeDataFromRows([[' Name '], ['  padded  ']])
    expect(data.fields).toEqual(['Name'])
    expect(data.records[0]).toEqual({ Name: '  padded  ' })
  })

  it('pads a short row with empty strings rather than leaving holes', () => {
    const data = mergeDataFromRows([['A', 'B', 'C'], ['1']])
    expect(data.records[0]).toEqual({ A: '1', B: '', C: '' })
  })

  it('reports a header-only file as no records', () => {
    const data = mergeDataFromRows([['A', 'B']])
    expect(data.fields).toEqual(['A', 'B'])
    expect(data.records).toEqual([])
  })

  it('reports an empty source as no fields rather than throwing', () => {
    const data = mergeDataFromRows([])
    expect(data.fields).toEqual([])
    expect(data.records).toEqual([])
  })
})

describe('mergeRecords — field substitution', () => {
  const data: MergeData = { fields: ['Name'], records: [{ Name: 'Ada' }] }

  it('replaces a MERGEFIELD run with the record value and drops the field', () => {
    const { blocks } = mergeRecords([p('Dear '), mergeName(), p('!')], data)
    expect(lines(blocks)).toEqual(['Dear ', 'Ada', '!'])
  })

  it('an unknown field becomes empty text, not the placeholder', () => {
    const { blocks } = mergeRecords([mergeName('Nope')], data)
    expect(blocks[0]!.runs).toEqual([{ text: '' }])
  })

  it('merges every occurrence of a field in one paragraph', () => {
    const { blocks } = mergeRecords(
      [
        blk([
          { text: '«Name»', instrField: 'MERGEFIELD Name' },
          { text: ' / ' },
          { text: '«Name»', instrField: 'MERGEFIELD Name' },
        ]),
      ],
      data,
    )
    expect(blocks[0]!.runs?.map((r) => r.text)).toEqual(['Ada', ' / ', 'Ada'])
  })

  it('reaches into table cells', () => {
    const { blocks } = mergeRecords(
      [
        {
          id: '',
          docxIndex: 0,
          originalXml: '',
          type: 'table',
          table: {
            // TableModel.rows is a grid of cells (w:tr -> w:tc), not row objects
            rows: [
              [
                {
                  paras: [],
                  richParas: [{ runs: [{ text: '«Name»', instrField: 'MERGEFIELD Name' }] }],
                },
              ],
            ],
          },
        } as unknown as Block,
      ],
      data,
    )
    const table = (blocks[0] as { table?: NonNullable<Block['table']> }).table!
    const cell = table.rows[0]![0]!
    expect(cell.richParas?.[0]!.runs[0]!.text).toBe('Ada')
  })

  it('leaves a non-merge field (a PAGE, say) untouched', () => {
    const { blocks } = mergeRecords([blk([{ text: 'x', instrField: 'PAGE' }])], data)
    expect(blocks[0]!.runs?.[0]!.instrField).toBe('PAGE')
  })

  it('keeps the run formatting the field carried', () => {
    const { blocks } = mergeRecords(
      [blk([{ text: '«Name»', instrField: 'MERGEFIELD Name', bold: true }])],
      data,
    )
    expect(blocks[0]!.runs?.[0]!.bold).toBe(true)
  })
})

describe('mergeRecords — control fields', () => {
  const data: MergeData = {
    fields: ['Name', 'Skip'],
    records: [
      { Name: 'Ada', Skip: 'no' },
      { Name: 'Grace', Skip: 'yes' },
      { Name: 'Linus', Skip: 'no' },
    ],
  }

  it('NEXT advances the cursor, so a field after it reads the following record', () => {
    const { blocks, cursor } = mergeRecords([mergeName(), control('NEXT'), mergeName()], data, 0)
    expect(lines(blocks)).toEqual(['Ada', 'Grace'])
    expect(cursor.next).toBe(true)
  })

  it('a NEXT past the last record reads empty fields, like a blank trailing label', () => {
    const { blocks } = mergeRecords([mergeName(), control('NEXT'), mergeName()], data, 2)
    expect(lines(blocks)).toEqual(['Linus', ''])
  })

  it('a paragraph that held only a control field does not survive the merge', () => {
    const { blocks } = mergeRecords([mergeName(), control('NEXT')], data, 0)
    expect(lines(blocks)).toEqual(['Ada'])
  })

  it('SKIPIF marks the record skipped when the test is true', () => {
    const { cursor } = mergeRecords([mergeName(), control('SKIPIF «Skip» = "yes"')], data, 1)
    expect(cursor.skipped).toBe(true)
  })

  it('SKIPIF leaves the record alone when the test is false', () => {
    // record 0 is { Skip: 'no' }, so the condition must ask for something it is not
    const { cursor } = mergeRecords([mergeName(), control('SKIPIF «Skip» = "yes"')], data, 0)
    expect(cursor.skipped).toBe(false)
  })

  it('SKIPIF compares numerically when both sides are numbers, as Word does', () => {
    const nums: MergeData = { fields: ['N'], records: [{ N: '9' }, { N: '50' }, { N: '5' }] }
    expect(mergeRecords([control('SKIPIF «N» < "50"')], nums, 0).cursor.skipped).toBe(true)
    expect(mergeRecords([control('SKIPIF «N» < "50"')], nums, 1).cursor.skipped).toBe(false)
    expect(mergeRecords([control('SKIPIF «N» = "5.0"')], nums, 2).cursor.skipped).toBe(true)
    // one side is text, so the comparison stays textual
    expect(compareOperands('"9"', '<', '"50a"', {})).toBe(false)
  })

  it('IF picks the matching branch for each record', () => {
    const out = mergeBlockTree([control('IF «Skip» = "yes" "hidden" "shown"')], data)
    expect(out.map(lines)).toEqual([['shown'], ['hidden'], ['shown']])
  })

  it('IF resolves a merge-field operand against the current record', () => {
    const out = mergeBlockTree([control('IF «Skip» = "yes" "skip me" "keep me"')], data)
    expect(out.map(lines)).toEqual([['keep me'], ['skip me'], ['keep me']])
  })

  it('a bare word in a comparison is a literal, not a column', () => {
    // Word wraps a column in guillemets: `SKIPIF Status = "yes"` compares the
    // text "Status" with "yes" and is false for every record. Reading `Status`
    // as a column here would make a template skip (or keep) rows at random.
    const rows: MergeData = { fields: ['Status'], records: [{ Status: 'yes' }] }
    const bare = mergeRecords([control('SKIPIF Status = "yes"')], rows, 0)
    expect(bare.cursor.skipped).toBe(false)
    const wrapped = mergeRecords([control('SKIPIF «Status» = "yes"')], rows, 0)
    expect(wrapped.cursor.skipped).toBe(true)
  })

  it('an IF with no comparison tests the field itself, not the whole instruction', () => {
    // `IF «Paid» "got" "missed"`: the first operand is the test, so an empty
    // value takes the false branch — passing the whole instruction instead
    // would be non-empty text for every record
    const rows: MergeData = { fields: ['Paid'], records: [{ Paid: '' }, { Paid: 'yes' }] }
    const out = mergeBlockTree([control('IF «Paid» "got" "missed"')], rows)
    expect(out.map(lines)).toEqual([['missed'], ['got']])
  })

  it('keeps the run formatting an IF branch inherits from its field', () => {
    const { blocks } = mergeRecords(
      [blk([{ text: 'x', instrField: 'IF «Name» = "Ada" "she" "they"', bold: true }])],
      data,
    )
    expect(blocks[0]!.runs?.[0]!.bold).toBe(true)
  })

  it('keeps a run that carries no text — a logo, an index entry', () => {
    // image and index runs hold their payload beside an empty text; the merge
    // must copy them through rather than drop "empty" runs
    const logo = blk([{ text: '', image: { rid: 'r1', name: 'logo' } }])
    const { blocks } = mergeRecords([p('x'), logo, mergeName()], data)
    expect(blocks[1]!.runs?.[0]!.image).toBeDefined()
  })
})

describe('mergeBlockTree', () => {
  it('emits one block set per record', () => {
    const out = mergeBlockTree([p('x'), mergeName()], {
      fields: ['Name'],
      records: [{ Name: 'Ada' }, { Name: 'Grace' }],
    })
    expect(out).toHaveLength(2)
  })

  it('returns nothing when the recipient list is empty', () => {
    expect(mergeBlockTree([p('x')], { fields: [], records: [] })).toEqual([])
  })

  it('produces one set per record when the template has no NEXT', () => {
    const out = mergeBlockTree([p('Dear '), mergeName()], {
      fields: ['Name'],
      records: [{ Name: 'A' }, { Name: 'B' }, { Name: 'C' }],
    })
    expect(out.map(lines)).toEqual([
      ['Dear ', 'A'],
      ['Dear ', 'B'],
      ['Dear ', 'C'],
    ])
  })

  it('a NEXT-driven template lays the records end to end, one sheet per batch', () => {
    const out = mergeBlockTree([mergeName(), control('NEXT'), mergeName()], {
      fields: ['Name'],
      records: [{ Name: 'A' }, { Name: 'B' }, { Name: 'C' }],
    })
    expect(out.map(lines)).toEqual([
      ['A', 'B'],
      ['C', ''],
    ])
  })

  it('a SKIPIF on a NEXT sheet drops that sheet and still moves on', () => {
    const out = mergeBlockTree(
      [mergeName(), control('SKIPIF «Name» = "A"'), control('NEXT'), mergeName()],
      {
        fields: ['Name'],
        records: [{ Name: 'A' }, { Name: 'B' }, { Name: 'C' }, { Name: 'D' }],
      },
    )
    expect(out.map(lines)).toEqual([['C', 'D']])
  })

  it('drops a record a SKIPIF rejected', () => {
    const out = mergeBlockTree([mergeName(), control('SKIPIF «Skip» = "yes"')], {
      fields: ['Name', 'Skip'],
      records: [
        { Name: 'Ada', Skip: 'no' },
        { Name: 'Grace', Skip: 'yes' },
        { Name: 'Linus', Skip: 'no' },
      ],
    })
    expect(out.map(lines)).toEqual([['Ada'], ['Linus']])
  })

  it('stops at the last record rather than emitting a blank tail', () => {
    const out = mergeBlockTree([mergeName()], {
      fields: ['Name'],
      records: [{ Name: 'only' }],
    })
    expect(out).toHaveLength(1)
  })

  it('does not mutate the input blocks', () => {
    const input = [mergeName()]
    const before = JSON.stringify(input)
    mergeBlockTree(input, { fields: ['Name'], records: [{ Name: 'Ada' }] })
    expect(JSON.stringify(input)).toBe(before)
  })
})
