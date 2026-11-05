import { describe, expect, it } from 'vitest'
import type { PmMark, PmNode } from '../src/renderer/editor/convert'
import {
  countRecordSlots,
  describeRecordSlots,
  documentText,
  mergeAllRecords,
  mergeFieldNames,
  mergeNextPages,
  mergePageCount,
  recordPage,
  substituteMergeFields,
} from '../src/renderer/mail-merge/template'

/**
 * The template half of the mail merge: the document with its merge fields
 * filled from a record. genoffice#2097.
 */

const mark = (instr: string): PmMark => ({ type: 'instrField', attrs: { instr } })
const mergeMark = (name: string): PmMark => mark(`MERGEFIELD "${name}" \\* MERGEFORMAT`)

const text = (value: string, marks?: PmMark[]): PmNode => ({ type: 'text', text: value, marks })
const para = (...content: PmNode[]): PmNode => ({ type: 'docParagraph', content })
const doc = (...content: PmNode[]): PmNode => ({ type: 'doc', content })

const record = { FirstName: 'Ada', LastName: 'Lovelace' }

describe('substituteMergeFields', () => {
  it('fills a merge field with the record value', () => {
    const template = doc(para(text('Dear '), text('«FirstName»', [mergeMark('FirstName')])))
    const merged = substituteMergeFields(template, record)
    expect(merged.content?.[0]?.content?.map((n) => n.text)).toEqual(['Dear ', 'Ada'])
  })

  it('leaves no field mark behind', () => {
    const template = doc(para(text('«FirstName»', [mergeMark('FirstName')])))
    const merged = substituteMergeFields(template, record)
    expect(merged.content?.[0]?.content?.[0]?.marks).toEqual([])
  })

  it('fills a field whose result is split over several nodes only once', () => {
    const template = doc(
      para(text('Ada', [mergeMark('FirstName')]), text(' Lovelace', [mergeMark('FirstName')])),
    )
    const merged = substituteMergeFields(template, record)
    expect(merged.content?.[0]?.content?.map((n) => n.text)).toEqual(['Ada'])
  })

  it('leaves a field that is not a merge field alone', () => {
    const template = doc(para(text('1', [mark('PAGE')])))
    const merged = substituteMergeFields(template, record)
    expect(merged.content?.[0]?.content?.[0]?.text).toBe('1')
    expect(merged.content?.[0]?.content?.[0]?.marks).toEqual([mark('PAGE')])
  })

  it('takes the branch an IF field selects', () => {
    const template = doc(para(text('?', [mark('IF «Paid» = "yes" "paid" "unpaid"')])))
    const merged = substituteMergeFields(template, { Paid: 'yes' })
    expect(merged.content?.[0]?.content?.[0]?.text).toBe('paid')
    const other = substituteMergeFields(template, { Paid: 'no' })
    expect(other.content?.[0]?.content?.[0]?.text).toBe('unpaid')
  })

  it('fills a field inside a table cell', () => {
    const template = doc({
      type: 'docTable',
      content: [
        {
          type: 'docTableRow',
          content: [
            {
              type: 'docTableCell',
              content: [para(text('«FirstName»', [mergeMark('FirstName')]))],
            },
          ],
        },
      ],
    })
    const merged = substituteMergeFields(template, record)
    const cell = merged.content?.[0]?.content?.[0]?.content?.[0]
    expect(cell?.content?.[0]?.content?.[0]?.text).toBe('Ada')
  })
})

describe('documentText', () => {
  it('reads the document as one line per block', () => {
    const template = doc(
      para(text('Dear '), text('«FirstName»', [mergeMark('FirstName')])),
      para(text('!')),
    )
    expect(documentText(substituteMergeFields(template, record))).toBe('Dear Ada\n!')
  })
})

describe('mergeFieldNames', () => {
  it('lists the fields the template holds, once each', () => {
    const template = doc(
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(
        text('«FirstName»', [mergeMark('FirstName')]),
        text(' '),
        text('«LastName»', [mergeMark('LastName')]),
      ),
    )
    expect(mergeFieldNames(template)).toEqual(['FirstName', 'LastName'])
  })
})

describe('mergeAllRecords', () => {
  it('repeats the template per record, each letter on its own page', () => {
    const template = doc(para(text('«FirstName»', [mergeMark('FirstName')])))
    const merged = mergeAllRecords(template, [record, { FirstName: 'Grace' }])
    expect(merged.content).toHaveLength(2)
    expect(merged.content?.[0]?.content?.[0]?.text).toBe('Ada')
    expect(merged.content?.[1]?.content?.[0]?.text).toBe('Grace')
    expect(merged.content?.[0]?.attrs?.pageBreakBefore).toBeUndefined()
    expect(merged.content?.[1]?.attrs?.pageBreakBefore).toBe(true)
  })

  it('keeps bookmarks, comment anchors and note references in the first letter only', () => {
    const block = para(
      text('Dear ', [{ type: 'comment', attrs: { ids: 'c1' } }]),
      text('«FirstName»', [mergeMark('FirstName')]),
      { type: 'docNoteRef', attrs: { id: '1', kind: 'footnote' } },
    )
    block.attrs = { ...(block.attrs ?? {}), bookmarks: ['intro'] }
    const merged = mergeAllRecords(doc(block), [record, { FirstName: 'Grace' }])
    const [first, second] = merged.content ?? []
    expect(first?.attrs?.bookmarks).toEqual(['intro'])
    expect(first?.content?.[0]?.marks?.map((m) => m.type)).toEqual(['comment'])
    expect(first?.content?.some((n) => n.type === 'docNoteRef')).toBe(true)
    expect(second?.attrs?.bookmarks).toBeNull()
    expect(second?.content?.[0]?.marks ?? []).toEqual([])
    expect(second?.content?.some((n) => n.type === 'docNoteRef')).toBe(false)
  })
})

describe('recordPage / mergePageCount', () => {
  const nextMark = (): PmMark => mark('NEXT')
  /** three records per page: each block followed by a NEXT */
  const threeBlocks = doc(
    para(text('«FirstName»', [mergeMark('FirstName')])),
    para(text(' ', [nextMark()])),
    para(text('«FirstName»', [mergeMark('FirstName')])),
    para(text(' ', [nextMark()])),
    para(text('«FirstName»', [mergeMark('FirstName')])),
    para(text(' ', [nextMark()])),
  )
  const slots = countRecordSlots(threeBlocks)

  /** the pages mergeNextPages actually writes, counted by page breaks */
  const writtenPages = (recordCount: number): number => {
    const records = Array.from({ length: recordCount }, (_, i) => ({ FirstName: `R${i + 1}` }))
    const merged = mergeNextPages(threeBlocks, records, slots)
    return (merged.content ?? []).filter((block) => block.attrs?.pageBreakBefore).length + 1
  }

  it('shows three records per page for a 3-NEXT template', () => {
    expect(slots).toBe(3)
    expect(recordPage(10, 0, slots)).toEqual({ start: 0, end: 2, count: 3 })
  })

  it('keeps the displayed page in step with the exported page for 1..10 records', () => {
    for (let recordCount = 1; recordCount <= 10; recordCount++) {
      expect(mergePageCount(recordCount, slots)).toBe(Math.ceil(recordCount / 3))
      expect(mergePageCount(recordCount, slots)).toBe(writtenPages(recordCount))
      // every page's window covers its own records and nothing earlier
      const lastStart = Math.floor((recordCount - 1) / slots) * slots
      expect(recordPage(recordCount, recordCount - 1, slots).start).toBe(lastStart)
      const page = Math.floor(lastStart / slots)
      expect(recordPage(recordCount, recordCount - 1, slots).start).toBe(page * slots)
    }
  })

  it('shows the short last page as one record when 7 records fill 2 pages and 1', () => {
    expect(mergePageCount(7, slots)).toBe(3)
    expect(writtenPages(7)).toBe(3)
    expect(recordPage(7, 6, slots)).toEqual({ start: 6, end: 6, count: 1 })
  })
})

describe('unevaluated rules and empty record sets', () => {
  it('drops NEXTIF / SKIPIF instruction text from a merged letter', () => {
    const template = doc(
      para(
        text('Dear'),
        text('«NEXTIF «Age» > 30»', [mark('NEXTIF «Age» > 30')]),
        text('«SKIPIF «Paid» = "no"»', [mark('SKIPIF «Paid» = "no"')]),
      ),
    )
    const merged = substituteMergeFields(template, record)
    expect(documentText(merged)).toBe('Dear')
    expect(merged.content?.[0]?.content?.some((n) => n.marks?.length)).toBe(false)
  })

  it('leaves the template alone when there are no records', () => {
    const template = doc(para(text('«FirstName»', [mergeMark('FirstName')])))
    expect(mergeNextPages(template, [], 3)).toEqual(template)
  })
})

describe('countRecordSlots', () => {
  const nextMark = (): PmMark => mark('NEXT')

  /** three record blocks, each followed by a NEXT — the last one ends the page */
  const threeBlocks = doc(
    para(text('«FirstName»', [mergeMark('FirstName')])),
    para(text(' ', [nextMark()])),
    para(text('«FirstName»', [mergeMark('FirstName')])),
    para(text(' ', [nextMark()])),
    para(text('«FirstName»', [mergeMark('FirstName')])),
    para(text(' ', [nextMark()])),
  )

  it('counts the blocks a page renders, not the NEXT fields', () => {
    // the NEXT after the third block hands off to the next page and renders
    // nothing, so a page holds 3 records rather than 4
    expect(countRecordSlots(threeBlocks)).toBe(3)
  })

  it('gives a page one slot when the only NEXT is at the end', () => {
    const template = doc(
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()])),
    )
    expect(countRecordSlots(template)).toBe(1)
  })

  it('walks into a table cell', () => {
    const template = doc({
      type: 'docTable',
      content: [
        {
          type: 'docTableRow',
          content: [
            {
              type: 'docTableCell',
              content: [para(text('«FirstName»', [mergeMark('FirstName')]))],
            },
          ],
        },
      ],
    })
    expect(countRecordSlots(template)).toBe(1)
  })

  it('does not count a trailing empty paragraph after the last NEXT', () => {
    const template = doc(
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()])),
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()])),
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()])),
      // Word leaves the empty line after the page's last NEXT in the body
      para(),
      para(text('   ')),
    )
    expect(countRecordSlots(template)).toBe(3)
    expect(mergePageCount(147, countRecordSlots(template))).toBe(49)
    // the display window and the pages the merge actually writes agree
    expect(recordPage(147, 0, countRecordSlots(template))).toEqual({ start: 0, end: 2, count: 3 })
    const records = Array.from({ length: 147 }, (_, i) => ({ FirstName: `R${i + 1}` }))
    const merged = mergeNextPages(template, records, countRecordSlots(template))
    const written = (merged.content ?? []).filter((b) => b.attrs?.pageBreakBefore).length + 1
    expect(written).toBe(49)
  })

  it('still counts a record block that follows a NEXT', () => {
    const template = doc(
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()])),
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()])),
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()])),
      // a real record block, not a blank line: this one does hold record 4
      para(text('«FirstName»', [mergeMark('FirstName')])),
    )
    expect(countRecordSlots(template)).toBe(4)
  })

  it('counts one NEXT split over several runs once', () => {
    const template = doc(
      para(text('«FirstName»', [mergeMark('FirstName')])),
      para(text(' ', [nextMark()]), text(' ', [nextMark()])),
      para(text('«FirstName»', [mergeMark('FirstName')])),
    )
    expect(countRecordSlots(template)).toBe(2)
  })

  it('reports which blocks take a slot (dev diagnostics)', () => {
    const trace = describeRecordSlots(threeBlocks)
    expect(trace.slots).toBe(3)
    expect(trace.blocks.filter((block) => block.contributed).map((block) => block.index)).toEqual([0, 2, 4])
    expect(trace.nextFields.map((entry) => entry.block)).toEqual([1, 3, 5])
  })

  it('does not skip a record between pages', () => {
    const records = Array.from({ length: 10 }, (_, index) => ({ FirstName: `R${index + 1}` }))
    const merged = mergeNextPages(threeBlocks, records, countRecordSlots(threeBlocks))
    // a NEXT leaves its (now empty) paragraph behind, so only the filled slots
    // carry text: three per page, and page ten of nothing after the records end
    const rendered = (merged.content ?? [])
      .map((block) => block.content?.[0]?.text ?? '')
      .filter((value) => value !== '')
    expect(rendered).toEqual(['R1', 'R2', 'R3', 'R4', 'R5', 'R6', 'R7', 'R8', 'R9', 'R10'])
  })
})
