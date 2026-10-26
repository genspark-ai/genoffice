import { describe, expect, it } from 'vitest'
import type { PmMark, PmNode } from '../src/renderer/editor/convert'
import {
  documentText,
  mergeAllRecords,
  mergeFieldNames,
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
    expect(merged.content?.[0]?.content?.map((n) => n.text)).toEqual(['Ada', ''])
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
