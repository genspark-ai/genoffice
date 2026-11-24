import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import type { PmMark, PmNode } from '../src/renderer/editor/convert'
import {
  documentText,
  mergeAllRecords,
  mergeDocuments,
  previewDocument,
  previewFieldTexts,
} from '../src/renderer/mail-merge/template'
import {
  askInstruction,
  fillInInstruction,
  nextIfInstruction,
  setInstruction,
  skipIfInstruction,
} from '../src/renderer/mail-merge/rules'
import { firstRecordInRange } from '../src/renderer/mail-merge/range'

/** The rule fields of Mailings > Rules, worked out when the merge runs. */

const mark = (instr: string): PmMark => ({ type: 'instrField', attrs: { instr } })
const field = (instr: string, shown = ' '): PmNode => ({
  type: 'text',
  text: shown,
  marks: [mark(instr)],
})
const merge = (name: string) => field(`MERGEFIELD "${name}" \\* MERGEFORMAT`, `«${name}»`)
const text = (value: string): PmNode => ({ type: 'text', text: value })
const para = (...content: PmNode[]): PmNode => ({ type: 'docParagraph', content })
const doc = (...content: PmNode[]): PmNode => ({ type: 'doc', content })

const PEOPLE = [
  { Name: 'Ada', City: 'London' },
  { Name: 'Alan', City: 'Wilmslow' },
  { Name: 'Grace', City: 'London' },
  { Name: 'Edsger', City: 'Austin' },
]

const texts = (template: PmNode, records = PEOPLE, options = {}) =>
  mergeDocuments(template, records, options).documents.map((d) => documentText(d.doc))

describe('the rule instructions', () => {
  it('are the fields Word writes', () => {
    expect(nextIfInstruction('City', '=', 'London')).toBe('NEXTIF «City» = "London"')
    expect(setInstruction('Greeting', 'Hello "you"')).toBe('SET Greeting "Hello ”you”"')
    expect(askInstruction('Reason', 'Why?', 'none', true)).toBe('ASK Reason "Why?" \\d "none" \\o')
    expect(fillInInstruction('Who signs?', 'The team', false)).toBe(
      'FILLIN "Who signs?" \\d "The team"',
    )
  })

  it('Merge Record # counts from the first record of the range', () => {
    expect(firstRecordInRange({ kind: 'all', from: 1, to: null }, 2)).toBe(1)
    expect(firstRecordInRange({ kind: 'current', from: 1, to: null }, 2)).toBe(3)
    expect(firstRecordInRange({ kind: 'custom', from: 4, to: null }, 0)).toBe(4)
  })
})

describe('the merge', () => {
  it('Skip Record If leaves the record out, and its paragraph out of the others', () => {
    const template = doc(para(field(skipIfInstruction('City', '=', 'London'))), para(merge('Name')))
    expect(texts(template)).toEqual(['Alan', 'Edsger'])
  })

  it('Merge Record # is the place in the list, Merge Sequence # counts what is written', () => {
    const template = doc(
      para(field(skipIfInstruction('City', '=', 'Wilmslow'))),
      para(merge('Name'), text(' '), field('MERGEREC'), text('/'), field('MERGESEQ')),
    )
    expect(texts(template)).toEqual(['Ada 1/1', 'Grace 3/2', 'Edsger 4/3'])
    expect(texts(template, PEOPLE.slice(2), { firstRecord: 3 })).toEqual([
      'Grace 3/1',
      'Edsger 4/2',
    ])
  })

  it('Next Record puts the following record in the same document', () => {
    const template = doc(para(merge('Name')), para(field('NEXT')), para(merge('Name')))
    expect(texts(template)).toEqual(['Ada\nAlan', 'Grace\nEdsger'])
  })

  it('two Next Record fields side by side step twice; one whose result is split over runs steps once', () => {
    const next = (id: number, shown = ' ', bold = false): PmNode => ({
      type: 'text',
      text: shown,
      marks: [
        ...(bold ? [{ type: 'bold' }] : []),
        { type: 'instrField', attrs: { instr: 'NEXT', fieldId: id } },
      ],
    })
    const two = doc(para(merge('Name'), text('>'), next(1), next(2), merge('Name')))
    expect(texts(two)).toEqual(['Ada>Grace', 'Edsger>'])
    const split = doc(para(merge('Name'), text('>'), next(1), next(1, ' ', true), merge('Name')))
    expect(texts(split)).toEqual(['Ada>Alan', 'Grace>Edsger'])
    // the same holds for fields that show something
    const counted = (id: number): PmNode => ({
      type: 'text',
      text: '0',
      marks: [{ type: 'instrField', attrs: { instr: 'MERGEREC', fieldId: id } }],
    })
    expect(texts(doc(para(counted(1), counted(2))), PEOPLE.slice(0, 1))).toEqual(['11'])
    expect(texts(doc(para(counted(1), counted(1))), PEOPLE.slice(0, 1))).toEqual(['1'])
  })

  it('Next Record If moves on only when its condition holds', () => {
    const template = doc(
      para(
        merge('Name'),
        field(nextIfInstruction('City', '=', 'London')),
        text('+'),
        merge('Name'),
      ),
    )
    expect(texts(template)).toEqual(['Ada+Alan', 'Grace+Edsger'])
    expect(texts(template, PEOPLE.slice(1, 2))).toEqual(['Alan+Alan'])
  })

  it('Set Bookmark gives its text to the REF fields that name it', () => {
    const template = doc(
      para(field(setInstruction('Greeting', 'Hello'))),
      para(
        { type: 'text', text: 'old', marks: [{ type: 'refField', attrs: { name: 'Greeting' } }] },
        text(' '),
        merge('Name'),
        text(' '),
        field('REF Greeting'),
        field('REF Elsewhere', 'kept'),
      ),
    )
    expect(texts(template, PEOPLE.slice(0, 1))).toEqual(['Hello Ada Hellokept'])
  })

  it('Fill-in and Ask come back as questions, once or per record, and take their answers', () => {
    const template = doc(
      para(field(askInstruction('Reason', 'Why?', 'no reason', true))),
      para(
        merge('Name'),
        text(': '),
        field(fillInInstruction('Gift?', 'a card', false), 'a card'),
        text(' / '),
        field('REF Reason'),
      ),
    )
    const first = mergeDocuments(template, PEOPLE.slice(0, 2))
    expect(first.prompts).toEqual([
      { key: '0', prompt: 'Why?', fallback: 'no reason', record: null },
      { key: '1:1', prompt: 'Gift?', fallback: 'a card', record: 1 },
      { key: '1:2', prompt: 'Gift?', fallback: 'a card', record: 2 },
    ])
    expect(first.documents.map((d) => documentText(d.doc))).toEqual([
      'Ada: a card / no reason',
      'Alan: a card / no reason',
    ])
    const answers = new Map([
      ['0', 'thanks'],
      ['1:2', 'a book'],
    ])
    const second = mergeDocuments(template, PEOPLE.slice(0, 2), { answers })
    expect(second.documents.map((d) => documentText(d.doc))).toEqual([
      'Ada: a card / thanks',
      'Alan: a book / thanks',
    ])
  })

  it('a record that is skipped asks nothing', () => {
    const template = doc(
      para(field(fillInInstruction('Gift?', '', false))),
      para(field(skipIfInstruction('City', '=', 'London'))),
    )
    expect(mergeDocuments(template, PEOPLE).prompts.map((p) => p.record)).toEqual([2, 4])
  })

  it('a paragraph that holds a merge field with nothing in it stays', () => {
    const template = doc(para(merge('Missing')), para(merge('Name')))
    expect(mergeDocuments(template, PEOPLE.slice(0, 1)).documents[0]!.doc.content).toHaveLength(2)
  })

  it('one document for all the records follows the rules too', () => {
    const template = doc(para(field(skipIfInstruction('City', '=', 'London'))), para(merge('Name')))
    const all = mergeAllRecords(template, PEOPLE)
    expect(documentText(all)).toBe('Alan\nEdsger')
    expect(all.content?.[1]?.attrs?.pageBreakBefore).toBe(true)
  })
})

describe('rule fields read from a file', () => {
  const fld = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
  const next = fld('begin') + '<w:r><w:instrText> NEXT </w:instrText></w:r>' + fld('end')
  const name =
    fld('begin') +
    '<w:r><w:instrText> MERGEFIELD Name </w:instrText></w:r>' +
    fld('separate') +
    '<w:r><w:t>«Name»</w:t></w:r>' +
    fld('end')

  it('Preview Results reads one record through the same rules', () => {
    const template = doc(
      para(field(skipIfInstruction('City', '=', 'Wilmslow'))),
      para(
        merge('Name'),
        text(' #'),
        field('MERGEREC'),
        text(' seq '),
        field('MERGESEQ'),
        text(' '),
        field(fillInInstruction('Who?', 'Clerk', false)),
      ),
      para(field('NEXT'), merge('Name')),
    )
    expect(documentText(previewDocument(template, PEOPLE, 0))).toBe('Ada #1 seq 1 Clerk\nAlan')
    // Ada took Alan along, so the next document opens with Grace; the preview of Alan still shows Alan
    expect(documentText(previewDocument(template, PEOPLE, 1))).toBe('Alan #2 seq 2 Clerk\nGrace')
    expect(documentText(previewDocument(template, PEOPLE, 2))).toBe('Grace #3 seq 2 Clerk\nEdsger')
    const instrs = ['MERGEFIELD Name', 'PAGE', 'MERGEREC', 'MERGESEQ', 'NEXTIF «City» = "London"']
    expect(previewFieldTexts([...instrs, 'MERGEFIELD Name'], PEOPLE, 2)).toEqual([
      'Grace',
      null,
      '3',
      '2',
      '',
      'Edsger',
    ])
  })

  it('two Next Record fields with nothing between them stay two fields: in the editor, in the merge, in the file', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: `<w:p>${name}${next}${next}${name}</w:p>` }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const template = editor.getJSON() as PmNode
    const ids = (template.content?.[0]?.content ?? [])
      .map((n) => n.marks?.find((m) => m.type === 'instrField')?.attrs)
      .filter((attrs) => attrs?.instr === 'NEXT')
      .map((attrs) => attrs?.fieldId)
    expect(ids).toHaveLength(2)
    expect(ids[0]).not.toBe(ids[1])
    expect(texts(template)).toEqual(['AdaGrace', 'Edsger'])
    // a paragraph typed into is written again from its runs
    editor.commands.insertContentAt(1, 'To ')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const out = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(xml).toContain('To ')
    expect(xml.match(/> NEXT </g)).toHaveLength(2)
    editor.destroy()
  })
})
