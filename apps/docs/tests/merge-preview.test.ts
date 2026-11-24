/** Preview Results shows a recipient's values in the merge fields without touching the document. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { toggleAllFieldCodes } from '../src/renderer/editor/field-codes'
import {
  mergePreviewPrompts,
  mergePreviewRecord,
  setMergePreview,
} from '../src/renderer/editor/merge-preview'

const field = (instr: string, text: string) => ({
  type: 'text',
  text,
  marks: [{ type: 'instrField', attrs: { instr } }],
})

describe('merge preview', () => {
  let editor: Editor
  const shown = () =>
    Array.from(editor.view.dom.querySelectorAll('.doc-merge-preview')).map((el) => el.textContent)
  const hidden = () =>
    Array.from(editor.view.dom.querySelectorAll('.doc-field-result-hidden')).map(
      (el) => el.textContent,
    )
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Dear ' },
              field(' MERGEFIELD First ', '«First»'),
              { type: 'text', text: ', page ' },
              field(' PAGE ', '1'),
            ],
          },
        ],
      },
    })
  })
  afterEach(() => editor.destroy())

  it('shows the record in the merge fields, leaves other fields alone, and goes back', () => {
    const before = editor.state.doc
    setMergePreview(editor, { First: 'Ada' })
    expect(shown()).toEqual(['Ada'])
    expect(hidden()).toEqual(['«First»'])
    expect(mergePreviewRecord(editor.state)).toEqual({ First: 'Ada' })
    // a display only: the document is the one it was
    expect(editor.state.doc.eq(before)).toBe(true)
    expect(editor.can().undo()).toBe(false)

    setMergePreview(editor, { First: 'Grace' })
    expect(shown()).toEqual(['Grace'])
    setMergePreview(editor, null)
    expect(shown()).toEqual([])
    expect(hidden()).toEqual([])
  })

  it('follows the document: a field typed in while previewing shows its value too', () => {
    setMergePreview(editor, { First: 'Ada', Last: 'Lovelace' })
    editor.commands.insertContentAt(6, [field(' MERGEFIELD Last ', '«Last»')] as never)
    expect(shown().sort()).toEqual(['Ada', 'Lovelace'])
  })

  it('shows the rule fields as the merge works them out for that record', () => {
    const people = [
      { First: 'Ada', City: 'London' },
      { First: 'Alan', City: 'Wilmslow' },
      { First: 'Grace', City: 'London' },
    ]
    const rule = (instr: string, fieldId: number) => ({
      type: 'text',
      text: '\u200b',
      marks: [{ type: 'instrField', attrs: { instr, fieldId } }],
    })
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            rule('SKIPIF «City» = "Wilmslow"', 1),
            rule('SET Greeting "Hello"', 2),
            {
              type: 'text',
              text: 'ref ',
              marks: [{ type: 'refField', attrs: { name: 'Greeting' } }],
            },
            field(' MERGEFIELD First ', '«First»'),
            { type: 'text', text: ' no ' },
            field('MERGEREC', ' '),
            { type: 'text', text: ' of ' },
            field('MERGESEQ', ' '),
            { type: 'text', text: ' by ' },
            field('FILLIN "Who signs?" \\d "The Clerk"', ' '),
            rule('NEXT', 3),
            { type: 'text', text: ' then ' },
            field(' MERGEFIELD First ', '«First»'),
          ],
        },
      ],
    } as never)
    setMergePreview(editor, people[0]!, people, 0)
    expect(shown()).toEqual(['Hello', 'Ada', '1', '1', 'The Clerk', 'Alan'])
    // the third record: the second is skipped, so it is the second document written
    setMergePreview(editor, people[2]!, people, 2)
    expect(shown()).toEqual(['Hello', 'Grace', '3', '2', 'The Clerk', ''])
    expect(mergePreviewRecord(editor.state)).toEqual(people[2])
  })

  it('asks the Ask and Fill-in fields of the record once and shows what was answered', () => {
    const people = [{ First: 'Ada' }, { First: 'Alan' }]
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            field(' MERGEFIELD First ', '\u00abFirst\u00bb'),
            { type: 'text', text: ' signed by ' },
            field('FILLIN "Who signs?" \\d "The Clerk" \\o', ' '),
            { type: 'text', text: ' at ' },
            field('FILLIN "What time?" \\d "noon"', ' '),
            {
              type: 'text',
              text: '\u200b',
              marks: [
                { type: 'instrField', attrs: { instr: 'ASK Room "Which room?"', fieldId: 1 } },
              ],
            },
            { type: 'text', text: ' in ' },
            { type: 'text', text: 'ref', marks: [{ type: 'refField', attrs: { name: 'Room' } }] },
          ],
        },
      ],
    } as never)
    // nothing answered yet: the defaults show, and three questions are due for this record
    setMergePreview(editor, people[0]!, people, 0)
    expect(shown()).toEqual(['Ada', 'The Clerk', 'noon', ''])
    const first = mergePreviewPrompts(editor.state, people, 0)
    expect(first.map((p) => [p.prompt, p.fallback, p.record])).toEqual([
      ['Who signs?', 'The Clerk', null],
      ['What time?', 'noon', 1],
      ['Which room?', '', 1],
    ])
    const answers = new Map(first.map((p, i) => [p.key, ['Grace', '9 am', 'Blue'][i]!]))
    setMergePreview(editor, people[0]!, people, 0, answers)
    expect(shown()).toEqual(['Ada', 'Grace', '9 am', 'Blue'])
    expect(mergePreviewPrompts(editor.state, people, 0, answers)).toEqual([])
    // the next record: the one asked once keeps its answer, the others are asked again
    setMergePreview(editor, people[1]!, people, 1, answers)
    expect(shown()).toEqual(['Alan', 'Grace', 'noon', ''])
    expect(mergePreviewPrompts(editor.state, people, 1, answers).map((p) => p.prompt)).toEqual([
      'What time?',
      'Which room?',
    ])
  })

  it("draws the value in the field's own formatting", () => {
    editor.commands.setTextSelection({ from: 6, to: 13 })
    editor.commands.setMark('bold')
    setMergePreview(editor, { First: 'Ada' })
    const value = editor.view.dom.querySelector('.doc-merge-preview')!
    expect(value.closest('strong, b')).not.toBeNull()
  })

  it('steps aside for a field that is showing its code, and comes back when the code is hidden', () => {
    setMergePreview(editor, { First: 'Ada' })
    toggleAllFieldCodes(editor.view)
    expect(shown()).toEqual([])
    expect(editor.view.dom.querySelectorAll('.doc-field-code')).toHaveLength(2)
    toggleAllFieldCodes(editor.view)
    expect(shown()).toEqual(['Ada'])
  })
})
