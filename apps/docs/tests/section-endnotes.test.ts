/** Endnotes at the end of each section: which section's end prints a note, and the rows put in the flow there. */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  endnoteHomesOf,
  endnotePrintSections,
  groupSectionEndnotes,
  sectionNoteSlot,
  setSectionEndnotes,
  topBlockSections,
} from '../src/renderer/editor/section-endnotes'
import type { PageNoteItem } from '../src/renderer/pagination-types'

const para = (docxIndex: number | null, text: string, noteId?: string) => ({
  type: 'docParagraph',
  attrs: { docxIndex },
  content: [
    { type: 'text', text },
    ...(noteId ? [{ type: 'docNoteRef', attrs: { kind: 'endnote', id: noteId, num: 1 } }] : []),
  ],
})

// three sections closed by the blocks 1, 3 and 5
const SECTIONS = [1, 3, 5].map((lastBlockIndex, i) => ({
  lastBlockIndex,
  sectPrXml: `<w:sectPr><w:pgSz w:w="${12240 + i}"/></w:sectPr>`,
}))

const item = (id: string, no: number): PageNoteItem => ({ id, no, text: `note ${id}`, height: 16 })

describe('section endnotes', () => {
  let editor: Editor | undefined
  afterEach(() => editor?.destroy())
  const open = (content: Array<ReturnType<typeof para>>) => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content },
    })
    return editor
  }

  it('a section prints its own notes unless it suppresses them or its break is gone', () => {
    expect(endnotePrintSections([false, false, false], [true, true, true])).toEqual([0, 1, 2])
    expect(endnotePrintSections([true, false, false], [true, true, true])).toEqual([1, 1, 2])
    expect(endnotePrintSections([true, true, false], [true, true, true])).toEqual([2, 2, 2])
    // the last section prints whatever it says
    expect(endnotePrintSections([false, true, true], [true, true, true])).toEqual([0, 2, 2])
    expect(endnotePrintSections([false, false, false], [true, false, true])).toEqual([0, 2, 2])
  })

  it('finds the section of each block and of each endnote reference', () => {
    const { state } = open([
      para(0, 'one', 'a'),
      para(1, 'break one'),
      para(null, 'typed after the break'),
      para(2, 'two', 'b'),
      para(3, 'break two', 'c'),
      para(4, 'three', 'd'),
      para(5, 'end'),
    ])
    const blocks = topBlockSections(state.doc, SECTIONS)
    expect(blocks.sectionOf).toEqual([0, 0, 0, 1, 1, 2, 2])
    expect(blocks.closed).toEqual([true, true, true])
    const homes = endnoteHomesOf(state.doc, SECTIONS)
    expect([...homes.sectionOf]).toEqual([
      ['a', 0],
      ['b', 1],
      ['c', 1],
      ['d', 2],
    ])
    const items = [item('a', 1), item('b', 2), item('c', 3), item('d', 4), item('lost', 5)]
    const all = groupSectionEndnotes(items, homes, [false, false, false])
    expect(all.groups.map((g) => [g.section, g.items.map((i) => i.id)])).toEqual([
      [0, ['a']],
      [1, ['b', 'c']],
    ])
    // a note whose reference is not in the body stays in the end area
    expect(all.tail.map((i) => i.id)).toEqual(['d', 'lost'])
    // a note added later sits where its reference is, not where the list has it
    const late = groupSectionEndnotes([...items].reverse(), homes, [false, false, false])
    expect(late.groups[1]!.items.map((i) => i.id)).toEqual(['b', 'c'])
    expect(late.tail.map((i) => i.id)).toEqual(['d', 'lost'])
    const handedOn = groupSectionEndnotes(items, homes, [true, false, false])
    expect(handedOn.groups.map((g) => [g.section, g.items.map((i) => i.id)])).toEqual([
      [1, ['a', 'b', 'c']],
    ])
    const toTheEnd = groupSectionEndnotes(items, homes, [false, true, false])
    expect(toTheEnd.groups.map((g) => g.section)).toEqual([0])
    expect(toTheEnd.tail.map((i) => i.id)).toEqual(['b', 'c', 'd', 'lost'])
  })

  it('a section whose break paragraph was deleted hands its notes on', () => {
    const { state } = open([para(0, 'one', 'a'), para(2, 'two', 'b'), para(3, 'break two')])
    const homes = endnoteHomesOf(state.doc, SECTIONS)
    expect(homes.closed).toEqual([false, true, true])
    const { groups } = groupSectionEndnotes([item('a', 1), item('b', 2)], homes, [
      false,
      false,
      false,
    ])
    expect(groups.map((g) => [g.section, g.items.map((i) => i.id)])).toEqual([[1, ['a', 'b']]])
  })

  it('draws the rows after the last block of their section and follows edits', () => {
    const ed = open([
      para(0, 'one', 'a'),
      para(1, 'break one'),
      para(2, 'two'),
      para(3, 'break two'),
      para(4, 'three'),
    ])
    const edited: string[] = []
    setSectionEndnotes(
      ed.view,
      [{ section: 0, items: [item('a', 1), item('b', 2)] }],
      SECTIONS,
      (id) => edited.push(id),
    )
    const order = () =>
      Array.from(ed.view.dom.children).map((el) =>
        el.classList.contains('page-sect-note')
          ? `#${(el as HTMLElement).dataset.noteId}`
          : (el.textContent ?? '').replace(/i$/, ''),
      )
    expect(order()).toEqual(['one', 'break one', '#a', '#b', 'two', 'break two', 'three'])
    const rows = Array.from(ed.view.dom.querySelectorAll<HTMLElement>('.page-sect-note'))
    expect(rows[0]!.classList.contains('page-sect-note-first')).toBe(true)
    expect(rows[1]!.classList.contains('page-sect-note-first')).toBe(false)
    expect(rows[0]!.querySelector('sup')?.textContent).toBe('i')
    expect(rows[1]!.textContent).toContain('note b')
    // a page gap opening at the second row sorts between the two rows
    const first = sectionNoteSlot(ed.view, rows[0]!)!
    const second = sectionNoteSlot(ed.view, rows[1]!)!
    expect(second.pos).toBe(first.pos)
    expect(Number(rows[0]!.dataset.side)).toBeLessThan(second.side)
    expect(second.side).toBeLessThan(Number(rows[1]!.dataset.side))
    expect(sectionNoteSlot(ed.view, ed.view.dom.firstElementChild as HTMLElement)).toBeNull()
    rows[1]!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
    expect(edited).toEqual(['b'])
    // a paragraph typed at the end of the first section stays above its notes
    ed.commands.insertContentAt(ed.state.doc.child(0).nodeSize, para(null, 'added'))
    expect(order()).toEqual(['one', 'added', 'break one', '#a', '#b', 'two', 'break two', 'three'])
    setSectionEndnotes(ed.view, [], SECTIONS, () => undefined)
    expect(ed.view.dom.querySelector('.page-sect-note')).toBeNull()
  })
})
