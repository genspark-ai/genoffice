import { describe, expect, it } from 'vitest'
import {
  commentRows,
  noteRows,
  revisionLabelKeys,
  spellRows,
  type CtxEntry,
  type CtxRow,
} from '../src/renderer/components/context-menu-model'

const ids = (rows: CtxRow[]) => rows.map((r) => (r === 'sep' ? '---' : r.id))
const byId = (rows: CtxRow[], id: string) =>
  rows.find((r): r is CtxEntry => r !== 'sep' && r.id === id)!

describe('right-click on a misspelled word', () => {
  const target = { word: 'mispeled', pos: 12, marked: false, suggestions: ['misspelled', 'misled'] }

  it('lists the suggestions first, then Ignore, Ignore All, Add to Dictionary, AutoCorrect, Spelling…', () => {
    const rows = spellRows(target, true)
    expect(ids(rows)).toEqual([
      'spell-suggestion',
      'spell-suggestion',
      '---',
      'spell-ignore-once',
      'spell-ignore-all',
      'spell-add',
      'spell-autocorrect',
      'spell-dialog',
    ])
    const [first] = rows as CtxEntry[]
    expect(first).toMatchObject({ label: 'misspelled', strong: true, command: 'spell-replace' })
    expect(JSON.parse(first.payload!)).toEqual({
      pos: 12,
      word: 'mispeled',
      to: 'misspelled',
      marked: false,
    })
    expect(byId(rows, 'spell-ignore-all')).toMatchObject({
      command: 'spell-ignore-all',
      payload: 'mispeled',
    })
    expect(byId(rows, 'spell-add')).toMatchObject({
      command: 'spell-add-word',
      payload: 'mispeled',
    })
    expect(byId(rows, 'spell-dialog')).toMatchObject({
      command: 'tools-spelling',
      payload: '12',
      ellipsis: true,
    })
  })

  it('AutoCorrect repeats the suggestions, each adding a replace pair', () => {
    const auto = byId(spellRows(target, true), 'spell-autocorrect')
    expect(auto.children!.map((c) => c.label)).toEqual(['misspelled', 'misled'])
    expect(auto.children![1]).toMatchObject({ command: 'spell-autocorrect' })
    expect(JSON.parse(auto.children![1].payload!)).toMatchObject({ word: 'mispeled', to: 'misled' })
  })

  it('shows five suggestions at most and says so when there is none', () => {
    const many = spellRows({ ...target, suggestions: ['a', 'b', 'c', 'd', 'e', 'f'] }, true)
    expect(ids(many).filter((id) => id === 'spell-suggestion')).toHaveLength(5)
    const none = spellRows({ ...target, suggestions: [] }, true)
    expect(none[0]).toMatchObject({ id: 'spell-none', disabled: true })
    expect(byId(none, 'spell-autocorrect').disabled).toBe(true)
  })

  it('a read-only document keeps the dictionary items and greys what would change the text', () => {
    const rows = spellRows(target, false)
    expect((rows[0] as CtxEntry).disabled).toBe(true)
    expect(byId(rows, 'spell-autocorrect').disabled).toBe(true)
    expect(byId(rows, 'spell-ignore-all').disabled).toBeUndefined()
  })
})

describe('right-click on a comment', () => {
  it('a balloon gets Reply, Delete and Resolve alone', () => {
    const rows = commentRows({ id: 'c1', done: false })
    expect(ids(rows)).toEqual(['comment-reply', 'comment-delete', 'comment-resolve'])
    expect(rows.every((r) => r !== 'sep' && r.payload === 'c1' && !r.disabled)).toBe(true)
  })

  it('a resolved thread offers Reopen and takes no reply', () => {
    const rows = commentRows({ id: 'c1', done: true })
    expect(ids(rows)).toEqual(['comment-reply', 'comment-delete', 'comment-reopen'])
    expect(byId(rows, 'comment-reply').disabled).toBe(true)
    expect(byId(rows, 'comment-reopen')).toMatchObject({ command: 'comment-reopen', payload: 'c1' })
  })

  it('off a comment the items stay listed, greyed', () => {
    const off = commentRows({ id: null, done: false })
    expect(ids(off)).toEqual(['comment-reply', 'comment-delete', 'comment-resolve'])
    expect(off.every((r) => r !== 'sep' && r.disabled)).toBe(true)
  })
})

describe('right-click on a footnote or endnote', () => {
  it('the note text goes to its mark, opens the options, converts to an endnote', () => {
    const rows = noteRows({ kind: 'footnote', id: '3', at: 'text' }, true)
    expect(rows).toMatchObject([
      { labelKey: 'appCtxGoToFootnote', command: 'note-goto', payload: 'footnote:ref:3' },
      { labelKey: 'appCtxNoteOptions', command: 'insert-footnote-dialog' },
      { labelKey: 'appCtxConvertToEndnote', command: 'note-convert', payload: 'footnote:3' },
    ])
  })

  it('the mark of an endnote goes to its text and converts to a footnote', () => {
    const rows = noteRows({ kind: 'endnote', id: '7', at: 'ref' }, false)
    expect(rows).toMatchObject([
      { labelKey: 'appCtxGoToEndnote', payload: 'endnote:text:7' },
      { disabled: true },
      { labelKey: 'appCtxConvertToFootnote', payload: 'endnote:7', disabled: true },
    ])
  })
})

describe('right-click on a tracked change', () => {
  it('names the change the way Word does', () => {
    expect(revisionLabelKeys('ins')).toEqual({
      accept: 'appCtxAcceptInsertion',
      reject: 'appCtxRejectInsertion',
    })
    expect(revisionLabelKeys('paraMarkDel')).toEqual({
      accept: 'appCtxAcceptDeletion',
      reject: 'appCtxRejectDeletion',
    })
    expect(revisionLabelKeys('rPrChange')).toEqual({
      accept: 'appCtxAcceptChange',
      reject: 'appCtxRejectChange',
    })
  })
})
