import { describe, expect, it } from 'vitest'
import { Schema } from '@tiptap/pm/model'
import { EditorState, type Transaction } from '@tiptap/pm/state'
import { StepMap } from '@tiptap/pm/transform'
import { appendsAtEnd, touchedTopLevelBlocks } from '../src/renderer/editor/touched-blocks'

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      content: 'text*',
      group: 'block',
      attrs: { align: { default: null } },
      toDOM: () => ['p', 0],
    },
    text: {},
  },
  marks: {
    strong: { toDOM: () => ['strong', 0] },
  },
})

const para = (text: string) => schema.node('paragraph', null, text ? schema.text(text) : undefined)
const makeDoc = (texts: string[]) => schema.node('doc', null, texts.map(para))

/** offset of top-level child `index` in `doc` */
const blockOffset = (doc: ReturnType<Schema['node']>, index: number): number => {
  let offset = 0
  for (let i = 0; i < index; i++) offset += doc.child(i).nodeSize
  return offset
}

const stateOf = (texts: string[]) => EditorState.create({ doc: makeDoc(texts) })

describe('touchedTopLevelBlocks', () => {
  it('touches only the edited block for a single text insertion', () => {
    const state = stateOf(['first', 'second', 'third'])
    const tr = state.tr.insertText('X', 2)
    const touched = touchedTopLevelBlocks(tr)
    expect(touched).not.toBeNull()
    expect([...touched!]).toEqual([0])
  })

  it('carries early step ranges through later steps into the final doc', () => {
    // inserts run last-block-first so every position stays valid against the
    // original doc while early ranges must shift through the later maps
    const state = stateOf(['alpha', 'beta', 'gamma'])
    const secondStart = blockOffset(state.doc, 1) + 1
    const thirdStart = blockOffset(state.doc, 2) + 1
    const tr = state.tr.insertText('C', thirdStart + 2)
    tr.insertText('B', secondStart + 1)
    tr.insertText('A', 2)
    const touched = touchedTopLevelBlocks(tr)
    // every block was edited, but 3 of 3 trips the 50% fallback
    expect(touched).toBeNull()
  })

  it('reports two touched blocks for inserts in two of four blocks', () => {
    const state = stateOf(['alpha', 'beta', 'gamma', 'delta'])
    const secondStart = blockOffset(state.doc, 1) + 1
    // later block first: the first range shifts by one through the second map
    const tr = state.tr.insertText('B', secondStart + 1)
    tr.insertText('A', 2)
    const touched = touchedTopLevelBlocks(tr)
    expect(touched).not.toBeNull()
    expect(touched!.size).toBe(2)
    expect(touched!.has(0)).toBe(true)
    const secondOffset = blockOffset(tr.doc, 1)
    expect(touched!.has(secondOffset)).toBe(true)
  })

  it('covers mark steps through their from/to range', () => {
    const state = stateOf(['hello world', 'untouched'])
    const mark = schema.marks.strong.create()
    const tr = state.tr.addMark(2, 5, mark)
    const touched = touchedTopLevelBlocks(tr)
    expect(touched).not.toBeNull()
    expect([...touched!]).toEqual([0])
  })

  it('covers attribute steps through their pos', () => {
    const state = stateOf(['first', 'second'])
    const secondOffset = blockOffset(state.doc, 1)
    const tr = state.tr.setNodeAttribute(secondOffset, 'align', 'center')
    const touched = touchedTopLevelBlocks(tr)
    expect(touched).not.toBeNull()
    expect([...touched!]).toEqual([secondOffset])
  })

  it('returns null for a step with neither map coverage nor from/to/pos', () => {
    const doc = makeDoc(['first', 'second'])
    const fakeTr = {
      doc,
      steps: [{}],
      mapping: { maps: [StepMap.empty] },
    } as unknown as Transaction
    expect(touchedTopLevelBlocks(fakeTr)).toBeNull()
  })

  it('returns a set below the 50% threshold and null above it', () => {
    // 4 blocks: touching 2 keeps the set (2*2 > 4 is false); inserts run
    // last-block-first so the early range shifts through the later map
    const below = stateOf(['a', 'b', 'c', 'd'])
    const trBelow = below.tr.insertText('Y', blockOffset(below.doc, 1) + 1)
    trBelow.insertText('X', 2)
    const kept = touchedTopLevelBlocks(trBelow)
    expect(kept).not.toBeNull()
    expect(kept!.size).toBe(2)

    // touching 3 of 4 falls back to a full recompute (3*2 > 4)
    const above = stateOf(['a', 'b', 'c', 'd'])
    const trAbove = above.tr.insertText('Z', blockOffset(above.doc, 2) + 1)
    trAbove.insertText('Y', blockOffset(above.doc, 1) + 1)
    trAbove.insertText('X', 2)
    expect(touchedTopLevelBlocks(trAbove)).toBeNull()
  })
})

describe('appendsAtEnd', () => {
  it('accepts a single insertion at the end of the document', () => {
    const state = stateOf(['first', 'second'])
    const at = state.doc.content.size
    const tr = state.tr.insert(at, para('tail'))
    expect(appendsAtEnd(tr)).toBe(true)
  })

  it('accepts an append onto a single empty paragraph', () => {
    const state = stateOf([''])
    const at = state.doc.content.size
    const tr = state.tr.insert(at, para('tail'))
    expect(appendsAtEnd(tr)).toBe(true)
  })

  it('rejects a replacement at the end', () => {
    const state = stateOf(['first', 'second'])
    const tr = state.tr.delete(1, 3)
    expect(appendsAtEnd(tr)).toBe(false)
  })

  it('rejects an insertion in the middle of the document', () => {
    const state = stateOf(['first', 'second'])
    const tr = state.tr.insertText('X', 2)
    expect(appendsAtEnd(tr)).toBe(false)
  })

  it('rejects multi-step transactions even when the last step appends', () => {
    const state = stateOf(['first', 'second'])
    const tr = state.tr.insertText('X', 2)
    tr.insert(state.doc.content.size + 1, para('tail'))
    expect(tr.steps.length).toBe(2)
    expect(appendsAtEnd(tr)).toBe(false)
  })

  it('rejects mark steps and empty transactions', () => {
    const state = stateOf(['hello', 'world'])
    const markTr = state.tr.addMark(1, 3, schema.marks.strong.create())
    expect(appendsAtEnd(markTr)).toBe(false)
    expect(appendsAtEnd(state.tr)).toBe(false)
  })
})
