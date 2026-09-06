import { describe, expect, it } from 'vitest'
import type { TextBlock } from '../src/renderer/text-block'
import {
  blockMoveInput,
  editCarriesBlock,
  isBlockEditOf,
  patchPendingEdits,
  resolveTextEdit,
  shiftPendingEdit,
} from '../src/renderer/text-edit-preview'
import type { LocalTextEdit } from '../src/renderer/text-edit-preview'

const block: TextBlock = {
  rect: [50, 676, 250, 712],
  fontSize: 10,
  lineHeight: 12,
  align: 'left',
  lines: [
    { text: 'first line', y: 700, rect: [50, 698, 250, 712], fontSize: 10 },
    { text: 'second line', y: 688, rect: [50, 686, 250, 700], fontSize: 10 },
    { text: 'third', y: 676, rect: [50, 674, 250, 688], fontSize: 10 },
  ],
}

describe('blockMoveInput', () => {
  it('builds a pure translate edit anchored at the shifted block corner', () => {
    const input = blockMoveInput(0, block, [5, -30])
    expect(input.translate).toEqual([5, -30])
    expect(input.origin).toEqual([55, 670])
    expect(input.rect).toEqual(block.rect)
    expect(input.newText).toBe('first line\nsecond line\nthird')
    expect(input.oldText).toBe(input.blockSource)
    expect(input.align).toBeUndefined()
  })
})

describe('shiftPendingEdit', () => {
  it('stacks a second move onto a pending pure move', () => {
    const te: LocalTextEdit = {
      id: 'a',
      input: blockMoveInput(0, block, [0, -30]),
      moveBy: [0, -30],
    }
    const shifted = shiftPendingEdit(te, block, [5, -10])
    expect(shifted.input.translate).toEqual([5, -40])
    expect(shifted.input.origin).toEqual([55, 660])
    expect(shifted.moveBy).toEqual([5, -40])
    expect(shifted.input.rect).toEqual(block.rect)
    expect(te.input.translate).toEqual([0, -30])
  })

  it('moves a pending paragraph rebuild by its origin without adding a translate', () => {
    const te: LocalTextEdit = {
      id: 'b',
      input: {
        pageIndex: 0,
        rect: block.rect,
        oldText: 'first line second line third',
        newText: 'rewritten',
        fontSize: 10,
        origin: [50, 700],
        lineLeading: 12,
      },
    }
    const shifted = shiftPendingEdit(te, block, [0, -30])
    expect(shifted.input.origin).toEqual([50, 670])
    expect(shifted.input.translate).toBeUndefined()
    expect(shifted.moveBy).toEqual([0, -30])
  })

  it('anchors a pending line edit at its own shifted row', () => {
    const te: LocalTextEdit = {
      id: 'c',
      input: {
        pageIndex: 0,
        rect: [80, 686, 200, 700],
        oldText: 'second line',
        newText: 'second LINE',
        fontSize: 10,
      },
    }
    const shifted = shiftPendingEdit(te, block, [0, -30])
    expect(shifted.input.origin).toEqual([80, 658])
    expect(shifted.input.lineLeading).toBe(12)
    expect(shifted.input.translate).toBeUndefined()
  })
})

describe('resolveTextEdit', () => {
  const rewrite = (origin: [number, number], newText = 'rewritten') => ({
    pageIndex: 0,
    rect: block.rect,
    oldText: 'first line second line third',
    newText,
    fontSize: 10,
    origin,
    lineLeading: 12,
    blockSource: newText,
  })
  const lineEdit = (id: string): LocalTextEdit => ({
    id,
    input: {
      pageIndex: 0,
      rect: [80, 686, 200, 700],
      oldText: 'second line',
      newText: 'second LINE',
      fontSize: 10,
    },
  })
  const ok = (r: ReturnType<typeof resolveTextEdit>) => {
    if ('reason' in r) throw new Error(r.reason)
    return r
  }

  it('appends when nothing pending claims the paragraph', () => {
    const r = ok(resolveTextEdit([], rewrite([50, 700])))
    expect(r.replaces).toEqual([])
    expect(r.moveBy).toBeUndefined()
  })

  it('folds a rewrite onto a pending move: rebuild at the moved corner, move superseded', () => {
    const moved: LocalTextEdit = {
      id: 'm',
      input: blockMoveInput(0, block, [0, -30]),
      moveBy: [0, -30],
    }
    const r = ok(resolveTextEdit([moved], rewrite([50, 700])))
    expect(r.replaces.map((e) => e.id)).toEqual(['m'])
    expect(r.moveBy).toEqual([0, -30])
    expect(r.input.origin).toEqual([50, 670])
    expect(r.input.translate).toBeUndefined()
    expect(r.input.newText).toBe('rewritten')
  })

  it('supersedes a pending rewrite of the same paragraph, keeping its displacement', () => {
    const prior: LocalTextEdit = { id: 'p', input: rewrite([50, 690], 'old'), moveBy: [0, -10] }
    const r = ok(resolveTextEdit([prior], rewrite([50, 700])))
    expect(r.replaces.map((e) => e.id)).toEqual(['p'])
    expect(r.input.origin).toEqual([50, 690])
  })

  it('lets a rewrite supersede a moved line edit inside the paragraph instead of refusing', () => {
    const movedLine = shiftPendingEdit(lineEdit('l'), block, [0, -30])
    expect(movedLine.input.origin).toBeDefined()
    const r = ok(resolveTextEdit([movedLine], rewrite([50, 700])))
    expect(r.replaces.map((e) => e.id)).toEqual(['l'])
    expect(r.input.origin).toEqual([50, 670])
  })

  it('supersedes plain line edits inside the paragraph without shifting it', () => {
    const r = ok(resolveTextEdit([lineEdit('l')], rewrite([50, 700])))
    expect(r.replaces.map((e) => e.id)).toEqual(['l'])
    expect(r.input.origin).toEqual([50, 700])
    expect(r.moveBy).toBeUndefined()
  })

  it('refuses a line edit inside a moved or rewritten paragraph', () => {
    const moved: LocalTextEdit = {
      id: 'm',
      input: blockMoveInput(0, block, [0, -30]),
      moveBy: [0, -30],
    }
    const r = resolveTextEdit([moved], lineEdit('x').input)
    expect('reason' in r && r.reason).toContain('edit_block')
    const rewritten: LocalTextEdit = { id: 'p', input: rewrite([50, 700]) }
    expect('reason' in resolveTextEdit([rewritten], lineEdit('x').input)).toBe(true)
  })

  it('does not treat a moved line edit as a blocker for another line edit', () => {
    const movedLine = shiftPendingEdit(lineEdit('l'), block, [0, -30])
    const r = ok(
      resolveTextEdit([movedLine], {
        pageIndex: 0,
        rect: [50, 674, 200, 688],
        oldText: 'third',
        newText: 'THIRD',
        fontSize: 10,
      }),
    )
    expect(r.replaces).toEqual([])
  })

  it('ignores pending edits on other pages', () => {
    const other: LocalTextEdit = {
      id: 'o',
      input: { ...blockMoveInput(0, block, [0, -30]), pageIndex: 1 },
    }
    const r = ok(resolveTextEdit([other], rewrite([50, 700])))
    expect(r.replaces).toEqual([])
  })
})

describe('editCarriesBlock / isBlockEditOf', () => {
  const line: LocalTextEdit = {
    id: 'l',
    input: {
      pageIndex: 0,
      rect: [80, 686, 200, 700],
      oldText: 'second line',
      newText: 'second LINE',
      fontSize: 10,
    },
  }

  it('a pure move or rewrite of the block carries it and counts as its block-level edit', () => {
    const move: LocalTextEdit = { id: 'm', input: blockMoveInput(0, block, [0, -30]) }
    expect(isBlockEditOf(move, block)).toBe(true)
    expect(editCarriesBlock(move, block)).toBe(true)
  })

  it('a line edit inside a longer paragraph carries only its row, even once shifted', () => {
    expect(editCarriesBlock(line, block)).toBe(false)
    const shifted = shiftPendingEdit(line, block, [0, -30])
    expect(editCarriesBlock(shifted, block)).toBe(false)
    expect(isBlockEditOf(shifted, block)).toBe(false)
  })

  it('a line edit covering all of a single-line block carries it without being block-level', () => {
    const single: TextBlock = {
      rect: [50, 698, 250, 712],
      fontSize: 10,
      lineHeight: 12,
      align: 'left',
      lines: [{ text: 'first line', y: 700, rect: [50, 698, 250, 712], fontSize: 10 }],
    }
    const whole: LocalTextEdit = {
      id: 'w',
      input: {
        pageIndex: 0,
        rect: [50, 700, 250, 712],
        oldText: 'first  line',
        newText: 'x',
        fontSize: 10,
      },
    }
    expect(editCarriesBlock(whole, single)).toBe(true)
    expect(isBlockEditOf(whole, single)).toBe(false)
  })
})

describe('patchPendingEdits', () => {
  const mk = (id: string, newText = id): LocalTextEdit => ({
    id,
    input: { pageIndex: 0, rect: [0, 0, 10, 10], oldText: 'o', newText, fontSize: 10 },
  })

  it('replaces an existing edit in place and keeps everything else', () => {
    const out = patchPendingEdits([mk('a'), mk('b'), mk('c')], mk('b', 'B'))
    expect(out.map((e) => e.id)).toEqual(['a', 'b', 'c'])
    expect(out[1]!.input.newText).toBe('B')
  })

  it('appends a new edit and drops the ids it supersedes', () => {
    const out = patchPendingEdits([mk('a'), mk('b')], mk('n'), new Set(['a']))
    expect(out.map((e) => e.id)).toEqual(['b', 'n'])
  })

  it('never removes the edit being landed even when listed', () => {
    const out = patchPendingEdits([mk('a')], mk('a', 'A'), new Set(['a']))
    expect(out.map((e) => e.input.newText)).toEqual(['A'])
  })

  it('in replace mode leaves the list untouched when the owner is gone', () => {
    const prev = [mk('b')]
    const out = patchPendingEdits(prev, mk('a', 'A'), new Set(['b']), 'replace')
    expect(out).toBe(prev)
    expect(patchPendingEdits([mk('a'), mk('b')], mk('a', 'A'), new Set(['b']), 'replace')).toEqual([
      mk('a', 'A'),
    ])
  })
})
