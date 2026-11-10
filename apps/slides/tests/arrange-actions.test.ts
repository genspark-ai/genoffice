import { afterEach, describe, expect, it, vi } from 'vitest'
import type { ActionCtx } from '../src/renderer/action-context'
import type {
  BatchEditTransformOp,
  DeleteElementsOp,
  EditTransformMultiOp,
} from '../src/shared/ipc'
import {
  eraseInk,
  regroupCandidates,
  regroupSelected,
  rotateSelected,
  ungroupSelected,
} from '../src/renderer/arrange-actions'
import type { UngroupedSet } from '../src/renderer/action-context'

vi.mock('../src/renderer/i18n/locale', () => ({ t: (key: string) => key }))

type Box = { x: number; y: number; w: number; h: number; rotationDeg: number }

function setup(selectedIds: string[], boxes: Record<string, Box>, groupOf: Record<string, string>) {
  const api = {
    batchEditTransform: vi.fn(async (_op: BatchEditTransformOp) => ({ nodes: [] })),
    editTransformMulti: vi.fn(async (_op: EditTransformMultiOp) => ({ nodes: [] })),
    deleteElements: vi.fn(async (_op: DeleteElementsOp) => ({ nodes: [] })),
  }
  vi.stubGlobal('window', { slidesApi: api })
  const ctx = {
    current: 1,
    selectedIds,
    slide: { nodes: [], widthPx: 1280, heightPx: 720 },
    findNodeCtx: (id: string) =>
      boxes[id] ? { node: { type: 'shape', box: boxes[id] }, groupId: groupOf[id] } : null,
    applySlide: vi.fn(),
    setDirty: vi.fn(),
  } as unknown as ActionCtx
  return { api, ctx }
}

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('multi-selection rotate', () => {
  const boxes: Record<string, Box> = {
    a: { x: 0, y: 0, w: 100, h: 50, rotationDeg: 0 },
    b: { x: 300, y: 100, w: 40, h: 40, rotationDeg: 45 },
  }

  it('rotates in-group children through one transform batch', async () => {
    const { api, ctx } = setup(['a', 'b'], boxes, { a: 'grp', b: 'grp' })
    await rotateSelected(ctx, 90)
    expect(api.batchEditTransform).not.toHaveBeenCalled()
    expect(api.editTransformMulti).toHaveBeenCalledTimes(1)
    const op = api.editTransformMulti.mock.calls[0]![0]
    expect(op.slideIndex).toBe(1)
    expect(op.items).toHaveLength(2)
    expect(op.items[0]).toMatchObject({ sourceId: 'a', groupId: 'grp', rotationDeg: 90 })
    expect(op.items[0]!.xPx).toBeCloseTo(75)
    expect(op.items[0]!.yPx).toBeCloseTo(-25)
    expect(op.items[1]).toMatchObject({ sourceId: 'b', groupId: 'grp', rotationDeg: 135 })
    expect(ctx.applySlide).toHaveBeenCalledTimes(1)
    expect(ctx.setDirty).toHaveBeenCalledWith(true)
  })

  it('keeps top-level elements on the legacy batch path', async () => {
    const { api, ctx } = setup(['a', 'b'], boxes, {})
    await rotateSelected(ctx, -90)
    expect(api.editTransformMulti).not.toHaveBeenCalled()
    expect(api.batchEditTransform).toHaveBeenCalledTimes(1)
    expect(ctx.applySlide).toHaveBeenCalledTimes(1)
  })
})

describe('ink eraser', () => {
  it('erases every hit stroke in one delete batch', async () => {
    const { api, ctx } = setup([], {}, {})
    await eraseInk(ctx, ['s1', 's2', 's3'])
    expect(api.deleteElements).toHaveBeenCalledTimes(1)
    expect(api.deleteElements).toHaveBeenCalledWith({
      slideIndex: 1,
      sourceIds: ['s1', 's2', 's3'],
    })
    expect(ctx.applySlide).toHaveBeenCalledTimes(1)
  })

  it('skips the IPC when no stroke was hit', async () => {
    const { api, ctx } = setup([], {}, {})
    await eraseInk(ctx, [])
    expect(api.deleteElements).not.toHaveBeenCalled()
  })
})

describe('regroup', () => {
  const node = (sourceId: string, type = 'shape') => ({ sourceId, type }) as never
  const A = 'ppt/slides/slide1.xml'
  const B = 'ppt/slides/slide2.xml'
  const sets: UngroupedSet[] = [{ slidePath: A, sourceIds: ['a', 'b', 'c'] }]
  const onA = (...nodes: never[]) => ({ partPath: A, nodes })

  it('collects every surviving former member when any one of them is selected', () => {
    const found = regroupCandidates(sets, onA(node('a'), node('b'), node('c'), node('z')), ['b'])
    expect(found).toEqual(['a', 'b', 'c'])
  })

  it('drops members that were deleted or are no longer groupable', () => {
    expect(regroupCandidates(sets, onA(node('a'), node('b', 'table')), ['a'])).toBeNull()
    expect(
      regroupCandidates(sets, onA(node('a'), node('b', 'table'), node('c', 'text')), ['a']),
    ).toEqual(['a', 'c'])
  })

  it('ignores selections outside any set and slides without an identity', () => {
    const nodes = [node('a'), node('b'), node('c')]
    expect(regroupCandidates(sets, onA(...nodes), ['z'])).toBeNull()
    expect(regroupCandidates([], onA(...nodes), ['a'])).toBeNull()
    expect(regroupCandidates(sets, { nodes }, ['a'])).toBeNull()
  })

  it('does not match another slide that reuses the same element ids', () => {
    const lookalike = { partPath: B, nodes: [node('a'), node('b'), node('c')] }
    expect(regroupCandidates(sets, lookalike, ['a'])).toBeNull()
  })

  function regroupCtx(stored: UngroupedSet[]) {
    const api = {
      groupElements: vi.fn(async () => ({ slide: { nodes: [node('g')] }, groupId: 'g' })),
    }
    vi.stubGlobal('window', { slidesApi: api })
    const ctx = {
      current: 0,
      selectedIds: ['c'],
      slide: onA(node('a'), node('b'), node('c')),
      ungroupedSets: stored,
      setUngroupedSets: vi.fn(),
      applySlide: vi.fn(),
      setSelectedIds: vi.fn(),
      setDirty: vi.fn(),
      setStatus: vi.fn(),
    } as unknown as ActionCtx
    return { api, ctx }
  }

  it('regroups the whole remembered set and selects the group', async () => {
    const { api, ctx } = regroupCtx(sets)
    await regroupSelected(ctx)
    expect(api.groupElements).toHaveBeenCalledWith({ slideIndex: 0, sourceIds: ['a', 'b', 'c'] })
    expect(ctx.setSelectedIds).toHaveBeenCalledWith(['g'])
  })

  it('keeps the set after regrouping so an undo re-enables Regroup', async () => {
    const { ctx } = regroupCtx(sets)
    await regroupSelected(ctx)
    expect(ctx.setUngroupedSets).not.toHaveBeenCalled()
    const afterUndo = onA(node('a'), node('b'), node('c'))
    expect(regroupCandidates(sets, afterUndo, ['a'])).toEqual(['a', 'b', 'c'])
  })

  function ungroupCtx(stored: UngroupedSet[], partPath: string) {
    const api = {
      ungroupElement: vi.fn(async () => ({ partPath, nodes: [node('a'), node('b')] })),
    }
    vi.stubGlobal('window', { slidesApi: api })
    const ctx = {
      current: 1,
      selectedIds: ['g'],
      slide: {
        partPath,
        nodes: [{ sourceId: 'g', type: 'group', children: [node('a'), node('b'), node('gone')] }],
      },
      setUngroupedSets: vi.fn((f: (s: UngroupedSet[]) => UngroupedSet[]) => {
        stored = f(stored)
      }),
      applySlide: vi.fn(),
      setSelectedIds: vi.fn(),
      setDirty: vi.fn(),
      setStatus: vi.fn(),
    } as unknown as ActionCtx
    return { ctx, stored: () => stored }
  }

  it('ungroup remembers the surviving children and replaces a same-slide set sharing an id', async () => {
    const { ctx, stored } = ungroupCtx(
      [
        { slidePath: A, sourceIds: ['a', 'old'] },
        { slidePath: A, sourceIds: ['p', 'q'] },
      ],
      A,
    )
    await ungroupSelected(ctx)
    expect(ctx.setSelectedIds).toHaveBeenCalledWith(['a', 'b'])
    expect(stored()).toEqual([
      { slidePath: A, sourceIds: ['p', 'q'] },
      { slidePath: A, sourceIds: ['a', 'b'] },
    ])
  })

  it('ungroup on another slide keeps sets whose slide merely reuses the ids', async () => {
    const { ctx, stored } = ungroupCtx([{ slidePath: A, sourceIds: ['a', 'b'] }], B)
    await ungroupSelected(ctx)
    expect(stored()).toEqual([
      { slidePath: A, sourceIds: ['a', 'b'] },
      { slidePath: B, sourceIds: ['a', 'b'] },
    ])
  })
})
