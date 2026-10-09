import { describe, expect, it } from 'vitest'
import {
  BUCKETS,
  GuidedError,
  lookup,
  planEditOps,
  reduceBucket,
  reduceEditOps,
  runEditOps,
  type Op,
  type OpContext,
} from '../src/renderer/edit-ops'
import type { EditSnapshot } from '../src/renderer/edit-state'

const empty = (): EditSnapshot => ({
  markups: [],
  annotDeletes: [],
  noteEdits: [],
  drawings: [],
  textEdits: [],
  textInserts: [],
  imageEdits: [],
  stampCfg: null,
  formEdits: new Map(),
  formFields: [],
  formWidgetEdits: new Map(),
  blankPages: [],
  rotations: new Map(),
  deleted: new Set(),
  order: null,
  metadata: null,
})

const ctx = (over: Partial<OpContext> = {}): OpContext => ({
  readOnly: false,
  pageCount: 3,
  deleted: new Set(),
  claimedImages: new Set(),
  ...over,
})

let n = 0
const newId = () => `id${++n}`

const run = (state: EditSnapshot, ops: Op[], c = ctx()) => runEditOps(state, ops, c, newId)

const markup = (pageIndex = 0) => ({
  pageIndex,
  type: 'highlight' as const,
  color: [1, 0.87, 0.35] as [number, number, number],
  quads: [[0, 10, 50, 10, 0, 0, 50, 0]],
})

const savedNote = (objNum: number, contents = 'old') => ({
  pageIndex: 0,
  objNum,
  type: 'note' as const,
  rect: [0, 0, 20, 20] as [number, number, number, number],
  color: null,
  author: 'A',
  contents,
  timeMs: null,
  inReplyTo: null,
})

describe('plan', () => {
  it('stamps ids on additive ops and reports them as created', () => {
    const { plan } = run(empty(), [{ op: 'addMarkup', markup: markup() }])
    expect(plan.failures).toEqual([])
    expect(plan.ops[0]!.id).toMatch(/^id/)
    expect(plan.records[0]!.created).toEqual([plan.ops[0]!.id])
    expect([...plan.touched]).toEqual(['markups'])
  })

  it('keeps a caller-supplied id', () => {
    const { plan } = run(empty(), [{ op: 'addMarkup', id: 'mine', markup: markup() }])
    expect(plan.ops[0]!.id).toBe('mine')
  })

  it('is atomic: one bad op rejects the whole batch and touches nothing', () => {
    const { state, plan } = run(empty(), [
      { op: 'addMarkup', markup: markup() },
      { op: 'rotatePages', pages: [9], dir: 90 },
    ])
    expect(plan.failures.map((f) => f.index)).toEqual([1])
    expect(plan.failures[0]!.error).toContain('out of range')
    expect(plan.ops).toEqual([])
    expect(state.markups).toEqual([])
  })

  it('rejects everything on a read-only document', () => {
    const { plan } = run(
      empty(),
      [{ op: 'setMetadata', metadata: { title: 'x' } }],
      ctx({ readOnly: true }),
    )
    expect(plan.failures[0]!.error).toContain('read-only')
  })

  it('rejects edits on deleted pages with guidance', () => {
    const { plan } = run(
      empty(),
      [{ op: 'addTextInsert', input: { pageIndex: 1, text: 'hi' } }],
      ctx({ deleted: new Set([1]) }),
    )
    expect(plan.failures[0]!.error).toContain('deleted')
  })
})

describe('annotate', () => {
  it('adds and removes pending markups by id', () => {
    const a = run(empty(), [
      { op: 'addMarkup', markup: markup() },
      { op: 'addMarkup', markup: markup(1) },
    ])
    expect(a.state.markups.map((m) => m.pageIndex)).toEqual([0, 1])
    const b = run(a.state, [{ op: 'removeMarkup', id: a.state.markups[0]!.id }])
    expect(b.state.markups.map((m) => m.pageIndex)).toEqual([1])
  })

  it('deleting a saved note voids its pending content edit', () => {
    const a = run(empty(), [{ op: 'editSavedNote', annot: savedNote(7), contents: 'new' }])
    expect(a.state.noteEdits).toHaveLength(1)
    const b = run(a.state, [{ op: 'deleteSavedAnnot', annot: savedNote(7) }])
    expect(b.state.annotDeletes).toHaveLength(1)
    expect(b.state.noteEdits).toEqual([])
  })

  it('editSavedNote replaces the entry per note and drops it when the text is restored', () => {
    const a = run(empty(), [
      { op: 'editSavedNote', annot: savedNote(7), contents: 'v1' },
      { op: 'editSavedNote', annot: savedNote(7), contents: 'v2' },
    ])
    expect(a.state.noteEdits.map((e) => e.contents)).toEqual(['v2'])
    const b = run(a.state, [{ op: 'editSavedNote', annot: savedNote(7), contents: 'old' }])
    expect(b.state.noteEdits).toEqual([])
    const c = run(a.state, [
      { op: 'editSavedNote', annot: savedNote(7), contents: 'old', force: true },
    ])
    expect(c.state.noteEdits.map((e) => e.contents)).toEqual(['old'])
  })

  it('note drawings carry their record id as localId; setNoteContents rewrites them', () => {
    const note = {
      kind: 'note' as const,
      pageIndex: 0,
      color: [1, 0, 0] as [number, number, number],
      at: [1, 2] as [number, number],
      contents: 'a',
    }
    const a = run(empty(), [{ op: 'addDrawing', drawing: note }])
    const d = a.state.drawings[0]!
    expect(d.input.kind === 'note' && d.input.localId).toBe(d.id)
    const b = run(a.state, [{ op: 'setNoteContents', id: d.id, contents: 'b' }])
    expect(b.state.drawings[0]!.input.kind === 'note' && b.state.drawings[0]!.input.contents).toBe(
      'b',
    )
  })

  it('addDrawing keeps the form widget link', () => {
    const ink = {
      kind: 'ink' as const,
      pageIndex: 0,
      color: [0, 0, 0] as [number, number, number],
      width: 1,
      paths: [[0, 0, 1, 1]],
    }
    const a = run(empty(), [{ op: 'addDrawing', drawing: ink, formWidgetId: 'w1' }])
    expect(a.state.drawings[0]!.formWidgetId).toBe('w1')
  })
})

describe('text', () => {
  const input = (pageIndex = 0) => ({
    pageIndex,
    rect: [0, 0, 10, 10] as [number, number, number, number],
    oldText: 'a',
    newText: 'b',
  })

  it('putTextEdit appends new records and replaces by id', () => {
    const a = run(empty(), [
      { op: 'putTextEdit', input: input(), moveBy: [1, 1], paper: '#0b1220' },
    ])
    const id = a.state.textEdits[0]!.id
    expect(a.state.textEdits[0]!.moveBy).toEqual([1, 1])
    expect(a.state.textEdits[0]!.paper).toBe('#0b1220')
    const b = run(a.state, [{ op: 'putTextEdit', id, input: input(1) }])
    expect(b.state.textEdits).toHaveLength(1)
    expect(b.state.textEdits[0]!.input.pageIndex).toBe(1)
    expect(b.state.textEdits[0]!.moveBy).toBeUndefined()
    expect(b.state.textEdits[0]!.paper).toBeUndefined()
  })

  it('text inserts add and remove', () => {
    const a = run(empty(), [{ op: 'addTextInsert', input: { pageIndex: 0, text: 'x' } }])
    const b = run(a.state, [{ op: 'removeTextInsert', id: a.state.textInserts[0]!.id }])
    expect(b.state.textInserts).toEqual([])
  })
})

describe('image', () => {
  const oldRect: [number, number, number, number] = [0, 0, 100, 50]
  const transform = {
    kind: 'transformImage' as const,
    pageIndex: 0,
    oldRect,
    rect: [10, 10, 110, 60] as [number, number, number, number],
  }

  it('refuses a second edit of an already claimed image', () => {
    const { plan } = run(
      empty(),
      [{ op: 'addImageEdit', input: transform }],
      ctx({ claimedImages: new Set(['0:0.00,0.00,100.00,50.00']) }),
    )
    expect(plan.failures[0]!.error).toContain('already has a pending edit')
  })

  it('removing a static-form-fill transform becomes a delete of the image', () => {
    const a = run(empty(), [
      {
        op: 'addImageEdit',
        input: transform,
        staticFill: { id: 'f', kind: 'signature', pageIndex: 0, rect: oldRect },
      },
    ])
    const b = run(a.state, [{ op: 'removeImageEdit', id: a.state.imageEdits[0]!.id }])
    expect(b.state.imageEdits[0]!.input).toEqual({ kind: 'deleteImage', pageIndex: 0, oldRect })
    const c = run(run(empty(), [{ op: 'addImageEdit', input: transform }]).state, [
      { op: 'removeImageEdit', id: 'id' + n },
    ])
    expect(c.state.imageEdits).toEqual([])
  })

  it('bakeImageEdit morphs a transform into a replace and patches an insert in place', () => {
    const insert = {
      kind: 'insertImage' as const,
      pageIndex: 0,
      image: 'AAA',
      rect: oldRect,
      layer: 'belowText' as const,
    }
    const a = run(empty(), [
      { op: 'addImageEdit', input: transform },
      { op: 'addImageEdit', input: insert },
    ])
    const [t, i] = a.state.imageEdits
    const b = run(a.state, [
      {
        op: 'bakeImageEdit',
        id: t!.id,
        image: 'BBB',
        rect: t!.input.kind === 'transformImage' ? t!.input.rect : oldRect,
        opacityBase: 'base',
      },
      { op: 'bakeImageEdit', id: i!.id, image: 'CCC', rect: [1, 1, 2, 2] },
    ])
    expect(b.state.imageEdits[0]!.input).toMatchObject({
      kind: 'replaceImage',
      oldRect,
      image: 'BBB',
    })
    expect(b.state.imageEdits[0]!.opacityBase).toBe('base')
    expect(b.state.imageEdits[1]!.input).toMatchObject({
      kind: 'insertImage',
      image: 'CCC',
      rect: [1, 1, 2, 2],
    })
  })

  it('setImageEditRect skips delete ops', () => {
    const a = run(empty(), [
      { op: 'addImageEdit', input: { kind: 'deleteImage', pageIndex: 0, oldRect } },
    ])
    const b = run(a.state, [
      { op: 'setImageEditRect', id: a.state.imageEdits[0]!.id, rect: [1, 1, 2, 2] },
    ])
    expect(b.state.imageEdits[0]!.input).toEqual({ kind: 'deleteImage', pageIndex: 0, oldRect })
  })
})

describe('page', () => {
  it('rotatePages accumulates, clears at 0, and swaps image stamp rects', () => {
    const stamp = {
      kind: 'image' as const,
      pageIndex: 0,
      image: 'AAA',
      rect: [0, 0, 40, 20] as [number, number, number, number],
    }
    const a = run(empty(), [
      { op: 'addDrawing', drawing: stamp },
      { op: 'rotatePages', pages: [0, 1], dir: 90 },
    ])
    expect([...a.state.rotations]).toEqual([
      [0, 90],
      [1, 90],
    ])
    expect(a.state.drawings[0]!.input.kind === 'image' && a.state.drawings[0]!.input.rect).toEqual([
      10, -10, 30, 30,
    ])
    const b = run(a.state, [{ op: 'rotatePages', pages: [0], dir: -90 }])
    expect(b.state.rotations.has(0)).toBe(false)
    const c = run(a.state, [{ op: 'rotatePages', pages: [0], dir: 180 }])
    expect(c.state.rotations.get(0)).toBe(270)
    expect(c.state.drawings[0]).toBe(a.state.drawings[0])
    expect(run(empty(), [{ op: 'rotatePages', pages: [0], dir: 45 }]).plan.failures).toHaveLength(1)
  })

  it('a batch plans against what its earlier ops deleted or claimed', () => {
    const wipe = run(empty(), [
      { op: 'deletePage', pageIndex: 0 },
      { op: 'deletePage', pageIndex: 1 },
      { op: 'deletePage', pageIndex: 2 },
    ])
    expect(wipe.plan.failures.map((f) => f.index)).toEqual([2])
    expect(wipe.plan.failures[0]!.error).toContain('at least one page must remain')
    expect(wipe.state.deleted.size).toBe(0)
    const twice = run(empty(), [
      { op: 'deletePage', pageIndex: 0 },
      { op: 'deletePage', pageIndex: 0 },
    ])
    expect(twice.plan.failures.map((f) => f.index)).toEqual([1])
    const oldRect: [number, number, number, number] = [0, 0, 10, 10]
    const claim = run(empty(), [
      { op: 'addImageEdit', input: { kind: 'deleteImage', pageIndex: 0, oldRect } },
      { op: 'addImageEdit', input: { kind: 'deleteImage', pageIndex: 0, oldRect } },
    ])
    expect(claim.plan.failures.map((f) => f.index)).toEqual([1])
    expect(claim.plan.failures[0]!.error).toContain('already has a pending edit')
  })

  it('deletePage drops the page markups and drawings and keeps one page', () => {
    const a = run(empty(), [
      { op: 'addMarkup', markup: markup(1) },
      { op: 'addMarkup', markup: markup(2) },
      { op: 'deletePage', pageIndex: 1 },
    ])
    expect([...a.state.deleted]).toEqual([1])
    expect(a.state.markups.map((m) => m.pageIndex)).toEqual([2])
    const { plan } = run(
      a.state,
      [{ op: 'deletePage', pageIndex: 0 }],
      ctx({ pageCount: 2, deleted: new Set([1]) }),
    )
    expect(plan.failures[0]!.error).toContain('at least one page must remain')
  })

  it('setPageOrder demands a full permutation', () => {
    expect(run(empty(), [{ op: 'setPageOrder', order: [2, 0, 1] }]).state.order).toEqual([2, 0, 1])
    expect(run(empty(), [{ op: 'setPageOrder', order: [0, 0, 1] }]).plan.failures).toHaveLength(1)
    expect(run(empty(), [{ op: 'setPageOrder', order: [0, 1] }]).plan.failures).toHaveLength(1)
    expect(run(empty(), [{ op: 'setPageOrder', order: null }]).state.order).toBeNull()
  })
})

describe('blank pages', () => {
  const blank = (after: number | null, pageIndex: number) => ({
    op: 'insertBlankPage',
    after,
    pageIndex,
    width: 612,
    height: 792,
  })

  it('appends a virtual page and places it in the order', () => {
    const { state, plan } = run(empty(), [blank(0, 3)])
    expect(plan.failures).toEqual([])
    expect(state.blankPages.map((b) => b.input)).toEqual([
      { pageIndex: 3, width: 612, height: 792 },
    ])
    expect(state.order).toEqual([0, 3, 1, 2])
  })

  it('inserts at the front with after=null and chains indices within a batch', () => {
    const { state } = run(empty(), [blank(null, 3), blank(2, 4)])
    expect(state.order).toEqual([3, 0, 1, 2, 4])
    expect(state.blankPages.map((b) => b.input.pageIndex)).toEqual([3, 4])
  })

  it('rejects a stale pageIndex and a bad size', () => {
    expect(run(empty(), [blank(0, 5)]).plan.failures[0]!.error).toMatch(/next page index/)
    expect(run(empty(), [{ ...blank(0, 3), width: 0 }]).plan.failures[0]!.error).toMatch(/width/)
  })

  it('can be deleted again like any page', () => {
    const first = run(empty(), [blank(0, 3)])
    const { state } = run(first.state, [{ op: 'deletePage', pageIndex: 3 }], ctx({ pageCount: 4 }))
    expect(state.deleted.has(3)).toBe(true)
    expect(state.blankPages).toHaveLength(1)
  })
})

describe('form / document', () => {
  it('setFormValue upserts by field name', () => {
    const a = run(empty(), [
      { op: 'setFormValue', value: { name: 'f', kind: 'text', value: 'a' } },
      { op: 'setFormValue', value: { name: 'f', kind: 'text', value: 'b' } },
    ])
    expect(a.state.formEdits.get('f')?.value).toBe('b')
  })

  it('setStamps and setMetadata replace wholesale', () => {
    const a = run(empty(), [
      { op: 'setStamps', cfg: { wm: null, hf: { text: 'x' } } },
      { op: 'setMetadata', metadata: { title: 't' } },
    ])
    expect(a.state.stampCfg?.hf).toEqual({ text: 'x' })
    expect(a.state.metadata).toEqual({ title: 't' })
    const b = run(a.state, [{ op: 'setStamps', cfg: { wm: null, hf: null } }])
    expect(b.plan.failures).toHaveLength(0)
    expect(b.state.stampCfg).toBeNull()
  })
})

describe('form design', () => {
  const field = (over: Record<string, unknown> = {}) => ({
    name: 'full_name',
    kind: 'text',
    pageIndex: 0,
    rect: [10, 10, 110, 30],
    ...over,
  })

  it('applies a batch that re-points a radio default in one step', () => {
    const a = run(empty(), [
      {
        op: 'addFormField',
        field: field({ name: 'c', kind: 'radio', exportValue: 'r', checked: true }),
      },
      { op: 'addFormField', field: field({ name: 'c', kind: 'radio', exportValue: 'b' }) },
    ])
    const [r, b] = a.state.formFields.map((f) => f.id)
    const c = run(a.state, [
      { op: 'patchFormField', id: b, kind: 'radio', input: { checked: true } },
      { op: 'patchFormField', id: r, kind: 'radio', input: { checked: false } },
    ])
    expect(c.plan.failures).toEqual([])
    expect(c.state.formFields.map((f) => !!f.input.checked)).toEqual([false, true])
  })

  it('lets radio buttons share a group name but no one else', () => {
    const a = run(empty(), [
      { op: 'addFormField', field: field({ name: 'color', kind: 'radio', exportValue: 'red' }) },
      { op: 'addFormField', field: field({ name: 'color', kind: 'radio', exportValue: 'blue' }) },
    ])
    expect(a.plan.failures).toEqual([])
    expect(a.state.formFields.map((f) => f.input.exportValue)).toEqual(['red', 'blue'])
    expect(
      run(
        a.state,
        [{ op: 'addFormField', field: field({ name: 'color' }) }],
        ctx({
          fieldNames: new Set(['color']),
          radioGroups: new Set(['color']),
        }),
      ).plan.failures[0]!.error,
    ).toMatch(/already exists/)
    expect(
      run(empty(), [{ op: 'addFormField', field: field({ kind: 'radio' }) }]).plan.failures[0]!
        .error,
    ).toMatch(/exportValue/)
    expect(
      run(empty(), [{ op: 'addFormField', field: field({ kind: 'choice', options: [] }) }]).plan
        .failures[0]!.error,
    ).toMatch(/options/)
  })

  it('accumulates edits of an existing widget under its widget id', () => {
    const a = run(empty(), [
      { op: 'editFormWidget', widgetId: '12R', fieldName: 'name', input: { rect: [1, 1, 50, 20] } },
    ])
    expect(a.plan.failures).toEqual([])
    const b = run(a.state, [
      { op: 'editFormWidget', widgetId: '12R', fieldName: 'name', input: { required: true } },
    ])
    expect(b.state.formWidgetEdits.get('12R')).toEqual({
      widgetId: '12R',
      fieldName: 'name',
      rect: [1, 1, 50, 20],
      required: true,
    })
    const c = run(b.state, [
      { op: 'editFormWidget', widgetId: '12R', fieldName: 'name', input: { remove: true } },
    ])
    expect(c.state.formWidgetEdits.get('12R')?.remove).toBe(true)
    expect(
      run(empty(), [
        { op: 'editFormWidget', widgetId: '12R', fieldName: 'name', input: { rect: [0, 0, 0, 0] } },
      ]).plan.failures[0]!.error,
    ).toMatch(/at least 1pt/)
  })

  it('adds, patches and removes authored fields', () => {
    const first = run(empty(), [{ op: 'addFormField', field: field() }])
    const plan = first.plan
    let state = first.state
    expect(plan.failures).toEqual([])
    const fid = plan.records[0]!.created![0]!
    expect(state.formFields.map((f) => f.input.name)).toEqual(['full_name'])
    ;({ state } = run(state, [
      { op: 'patchFormField', id: fid, input: { rect: [20, 20, 120, 40], required: true } },
    ]))
    expect(state.formFields[0]!.input).toMatchObject({
      rect: [20, 20, 120, 40],
      required: true,
      kind: 'text',
    })
    ;({ state } = run(state, [{ op: 'removeFormField', id: fid }]))
    expect(state.formFields).toEqual([])
  })

  it('rejects colliding names, including within one batch, and bad kinds', () => {
    const taken = ctx({ fieldNames: new Set(['agree']) })
    expect(
      run(empty(), [{ op: 'addFormField', field: field({ name: 'agree' }) }], taken).plan
        .failures[0]!.error,
    ).toMatch(/already exists/)
    expect(
      run(empty(), [
        { op: 'addFormField', field: field({ name: 'x' }) },
        { op: 'addFormField', field: field({ name: 'x', kind: 'checkbox' }) },
      ]).plan.failures[0]!.error,
    ).toMatch(/already exists/)
    expect(
      run(empty(), [{ op: 'addFormField', field: field({ kind: 'button' }) }]).plan.failures[0]!
        .error,
    ).toMatch(/kind/)
    expect(
      run(empty(), [{ op: 'addFormField', field: field({ name: 'a.b' }) }]).plan.failures[0]!.error,
    ).toMatch(/name/)
  })

  it('drops authored fields with their deleted page', () => {
    const { state } = run(empty(), [
      { op: 'addFormField', field: field({ pageIndex: 1 }) },
      { op: 'addFormField', field: field({ name: 'other', pageIndex: 2 }) },
      { op: 'deletePage', pageIndex: 1 },
    ])
    expect(state.formFields.map((f) => f.input.name)).toEqual(['other'])
  })
})

describe('in-place updates of pending records', () => {
  it('moveDrawing shifts every drawing kind and setDrawingRect resizes image stamps only', () => {
    const color: [number, number, number] = [0, 0, 0]
    const a = run(empty(), [
      {
        op: 'addDrawing',
        drawing: { kind: 'ink', pageIndex: 0, color, width: 1, paths: [[0, 0, 2, 2]] },
      },
      {
        op: 'addDrawing',
        drawing: { kind: 'arrow', pageIndex: 0, color, width: 1, from: [0, 0], to: [1, 1] },
      },
      {
        op: 'addDrawing',
        drawing: { kind: 'image', pageIndex: 0, image: 'A', rect: [0, 0, 4, 4] },
      },
    ])
    const [ink, arrow, img] = a.state.drawings.map((d) => d.id)
    const b = run(a.state, [
      { op: 'moveDrawing', id: ink, dx: 1, dy: 2 },
      { op: 'moveDrawing', id: arrow, dx: 1, dy: 2 },
      { op: 'moveDrawing', id: img, dx: 1, dy: 2 },
      { op: 'setDrawingRect', id: ink, rect: [9, 9, 9, 9] },
      { op: 'setDrawingRect', id: img, rect: [5, 5, 6, 6] },
    ])
    const [i, ar, im] = b.state.drawings.map((d) => d.input)
    expect(i!.kind === 'ink' && i!.paths).toEqual([[1, 2, 3, 4]])
    expect(ar!.kind === 'arrow' && [ar!.from, ar!.to]).toEqual([
      [1, 2],
      [2, 3],
    ])
    expect(im!.kind === 'image' && im!.rect).toEqual([5, 5, 6, 6])
  })

  it('patchTextInsert merges fields', () => {
    const a = run(empty(), [
      { op: 'addTextInsert', input: { pageIndex: 0, text: 'x', origin: [1, 1] } },
    ])
    const b = run(a.state, [
      { op: 'patchTextInsert', id: a.state.textInserts[0]!.id, input: { origin: [5, 6] } },
    ])
    expect(b.state.textInserts[0]!.input).toMatchObject({ text: 'x', origin: [5, 6] })
  })

  it('patchImageEdit merges input fields, sets or clears opacityBase, and skips deletes', () => {
    const oldRect: [number, number, number, number] = [0, 0, 10, 10]
    const a = run(empty(), [
      {
        op: 'addImageEdit',
        input: { kind: 'transformImage', pageIndex: 0, oldRect, rect: oldRect },
        opacityBase: 'b',
      },
      { op: 'addImageEdit', input: { kind: 'deleteImage', pageIndex: 0, oldRect: [1, 1, 2, 2] } },
    ])
    const [t, del] = a.state.imageEdits.map((e) => e.id)
    const b = run(a.state, [
      { op: 'patchImageEdit', id: t, input: { layer: 'aboveText', quarterTurns: 1 } },
      { op: 'patchImageEdit', id: del, input: { rect: [0, 0, 1, 1] } },
    ])
    expect(b.state.imageEdits[0]!.input).toMatchObject({
      kind: 'transformImage',
      layer: 'aboveText',
      quarterTurns: 1,
    })
    expect(b.state.imageEdits[0]!.opacityBase).toBe('b')
    expect(b.state.imageEdits[1]!.input).toEqual({
      kind: 'deleteImage',
      pageIndex: 0,
      oldRect: [1, 1, 2, 2],
    })
    const c = run(b.state, [{ op: 'patchImageEdit', id: t, opacityBase: null }])
    expect('opacityBase' in c.state.imageEdits[0]!).toBe(false)
    const d = run(c.state, [{ op: 'patchImageEdit', id: t, opacityBase: 'z' }])
    expect(d.state.imageEdits[0]!.opacityBase).toBe('z')
  })

  it('setStaticFillImage swaps pixels, morphs transforms into replaces, and re-rects the fill', () => {
    const oldRect: [number, number, number, number] = [0, 0, 10, 10]
    const fill = { id: 'f', kind: 'text' as const, pageIndex: 0, rect: oldRect, text: 'old' }
    const a = run(empty(), [
      {
        op: 'addImageEdit',
        input: {
          kind: 'transformImage',
          pageIndex: 0,
          oldRect,
          rect: [5, 5, 15, 15],
          layer: 'aboveText',
        },
        staticFill: fill,
      },
    ])
    const b = run(a.state, [
      {
        op: 'setStaticFillImage',
        id: a.state.imageEdits[0]!.id,
        image: 'NEW',
        staticFill: { ...fill, text: 'new' },
      },
    ])
    const e = b.state.imageEdits[0]!
    expect(e.input).toEqual({
      kind: 'replaceImage',
      pageIndex: 0,
      oldRect,
      rect: [5, 5, 15, 15],
      image: 'NEW',
      layer: 'aboveText',
      quarterTurns: undefined,
    })
    expect(e.staticFill).toMatchObject({ text: 'new', rect: [5, 5, 15, 15] })
  })
})

describe('bucket reduction matches whole-snapshot reduction', () => {
  it('reducing every bucket independently yields the same snapshot and leaves untouched buckets by reference', () => {
    const base = run(empty(), [
      { op: 'addMarkup', markup: markup(1) },
      { op: 'editSavedNote', annot: savedNote(7), contents: 'v' },
      {
        op: 'addDrawing',
        drawing: { kind: 'image', pageIndex: 1, image: 'A', rect: [0, 0, 40, 20] },
      },
    ]).state
    const ops = planEditOps(
      [
        { op: 'deletePage', pageIndex: 1 },
        { op: 'deleteSavedAnnot', annot: savedNote(7) },
        { op: 'setFormValue', value: { name: 'f', kind: 'checkbox', checked: true } },
      ],
      ctx(),
      newId,
    ).ops
    const whole = reduceEditOps(base, ops)
    for (const b of BUCKETS) {
      const viaBucket = reduceBucket(b, base[b], ops, base)
      expect(viaBucket).toEqual(whole[b])
      const touched = [
        'deleted',
        'markups',
        'drawings',
        'formFields',
        'formWidgetEdits',
        'annotDeletes',
        'noteEdits',
        'formEdits',
      ]
      if (!touched.includes(b)) expect(viaBucket).toBe(base[b])
    }
  })
})

describe('edit-op geometry rejects non-finite numbers', () => {
  const NON_FINITE = [Number.NaN, Infinity, -Infinity]

  for (const bad of NON_FINITE) {
    it(`setDrawingRect rejects ${bad}`, () => {
      expect(() =>
        lookup('setDrawingRect').validate(
          { op: 'setDrawingRect', id: 'd', rect: [bad, 10, 200, 50] },
          ctx(),
        ),
      ).toThrow(GuidedError)
    })

    it(`moveDrawing rejects dx=${bad} and dy=${bad}`, () => {
      expect(() =>
        lookup('moveDrawing').validate({ op: 'moveDrawing', id: 'd', dx: bad, dy: 1 }, ctx()),
      ).toThrow(GuidedError)
      expect(() =>
        lookup('moveDrawing').validate({ op: 'moveDrawing', id: 'd', dx: 1, dy: bad }, ctx()),
      ).toThrow(GuidedError)
    })
  }

  it('still accepts finite geometry, including negative and fractional coords', () => {
    expect(() =>
      lookup('setDrawingRect').validate(
        { op: 'setDrawingRect', id: 'd', rect: [-1.5, 0, 200.25, 50] },
        ctx(),
      ),
    ).not.toThrow()
    expect(() =>
      lookup('moveDrawing').validate({ op: 'moveDrawing', id: 'd', dx: -0.5, dy: 0 }, ctx()),
    ).not.toThrow()
  })

  it('a non-finite rect is a planned failure, not a silently applied op', () => {
    const plan = planEditOps(
      [{ op: 'setDrawingRect', id: 'd', rect: [NaN, 10, 200, 50] }],
      ctx(),
      newId,
    )
    expect(plan.ops).toEqual([])
    expect(plan.failures).toHaveLength(1)
  })
})
