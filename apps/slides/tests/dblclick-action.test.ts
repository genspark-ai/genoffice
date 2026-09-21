import type { RenderNode } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import { dblClickActionFor, type DblClickContext } from '../src/renderer/dblclick-action'

const node = (n: object) => n as unknown as RenderNode
const topLevel: DblClickContext = {
  editable: false,
  insideGroup: false,
  canEnterGroup: true,
  canPlayMedia: true,
}

describe('slides double-click action', () => {
  it('keeps the existing edit gestures ahead of the tab switch', () => {
    expect(dblClickActionFor(node({ type: 'shape' }), { ...topLevel, editable: true })).toEqual({
      kind: 'editText',
    })
    expect(dblClickActionFor(node({ type: 'group', children: [] }), topLevel)).toEqual({
      kind: 'enterGroup',
    })
    expect(dblClickActionFor(node({ type: 'table' }), { ...topLevel, hitCell: true })).toEqual({
      kind: 'editCell',
    })
    expect(dblClickActionFor(node({ type: 'picture', media: 'audio' }), topLevel)).toEqual({
      kind: 'playMedia',
    })
  })

  it('opens the object tools tab for pictures, charts and connectors', () => {
    expect(dblClickActionFor(node({ type: 'picture' }), topLevel)).toEqual({
      kind: 'openTab',
      tab: 'pictureFormat',
    })
    expect(dblClickActionFor(node({ type: 'chart' }), topLevel)).toEqual({
      kind: 'openTab',
      tab: 'chartDesign',
    })
    expect(dblClickActionFor(node({ type: 'shape', line: {} }), topLevel)).toEqual({
      kind: 'openTab',
      tab: 'shapeFormat',
    })
  })

  it('opens Table Design when the double-click misses every cell', () => {
    expect(dblClickActionFor(node({ type: 'table' }), { ...topLevel, hitCell: false })).toEqual({
      kind: 'openTab',
      tab: 'tableDesign',
    })
  })

  it('falls back to the tab when the gesture is unavailable', () => {
    expect(
      dblClickActionFor(node({ type: 'group', children: [] }), {
        ...topLevel,
        canEnterGroup: false,
      }),
    ).toEqual({ kind: 'openTab', tab: 'shapeFormat' })
    expect(
      dblClickActionFor(node({ type: 'picture', media: 'video' }), {
        ...topLevel,
        canPlayMedia: false,
      }),
    ).toEqual({ kind: 'openTab', tab: 'pictureFormat' })
  })

  it('inside an entered group only text editing and the tab remain', () => {
    const inGroup = { ...topLevel, insideGroup: true }
    expect(dblClickActionFor(node({ type: 'group', children: [] }), inGroup)).toEqual({
      kind: 'openTab',
      tab: 'shapeFormat',
    })
    expect(dblClickActionFor(node({ type: 'table' }), { ...inGroup, hitCell: true })).toEqual({
      kind: 'openTab',
      tab: 'tableDesign',
    })
    expect(dblClickActionFor(node({ type: 'picture', media: 'video' }), inGroup)).toEqual({
      kind: 'openTab',
      tab: 'pictureFormat',
    })
  })

  it('does nothing for nodes without a contextual tab', () => {
    expect(dblClickActionFor(node({ type: 'placeholder-chip' }), topLevel)).toBeNull()
  })
})
