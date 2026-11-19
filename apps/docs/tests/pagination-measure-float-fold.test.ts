import { describe, expect, it } from 'vitest'
import { measureBlocks } from '../src/renderer/pagination'

const rectOf = (top: number, height: number, width = 100) =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: width,
    width,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

function flowOf(
  specs: Array<{
    top: number
    height: number
    inner: string
    tag?: string
    cls?: string
    data?: Record<string, string>
  }>,
) {
  const pm = document.createElement('div')
  for (const spec of specs) {
    const el = document.createElement(spec.tag ?? 'p')
    el.innerHTML = spec.inner
    if (spec.cls) el.className = spec.cls
    for (const [k, v] of Object.entries(spec.data ?? {})) el.dataset[k] = v
    el.getBoundingClientRect = () => rectOf(spec.top, spec.height)
    pm.appendChild(el)
  }
  return pm
}

const bottomFloat = (top: number, dy: string) => ({
  top,
  height: 50,
  inner: '<tbody><tr><td>float</td></tr></tbody>',
  tag: 'table',
  cls: 'doc-table doc-table-float-left',
  data: { tblpVy: '0', tblpVanchor: 'margin', tblpVspec: 'bottom', tblpDy: dy },
})

describe('measureBlocks — page-anchored float and the inter-block margin fold', () => {
  it('keeps the applied --tblp-dy shift out of the float height and space after', () => {
    // the float sits 206px below its flow position (shift applied) and the
    // section opener clears it, 6px of float text distance further down
    const pm = flowOf([
      { top: 0, height: 20, inner: 'cover' },
      bottomFloat(226, '206'),
      { top: 282, height: 20, inner: 'opener', cls: 'doc-sect-page-start' },
      { top: 302, height: 20, inner: 'body' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[1].pageRelVyPx).toBe(0)
    expect(blocks[1].top).toBe(20)
    expect(blocks[1].height).toBe(50)
    expect(blocks[1].spaceAfterPx).toBeUndefined()
    // the opener starts where the shifted float ends; the text distance
    // becomes its own space after
    expect(blocks[2].top).toBe(276)
    expect(blocks[2].spaceAfterPx).toBe(6)
    expect(blocks[3].top).toBe(302)
  })

  it('does not fold the clearing block distance into an unshifted float either', () => {
    const pm = flowOf([
      { top: 0, height: 20, inner: 'cover' },
      bottomFloat(20, '0'),
      { top: 90, height: 20, inner: 'opener', cls: 'doc-sect-page-start' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[1].height).toBe(50)
    expect(blocks[1].spaceAfterPx).toBeUndefined()
  })

  it('still folds the CSS margin after an in-flow block', () => {
    const pm = flowOf([
      { top: 0, height: 20, inner: 'one' },
      { top: 30, height: 20, inner: 'two' },
    ])
    const { blocks } = measureBlocks(pm, 0, 1)
    expect(blocks[0].spaceAfterPx).toBe(10)
    expect(blocks[0].height).toBe(30)
  })
})
