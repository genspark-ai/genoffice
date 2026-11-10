import type { TableModel } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

// raw spec attrs: jsdom's style parser drops shape-outside
const styleOf = (model: TableModel, nested: boolean): string =>
  ((renderTableSpec(model, nested) as unknown[])[1] as { style?: string }).style ?? ''

const anchored: TableModel = {
  rows: [[{ paras: ['a'] }], [{ paras: ['b'] }]],
  colWidthsTwips: [9392],
  fixedLayout: true,
  floatSide: 'left',
  floatPos: {
    xTwips: 0,
    yTwips: 261,
    horzAnchor: 'margin',
    vertAnchor: 'page',
    distanceTwips: { left: 180, right: 180 },
  },
}

describe('floating table nested in a cell', () => {
  it('floats at tblpY below the cell content top even under a page anchor', () => {
    const style = styleOf(anchored, true)
    expect(style).toContain('float:left')
    expect(style).toContain('margin-top:17.4px')
    expect(style).toContain('shape-outside:inset(17.4px 0 0 0)')
    expect(style).toContain('margin-right:12.0px')
    expect(style).not.toContain('margin-left')
  })

  it('a top-level table keeps the DocTable float path (no inline float here)', () => {
    expect(styleOf(anchored, false)).not.toContain('float:')
  })

  it('a nested table too tall to float on one page flows inline', () => {
    const tall: TableModel = {
      ...anchored,
      rows: Array.from({ length: 60 }, () => [{ paras: ['r'] }]),
    }
    expect(styleOf(tall, true)).not.toContain('float:')
  })

  it('a keyword Y alignment carries no offset', () => {
    const spec: TableModel = {
      ...anchored,
      floatPos: { ...anchored.floatPos!, ySpec: 'top' },
    }
    const style = styleOf(spec, true)
    expect(style).toContain('float:left')
    expect(style).not.toContain('margin-top')
  })
})
