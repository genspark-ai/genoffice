import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

const render = (spec: unknown): HTMLElement =>
  DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement

describe('nested table borders', () => {
  it('a borderless nested table resets the host table border variables', () => {
    const nested: TableModel = { rows: [[{ paras: ['inner'] }]], cellSpacingTwips: 15 }
    const style = render(renderTableSpec(nested, true)).getAttribute('style') ?? ''
    expect(style).toMatch(/--doc-b-t:\s*none/)
    expect(style).toMatch(/--doc-b-v:\s*none/)
  })

  it('a top-level table without borders declares nothing', () => {
    const style =
      render(renderTableSpec({ rows: [[{ paras: ['x'] }]] })).getAttribute('style') ?? ''
    expect(style).not.toContain('--doc-b-t')
  })
})
