/**
 * Word keeps hangul words whole (the renderer's word-break:keep-all) unless the
 * paragraph turns w:wordWrap off, when lines may break between syllables
 * (Word probe 2026-09-03: "방위|군의"; Latin words still stay whole, which is
 * Chromium's word-break:normal).
 */
import { describe, expect, it } from 'vitest'
import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel } from '@genoffice/docx-engine'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

const render = (spec: unknown): HTMLElement =>
  DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement

function cellStyle(wordWrap: boolean | undefined): string {
  const model: TableModel = {
    rows: [
      [
        {
          paras: ['가맹점 PG사'],
          richParas: [
            { runs: [{ text: '가맹점 PG사' }], ...(wordWrap === undefined ? {} : { wordWrap }) },
          ],
        },
      ],
    ],
  }
  return render(renderTableSpec(model)).querySelector('td div')!.getAttribute('style') ?? ''
}

describe('hangul word breaking follows w:wordWrap', () => {
  it('keeps words whole by default', () => {
    expect(cellStyle(undefined)).toMatch(/word-break:\s*keep-all/)
  })

  it('drops keep-all when the paragraph allows character-level breaking', () => {
    expect(cellStyle(false)).not.toContain('keep-all')
  })

  it('an explicit on keeps words whole', () => {
    expect(cellStyle(true)).toMatch(/word-break:\s*keep-all/)
  })
})
