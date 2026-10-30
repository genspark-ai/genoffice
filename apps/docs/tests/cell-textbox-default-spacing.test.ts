/** Normal/docDefaults paragraph spacing reaches nested-table cell and textbox paragraphs */
import { DOMSerializer } from '@tiptap/pm/model'
import type { ParsedDocFull, StyleInfo, TableModel } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { setAdditiveParagraphSpacing } from '../src/renderer/line-metrics'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const render = (spec: unknown): HTMLElement =>
  DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement

const cellDivs = (model: TableModel) =>
  Array.from(render(renderTableSpec(model, true)).querySelectorAll<HTMLElement>('td > div'))

describe('nested-table cell paragraph spacing', () => {
  it('leaves undeclared spacing to the stylesheet and tags the block', () => {
    const [div] = cellDivs({ rows: [[{ paras: [], richParas: [{ runs: [{ text: 'x' }] }] }]] })
    expect(div.classList.contains('doc-cell-para')).toBe(true)
    expect(div.style.marginBottom).toBe('')
    expect(div.style.marginTop).toBe('')
  })

  it('emits a direct 0 so it beats the inherited default', () => {
    const [div] = cellDivs({
      rows: [
        [{ paras: [], richParas: [{ runs: [{ text: 'x' }], spaceAfter: 0, spaceBefore: 0 }] }],
      ],
    })
    expect(div.style.marginBottom).toBe('0pt')
    expect(div.style.marginTop).toBe('0pt')
  })

  it('inlines table-style spacing and exposes the paragraph style', () => {
    const effective = {
      spacing: { before: 0, after: 0, beforeAuto: false, afterAuto: false },
      source: { spaceAfter: 'tableStyle', spaceBefore: 'docDefaults' },
    }
    const [div] = cellDivs({
      rows: [
        [
          {
            paras: [],
            richParas: [{ runs: [{ text: 'x' }], styleId: 'Caption', effective } as never],
          },
        ],
      ],
    })
    expect(div.style.marginBottom).toBe('0pt')
    expect(div.style.marginTop).toBe('')
    expect(div.getAttribute('data-style')).toBe('Caption')
  })

  describe('additive documents (doNotUseHTMLParagraphAutoSpacing)', () => {
    afterEach(() => setAdditiveParagraphSpacing(false))

    it('inlines the direct before as padding-top so it replaces the stylesheet default', () => {
      setAdditiveParagraphSpacing(true)
      const parsed = {
        styles: new Map(),
        docDefaults: { spaceBeforeTwips: 200 },
        doNotUseHTMLParagraphAutoSpacing: true,
        blocks: [],
      } as unknown as ParsedDocFull
      const rule = docStyleCss(parsed)
        .split('\n')
        .find((r) => r.startsWith('.doc-page p, '))
      expect(rule).toContain('.doc-cell-para {')
      expect(rule).toContain('margin-top:0;padding-top:10.0pt')
      const [div] = cellDivs({
        rows: [[{ paras: [], richParas: [{ runs: [{ text: 'x' }], spaceBefore: 120 }] }]],
      })
      expect(div.style.paddingTop).toBe('6pt')
      expect(div.style.marginTop).toBe('')
    })
  })
})

describe('docStyleCss default block spacing', () => {
  it('covers cell and textbox paragraph divs with the docDefaults after', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('Normal', {
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: { lineSpacing: 1.15 },
    } as StyleInfo)
    const parsed = {
      styles,
      docDefaults: { spaceAfterTwips: 160 },
      blocks: [],
    } as unknown as ParsedDocFull
    const rule = docStyleCss(parsed)
      .split('\n')
      .find((r) => r.startsWith('.doc-page p, '))
    expect(rule).toContain('.doc-page .doc-textbox-para, .doc-page .doc-cell-para {')
    expect(rule).toContain('margin-bottom:8.0pt')
  })
})
