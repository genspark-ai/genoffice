import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import type { HfImage, HfParagraph, HfTextBox } from '@genoffice/docx-engine'
import { HfContent, hfEditorSlot } from '../src/renderer/components/HeaderFooterArea'
import type { HfStripGeom } from '../src/renderer/editor/hf-dom'

/**
 * While a header is edited the strip renders its display-only paragraphs
 * around the editor; the editor stands in for the editable ones.
 */

const geom: HfStripGeom = {
  pageW: 816,
  pageH: 1056,
  marginLeft: 96,
  marginRight: 96,
  marginTop: 96,
  marginBottom: 96,
  headerStripTop: 48,
  footerDist: 48,
}

const line = (text: string): HfParagraph => ({ runs: [{ text }] })
const row: HfParagraph = {
  runs: [],
  row: { heightTwips: 600, heightRule: 'atLeast' },
  // flattened out of a nested table: the one kind of row that stays display-only
  cells: [{ paras: [[{ text: 'LOGO' }]], widthTwips: 4000, locked: true }],
}

const render = (
  paras: HfParagraph[],
  only?: (p: HfParagraph, i: number) => boolean,
  images?: HfImage[],
) =>
  renderToStaticMarkup(
    createElement(HfContent, {
      kind: 'header',
      paras,
      images,
      display: (t) => t,
      boxGeom: geom,
      only,
    }),
  )
const halves = (paras: HfParagraph[], images?: HfImage[]) => {
  const slot = hfEditorSlot(paras)
  return [
    render(paras, (_, i) => i < slot, images),
    render(paras, (_, i) => i >= slot, images),
  ] as const
}

describe('strip runs carry their revision classes', () => {
  it('marks inserted and deleted runs so the Markup view CSS applies', () => {
    const rev = { author: 'a', date: '2026-01-01T00:00:00Z' }
    const html = render([
      { runs: [{ text: 'kept' }, { text: 'gone', del: rev }, { text: 'new', ins: rev }] },
    ])
    expect(html).toContain('class="doc-del"')
    expect(html).toContain('class="doc-ins"')
    expect(html).toContain('gone')
  })
})

describe('display-only strip halves around the header editor', () => {
  it('keeps a paragraph-anchored box whose host paragraph is in the editor', () => {
    const box: HfTextBox = {
      id: 1,
      widthPx: 200,
      posHRel: 'margin',
      posXPx: 10,
      posVRel: 'paragraph',
      posYPx: 12,
      anchorPara: 0,
    }
    const paras = [line('Company'), { runs: [{ text: 'Confidential' }], boxAnchored: true, box }]
    const [above, below] = halves(paras)
    expect(above).toBe('')
    expect(below).toContain('Confidential')
    expect(below).not.toContain('Company')
  })

  it('draws a logo hanging on a display-only row once, in the row half', () => {
    const logo: HfImage = {
      dataUrl: 'data:image/png;base64,AA==',
      floating: true,
      anchorPara: 0,
      posVRel: 'paragraph',
      posYPx: 4,
      widthPx: 40,
      heightPx: 40,
    }
    const [above, below] = halves([row, line('Company')], [logo])
    expect(above.match(/<img /g)).toHaveLength(1)
    expect(below).not.toContain('<img ')
  })

  it('puts the editor after a locked-table-only strip, where typed text is appended', () => {
    expect(hfEditorSlot([row])).toBe(1)
    expect(hfEditorSlot([row, line('x')])).toBe(1)
    expect(hfEditorSlot([line('x'), row])).toBe(0)
    const [above, below] = halves([row])
    expect(above).toContain('LOGO')
    expect(below).toBe('')
  })

  it('an ordinary layout-table row belongs to the editor, not to either half', () => {
    const editable: HfParagraph = { runs: [], cells: [{ paras: [[{ text: 'LOGO' }]] }] }
    expect(hfEditorSlot([editable, line('x')])).toBe(0)
    const [above, below] = halves([editable, line('x')])
    expect(above).toBe('')
    expect(below).toBe('')
  })
})
