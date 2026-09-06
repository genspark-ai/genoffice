import { describe, expect, it } from 'vitest'
import type { HfParagraph, HfTextBox } from '@genoffice/docx-engine'
import {
  hfTextBoxClass,
  hfTextBoxStyle,
  makeGapHfEl,
  type HfStripGeom,
} from '../src/renderer/editor/hf-dom'

/**
 * Header/footer text boxes that share their paragraph with text hang off that
 * paragraph (paragraph-relative offsets measure from its top); boxes are kept
 * on the paper horizontally and honour wrap="none".
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

const strip = (paras: HfParagraph[]) =>
  makeGapHfEl({ kind: 'footer', value: { text: '', paras }, pageNo: 7, pageTotal: 9, geom })

const line = (text: string): HfParagraph => ({ runs: [{ text }] })
const boxed = (text: string, box: HfTextBox): HfParagraph => ({
  runs: [{ text }],
  boxAnchored: true,
  box,
})

describe('anchored footer boxes hosted by their paragraph', () => {
  it('hangs an offset box off its anchor paragraph, clamped to the paper edge', () => {
    const box: HfTextBox = {
      id: 1,
      widthPx: 248,
      posHRel: 'margin',
      posXPx: 542,
      posVRel: 'paragraph',
      posYPx: 32,
      anchorPara: 1,
    }
    const el = strip([line('first'), line('second'), boxed('Prepared by', box)])
    const paras = el.querySelectorAll<HTMLElement>(':scope > .page-hf-para')
    expect(paras).toHaveLength(2)
    expect(paras[0].classList.contains('page-hf-anchor')).toBe(false)
    expect(paras[1].classList.contains('page-hf-anchor')).toBe(true)
    const host = paras[1].querySelector<HTMLElement>(':scope > .page-hf-textbox')!
    expect(host.style.top).toBe('32px')
    // 96 + 542 + 248 runs 70px past the paper: pulled back to end on its edge
    expect(host.style.left).toBe(`${816 - 248 - 96}px`)
    expect(host.textContent).toBe('Prepared by')
    expect(el.querySelectorAll('.page-hf-textbox')).toHaveLength(1)
  })

  it('applies paragraph-relative bottom/center aligns against the host', () => {
    const bottom = hfTextBoxStyle(
      {
        id: 2,
        posVRel: 'paragraph',
        posV: 'bottom',
        posHRel: 'margin',
        posH: 'right',
        anchorPara: 0,
      },
      'footer',
      geom,
    )!
    expect(bottom.top).toBe('100%')
    expect(bottom.bottom).toBeUndefined()
    expect(bottom.transform).toBe('translate(-100%, -100%)')
    const center = hfTextBoxStyle({ id: 3, posVRel: 'paragraph', posV: 'center' }, 'header', geom)!
    expect(center.top).toBe('50%')
    expect(center.transform).toBe('translate(0%, -50%)')
  })

  it('clips a fixed-size box but lets an autofit box grow', () => {
    const fixed = { id: 8, widthPx: 215, heightPx: 16, posXPx: 512, posYPx: 33 }
    expect(hfTextBoxStyle(fixed, 'header', geom)!.overflow).toBe('hidden')
    expect(hfTextBoxStyle({ ...fixed, autofit: true }, 'header', geom)!.overflow).toBeUndefined()
    expect(
      hfTextBoxStyle({ ...fixed, heightPx: undefined }, 'header', geom)!.overflow,
    ).toBeUndefined()
  })

  it('keeps a wrap="none" label on one line and page-aligns against the paper', () => {
    const style = hfTextBoxStyle(
      {
        id: 4,
        widthPx: 287,
        posH: 'left',
        posHRel: 'page',
        posV: 'bottom',
        posVRel: 'page',
        nowrap: true,
      },
      'footer',
      geom,
    )!
    expect(hfTextBoxClass({ id: 4, nowrap: true })).toBe('page-hf-textbox page-hf-nowrap')
    expect(hfTextBoxClass({ id: 4 })).toBe('page-hf-textbox')
    const el = strip([
      boxed('One line label', {
        id: 6,
        posH: 'left',
        posHRel: 'page',
        posV: 'bottom',
        posVRel: 'page',
        nowrap: true,
      }),
    ])
    expect(el.querySelector('.page-hf-textbox')!.classList.contains('page-hf-nowrap')).toBe(true)
    // page-relative left: the paper edge, 96px left of the strip
    expect(style.left).toBe('-96px')
    // page-relative bottom: the box bottom on the paper edge (footerDist below the strip bottom)
    expect(style.bottom).toBe(`${-geom.footerDist}px`)
  })

  it('a box without a shared paragraph stays on the strip', () => {
    const el = strip([
      boxed('Classified', {
        id: 5,
        posH: 'left',
        posHRel: 'page',
        posV: 'bottom',
        posVRel: 'page',
      }),
    ])
    expect(el.querySelectorAll(':scope > .page-hf-textbox')).toHaveLength(1)
    expect(el.querySelectorAll('.page-hf-anchor')).toHaveLength(0)
  })
})
