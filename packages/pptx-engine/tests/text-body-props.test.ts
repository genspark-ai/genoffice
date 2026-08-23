/** setElementTextBodyProps: bodyPr direction/autofit/insets/wrap byte surgery. */
import { describe, it, expect } from 'vitest'
import { addElement, createBlankPptx, openPptx, setElementTextBodyProps } from '../src/index'
import type { TextElement } from '../src/types'

async function textboxSlide() {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const el = addElement(slide, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1000, cy: 1000 },
    paragraphs: [{ runs: [{ text: 'x' }] }],
  })
  return { slide, el: el as TextElement }
}

describe('setElementTextBodyProps', () => {
  it('vert attribute set and cleared, model kept in sync', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { vert: 'eaVert' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr vert="eaVert"/)
    expect(el.text!.vert).toBe('eaVert')
    expect(setElementTextBodyProps(slide, el.id, { vert: 'horz' })).toBe(true)
    expect(el.anchor.originalXml).not.toMatch(/<a:bodyPr[^>]*\svert="/)
    expect(el.text!.vert).toBeUndefined()
  })

  it('wrap toggles wrap="none" / wrap="square"', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { wrap: false })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr wrap="none"/)
    expect(el.text!.wrap).toBe(false)
    expect(setElementTextBodyProps(slide, el.id, { wrap: true })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr wrap="square"/)
    expect(el.text!.wrap).toBe(true)
  })

  it('insets written per side (EMU), previously written sides kept', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { insets: { l: 180000 } })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/lIns="180000"/)
    expect(el.text!.insets).toMatchObject({ l: 180000, t: 45720 })
    expect(setElementTextBodyProps(slide, el.id, { insets: { t: 0 } })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/tIns="0"/)
    expect(el.anchor.originalXml).toMatch(/lIns="180000"/)
    expect(el.text!.insets).toMatchObject({ l: 180000, t: 0 })
  })

  it('autofit child swapped in place (self-closing bodyPr expands)', async () => {
    const { slide, el } = await textboxSlide()
    expect(setElementTextBodyProps(slide, el.id, { autofit: 'shrink' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/)
    expect(el.text!.autofit).toBe('shrink')
    expect(setElementTextBodyProps(slide, el.id, { autofit: 'resize' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:spAutoFit\/>/)
    expect(el.anchor.originalXml).not.toMatch(/normAutofit/)
    expect(el.text!.autofit).toBe('resize')
    expect(setElementTextBodyProps(slide, el.id, { autofit: 'none' })).toBe(true)
    expect(el.anchor.originalXml).toMatch(/<a:noAutofit\/>/)
    expect(el.text!.autofit).toBe('none')
  })

  it('combined patch applies attributes and autofit together', async () => {
    const { slide, el } = await textboxSlide()
    expect(
      setElementTextBodyProps(slide, el.id, {
        vert: 'vert270',
        wrap: false,
        insets: { b: 91440 },
        autofit: 'shrink',
      }),
    ).toBe(true)
    const xml = el.anchor.originalXml
    expect(xml).toMatch(/vert="vert270"/)
    expect(xml).toMatch(/wrap="none"/)
    expect(xml).toMatch(/bIns="91440"/)
    expect(xml).toMatch(/<a:normAutofit\/>/)
  })

  it('rejects elements without a text body', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    expect(setElementTextBodyProps(slide, 'nope', { wrap: false })).toBe(false)
  })
})
