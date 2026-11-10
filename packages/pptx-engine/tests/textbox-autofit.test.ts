/**
 * Text boxes grow with their text like PowerPoint's own: generated boxes default to
 * spAutoFit, and landing normalizes cloud html→pptx pages (pptxgenjs-style output with
 * no autofit child) the same way.
 */
import { describe, it, expect } from 'vitest'
import PptxGenJS from 'pptxgenjs'
import {
  addElement,
  autofitGeneratedTextBoxes,
  createBlankPptx,
  openPptx,
  savePptx,
  type TextElement,
} from '../src/index'

const OFF = { x: 914400, y: 914400, cx: 4572000, cy: 914400 }

async function cloudLikePage(): Promise<Uint8Array> {
  const p = new PptxGenJS()
  const s = p.addSlide()
  s.addText('Title', { x: 1, y: 1, w: 8, h: 1, fontSize: 32, isTextBox: true })
  s.addText('Shrunk', { x: 1, y: 2, w: 8, h: 1, fontSize: 32, isTextBox: true, fit: 'shrink' })
  s.addText('Card', {
    x: 1,
    y: 3,
    w: 8,
    h: 1,
    fontSize: 32,
    shape: 'roundRect',
    fill: { color: 'DDDDDD' },
  })
  const buf = (await p.write({ outputType: 'nodebuffer' })) as Buffer
  return new Uint8Array(buf)
}

describe('text box autofit defaults', () => {
  it('addElement textbox writes spAutoFit; a preset shape stays without autofit', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const box = addElement(slide, {
      kind: 'textbox',
      offset: OFF,
      paragraphs: [{ runs: [{ text: 'hi' }] }],
    })
    const card = addElement(slide, {
      kind: 'roundRect',
      offset: OFF,
      paragraphs: [{ runs: [{ text: 'hi' }] }],
    })
    expect(box.text?.autofit).toBe('resize')
    expect(box.anchor.originalXml).toContain('<a:spAutoFit/>')
    expect(card.text?.autofit).toBeUndefined()
    expect(card.anchor.originalXml).not.toMatch(/Autofit|AutoFit/)

    const reopened = await openPptx(await savePptx(opened))
    const [b2, c2] = reopened.deck.slides[0]!.elements as TextElement[]
    expect(b2!.text!.autofit).toBe('resize')
    expect(c2!.text!.autofit).toBe('none')
  })

  it('an explicit autoFit still wins', async () => {
    const opened = await openPptx(await createBlankPptx())
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'textbox',
      offset: OFF,
      bodyPr: { autoFit: 'shrink' },
    })
    expect(el.text?.autofit).toBe('shrink')
    expect(el.anchor.originalXml).toContain('<a:normAutofit/>')
  })
})

describe('autofitGeneratedTextBoxes', () => {
  it('gives bare txBox bodies spAutoFit, keeps normAutofit and autoshapes', async () => {
    const opened = await openPptx(await cloudLikePage())
    const slide = opened.deck.slides[0]!
    const byText = (t: string) =>
      slide.elements.find((e) =>
        (e as TextElement).text?.paragraphs.some((p) => p.runs.some((r) => r.text === t)),
      ) as TextElement
    expect(byText('Title').text!.autofit).toBe('none')
    expect(autofitGeneratedTextBoxes(slide)).toBe(1)
    expect(byText('Title').text!.autofit).toBe('resize')
    expect(byText('Shrunk').text!.autofit).toBe('shrink')
    expect(byText('Card').text!.autofit).toBe('none')
    // idempotent
    expect(autofitGeneratedTextBoxes(slide)).toBe(0)

    const reopened = await openPptx(await savePptx(opened))
    const title = reopened.deck.slides[0]!.elements.find(
      (e) => (e as TextElement).text?.paragraphs[0]?.runs[0]?.text === 'Title',
    ) as TextElement
    expect(title.text!.autofit).toBe('resize')
    expect(title.anchor.originalXml).toContain('<a:spAutoFit/>')
  })
})
