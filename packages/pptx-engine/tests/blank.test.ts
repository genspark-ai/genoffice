import { describe, it, expect } from 'vitest'
import { openPptx, savePptx, createBlankPptx, addElement, duplicateSlide } from '../src/index'

describe('createBlankPptx', () => {
  it('opens as a 16:9 single blank slide', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(opened.deck.slides.length).toBe(1)
    expect(opened.deck.size).toEqual({ cx: 12192000, cy: 6858000 })
    expect(opened.deck.slides[0]!.elements.length).toBe(0)
  })

  it('theme default fonts: Calibri latin + Microsoft YaHei ea (cross-platform CJK)', async () => {
    const opened = await openPptx(await createBlankPptx())
    const theme = opened.archive.readText('ppt/theme/theme1.xml')!
    expect(theme).toContain('<a:latin typeface="Calibri"/>')
    expect(theme.match(/<a:ea typeface="Microsoft YaHei"\/>/g)?.length).toBe(2)
  })

  it('supports the full edit pipeline: add element + add slide + save round-trip', async () => {
    const opened = await openPptx(await createBlankPptx())
    addElement(opened.deck.slides[0]!, {
      kind: 'textbox',
      offset: { x: 914400, y: 914400, cx: 6096000, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'Generated Title', bold: true, fontSize: 40 }] }],
    })
    duplicateSlide(opened, 0, { clearText: true })

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.slides.length).toBe(2)
    const texts = reopened.deck.slides[0]!.elements.flatMap((e: any) => e.text?.paragraphs ?? [])
      .flatMap((p: any) => p.runs)
      .map((r: any) => r.text)
      .join('')
    expect(texts).toBe('Generated Title')
  })
})
