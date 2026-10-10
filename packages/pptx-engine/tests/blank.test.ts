import { describe, it, expect } from 'vitest'
import { openPptx, savePptx, createBlankPptx, addElement, duplicateSlide } from '../src/index'
import { parseTheme, resolveFontRef } from '../src/theme'

describe('createBlankPptx', () => {
  it('opens as a 16:9 single blank slide', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(opened.deck.slides.length).toBe(1)
    expect(opened.deck.size).toEqual({ cx: 12192000, cy: 6858000 })
    expect(opened.deck.slides[0]!.elements.length).toBe(0)
  })

  it('theme default fonts: Calibri latin + empty ea carried by per-script entries', async () => {
    const opened = await openPptx(await createBlankPptx())
    const theme = opened.archive.readText('ppt/theme/theme1.xml')!
    expect(theme).toContain('<a:latin typeface="Calibri"/>')
    expect(theme.match(/<a:ea typeface=""\/>/g)?.length).toBe(2)
    // the per-script faces PowerPoint's own blank Office theme ships: Jpan/Hang/Hans/Hant
    expect(theme.match(/<a:font script=/g)?.length).toBe(8)
  })

  it('a theme ea ref resolves per script instead of a Windows-only pin', async () => {
    const opened = await openPptx(await createBlankPptx())
    const theme = parseTheme(opened.archive.readText('ppt/theme/theme1.xml')!)
    expect(theme.majorEaFont).toBeUndefined()
    expect(theme.minorEaFont).toBeUndefined()
    expect(resolveFontRef('+mj-ea', theme, 'sc')).toBe('等线')
    expect(resolveFontRef('+mj-ea', theme, 'tc')).toBe('新細明體')
    expect(resolveFontRef('+mj-ea', theme, 'ja')).toBe('ＭＳ ゴシック')
    expect(resolveFontRef('+mj-ea', theme, 'ko')).toBe('맑은 고딕')
    expect(resolveFontRef('+mn-ea', theme, 'ja')).toBe('ＭＳ Ｐゴシック')
    expect(resolveFontRef('+mn-ea', theme, 'sc')).toBe('等线')
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
