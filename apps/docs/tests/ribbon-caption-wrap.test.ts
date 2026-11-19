import { describe, expect, it } from 'vitest'
import { twoLineWidth, wrapBigCaptions } from '../src/renderer/components/ribbon-caption-wrap'

describe('twoLineWidth', () => {
  it('puts the break where the wider line is narrowest', () => {
    // "Insert Table of Figures": Insert Table | of Figures
    expect(twoLineWidth([34, 30, 12, 42], 3)).toBe(67)
  })

  it('gives each of two words its own line', () => {
    expect(twoLineWidth([30, 45], 3)).toBe(45)
  })
})

describe('wrapBigCaptions', () => {
  it('drops the width a longer caption left behind once the caption is one word', () => {
    const root = document.createElement('div')
    root.innerHTML =
      '<button class="rb-big"><span class="rb-big-icon"></span><span>Seitenzahl</span></button>'
    const caption = root.querySelector<HTMLElement>('.rb-big > span:not(.rb-big-icon)')!
    caption.style.maxWidth = '40px'
    wrapBigCaptions(root)
    expect(caption.style.maxWidth).toBe('')
  })
})
