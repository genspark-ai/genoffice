import { describe, expect, it } from 'vitest'
import { backgroundPickStyles, backgroundSwatch } from '../src/renderer/document/background-style'

const GRADIENT = 'linear-gradient(135deg, rgb(0, 0, 0) 0%, rgb(26, 26, 26) 100%)'
const PICTURE = 'url("https://x/y.png")'

describe('backgroundSwatch', () => {
  it('shows the colour when there is one, the gradient when the colour is transparent', () => {
    expect(backgroundSwatch('#ff0000', GRADIENT)).toBe('#ff0000')
    expect(backgroundSwatch('', GRADIENT)).toBe(GRADIENT)
    expect(backgroundSwatch('', 'none')).toBe('')
  })

  it('does not paint picture layers, with or without a gradient on top', () => {
    expect(backgroundSwatch('', PICTURE)).toBe('')
    expect(backgroundSwatch('', `${GRADIENT}, ${PICTURE}`)).toBe('')
  })
})

describe('backgroundPickStyles', () => {
  it('replaces a gradient by the picked colour', () => {
    expect(backgroundPickStyles('#112233', GRADIENT)).toEqual({
      'background-image': 'none',
      'background-color': '#112233',
    })
    expect(backgroundPickStyles(null, GRADIENT)).toEqual({
      'background-image': 'none',
      'background-color': 'transparent',
    })
  })

  it('keeps a picture and only writes the colour underneath', () => {
    expect(backgroundPickStyles('#112233', PICTURE)).toEqual({ 'background-color': '#112233' })
    expect(backgroundPickStyles('#112233', 'none')).toEqual({ 'background-color': '#112233' })
  })
})
