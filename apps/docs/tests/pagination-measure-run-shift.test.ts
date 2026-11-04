import { describe, expect, it } from 'vitest'
import { unshiftRunRect } from '../src/renderer/pagination-measure'

function textIn(html: string): Node {
  const p = document.createElement('p')
  p.innerHTML = html
  document.body.appendChild(p)
  return p.querySelector('span')!.firstChild!
}

describe('unshiftRunRect', () => {
  it('moves a raised run rect back onto its line box, scaled by the zoom', () => {
    const t = textIn('<span style="position:relative;top:-13px">up</span>')
    const r = new DOMRect(0, 100, 50, 20)
    expect(unshiftRunRect(r, t, 1).y).toBe(113)
    expect(unshiftRunRect(r, t, 2).y).toBe(126)
  })

  it('leaves unshifted runs alone and ignores the paragraph lead shift', () => {
    const t = textIn('<span>plain</span>')
    ;(t.parentElement!.parentElement as HTMLElement).style.cssText = 'position:relative;top:-4px'
    const r = new DOMRect(0, 100, 50, 20)
    expect(unshiftRunRect(r, t, 1)).toBe(r)
  })
})
