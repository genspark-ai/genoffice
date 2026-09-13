import { describe, expect, it, vi } from 'vitest'
import { pressOnEditFrame } from '../src/renderer/TextEditOverlay'

function editor(html: string) {
  const frame = document.createElement('div')
  const ed = document.createElement('div')
  ed.innerHTML = html
  frame.appendChild(ed)
  document.body.appendChild(frame)
  return { frame, ed }
}

function stubLineRects(rects: Array<[number, number, number, number]>) {
  const list = rects.map(([left, top, width, height]) => ({
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
  }))
  vi.spyOn(document, 'createRange').mockReturnValue({
    selectNodeContents: () => {},
    getClientRects: () => list,
  } as unknown as Range)
}

describe('pressOnEditFrame', () => {
  it('treats an empty body as text over the whole frame', () => {
    const { frame, ed } = editor('<div><br></div>')
    expect(pressOnEditFrame(ed, frame, 10, 10)).toBe(false)
  })

  it('a press on the frame or the editor padding grabs the shape', () => {
    const { frame, ed } = editor('<div><span>abc</span></div>')
    expect(pressOnEditFrame(ed, frame, 10, 300)).toBe(true)
    expect(pressOnEditFrame(ed, ed, 10, 300)).toBe(true)
  })

  it('splits a paragraph row into its line box (caret) and the space beside it (drag)', () => {
    const { ed } = editor('<div><span>abc</span></div><div><br></div>')
    stubLineRects([[100, 50, 200, 30]])
    const span = ed.querySelector('span')!
    expect(pressOnEditFrame(ed, span.firstChild, 150, 60)).toBe(false)
    expect(pressOnEditFrame(ed, span, 302, 82)).toBe(false) // inside the 4px pad
    expect(pressOnEditFrame(ed, ed.firstElementChild, 400, 60)).toBe(true)
  })

  it('an empty paragraph row stays a caret line beyond its <br> rect', () => {
    const { ed } = editor('<div><span>abc</span></div><div><br></div>')
    stubLineRects([[100, 80, 0, 30]])
    const empty = ed.children[1]!
    expect(pressOnEditFrame(ed, empty, 400, 95)).toBe(false)
    expect(pressOnEditFrame(ed, empty.firstChild, 100, 95)).toBe(false)
  })
})
