import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { renderStaticFormMark, renderStaticFormText } from '../src/renderer/static-form-fill'

interface FakeContext {
  font: string
  fillStyle: string
  strokeStyle: string
  lineWidth: number
  lineCap: string
  lineJoin: string
  textAlign: string
  textBaseline: string
  texts: string[]
  moves: number
  lines: number
  scale: () => void
  measureText: (text: string) => { width: number }
  fillText: (text: string) => void
  beginPath: () => void
  moveTo: () => void
  lineTo: () => void
  stroke: () => void
}

let contexts: FakeContext[]

function fakeContext(): FakeContext {
  const context: FakeContext = {
    font: '',
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 0,
    lineCap: '',
    lineJoin: '',
    textAlign: '',
    textBaseline: '',
    texts: [],
    moves: 0,
    lines: 0,
    scale: vi.fn(),
    measureText: (text) => ({ width: text.length * 10 }),
    fillText: (text) => context.texts.push(text),
    beginPath: vi.fn(),
    moveTo: () => {
      context.moves++
    },
    lineTo: () => {
      context.lines++
    },
    stroke: vi.fn(),
  }
  return context
}

beforeEach(() => {
  contexts = []
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => {
    const context = fakeContext()
    contexts.push(context)
    return context as unknown as CanvasRenderingContext2D
  })
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue(
    'data:image/png;base64,STATICFILL',
  )
})

afterEach(() => vi.restoreAllMocks())

describe('static form fill bitmaps', () => {
  it('ignores blank text and rasterizes multiline CJK text', () => {
    expect(renderStaticFormText('   ')).toBeNull()
    const result = renderStaticFormText('姓名\n王小明', 16)

    expect(result).toMatchObject({ kind: 'image', image: 'STATICFILL' })
    expect(result?.width).toBeGreaterThan(30)
    expect(result?.height).toBeGreaterThan(32)
    expect(contexts.at(-1)?.texts).toEqual(['姓名', '王小明'])
  })

  it('paints distinct vector paths for check and cross marks', () => {
    const check = renderStaticFormMark('check')
    const checkContext = contexts.at(-1)
    const cross = renderStaticFormMark('cross')
    const crossContext = contexts.at(-1)

    expect(check).toMatchObject({ image: 'STATICFILL', width: 22, height: 22 })
    expect(cross).toMatchObject({ image: 'STATICFILL', width: 22, height: 22 })
    expect(checkContext).toMatchObject({ moves: 1, lines: 2 })
    expect(crossContext).toMatchObject({ moves: 2, lines: 2 })
  })
})
