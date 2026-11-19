/**
 * Format painter pure-function unit tests — extract / apply / cross-type intersection.
 */
import { describe, it, expect } from 'vitest'
import { extractFormat, applyFormat } from '../src/format-brush'
import type { TextElement } from '../src/types'

// ── Helpers ─────────────────────────────────────────────────────────────

function makeTextEl(partial: Partial<TextElement>): TextElement {
  return {
    id: 'el1',
    type: 'text',
    anchor: { spIndex: 0, originalXml: '', range: [0, 0] },
    transform: { offset: { x: 0, y: 0, cx: 100, cy: 50 }, rot: 0, flipH: false, flipV: false },
    ...partial,
  } as TextElement
}

// ── extractFormat ────────────────────────────────────────────────────────

describe('extractFormat', () => {
  it('extracts fill / stroke / run format', () => {
    const el = makeTextEl({
      type: 'shape',
      fill: { type: 'solid', color: '#FF0000' },
      stroke: { fill: { type: 'solid', color: '#000000' }, width: 12700 },
      text: {
        paragraphs: [
          {
            align: 'center',
            runs: [
              {
                text: 'hello',
                bold: true,
                italic: false,
                fontSize: 24,
                color: '#333333',
                fontFamily: 'Arial',
              },
            ],
          },
        ],
        anchor: 'top',
        wrap: true,
      },
    })

    const fmt = extractFormat(el)

    expect(fmt.sourceType).toBe('shape')
    expect(fmt.shape?.fill).toEqual({ type: 'solid', color: '#FF0000' })
    expect(fmt.shape?.stroke?.width).toBe(12700)
    expect(fmt.run?.bold).toBe(true)
    expect(fmt.run?.fontSize).toBe(24)
    expect(fmt.run?.fontFamily).toBe('Arial')
    expect(fmt.run?.color).toBe('#333333')
    expect(fmt.paragraph?.align).toBe('center')
  })

  it('shape without text extracts only the spPr part', () => {
    const el = makeTextEl({
      type: 'shape',
      fill: { type: 'solid', color: '#0000FF' },
      text: undefined,
    })
    const fmt = extractFormat(el)
    expect(fmt.shape?.fill).toEqual({ type: 'solid', color: '#0000FF' })
    expect(fmt.run).toBeUndefined()
    expect(fmt.paragraph).toBeUndefined()
  })

  it('plain textbox without fill leaves the shape field undefined', () => {
    const el = makeTextEl({
      type: 'text',
      fill: undefined,
      stroke: undefined,
      text: {
        paragraphs: [{ runs: [{ text: 'abc', fontFamily: 'Calibri', fontSize: 18 }] }],
        wrap: true,
      },
    })
    const fmt = extractFormat(el)
    expect(fmt.shape).toBeUndefined()
    expect(fmt.run?.fontFamily).toBe('Calibri')
    expect(fmt.run?.fontSize).toBe(18)
  })
})

// ── applyFormat ──────────────────────────────────────────────────────────

describe('applyFormat', () => {
  const fmt = {
    sourceType: 'shape' as const,
    shape: {
      fill: { type: 'solid' as const, color: '#FF0000' },
      stroke: { fill: { type: 'solid' as const, color: '#000000' }, width: 12700 },
    },
    run: { bold: true, fontSize: 24, fontFamily: 'Arial', color: '#111111' },
    paragraph: { align: 'center' as const },
  }

  it('apply to shape: fill + stroke + run + alignment all applied', () => {
    const result = applyFormat(fmt, 'shape')
    expect(result.fill).toEqual({ type: 'solid', color: '#FF0000' })
    expect(result.stroke?.width).toBe(12700)
    expect(result.runFormat?.bold).toBe(true)
    expect(result.align).toBe('center')
  })

  it('apply to text: same full set as shape', () => {
    const result = applyFormat(fmt, 'text')
    expect(result.fill).toEqual({ type: 'solid', color: '#FF0000' })
    expect(result.runFormat?.fontSize).toBe(24)
    expect(result.align).toBe('center')
  })

  it('apply to picture: keeps only the stroke, does not apply fill / text', () => {
    const result = applyFormat(fmt, 'picture')
    expect(result.fill).toBeUndefined()
    expect(result.stroke?.width).toBe(12700)
    expect(result.runFormat).toBeUndefined()
    expect(result.align).toBeUndefined()
  })

  it('without shape format, both fill/stroke are undefined', () => {
    const fmtTextOnly = {
      sourceType: 'text' as const,
      run: { fontSize: 12 },
      paragraph: { align: 'left' as const },
    }
    const result = applyFormat(fmtTextOnly, 'shape')
    expect(result.fill).toBeUndefined()
    expect(result.stroke).toBeUndefined()
    expect(result.runFormat?.fontSize).toBe(12)
    expect(result.align).toBe('left')
  })
})
