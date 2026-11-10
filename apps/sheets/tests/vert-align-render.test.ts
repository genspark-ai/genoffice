/**
 * A whole-cell superscript/subscript must reach the canvas: Univer's plain
 * painter drops `va`, so the render cache is handed a rich-text twin that
 * takes the document layout path (which honors `va` on a run).
 */
import { BaselineOffset, BooleanNumber, CellValueType, type ICellData } from '@univerjs/core'
import { SpreadsheetSkeleton } from '@univerjs/engine-render'
import { describe, expect, it, vi } from 'vitest'

import {
  hasVertAlign,
  installVertAlignRender,
  vertAlignRunStyle,
  vertAlignTwin,
} from '../src/renderer/vert-align-render'

const SUPER = { va: BaselineOffset.SUPERSCRIPT }

describe('hasVertAlign', () => {
  it('is true only for a raised or lowered baseline', () => {
    expect(hasVertAlign(SUPER)).toBe(true)
    expect(hasVertAlign({ va: BaselineOffset.SUBSCRIPT })).toBe(true)
    expect(hasVertAlign({ va: BaselineOffset.NORMAL })).toBe(false)
    expect(hasVertAlign({ va: null })).toBe(false)
    expect(hasVertAlign({})).toBe(false)
    expect(hasVertAlign(undefined)).toBe(false)
  })
})

describe('vertAlignRunStyle', () => {
  it('copies the font keys the plain painter reads, plus va, and nothing else', () => {
    const run = vertAlignRunStyle({
      ff: 'Calibri',
      fs: 14,
      it: BooleanNumber.TRUE,
      bl: BooleanNumber.TRUE,
      ul: { s: BooleanNumber.TRUE },
      st: { s: BooleanNumber.TRUE },
      ol: { s: BooleanNumber.TRUE },
      cl: { rgb: '#ff0000' },
      va: BaselineOffset.SUBSCRIPT,
      // cell-level, not run-level: alignment, fill, border, number format
      ht: 2,
      bg: { rgb: '#ffff00' },
      bd: { t: { s: 1, cl: { rgb: '#000' } } },
      n: { pattern: '0.00' },
    })
    expect(run).toEqual({
      ff: 'Calibri',
      fs: 14,
      it: BooleanNumber.TRUE,
      bl: BooleanNumber.TRUE,
      ul: { s: BooleanNumber.TRUE },
      st: { s: BooleanNumber.TRUE },
      ol: { s: BooleanNumber.TRUE },
      cl: { rgb: '#ff0000' },
      va: BaselineOffset.SUBSCRIPT,
    })
  })

  it('leaves unset keys out so the document defaults apply, like the plain painter', () => {
    expect(vertAlignRunStyle({ va: BaselineOffset.SUPERSCRIPT, bl: BooleanNumber.FALSE })).toEqual({
      va: BaselineOffset.SUPERSCRIPT,
    })
  })
})

describe('vertAlignTwin', () => {
  it('keeps the cell on its own paint path when nothing needs the document path', () => {
    expect(vertAlignTwin(undefined, SUPER)).toBeUndefined()
    expect(vertAlignTwin({ v: 'x' }, undefined)).toBeUndefined()
    expect(vertAlignTwin({ v: 'x' }, {})).toBeUndefined()
    expect(vertAlignTwin({ v: 'x' }, { va: BaselineOffset.NORMAL })).toBeUndefined()
    expect(vertAlignTwin({ v: null }, SUPER)).toBeUndefined()
    expect(vertAlignTwin({}, SUPER)).toBeUndefined()
    expect(vertAlignTwin({ v: '' }, SUPER)).toBeUndefined()
    // real rich text already carries its own run styles
    const rich: ICellData = {
      p: { id: 'rich-cell', body: { dataStream: 'a\r\n' }, documentStyle: {} },
    }
    expect(vertAlignTwin(rich, SUPER)).toBeUndefined()
  })

  it('builds a rich-cell document whose single run carries the font and va', () => {
    const cell: ICellData = { v: 'H2O', t: CellValueType.STRING, s: 'style-id' }
    const twin = vertAlignTwin(cell, {
      va: BaselineOffset.SUBSCRIPT,
      bl: BooleanNumber.TRUE,
      fs: 12,
      bg: { rgb: '#eeeeee' },
    })
    expect(twin).toBeDefined()
    // the model fields survive untouched beside the twin document
    expect(twin?.v).toBe('H2O')
    expect(twin?.s).toBe('style-id')
    expect(twin?.p).toEqual({
      id: 'rich-cell',
      body: {
        dataStream: 'H2O\r\n',
        textRuns: [
          { st: 0, ed: 3, ts: { bl: BooleanNumber.TRUE, fs: 12, va: BaselineOffset.SUBSCRIPT } },
        ],
        paragraphs: [{ startIndex: 3 }],
        sectionBreaks: [{ startIndex: 4 }],
      },
      documentStyle: { renderConfig: { cellValueType: CellValueType.STRING } },
    })
    // the input cell is not mutated
    expect(cell.p).toBeUndefined()
  })

  it('tells the document layout the value type so numbers and booleans keep their default alignment', () => {
    expect(vertAlignTwin({ v: 42 }, SUPER)?.p?.documentStyle?.renderConfig?.cellValueType).toBe(
      CellValueType.NUMBER,
    )
    // a number-formatted display string still reports the cell's number type
    expect(
      vertAlignTwin({ v: '$42.00', t: CellValueType.NUMBER }, SUPER)?.p?.documentStyle?.renderConfig
        ?.cellValueType,
    ).toBe(CellValueType.NUMBER)
    const bool = vertAlignTwin({ v: 1, t: CellValueType.BOOLEAN }, SUPER)
    expect(bool?.p?.body?.dataStream).toBe('TRUE\r\n')
    expect(bool?.p?.documentStyle?.renderConfig?.cellValueType).toBe(CellValueType.BOOLEAN)
    expect(vertAlignTwin({ v: 'text' }, SUPER)?.p?.documentStyle?.renderConfig?.cellValueType).toBe(
      CellValueType.STRING,
    )
  })

  it('turns line breaks into paragraph breaks instead of section breaks', () => {
    const twin = vertAlignTwin({ v: 'ab\ncd\r\ne' }, SUPER)
    expect(twin?.p?.body).toEqual({
      dataStream: 'ab\rcd\re\r\n',
      textRuns: [{ st: 0, ed: 7, ts: { va: BaselineOffset.SUPERSCRIPT } }],
      paragraphs: [{ startIndex: 2 }, { startIndex: 5 }, { startIndex: 7 }],
      sectionBreaks: [{ startIndex: 8 }],
    })
  })
})

describe('installVertAlignRender', () => {
  it('hands the font cache the twin for a va cell and the cell itself otherwise', () => {
    const proto = SpreadsheetSkeleton.prototype as unknown as Record<string, unknown>
    const inner = vi.fn()
    proto._setFontStylesCache = inner
    installVertAlignRender()
    const patched = proto._setFontStylesCache as (
      this: unknown,
      row: number,
      col: number,
      cell: ICellData | undefined,
      style: unknown,
    ) => void
    expect(patched).not.toBe(inner)

    const plain: ICellData = { v: 'x' }
    patched.call({ _renderRawFormula: false }, 0, 0, plain, {})
    expect(inner).toHaveBeenLastCalledWith(0, 0, plain, {})

    const raised: ICellData = { v: 'x' }
    patched.call({ _renderRawFormula: false }, 1, 2, raised, SUPER)
    const [, , handed] = inner.mock.calls.at(-1) as [number, number, ICellData]
    expect(handed).not.toBe(raised)
    expect(handed.p?.body?.textRuns?.[0]?.ts?.va).toBe(BaselineOffset.SUPERSCRIPT)

    // Show Formulas keeps every cell on the plain path
    patched.call({ _renderRawFormula: true }, 1, 2, raised, SUPER)
    expect(inner).toHaveBeenLastCalledWith(1, 2, raised, SUPER)

    // installing twice does not wrap twice
    installVertAlignRender()
    expect(proto._setFontStylesCache).toBe(patched)
  })
})
