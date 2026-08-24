import { CellValueType, HorizontalAlign } from '@univerjs/core'
import { Documents, SpreadsheetSkeleton, Text } from '@univerjs/engine-render'
import { beforeAll, describe, expect, it, vi } from 'vitest'

import { installRtlTextDirectionFix, resolveBidiDirection } from '../src/renderer/rtl-text-fix'

describe('resolveBidiDirection', () => {
  it('is rtl when the first strong character is Arabic, even after digits', () => {
    expect(resolveBidiDirection('1446هـ')).toBe('rtl')
  })

  it('is rtl for Hebrew text wrapped in neutral punctuation', () => {
    expect(resolveBidiDirection('(שלום)')).toBe('rtl')
  })

  it('is ltr when the first strong character is Latin', () => {
    expect(resolveBidiDirection('abc عرب')).toBe('ltr')
  })

  it('is ltr for weak-only content (digits, Arabic-Indic digits, empty)', () => {
    expect(resolveBidiDirection('123.45')).toBe('ltr')
    expect(resolveBidiDirection('١٢٣')).toBe('ltr')
    expect(resolveBidiDirection('')).toBe('ltr')
  })

  it('honors explicit directional marks before any letter', () => {
    expect(resolveBidiDirection('\u200F123')).toBe('rtl')
    expect(resolveBidiDirection('\u200Eعرب')).toBe('ltr')
  })
})

describe('installRtlTextDirectionFix', () => {
  const makeCtx = () => {
    const events: string[] = []
    return {
      events,
      ctx: {
        direction: 'inherit' as CanvasDirection,
        textAlign: 'start' as CanvasTextAlign,
        save: () => events.push('save'),
        restore: () => events.push('restore'),
      },
    }
  }

  it('draws rtl-first text with rtl direction and a left anchor, then restores', () => {
    const textClass = Text as unknown as {
      drawWith(ctx: unknown, props: unknown, skeleton?: unknown): void
    }
    const seen: { direction: string; textAlign: string }[] = []
    const inner = vi.fn((ctx: { direction: string; textAlign: string }) => {
      seen.push({ direction: ctx.direction, textAlign: ctx.textAlign })
    })
    textClass.drawWith = inner as never
    installRtlTextDirectionFix()

    const rtl = makeCtx()
    textClass.drawWith(rtl.ctx, { text: '1446هـ' })
    expect(seen).toEqual([{ direction: 'rtl', textAlign: 'left' }])
    expect(rtl.events).toEqual(['save', 'restore'])

    const ltr = makeCtx()
    textClass.drawWith(ltr.ctx, { text: 'plain' })
    expect(seen[1]).toEqual({ direction: 'inherit', textAlign: 'start' })
    expect(ltr.events).toEqual([])
  })
})

describe('General alignment for RTL-first text (context reading order)', () => {
  beforeAll(() => installRtlTextDirectionFix())
  const skeletonProto = SpreadsheetSkeleton.prototype as unknown as {
    _calculateOverflowCell(row: number, column: number, config: unknown): boolean
  }
  // Minimal host: the original implementation only touches _cellData here
  // because an undefined wrapStrategy short-circuits both overflow branches.
  const host = { _cellData: { getValue: () => ({}) } }
  const docSkeleton = (text: string) => ({
    getViewModel: () => ({ getDataModel: () => ({ getBody: () => ({ dataStream: text }) }) }),
  })

  it('resolves the font-cache entry to RIGHT for Arabic-first General text', () => {
    const config = { horizontalAlign: HorizontalAlign.UNSPECIFIED, cellData: { v: 'مرحبا' } }
    skeletonProto._calculateOverflowCell.call(host, 0, 0, config)
    expect(config.horizontalAlign).toBe(HorizontalAlign.RIGHT)
  })

  it('reads rich-text cells through their document skeleton', () => {
    const config = {
      horizontalAlign: HorizontalAlign.UNSPECIFIED,
      cellData: { v: null, p: {} },
      documentSkeleton: docSkeleton('نص عربي'),
    }
    skeletonProto._calculateOverflowCell.call(host, 0, 0, config)
    expect(config.horizontalAlign).toBe(HorizontalAlign.RIGHT)
  })

  it('leaves Latin text, numbers, explicit alignment and rotated cells alone', () => {
    for (const config of [
      { horizontalAlign: HorizontalAlign.UNSPECIFIED, cellData: { v: 'latin' } },
      { horizontalAlign: HorizontalAlign.UNSPECIFIED, cellData: { v: 42 } },
      {
        horizontalAlign: HorizontalAlign.UNSPECIFIED,
        cellData: { v: 42, t: CellValueType.NUMBER },
      },
      { horizontalAlign: HorizontalAlign.LEFT, cellData: { v: 'مرحبا' } },
      { horizontalAlign: HorizontalAlign.UNSPECIFIED, cellData: { v: 'مرحبا' }, vertexAngle: 45 },
    ]) {
      const before = config.horizontalAlign
      skeletonProto._calculateOverflowCell.call(host, 0, 0, config)
      expect(config.horizontalAlign).toBe(before)
    }
  })

  it('right-aligns the Documents page offset for Arabic-first General cells', () => {
    const proto = Documents.prototype as unknown as {
      _horizontalHandler(
        pageWidth: number,
        pagePaddingLeft: number,
        pagePaddingRight: number,
        horizontalAlign: number,
        vertexAngleDeg?: number,
        centerAngleDeg?: number,
        cellValueType?: number,
      ): number
    }
    const arabicHost = { width: 200, getSkeleton: () => docSkeleton('مرحبا بالعالم') }
    const latinHost = { width: 200, getSkeleton: () => docSkeleton('hello') }
    const offsetArabic = proto._horizontalHandler.call(
      arabicHost,
      80,
      2,
      3,
      HorizontalAlign.UNSPECIFIED,
    )
    const offsetLatin = proto._horizontalHandler.call(
      latinHost,
      80,
      2,
      3,
      HorizontalAlign.UNSPECIFIED,
    )
    expect(offsetArabic).toBe(200 - 80 - 3)
    expect(offsetLatin).toBe(2)
  })
})
