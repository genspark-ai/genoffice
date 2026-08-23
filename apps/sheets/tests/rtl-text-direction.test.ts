import { Text } from '@univerjs/engine-render'
import { describe, expect, it, vi } from 'vitest'

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
