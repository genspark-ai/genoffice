/**
 * Word turns autoSpaceDE/DN off document-wide when the declared default East
 * Asian face does not resolve to an installed font (Word probe 2026-09-01,
 * SAS prod 049/047/044: docDefaults eastAsia "Noto Sans CJK KR/JP" renders
 * without any CJK-Latin gaps while an installed face keeps the ~1/4em gap).
 * jsdom has no canvas, so every non-bundled face measures as missing here.
 */
import { describe, expect, it } from 'vitest'
import type { DocDefaults, ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'
import { docAutospaceOff, docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedDoc(opts: { docDefaults?: DocDefaults; styles?: StyleInfo[] }): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  for (const s of opts.styles ?? []) styles.set(s.styleId, s)
  return {
    styles,
    docDefaults: opts.docDefaults ?? {},
    blocks: [],
  } as unknown as ParsedDocFull
}

const KILL_RULE =
  '.page-wrap, .doc-page, .pv-page { text-autospace:no-autospace; --doc-autospace-pad:0 }'

describe('docAutospaceOff', () => {
  it('suppresses when docDefaults declares an unavailable EA face', () => {
    const doc = parsedDoc({ docDefaults: { eastAsiaFont: 'Noto Sans CJK KR' } })
    expect(docAutospaceOff(doc)).toBe(true)
    expect(docStyleCss(doc)).toContain(KILL_RULE)
  })

  it('treats bundled subset faces as missing (Word never sees them)', () => {
    expect(docAutospaceOff(parsedDoc({ docDefaults: { eastAsiaFont: 'Noto Sans CJK SC' } }))).toBe(
      true,
    )
  })

  it('keeps the pads when the EA face is a w:lang backfill, not a document choice', () => {
    const doc = parsedDoc({ docDefaults: { eastAsiaFont: 'Noto Sans CJK KR', eaFromLang: true } })
    expect(docAutospaceOff(doc)).toBe(false)
    expect(docStyleCss(doc)).not.toContain('no-autospace')
  })

  it('keeps the pads for an empty EA theme slot', () => {
    const doc = parsedDoc({ docDefaults: { eastAsiaFont: 'Noto Sans CJK KR', eaSlotEmpty: true } })
    expect(docAutospaceOff(doc)).toBe(false)
  })

  it('keeps the pads for per-script Google names Word resolves as cloud fonts', () => {
    // SAS prod_098: Word keeps the gaps under docDefaults eastAsia "Noto Sans KR"
    expect(docAutospaceOff(parsedDoc({ docDefaults: { eastAsiaFont: 'Noto Sans KR' } }))).toBe(
      false,
    )
    expect(docAutospaceOff(parsedDoc({ docDefaults: { eastAsiaFont: 'Noto Serif JP' } }))).toBe(
      false,
    )
  })

  it('suppresses for the Source Han superfamily names', () => {
    expect(
      docAutospaceOff(parsedDoc({ docDefaults: { eastAsiaFont: 'Source Han Sans KR' } })),
    ).toBe(true)
  })

  it('trusts Word to resolve its private DFonts names the canvas cannot see', () => {
    expect(docAutospaceOff(parsedDoc({ docDefaults: { eastAsiaFont: '宋体' } }))).toBe(false)
    expect(docAutospaceOff(parsedDoc({ docDefaults: { eastAsiaFont: 'ＭＳ 明朝' } }))).toBe(false)
    expect(docAutospaceOff(parsedDoc({ docDefaults: { eastAsiaFont: 'Batang' } }))).toBe(false)
  })

  it('keeps the pads without any EA default', () => {
    expect(docAutospaceOff(parsedDoc({}))).toBe(false)
  })

  it('the default paragraph style EA face also gates', () => {
    const doc = parsedDoc({
      styles: [
        {
          styleId: 'Normal',
          name: 'Normal',
          type: 'paragraph',
          isDefault: true,
          display: { font: 'Noto Sans CJK JP', fontAscii: 'Calibri' },
        } as StyleInfo,
      ],
    })
    expect(docAutospaceOff(doc)).toBe(true)
  })
})
