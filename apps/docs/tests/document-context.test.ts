/**
 * The body East Asian face three-way rule (Word probes 2026-11-04 e1-e20 and
 * truth 429): an East Asian-named slot is the face, an empty theme slot takes
 * the themeFontLang face, a Latin-named slot falls back to SimSun / Batang.
 */
import { describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import {
  documentTextContext,
  eaFaceForCjkFactor,
  eaFaceForGridStrut,
  eaFaceForKrFactor,
} from '../src/renderer/text-layout/document-context'

type Display = { font?: string; fontAscii?: string; eastAsiaFont?: string; eaSlotEmpty?: boolean }
function parsedWith(
  normal: Display | undefined,
  docDefaults: Record<string, unknown>,
  themeFonts?: Record<string, unknown>,
  extra: Record<string, unknown> = {},
): ParsedDocFull {
  const styles = new Map()
  if (normal)
    styles.set('Normal', { styleId: 'Normal', type: 'paragraph', isDefault: true, display: normal })
  return { styles, docDefaults, blocks: [], themeFonts, ...extra } as unknown as ParsedDocFull
}
const face = (p: ParsedDocFull) => documentTextContext(p).bodyEaFace

describe('bodyEastAsianFace', () => {
  it('an East Asian face in the Normal slot is the face (e: declared)', () => {
    const f = face(
      parsedWith(
        { font: '\u5b8b\u4f53', fontAscii: 'Calibri' },
        { asciiFont: 'Calibri', eastAsiaFont: 'Calibri' },
      ),
    )
    expect(f).toMatchObject({
      face: '\u5b8b\u4f53',
      hangulFace: 'Batang',
      source: 'normal',
      declared: true,
    })
  })

  it('one East Asian name in every Normal slot is an East Asian choice', () => {
    expect(
      face(parsedWith({ font: 'Malgun Gothic', fontAscii: 'Malgun Gothic' }, {})),
    ).toMatchObject({
      face: 'Malgun Gothic',
      hangulFace: 'Malgun Gothic',
      source: 'normal',
    })
    expect(face(parsedWith({ font: 'Meiryo', fontAscii: 'Meiryo' }, {})).hangulFace).toBe(
      'Malgun Gothic',
    )
  })

  it('a Latin face in the eastAsia slot falls back to SimSun / Batang whatever themeFontLang says (e2, e5, e19, e20)', () => {
    const tnrAll = parsedWith(
      { font: 'Times New Roman', fontAscii: 'Times New Roman', eastAsiaFont: 'Times New Roman' },
      { asciiFont: 'Aptos', eastAsiaFont: 'Aptos', eaThemeSlot: true },
      { minor: 'Aptos', eaLang: 'ko-KR', minorScripts: { Hang: 'Malgun Gothic' } },
    )
    expect(face(tnrAll)).toEqual({
      face: 'SimSun',
      hangulFace: 'Batang',
      source: 'latinSlot',
      declared: false,
    })
    const ddCalibri = parsedWith(undefined, { asciiFont: 'Calibri', eastAsiaFont: 'Calibri' })
    expect(face(ddCalibri).source).toBe('latinSlot')
    expect(eaFaceForCjkFactor(documentTextContext(ddCalibri))).toBe('SimSun')
    expect(eaFaceForGridStrut(documentTextContext(ddCalibri))).toBeUndefined()
  })

  it('a Latin-only Normal inherits the docDefaults slot', () => {
    const f = face(
      parsedWith(
        { font: 'Calibri', fontAscii: 'Calibri' },
        { asciiFont: 'Calibri', eastAsiaFont: 'Malgun Gothic' },
      ),
    )
    expect(f).toMatchObject({ face: 'Malgun Gothic', source: 'docDefaults', declared: true })
    expect(
      eaFaceForKrFactor(
        documentTextContext(
          parsedWith(
            { font: 'Calibri', fontAscii: 'Calibri' },
            { asciiFont: 'Calibri', eastAsiaFont: 'Malgun Gothic' },
          ),
        ),
      ),
    ).toBe('Malgun Gothic')
  })

  it('an empty theme slot takes the themeFontLang face, SimSun without a language (e1, e6, e16-e18)', () => {
    const empty = (theme: Record<string, unknown> | undefined) =>
      face(
        parsedWith(
          { font: 'Calibri', fontAscii: 'Calibri' },
          { asciiFont: 'Calibri', eaThemeSlot: true },
          theme,
        ),
      )
    expect(empty({ minor: 'Calibri' })).toMatchObject({
      face: 'SimSun',
      hangulFace: 'Batang',
      source: 'emptySlot',
    })
    expect(
      empty({ minor: 'Calibri', eaLang: 'ko-KR', minorScripts: { Hang: 'Malgun Gothic' } }),
    ).toMatchObject({
      face: 'Malgun Gothic',
      hangulFace: 'Malgun Gothic',
    })
    expect(empty({ minor: 'Calibri', eaLang: 'ja-JP' })).toMatchObject({
      face: 'Yu Mincho',
      hangulFace: 'Batang',
    })
    expect(
      empty({
        minor: 'Calibri',
        eaLang: 'ja-JP',
        minorScripts: { Jpan: '\uff2d\uff33 \u660e\u671d' },
      }),
    ).toMatchObject({
      face: '\uff2d\uff33 \u660e\u671d',
      hangulFace: 'Malgun Gothic',
    })
    expect(
      empty({ minor: 'Calibri', eaLang: 'zh-CN', minorScripts: { Hans: '\u5b8b\u4f53' } }).face,
    ).toBe('\u5b8b\u4f53')
  })

  it('the balanced-space switch needs a declared East Asian face or useFELayout', () => {
    const declared = parsedWith(undefined, { eastAsiaFont: '\u5b8b\u4f53' }, undefined, {
      balanceDbcsSpacing: true,
    })
    expect(documentTextContext(declared).balancedSpaces).toBe(true)
    const latin = parsedWith(undefined, { eastAsiaFont: 'Calibri' }, undefined, {
      balanceDbcsSpacing: true,
    })
    expect(documentTextContext(latin).balancedSpaces).toBe(false)
    const fe = parsedWith(undefined, { eastAsiaFont: 'Calibri' }, undefined, {
      balanceDbcsSpacing: true,
      useFELayout: true,
    })
    expect(documentTextContext(fe).balancedSpaces).toBe(true)
  })
})
