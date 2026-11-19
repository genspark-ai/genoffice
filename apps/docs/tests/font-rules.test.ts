/**
 * The font-rules entry point returns the same answers as the functions it wraps
 * and classifies the route Word took (DESIGN-fonts §3.2). Shape and routing
 * only; the numbers themselves are pinned by line-metrics.test.ts and
 * packages/font-metrics/tests/resolve.test.ts.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { STAND_IN_ALIAS_CHAINS } from '../src/renderer/stand-in-chains.generated'
import {
  shadowedStandInRows,
  STAND_IN_CHAINS,
} from '../src/renderer/text-layout/font-rules/css-chains'
import {
  arabicAltNameFace,
  arabicCodePageFace,
  cssFontFamily,
  krCharsetRoute,
  lineHeightFactor,
  resolveFace,
  setDocFontTable,
  setEmbeddedLineMetrics,
  tableLineHeightFactor,
} from '../src/renderer/text-layout/font-rules'

afterEach(() => {
  setDocFontTable([])
  setEmbeddedLineMetrics([])
})

describe('resolveFace', () => {
  it('wraps the legacy answers for a face Word ships', () => {
    const d = resolveFace('Calibri')
    expect(d.declared).toBe('calibri')
    expect(d.face).toBe('Calibri')
    expect(d.source).toBe('shipped')
    expect(d.substituted).toBe(false)
    expect(d.lineFactor).toBe(lineHeightFactor('Calibri'))
    expect(d.chain).toBe(cssFontFamily('Calibri'))
    expect(d.standIn).toBe('Carlito GO')
    expect(d.kr).toBeNull()
    expect(d.missingLatin).toBeNull()
  })

  it('classifies the PT name-prefix rule (Word probe 2026-11-08)', () => {
    setDocFontTable([{ name: 'PT Astra Serif', panose: '020A0603040505020204' }])
    const d = resolveFace('PT Astra Serif')
    expect(d.face).toBe('Arial')
    expect(d.source).toBe('prefix')
    expect(d.substituted).toBe(true)
    expect(d.path).toEqual(['pt astra serif', 'arial'])
  })

  it('paces a style-suffixed name by its base family (Word probe 2026-10-22)', () => {
    setDocFontTable([
      { name: 'Bookman Old Style Bold', panose: '02050604050505020204', family: 'roman' },
      { name: 'Arial Bold', panose: '020B0604020202090204', family: 'swiss' },
    ])
    const bookman = resolveFace('Bookman Old Style Bold')
    expect(bookman.face).toBe('Bookman Old Style')
    expect(bookman.source).toBe('shipped')
    expect(bookman.lineFactor).toBe(lineHeightFactor('Bookman Old Style'))
    expect(bookman.lineFactor).toBe(1.1738)
    expect(lineHeightFactor('Arial Bold')).toBe(lineHeightFactor('Arial'))
    expect(lineHeightFactor('Monotype Corsiva Italic')).toBe(lineHeightFactor('Monotype Corsiva'))
  })

  it('classifies a fontTable altName route (Word probe 2026-10-22)', () => {
    setDocFontTable([{ name: 'Zebra Alt Probe', altName: 'Cambria' }])
    const d = resolveFace('Zebra Alt Probe')
    expect(d.face).toBe('Cambria')
    expect(d.source).toBe('altName')
  })

  it('classifies the charset-81 route and the Latin fallback', () => {
    setDocFontTable([
      { name: 'Zebra Hangul Probe', charset: '81', family: 'swiss' },
      { name: 'Zebra Plain Probe', family: 'swiss' },
    ])
    const kr = resolveFace('Zebra Hangul Probe')
    expect(kr.source).toBe('charset')
    expect(kr.kr).toBe(krCharsetRoute('Zebra Hangul Probe'))
    expect(kr.kr).toBe('malgun')
    const latin = resolveFace('Zebra Plain Probe')
    expect(latin.source).toBe('fallback')
    expect(latin.face).toBe('Calibri')
    expect(latin.missingLatin).toBe('calibri')
  })

  it('reads the Arabic code-page step from the resolver, chain and face agreeing (corpus 2026-10-27)', () => {
    setDocFontTable([
      { name: 'Zebra Kufi', charset: '00', csb0: '00000041' },
      { name: 'Zebra Kufi Alt', charset: '00', csb0: '00000041', altName: 'Verdana' },
    ])
    const sig = resolveFace('Zebra Kufi')
    expect(sig.face).toBe('Arial')
    expect(sig.path).toEqual(['zebra kufi', 'arial'])
    expect(arabicCodePageFace('Zebra Kufi')).toBe('Arial')
    expect(sig.chain.startsWith("'Arial','Arial Arabic GO'")).toBe(true)
    expect(arabicCodePageFace('Zebra Kufi Alt')).toBeNull()
    expect(arabicAltNameFace('Zebra Kufi Alt')).toBe('Verdana')
    expect(resolveFace('Zebra Kufi Alt').source).toBe('altName')
  })

  it('classifies a hangul-lettered name as a script route', () => {
    const d = resolveFace('한글테스트')
    expect(d.source).toBe('script')
    expect(d.face).toBe('Batang')
  })
})

describe('lineHeightFactor = pins U resolver (line F wave 3)', () => {
  it('a face Word lays out for real paces at its file metrics, not the old table digits', () => {
    expect(lineHeightFactor('Arial')).toBe(1.1499)
    expect(lineHeightFactor('Calibri')).toBe(1.2207)
    expect(lineHeightFactor('Courier New')).toBe(1.1328)
    expect(lineHeightFactor('Tahoma')).toBe(1.207)
    expect(lineHeightFactor('Verdana Pro')).toBe(1.2153)
    expect(resolveFace('Arial').factorSource).toBe('resolver')
  })

  it('a behavioural substitution paces as its target face', () => {
    expect(lineHeightFactor('Liberation Sans')).toBe(lineHeightFactor('Arial'))
    expect(lineHeightFactor('Carlito')).toBe(lineHeightFactor('Calibri'))
    expect(lineHeightFactor('Liberation Serif')).toBe(lineHeightFactor('Times New Roman'))
  })

  it('pins keep the probed value where the file metric is not what Word draws', () => {
    expect(lineHeightFactor('Times')).toBe(1.15)
    expect(lineHeightFactor('Courier')).toBe(1.1333)
    expect(lineHeightFactor('Helvetica')).toBe(1.2)
    expect(lineHeightFactor('Inter')).toBe(1.21)
    expect(resolveFace('Times').factorSource).toBe('pin')
  })

  it('the Calibri / Cambria fallback and the shadow table column carry the resolver digits', () => {
    expect(lineHeightFactor('Zebra Probe Noentry')).toBe(1.1724)
    setDocFontTable([{ name: 'Zebra Probe Sans', panose: '00000700' }])
    expect(lineHeightFactor('Zebra Probe Sans')).toBe(1.2207)
    expect(tableLineHeightFactor('Arial')).toBe(1.1724)
  })
})

describe('factorSource', () => {
  it('names the layer that answered: resolver, pin, embedded, table or fallback', () => {
    expect(resolveFace('Calibri').factorSource).toBe('resolver')
    expect(resolveFace('Helvetica').factorSource).toBe('pin')
    expect(resolveFace('Segoe UI Emoji').factorSource).toBe('pin')
    expect(resolveFace('Zebra Fixture Sans').factorSource).toBe('fallback')
    setDocFontTable([{ name: 'Zebra Fixture Sans', altName: 'Arial' }])
    expect(resolveFace('Zebra Fixture Sans').factorSource).toBe('resolver')
    setEmbeddedLineMetrics([{ family: 'Zebra Embedded', ascent: 0.9, descent: 0.25, lineGap: 0 }])
    expect(resolveFace('Zebra Embedded')).toMatchObject({
      lineFactor: 1.15,
      factorSource: 'embedded',
    })
    // a missing JP vendor name falls to Cambria in the resolver; the table's SimSun-class route answers
    expect(resolveFace('Kozuka Gothic Pro')).toMatchObject({
      lineFactor: lineHeightFactor('SimSun'),
      factorSource: 'table',
    })
  })
})

describe('STAND_IN_CHAINS', () => {
  it('keys by the canonical family: a localized spelling shares its English twin row', () => {
    const english = resolveFace('MS Gothic')
    const twin = resolveFace('\uff2d\uff33 \u30b4\u30b7\u30c3\u30af')
    expect(twin.face).toBe('\uff2d\uff33 \u30b4\u30b7\u30c3\u30af')
    expect(twin.standIn).toBe('MS Gothic GO')
    expect(twin.chain.replace(/^'[^']+',/, '')).toBe(english.chain.replace(/^'[^']+',/, ''))
    // the KO chain leads with Malgun Gothic itself, so the English head is deduplicated away
    expect(resolveFace('\ub9d1\uc740 \uace0\ub515').chain).toBe(
      `'\ub9d1\uc740 \uace0\ub515',${resolveFace('Malgun Gothic').chain}`,
    )
    for (const key of Object.keys(STAND_IN_CHAINS)) expect(key).toMatch(/^[\x20-\x7e]+$/)
  })

  it('every alias-led row is generated from its manifest or a hand row the manifest cannot express; none shadows', () => {
    const dir = join(__dirname, '../../../packages/font-metrics/measurements/families')
    const generated = new Map<string, string>()
    const hand = new Map<string, string>()
    for (const f of readdirSync(dir)) {
      const m = JSON.parse(readFileSync(join(dir, f), 'utf8')) as {
        family: string
        standsInFor?: string
        standInChain?: unknown
      }
      if (!m.standsInFor) continue
      ;(m.standInChain ? generated : hand).set(m.standsInFor.toLowerCase(), m.family)
    }
    for (const [key, lead] of generated) expect(STAND_IN_ALIAS_CHAINS[key]?.[0][0]).toBe(lead)
    const handLeads = new Set(hand.values())
    const rows = new Map<string, string>()
    for (const [key, [families]] of Object.entries(STAND_IN_CHAINS))
      if (handLeads.has(families[0])) rows.set(key, families[0])
    expect(Object.fromEntries(rows)).toEqual(Object.fromEntries(hand))
    expect(shadowedStandInRows()).toEqual([])
  })
})
