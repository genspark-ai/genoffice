import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

import {
  faceKey,
  lineFactorOf,
  panoseSerif,
  resolveFace,
  storeContext,
  SUBSTITUTION_RULES,
  SUBSTITUTION_STEPS,
  type FaceHints,
  type WordFacesData,
} from '../src/resolve'

const data = JSON.parse(
  readFileSync(new URL('../word-faces.json', import.meta.url), 'utf8'),
) as WordFacesData

function ctxWith(hints: Record<string, FaceHints> = {}, installed?: Record<string, number>) {
  return storeContext(data, {
    hints: (key) => hints[key],
    installed: installed
      ? (key) =>
          key in installed
            ? {
                family: key,
                source: 'installed',
                hheaTotal: installed[key],
                ascent: installed[key] - 0.25,
                descent: 0.25,
                xAvg: null,
                eastAsian: false,
              }
            : null
      : undefined,
  })
}

describe('word-faces.json', () => {
  it('is generated, not hand-edited, and carries both stores', () => {
    expect(Object.keys(data.faces).length).toBeGreaterThan(400)
    expect(Object.keys(data.system).length).toBeGreaterThan(100)
    expect(data.faces.calibri).toMatchObject({ source: 'dfonts', eastAsian: false })
    expect(data.faces['malgun gothic']).toMatchObject({ source: 'dfonts', eastAsian: true })
    expect(data.faces.montserrat.source).toBe('cloud')
  })

  it('carries the box and width fields the renderer tables derive from', () => {
    for (const face of [data.faces.calibri, data.faces.cambria, data.system['times new roman']]) {
      expect(face.descent).toBeGreaterThan(0)
      expect(face.ascent).toBeGreaterThan(face.descent)
      expect(face.xAvg).toBeGreaterThan(0)
    }
    expect(data.faces.cambria).toMatchObject({ descent: 0.2222, xAvg: 0.6152 })
  })
})

describe('measurements/substitutions.json', () => {
  it('every row has a unique id, a compiling match, a target the store knows and a dated source', () => {
    const ids = new Set<string>()
    for (const r of SUBSTITUTION_RULES) {
      expect(ids.has(r.id), r.id).toBe(false)
      ids.add(r.id)
      if (typeof r.match === 'string') expect(() => new RegExp(r.match as string)).not.toThrow()
      else expect(['HG_RE', 'ARABIC_RE', 'TAMIL_RE', 'THAI_RE']).toContain(r.match.ref)
      expect(data.faces[faceKey(r.target)] ?? data.system[faceKey(r.target)], r.target).toBeTruthy()
      expect(['probe', 'corpus']).toContain(r.source.kind)
      expect(r.source.date).toMatch(/^\d{4}-\d{2}-\d{2}$/)
      expect(r.source.probe.length).toBeGreaterThan(10)
    }
  })

  it("Word's name-prefix table: Helvetica* -> Arial, Segoe* -> Segoe UI (probe 2026-11-08)", () => {
    expect(resolveFace('Helvetica Zebra', ctxWith())).toMatchObject({
      family: 'Arial',
      exception: true,
    })
    expect(resolveFace('HelveticaZebra', ctxWith())?.family).toBe('Arial')
    expect(resolveFace('Segoe Zebra', ctxWith())).toMatchObject({
      family: 'Segoe UI',
      exception: true,
    })
    // the installed copy still lays out itself
    expect(resolveFace('Helvetica', ctxWith({}, { helvetica: 1.2 }))?.substituted).toBe(false)
    expect(resolveFace('Segoe UI', ctxWith())?.substituted).toBe(false)
  })
})

describe('measurements/substitution-steps.json', () => {
  it('every step has a unique id, a rule label, compiling patterns, a dated source; the fallback closes the table', () => {
    const ids = new Set<string>()
    const refs = [
      'HG_RE',
      'ARABIC_RE',
      'TAMIL_RE',
      'THAI_RE',
      'HAN_RE',
      'HANGUL_RE',
      'HAN_KR_RE',
      'ARABIC_CLOUD_RE',
    ]
    const pattern = (m: string | { ref: string }) =>
      typeof m === 'string'
        ? expect(() => new RegExp(m)).not.toThrow()
        : expect(refs).toContain(m.ref)
    for (const s of SUBSTITUTION_STEPS) {
      expect(ids.has(s.id), s.id).toBe(false)
      ids.add(s.id)
      expect(s.rule.length).toBeGreaterThan(0)
      for (const k of ['name', 'nameNot', 'altNameNot'] as const)
        if (s.when?.[k]) pattern(s.when[k]!)
      if (typeof s.target === 'string')
        expect(
          data.faces[faceKey(s.target)] ?? data.system[faceKey(s.target)],
          s.target,
        ).toBeTruthy()
      expect(['probe', 'corpus']).toContain(s.source.kind)
      expect(s.source.date).toMatch(/^\d{4}-\d{2}-\d{2}$/)
    }
    const last = SUBSTITUTION_STEPS[SUBSTITUTION_STEPS.length - 1]
    expect(last.id).toBe('fallback')
    expect(last.when).toBeUndefined()
  })

  it('keeps the order Word applies: charset before altName before the name table before script', () => {
    const order = SUBSTITUTION_STEPS.map((s) => s.id)
    const before = (a: string, b: string) => expect(order.indexOf(a)).toBeLessThan(order.indexOf(b))
    before('charset81', 'alt-name')
    before('charset80', 'alt-name')
    before('arabic-alt-name', 'name-table')
    before('arabic-code-page', 'name-table')
    before('hg-alt-name', 'name-table')
    before('name-table', 'alt-name')
    before('alt-name', 'latin-code-page')
    before('latin-code-page', 'han')
    before('han', 'hangul')
    // a charset-81 entry with an altName still takes the charset step
    expect(
      resolveFace('Fixture KR', ctxWith({ 'fixture kr': { charset: '81', altName: 'Arial' } }))
        ?.rules,
    ).toEqual(['charset81'])
  })
})

describe('lineFactorOf', () => {
  it('uses the hhea total for Latin faces and (ascent + descent) x 1.3 for East Asian ones', () => {
    expect(lineFactorOf(data.faces.calibri)).toBe(1.2207)
    expect(lineFactorOf(data.faces.cambria)).toBe(1.1724)
    expect(lineFactorOf(data.faces.symbol)).toBe(1.2251)
    expect(lineFactorOf(data.faces['malgun gothic'])).toBe(1.7291)
    expect(lineFactorOf(data.faces.simsun)).toBe(1.3)
    expect(lineFactorOf(data.faces['ms gothic'])).toBe(1.3)
  })
})

describe('resolveFace', () => {
  it('lays Word-private and cloud faces out for real', () => {
    expect(resolveFace('Poppins', ctxWith())).toMatchObject({
      family: 'Poppins',
      source: 'cloud',
      substituted: false,
      lineFactor: 1.5,
    })
    expect(resolveFace('Montserrat SemiBold', ctxWith())?.lineFactor).toBe(1.219)
    expect(resolveFace('Comic Sans MS', ctxWith())?.lineFactor).toBe(1.3936)
    expect(resolveFace('Mangal', ctxWith())?.lineFactor).toBe(1.6797)
  })

  it('prefers the installed copy except for the symbol faces', () => {
    // Times New Roman: Word lays out with the system copy (lineGap 87), not DFonts' (lineGap 0)
    const tnr = resolveFace('Times New Roman', ctxWith())
    expect(tnr?.source).toBe('system')
    expect(tnr?.lineFactor).toBe(1.1499)
    const symbol = resolveFace('Symbol', ctxWith({}, { symbol: 1.0 }))
    expect(symbol?.source).toBe('dfonts')
    expect(symbol?.lineFactor).toBe(1.2251)
    expect(resolveFace('Foo Sans', ctxWith({}, { 'foo sans': 1.3 }))?.lineFactor).toBe(1.3)
  })

  it('falls to Calibri with a fontTable entry, Cambria with w:family roman or no entry', () => {
    expect(resolveFace('Mont', ctxWith({ mont: { panose: '00000700' } }))).toMatchObject({
      family: 'Calibri',
      substituted: true,
      lineFactor: 1.2207,
      path: ['mont', 'calibri'],
    })
    // probe 2026-11-08: a serif PANOSE alone (no w:family) is Calibri, PANOSE never decides
    expect(resolveFace('Mont', ctxWith({ mont: { panose: '02030504' } }))?.family).toBe('Calibri')
    expect(
      resolveFace('Mont', ctxWith({ mont: { panose: '020A0603', csb0: '00000007' } }))?.family,
    ).toBe('Calibri')
    expect(resolveFace('Mont', ctxWith())?.family).toBe('Cambria')
    // probes 2026-10-22 / 2026-11-08: w:family roman is Cambria whatever the PANOSE, swiss / modern / auto Calibri
    expect(resolveFace('Mont', ctxWith({ mont: { family: 'roman' } }))?.family).toBe('Cambria')
    expect(
      resolveFace('Mont', ctxWith({ mont: { family: 'roman', panose: '020B0604' } }))?.family,
    ).toBe('Cambria')
    expect(resolveFace('Mont', ctxWith({ mont: { family: 'modern' } }))?.family).toBe('Calibri')
    expect(
      resolveFace('Mont', ctxWith({ mont: { family: 'swiss', panose: '020A0603' } }))?.family,
    ).toBe('Calibri')
  })

  it('a missing name starting with PT lays out in Arial whatever its entry (probe 2026-11-08)', () => {
    const hints = {
      'pt astra serif': { panose: '020A0603040505020204' },
      'pt zebra roman': { panose: '020A0603040505020204', family: 'roman' },
    }
    expect(resolveFace('PT Astra Serif', ctxWith(hints))).toMatchObject({
      family: 'Arial',
      substituted: true,
      exception: true,
      lineFactor: 1.1499,
      path: ['pt astra serif', 'arial'],
    })
    expect(resolveFace('PT Zebra Roman', ctxWith(hints))?.family).toBe('Arial')
    expect(resolveFace('PTZebra Serif', ctxWith())?.family).toBe('Arial')
    expect(resolveFace('Astra Serif', ctxWith(hints))?.family).toBe('Cambria')
  })

  it('lays the LibreOffice metric clones out with the faces they clone', () => {
    const hints = {
      'liberation sans': { altName: 'Liberation Sans', family: 'swiss' },
      'liberation serif': { altName: 'Liberation Serif', family: 'roman' },
      'liberation mono': { altName: 'Liberation Mono', family: 'modern' },
      carlito: { altName: 'Carlito', family: 'swiss' },
      caladea: { altName: 'Caladea', family: 'roman' },
    }
    expect(resolveFace('Liberation Sans', ctxWith(hints))).toMatchObject({
      family: 'Arial',
      substituted: true,
      lineFactor: 1.1499,
      path: ['liberation sans', 'arial'],
    })
    expect(resolveFace('Liberation Serif', ctxWith(hints))?.family).toMatch(/^times new roman$/i)
    expect(resolveFace('Liberation Mono', ctxWith(hints))?.family).toBe('Calibri')
    expect(resolveFace('Carlito', ctxWith(hints))?.family).toBe('Calibri')
    expect(resolveFace('Caladea', ctxWith(hints))?.family).toBe('Cambria')
  })

  it('follows altName for a missing non-cloud name, never for a cloud family', () => {
    const hints = {
      'montserrat-bold': { altName: 'Times New Roman' },
      montserrat: { altName: 'Times New Roman' },
    }
    expect(resolveFace('Montserrat-Bold', ctxWith(hints))?.family).toMatch(/^times new roman$/i)
    expect(resolveFace('Montserrat', ctxWith(hints))?.family).toBe('Montserrat')
  })

  it('lets a missing Arabic name follow its altName ahead of the Times New Roman rule', () => {
    const hints = {
      'fixture uthman naskh': { altName: 'Arial', charset: 'B2' },
      'fixture kufi': { altName: 'Fixture Naskh Alt' },
    }
    expect(resolveFace('Fixture Uthman Naskh', ctxWith(hints))?.family).toBe('Arial')
    expect(resolveFace('Fixture Kufi', ctxWith(hints))?.family).toMatch(/^times new roman$/i)
    expect(resolveFace('Fixture Uthman Naskh', ctxWith())?.family).toMatch(/^times new roman$/i)
  })

  it('sends a missing Latin-charset name covering the Arabic code page to Arial', () => {
    const hints = {
      'fixture sans ar': { panose: '00000500000000000000', charset: '00', csb0: '00000041' },
      'fixture naskh ar': { charset: 'B2', csb0: '00000040' },
      'fixture sans latin': { panose: '00000500000000000000', charset: '00', csb0: '00000001' },
    }
    expect(resolveFace('Fixture Sans AR', ctxWith(hints))).toMatchObject({
      family: 'Arial',
      substituted: true,
      exception: true,
    })
    expect(resolveFace('Fixture Naskh AR', ctxWith(hints))?.family).toMatch(/^times new roman$/i)
    expect(resolveFace('Fixture Sans Latin', ctxWith(hints))?.family).toBe('Calibri')
  })

  it('sends an Arabic name covering the Arabic code page to Arial ahead of the Times rule, cloud faces excepted', () => {
    const sig = { charset: '00', csb0: '00000041' }
    const hints = {
      'fixture kufi sig': sig,
      'fixture kufi alt': { ...sig, altName: 'Verdana' },
      'fixture kufi b2': { charset: 'B2', csb0: '00000041' },
      'traditional arabic': sig,
    }
    expect(resolveFace('Fixture Kufi Sig', ctxWith(hints))).toMatchObject({
      family: 'Arial',
      rules: ['arabicCodePage'],
    })
    expect(resolveFace('Fixture Kufi Alt', ctxWith(hints))).toMatchObject({
      family: 'Verdana',
      rules: ['altName'],
    })
    expect(resolveFace('Fixture Kufi B2', ctxWith(hints))).toMatchObject({
      family: 'Times New Roman',
      rules: ['table'],
    })
    expect(resolveFace('Traditional Arabic', ctxWith(hints))?.substituted).toBe(false)
  })

  it('names the rule behind every substitution step', () => {
    const hints = {
      'fixture kr': { charset: '81' },
      'fixture ja': { charset: '80' },
      'fixture alt': { altName: 'Fixture Target', panose: '020b0604' },
      'fixture target': { family: 'roman' },
      'fixture plain': { family: 'swiss' },
    }
    const rules = (name: string) => resolveFace(name, ctxWith(hints))?.rules
    expect(rules('Fixture KR')).toEqual(['charset81'])
    expect(rules('Fixture JA')).toEqual(['charset80'])
    expect(rules('Fixture Alt')).toEqual(['altName', 'fallback'])
    expect(rules('Fixture Plain')).toEqual(['fallback'])
    expect(rules('PT Zebra')).toEqual(['table'])
    expect(rules('Noto Sans KR')).toEqual(['hanKr'])
    expect(rules('\u6c49\u4eea\u4e66\u5b8b')).toEqual(['han'])
    expect(rules('\ud734\uba3c\uace0\ub515')).toEqual(['hangul'])
    expect(rules('Arial Bold')).toEqual(['styleSuffix'])
    expect(rules('Calibri')).toEqual([])
  })

  it('a missing name follows its fontTable altName in every slot, Arabic names included', () => {
    const alt = (altName: string) => ({ altName, charset: '00', family: 'roman' })
    expect(
      resolveFace('ae_AlMohanad', ctxWith({ ae_almohanad: alt('Sakkal Majalla') })),
    ).toMatchObject({
      family: 'Sakkal Majalla',
      substituted: true,
      lineFactor: 1.3965,
    })
    expect(resolveFace('ae_AlMohanad', ctxWith({ ae_almohanad: alt('Arial') }))?.family).toBe(
      'Arial',
    )
    // altName beats a sans PANOSE
    expect(
      resolveFace(
        'ProbeAltTimesSans',
        ctxWith({ probealttimessans: { altName: 'Times New Roman', panose: '020b0604' } }),
      )?.family,
    ).toBe('Times New Roman')
    // without altName a Latin-charset roman entry is Cambria
    expect(
      resolveFace('ae_AlMohanad', ctxWith({ ae_almohanad: { charset: '00', family: 'roman' } }))
        ?.family,
    ).toBe('Cambria')
  })

  it('routes the East Asian substitutions', () => {
    const HUMAN_GOTHIC = '\ud734\uba3c\uace0\ub515'
    const han = resolveFace('\u6c49\u4eea\u4e66\u5b8b\u4e8cKW', ctxWith())
    expect(han).toMatchObject({ family: 'SimSun', substituted: true, lineFactor: 1.3 })
    expect(resolveFace(HUMAN_GOTHIC, ctxWith())?.family).toBe('Batang')
    // a sans PANOSE does not move hangul off Batang (probe 2026-10-22)
    expect(
      resolveFace(HUMAN_GOTHIC, ctxWith({ [HUMAN_GOTHIC]: { panose: '020b0604' } }))?.family,
    ).toBe('Batang')
    expect(resolveFace('HCR Batang', ctxWith({ 'hcr batang': { charset: '81' } }))?.family).toBe(
      'Malgun Gothic',
    )
    expect(
      resolveFace('HCR Batang', ctxWith({ 'hcr batang': { charset: '81', family: 'roman' } }))
        ?.family,
    ).toBe('Batang')
    expect(resolveFace('Noto Sans CJK SC', ctxWith())?.family).toBe('SimSun')
    expect(resolveFace('ProbeNoAltEA', ctxWith({ probenoaltea: { charset: '80' } }))?.family).toBe(
      'Yu Gothic',
    )
    // Hiragino Kaku Gothic Pro W3 by its kana name: the HG rule without altName, the altName with one
    const hiragino = '\u30d2\u30e9\u30ae\u30ce\u89d2\u30b4 Pro W3'
    expect(resolveFace(hiragino, ctxWith())?.family).toBe('Yu Gothic')
    const viaAlt = resolveFace(
      hiragino,
      ctxWith({ [hiragino.toLowerCase()]: { altName: 'Cambria', charset: '00', family: 'roman' } }),
    )
    expect(viaAlt?.family).toBe('Cambria')
    expect(viaAlt?.exception).toBe(false)
    expect(
      resolveFace(
        'ProbeNoAltEA',
        ctxWith({ probenoaltea: { charset: '80', altName: 'ProbeNoAltEA' } }),
      )?.family,
    ).toBe('Yu Gothic')
    expect(
      resolveFace(
        'ProbeAltMincho',
        ctxWith({ probealtmincho: { charset: '80', altName: 'MS Mincho' } }),
      )?.family,
    ).toBe('MS Mincho')
    expect(resolveFace('BIZ UDPGothic', ctxWith())?.path.at(-1)).toBe('yu gothic')
    expect(resolveFace('DejaVu Sans', ctxWith())?.path.at(-1)).toBe('verdana')
    expect(resolveFace('Noto Naskh Arabic', ctxWith())?.path.at(-1)).toBe('times new roman')
    expect(resolveFace('TAU-Marutham', ctxWith())?.family).toBe('Vijaya')
    expect(
      resolveFace('Acme Myeongjo', ctxWith({ 'acme myeongjo': { altName: 'Batang' } }))?.family,
    ).toBe('Batang')
    // the macOS on-demand Nanum asset is a face Word lays out with
    expect(resolveFace('Nanum Myeongjo', ctxWith())?.source).toBe('asset')
    expect(resolveFace('NanumGothic', ctxWith())?.lineFactor).toBe(1.495)
    // hangul names: Batang whatever the PANOSE; a style suffix or MT tag drops to the base family
    expect(
      resolveFace(
        '\ud504\ub85c\ube0c',
        ctxWith({ '\ud504\ub85c\ube0c': { panose: '020B0604020202020204' } }),
      )?.family,
    ).toBe('Batang')
    expect(
      resolveFace('Arial Bold', ctxWith({ 'arial bold': { panose: '020B0604020202090204' } })),
    ).toMatchObject({ family: 'Arial', substituted: true, exception: false })
    expect(resolveFace('Courier New Regular', ctxWith())?.family).toBe('Courier New')
    // a weight-worded instance of a cloud family is not laid out as that family: Word for Mac
    // paces Bahnschrift Light Condensed at Calibri-like advances (truth 2026-10-28), so it
    // substitutes like any missing swiss name
    expect(
      resolveFace(
        'Bahnschrift Light Condensed',
        ctxWith({ 'bahnschrift light condensed': { panose: '020B0502040204020203' } }),
      ),
    ).toMatchObject({ family: 'Calibri', substituted: true, lineFactor: 1.2207 })
    expect(resolveFace('Roboto Condensed Light', ctxWith())?.family).toBe('Roboto Condensed Light')
    // a cloud family name wins over its fontTable altName (truth 2026-10-19)
    expect(
      resolveFace('Roboto', ctxWith({ roboto: { altName: 'Times New Roman', family: 'roman' } })),
    ).toMatchObject({ family: 'Roboto', substituted: false, lineFactor: 1.1719 })
    expect(resolveFace('Arial MT', ctxWith({ 'arial mt': { altName: 'Arial MT' } }))?.family).toBe(
      'Arial',
    )
  })

  it('reports an embedded face with its own box', () => {
    const ctx = storeContext(data, {
      embedded: (key) =>
        key === 'acme sans'
          ? { hheaTotal: 1.31, ascent: 1.0, descent: 0.31, xAvg: null, eastAsian: false }
          : null,
    })
    expect(resolveFace('Acme Sans', ctx)).toMatchObject({ source: 'embedded', lineFactor: 1.31 })
  })

  it('lays a family Word ships with its own copy, not the embed', () => {
    const box = { hheaTotal: 1.4282, ascent: 1.1216, descent: 0.3066, xAvg: null, eastAsian: false }
    const ctx = storeContext(data, {
      embedded: (key) => (key === 'century gothic' || key === 'roboto' ? box : null),
    })
    expect(resolveFace('Century Gothic', ctx)).toMatchObject({
      source: 'dfonts',
      lineFactor: 1.2261,
    })
    expect(resolveFace('Roboto', ctx)).toMatchObject({ source: 'embedded', lineFactor: 1.4282 })
  })

  it('normalises declared names like the renderer tables', () => {
    expect(faceKey(' "Segoe UI", sans-serif')).toBe('segoe ui')
    expect(faceKey('\uff2d\uff33 \u660e\u671d')).toBe('ms \u660e\u671d')
    expect(panoseSerif('020b0604020202020204')).toBe('sans')
    expect(panoseSerif('02040503050406030204')).toBe('serif')
    expect(panoseSerif('00000000000000000000')).toBeUndefined()
  })
})
