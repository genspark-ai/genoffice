import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

import {
  faceKey,
  lineFactorOf,
  panoseSerif,
  resolveFace,
  storeContext,
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

  it('falls to Calibri with a fontTable entry, Cambria with PANOSE serif or no entry', () => {
    expect(resolveFace('Mont', ctxWith({ mont: { panose: '00000700' } }))).toMatchObject({
      family: 'Calibri',
      substituted: true,
      lineFactor: 1.2207,
      path: ['mont', 'calibri'],
    })
    expect(resolveFace('Mont', ctxWith({ mont: { panose: '02030504' } }))?.family).toBe('Cambria')
    expect(resolveFace('Mont', ctxWith())?.family).toBe('Cambria')
    // probe 2026-10-22: w:family roman without PANOSE is Cambria, swiss / modern / auto Calibri
    expect(resolveFace('Mont', ctxWith({ mont: { family: 'roman' } }))?.family).toBe('Cambria')
    expect(resolveFace('Mont', ctxWith({ mont: { family: 'modern' } }))?.family).toBe('Calibri')
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

  it('normalises declared names like the renderer tables', () => {
    expect(faceKey(' "Segoe UI", sans-serif')).toBe('segoe ui')
    expect(faceKey('\uff2d\uff33 \u660e\u671d')).toBe('ms \u660e\u671d')
    expect(panoseSerif('020b0604020202020204')).toBe('sans')
    expect(panoseSerif('02040503050406030204')).toBe('serif')
    expect(panoseSerif('00000000000000000000')).toBeUndefined()
  })
})
