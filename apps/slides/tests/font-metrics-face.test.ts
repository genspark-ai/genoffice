import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import * as opentype from 'opentype.js'

// Vitest turns ?asset imports into /@fs URLs, while font metrics reads the files from disk.
vi.mock('@genoffice/ui/fonts/Carlito-Regular.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Regular.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Bold.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Bold.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Italic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Italic.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-BoldItalic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-BoldItalic.ttf'),
}))

// shaped-metrics pulls in the harfbuzz wasm (?asset only resolves in electron builds); disconnected in unit tests
vi.mock('../src/main/shaped-metrics', () => ({
  initShapedMetrics: () => {},
  shapedMeasure: () => null,
  shapedFamily: () => null,
  gtMeasure: () => null,
  complexScriptOf: () => null,
}))

import { extractFace } from '../src/main/fonts'

/**
 * The metrics-only face (gvar stubbed, cvar dropped, parsed with lowMemory) must measure
 * exactly like the full parse — at the default instance and at an instantiated weight —
 * while leaving no per-glyph variation data behind. Uses macOS system fonts: SF Pro is a
 * variable font whose real gvar costs ~220 MB of heap; Songti is a ttc.
 */
const SF = '/System/Library/Fonts/SFNS.ttf'
const SONGTI = '/System/Library/Fonts/Supplemental/Songti.ttc'
const sample = [...'Hello 3,173 \u4e2d\u6587']

type Parsed = ReturnType<typeof opentype.parse> & {
  variation?: { getTransform(g: unknown, c: Record<string, number>): { advanceWidth: number } }
  tables: Record<string, { glyphVariations?: unknown[] } | undefined>
}

const advances = (font: Parsed, wght?: number) =>
  sample.map((ch) => {
    const g = font.charToGlyph(ch)
    return wght && font.variation
      ? font.variation.getTransform(g, { wght }).advanceWidth
      : g.advanceWidth
  })

describe.runIf(process.platform === 'darwin' && existsSync(SF) && existsSync(SONGTI))(
  'metrics-only face extraction',
  () => {
    it('keeps default and weight-instantiated advances of a variable font with an empty gvar', () => {
      const buf = readFileSync(SF)
      const full = opentype.parse(extractFace(buf, 0)) as Parsed
      const lean = opentype.parse(extractFace(buf, 0, undefined, true), {
        lowMemory: true,
      }) as Parsed
      expect(lean.variation).toBeDefined()
      expect(lean.tables.hvar).toBeDefined()
      expect(Object.keys(lean.tables.gvar?.glyphVariations ?? {})).toHaveLength(0)
      expect(lean.tables.cvar).toBeUndefined()
      // getTransform writes the instantiated advance back onto the glyph, so read defaults first
      const leanDefault = advances(lean)
      const fullDefault = advances(full)
      expect(leanDefault).toEqual(fullDefault)
      expect(advances(lean, 700)).toEqual(advances(full, 700))
      expect(advances(lean, 700)).not.toEqual(leanDefault)
    })

    it('a ttc face without variation tables is unchanged apart from lazy glyphs', () => {
      const buf = readFileSync(SONGTI)
      const face = buf.readUInt32BE(12)
      const full = opentype.parse(extractFace(buf, face)) as Parsed
      const lean = opentype.parse(extractFace(buf, face, undefined, true), {
        lowMemory: true,
      }) as Parsed
      expect(lean.numGlyphs).toBe(full.numGlyphs)
      expect(advances(lean)).toEqual(advances(full))
    })
  },
)
