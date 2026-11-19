/**
 * Real-advance FontMetricsProvider for node: walks the renderer's CSS family
 * chain (`cssFontFamily`) the way Chromium does — per character, first family
 * with a glyph wins — reading the bundled @font-face files (fonts.css +
 * fonts.generated.css, ttf and woff2, size-adjust applied) and the installed
 * fonts through @genoffice/font-metrics. A character no chain face covers takes
 * the installed face Chromium on macOS would cascade to for its script
 * (`system` counter); when none is installed it falls back to 1 em (CJK /
 * fullwidth) or the heuristic estimate (`unresolved` counter).
 */
import { existsSync, readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { advanceWidths, fileAdvanceWidths } from '@genoffice/font-metrics'
import {
  HeuristicMetrics,
  type FontMetricsProvider,
  type RunStyle,
} from '../../src/renderer/line-metrics'
import { cssFontFamily } from '../../src/renderer/text-layout/font-rules'
import { GENERIC_FAMILIES } from '../../src/renderer/text-layout/font-rules/availability'

interface CssFace {
  weight: 'normal' | 'bold'
  italic: boolean
  srcs: Array<{ file: string } | { local: string }>
  sizeAdjust: number
  ranges: Array<[number, number]> | null
}

const here = dirname(fileURLToPath(import.meta.url))
const FONT_DIR = resolve(here, '../../src/renderer/fonts')
const UI_FONT_DIR = resolve(here, '../../../../packages/ui/src/fonts')

function parseRanges(value: string): Array<[number, number]> {
  return value.split(',').map((part) => {
    const m = /U\+([0-9A-Fa-f?]+)(?:-([0-9A-Fa-f]+))?/.exec(part.trim())!
    const lo = parseInt(m[1]!.replace(/\?/g, '0'), 16)
    const hi = m[2] ? parseInt(m[2], 16) : parseInt(m[1]!.replace(/\?/g, 'F'), 16)
    return [lo, hi]
  })
}

function parseFontFaces(css: string, faces: Map<string, CssFace[]>): void {
  for (const block of css.matchAll(/@font-face\s*\{([^}]*)\}/g)) {
    const body = block[1]!
    const prop = (name: string) =>
      new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'm').exec(body)?.[1]?.trim()
    const family = prop('font-family')?.replace(/^['"]|['"]$/g, '')
    const src = prop('src')
    if (!family || !src) continue
    const srcs: CssFace['srcs'] = []
    for (const s of src.matchAll(/(url|local)\(\s*['"]?([^'")]+)['"]?\s*\)/g)) {
      if (s[1] === 'url') {
        const ref = s[2]!
        const file = ref.startsWith('@genoffice/ui/fonts/')
          ? resolve(UI_FONT_DIR, ref.slice('@genoffice/ui/fonts/'.length))
          : resolve(FONT_DIR, ref)
        if (/\.(ttf|otf|ttc|woff2)$/i.test(file) && existsSync(file)) srcs.push({ file })
      } else srcs.push({ local: s[2]! })
    }
    if (!srcs.length) continue
    const weight = prop('font-weight') ?? 'normal'
    const adjust = prop('size-adjust')
    const range = prop('unicode-range')
    const face: CssFace = {
      weight: weight === 'bold' || parseInt(weight, 10) >= 600 ? 'bold' : 'normal',
      italic: (prop('font-style') ?? 'normal') !== 'normal',
      srcs,
      sizeAdjust: adjust ? parseFloat(adjust) / 100 : 1,
      ranges: range ? parseRanges(range) : null,
    }
    const key = family.toLowerCase()
    faces.set(key, [...(faces.get(key) ?? []), face])
  }
}

let cssFaces: Map<string, CssFace[]> | null = null
function fontFaces(): Map<string, CssFace[]> {
  if (!cssFaces) {
    cssFaces = new Map()
    for (const name of ['fonts.css', 'fonts.generated.css'])
      parseFontFaces(readFileSync(resolve(FONT_DIR, name), 'utf8'), cssFaces)
  }
  return cssFaces
}

const isFullwidth = (cp: number) =>
  (cp >= 0x1100 && cp <= 0x11ff) ||
  (cp >= 0x2e80 && cp <= 0x9fff) ||
  (cp >= 0xa960 && cp <= 0xa97f) ||
  (cp >= 0xac00 && cp <= 0xd7ff) ||
  (cp >= 0xf900 && cp <= 0xfaff) ||
  (cp >= 0xfe30 && cp <= 0xfe4f) ||
  (cp >= 0xff00 && cp <= 0xff60) ||
  (cp >= 0xffe0 && cp <= 0xffe6) ||
  (cp >= 0x20000 && cp <= 0x3fffd)

/**
 * Chromium's face choice inside one family: the cuts matching weight/style best (else the
 * regular cut — synthesised bold keeps its advances), those whose unicode-range covers the
 * character, later-declared first (a range overlay declared after the catch-all wins).
 */
function pickFaces(faces: CssFace[], style: RunStyle, cp: number): CssFace[] {
  const score = (f: CssFace) =>
    (f.weight === (style.bold ? 'bold' : 'normal') ? 2 : 0) + (f.italic === style.italic ? 1 : 0)
  const best = Math.max(...faces.map(score))
  return faces
    .filter((f) => score(f) === best)
    .filter((f) => !f.ranges || f.ranges.some(([lo, hi]) => cp >= lo && cp <= hi))
    .reverse()
}

/** CoreText's cascade list as Chromium on macOS sees it, by script: the face a run gets
 * when none of its own chain covers the character (first installed family wins) */
const SYSTEM_CASCADE: Array<[string, Array<[number, number]>, string[]]> = [
  [
    'hangul',
    [
      [0x1100, 0x11ff],
      [0x3130, 0x318f],
      [0xa960, 0xa97f],
      [0xac00, 0xd7ff],
    ],
    ['Apple SD Gothic Neo', 'Malgun Gothic', 'NanumGothic', 'Arial Unicode MS'],
  ],
  [
    'kana',
    [
      [0x3040, 0x30ff],
      [0x31f0, 0x31ff],
      [0xff66, 0xff9f],
    ],
    ['Hiragino Sans', 'Yu Gothic', 'Meiryo', 'MS Gothic', 'Arial Unicode MS'],
  ],
  [
    'han',
    [
      [0x2e80, 0x33ff],
      [0x3400, 0x9fff],
      [0xf900, 0xfaff],
      [0xfe30, 0xfe4f],
      [0xff00, 0xffef],
      [0x20000, 0x3fffd],
    ],
    ['PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', 'SimSun', 'Arial Unicode MS'],
  ],
  ['thai', [[0x0e00, 0x0e7f]], ['Thonburi', 'Tahoma', 'Leelawadee UI', 'Arial Unicode MS']],
  [
    'arabic',
    [
      [0x0600, 0x06ff],
      [0x0750, 0x077f],
      [0x08a0, 0x08ff],
      [0xfb50, 0xfdff],
      [0xfe70, 0xfeff],
    ],
    ['Geeza Pro', 'Arial', 'Tahoma', 'Arial Unicode MS'],
  ],
  ['hebrew', [[0x0590, 0x05ff]], ['Arial Hebrew', 'Arial', 'Tahoma']],
  [
    'devanagari',
    [[0x0900, 0x097f]],
    ['Kohinoor Devanagari', 'Devanagari Sangam MN', 'Mangal', 'Arial Unicode MS'],
  ],
  ['bengali', [[0x0980, 0x09ff]], ['Kohinoor Bangla', 'Bangla Sangam MN', 'Vrinda']],
  ['tamil', [[0x0b80, 0x0bff]], ['Tamil Sangam MN', 'Latha']],
]
const OTHER_CASCADE = ['Arial', 'Helvetica', 'Arial Unicode MS', 'Times New Roman']

function scriptOf(cp: number): [string, string[]] {
  for (const [name, ranges, families] of SYSTEM_CASCADE)
    if (ranges.some(([lo, hi]) => cp >= lo && cp <= hi)) return [name, families]
  return ['other', OTHER_CASCADE]
}

export interface NodeMetricsStats {
  /** characters a face of the run's own chain measured */
  resolved: number
  /** characters measured by the per-script system cascade outside the chain */
  system: number
  /** characters no installed face covers: 1 em or the heuristic estimate */
  unresolved: number
  /** families (chain heads) that reached the fallback at least once */
  unresolvedFamilies: Set<string>
  /** per script: how many characters took the system cascade / stayed unresolved */
  scripts: Map<string, { system: number; unresolved: number }>
}

export class NodeFontMetrics implements FontMetricsProvider {
  readonly stats: NodeMetricsStats = {
    resolved: 0,
    system: 0,
    unresolved: 0,
    unresolvedFamilies: new Set(),
    scripts: new Map(),
  }
  private heuristic = new HeuristicMetrics()
  private chains = new Map<string, string[]>()
  private cache = new Map<string, { em: number; via: 'chain' | 'system' | 'none' }>()

  metrics(style: RunStyle) {
    return this.heuristic.metrics(style)
  }

  /** chains and advances depend on the document fontTable: drop them when it changes */
  resetDocument(): void {
    this.chains.clear()
    this.cache.clear()
  }

  private script(cp: number) {
    const name = scriptOf(cp)[0]
    let entry = this.stats.scripts.get(name)
    if (!entry) this.stats.scripts.set(name, (entry = { system: 0, unresolved: 0 }))
    return entry
  }

  private chainOf(family: string): string[] {
    let chain = this.chains.get(family)
    if (!chain) {
      chain = cssFontFamily(family)
        .split(',')
        .map((f) => f.trim().replace(/^['"]|['"]$/g, ''))
        .filter((f) => f && !GENERIC_FAMILIES.has(f.toLowerCase()))
      this.chains.set(family, chain)
    }
    return chain
  }

  /** advance (twips at sizePt) of one character in one family, null when no face covers it */
  private familyAdvance(family: string, ch: string, cp: number, style: RunStyle, pt: number) {
    const faces = fontFaces().get(family.toLowerCase())
    if (faces) {
      for (const face of pickFaces(faces, style, cp)) {
        for (const src of face.srcs) {
          // a local() face loads the cut its own descriptors name, not the run's
          const w =
            'file' in src
              ? fileAdvanceWidths(src.file, ch, pt)
              : advanceWidths(src.local, ch, pt, {
                  bold: face.weight === 'bold',
                  italic: face.italic,
                })
          if (w && !Number.isNaN(w[0]!)) return w[0]! * face.sizeAdjust
        }
      }
      return null
    }
    const w = advanceWidths(family, ch, pt, { bold: style.bold, italic: style.italic })
    return w && !Number.isNaN(w[0]!) ? w[0]! : null
  }

  measure(text: string, style: RunStyle): number {
    const chain = this.chainOf(style.fontFamily)
    let px = 0
    for (const ch of text) {
      const cp = ch.codePointAt(0)!
      const key = `${style.fontFamily}|${style.bold}|${style.italic}|${ch}`
      let hit = this.cache.get(key)
      if (hit === undefined) {
        hit = { em: NaN, via: 'none' }
        // advances at 1000 pt in twips -> em, cached per character
        for (const family of chain) {
          const tw = this.familyAdvance(family, ch, cp, style, 1000)
          if (tw !== null) {
            hit = { em: tw / 20000, via: 'chain' }
            break
          }
        }
        if (hit.via === 'none') {
          for (const family of scriptOf(cp)[1]) {
            const w = advanceWidths(family, ch, 1000, { bold: style.bold, italic: style.italic })
            if (w && !Number.isNaN(w[0]!)) {
              hit = { em: w[0]! / 20000, via: 'system' }
              break
            }
          }
        }
        this.cache.set(key, hit)
      }
      if (hit.via === 'none') {
        this.stats.unresolved++
        this.stats.unresolvedFamilies.add(style.fontFamily)
        this.script(cp).unresolved++
        px += isFullwidth(cp) ? style.fontSizePx : this.heuristic.measure(ch, style)
      } else {
        if (hit.via === 'system') {
          this.stats.system++
          this.script(cp).system++
        } else this.stats.resolved++
        px += hit.em * style.fontSizePx
      }
    }
    return px
  }
}
