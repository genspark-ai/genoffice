#!/usr/bin/env node
// Emits the renderer's metric-alias @font-face blocks and the line-factor table
// from the measurement manifest, so every size-adjust / override / factor is a
// manifest entry with a source instead of a hand-edited literal.
//
//   node tools/build-font-aliases.mjs            write both generated files
//   node tools/build-font-aliases.mjs --check    exit 1 when they are stale
//   node tools/build-font-aliases.mjs --refit-aptos [--write] [--dfonts DIR] [--cloud-fonts DIR]
//        re-derive the Aptos bucket ratios from the Word font files over Carlito
//        (hmtx), report buckets that do not reproduce, optionally update the manifest
//   node tools/build-font-aliases.mjs --refit-buckets [--write] [--family NAME]
//        rebuild the faces of every family with refit.method 'hmtx-buckets' from
//        the font files named in refit.cuts (see fitBuckets)
//
// Inputs:  packages/font-metrics/measurements/families/*.json  (one per CSS family)
//          packages/font-metrics/measurements/line-factors.json
// Outputs: apps/docs/src/renderer/fonts/fonts.generated.css
//          apps/docs/src/renderer/line-factors.generated.ts
import { spawnSync } from 'node:child_process'
import { readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
export const MANIFEST_DIR = 'packages/font-metrics/measurements'
export const GENERATED_CSS = 'apps/docs/src/renderer/fonts/fonts.generated.css'
export const GENERATED_TS = 'apps/docs/src/renderer/line-factors.generated.ts'
const UI_FONTS = 'packages/ui/src/fonts'

export const SOURCE_KINDS = ['probe', 'hmtx', 'corpus', 'legacy']
const STAGES = ['first', 'arabic', 'latin', 'last']
const PATTERN_REFS = ['TAMIL_FONT_RE', 'SEGOE_UI_TEXT_RE', 'NOTO_HAN_REGIONAL_RE']
const DATE = /^\d{4}-\d{2}-\d{2}$/

export function loadManifest(root = ROOT) {
  const dir = join(root, MANIFEST_DIR, 'families')
  const families = readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .sort()
    .map((f) => ({
      file: `${MANIFEST_DIR}/families/${f}`,
      ...JSON.parse(readFileSync(join(dir, f), 'utf8')),
    }))
  const lineFactors = JSON.parse(
    readFileSync(join(root, MANIFEST_DIR, 'line-factors.json'), 'utf8'),
  )
  return { families, lineFactors }
}

/** problems as `where: message`; empty when the manifest is well-formed */
export function validateManifest({ families, lineFactors }) {
  const out = []
  const checkSource = (where, source) => {
    if (!source || typeof source !== 'object') return out.push(`${where}: missing source`)
    if (!SOURCE_KINDS.includes(source.kind))
      out.push(`${where}: source.kind must be one of ${SOURCE_KINDS.join('|')}`)
    if (!DATE.test(source.date || '')) out.push(`${where}: source.date must be YYYY-MM-DD`)
    if (source.kind === 'hmtx' && !(Array.isArray(source.fonts) && source.fonts.length))
      out.push(`${where}: hmtx source needs the font files in source.fonts`)
    if (source.kind !== 'hmtx' && !source.probe) out.push(`${where}: source needs a probe id`)
  }
  const seen = new Set()
  for (const fam of families) {
    const where = fam.file
    if (!fam.family) out.push(`${where}: missing family`)
    if (seen.has(fam.family)) out.push(`${where}: duplicate family ${fam.family}`)
    seen.add(fam.family)
    checkSource(where, fam.source)
    if (!Array.isArray(fam.faces) || !fam.faces.length) out.push(`${where}: no faces`)
    for (const [i, face] of (fam.faces || []).entries()) {
      if (!face.src) out.push(`${where} face ${i}: missing src`)
      for (const k of ['sizeAdjust', 'ascentOverride', 'descentOverride', 'lineGapOverride'])
        if (face[k] !== undefined && typeof face[k] !== 'number')
          out.push(`${where} face ${i}: ${k} must be a number (percent)`)
      if (face.pin && !face.pin.reason) out.push(`${where} face ${i}: pin needs a reason`)
    }
  }
  const ids = new Set()
  for (const rule of lineFactors.rules || []) {
    const where = `${MANIFEST_DIR}/line-factors.json rule ${rule.id}`
    if (!rule.id) out.push(`${where}: missing id`)
    if (ids.has(rule.id)) out.push(`${where}: duplicate id`)
    ids.add(rule.id)
    if (!STAGES.includes(rule.stage)) out.push(`${where}: stage must be one of ${STAGES.join('|')}`)
    for (const k of ['match', 'exclude']) {
      const m = rule[k]
      if (m === undefined) {
        if (k === 'match') out.push(`${where}: missing match`)
        continue
      }
      if (typeof m === 'string') {
        try {
          new RegExp(m)
        } catch (e) {
          out.push(`${where}: ${k} is not a regex: ${e.message}`)
        }
      } else if (!PATTERN_REFS.includes(m?.ref))
        out.push(`${where}: ${k}.ref must be one of ${PATTERN_REFS.join('|')}`)
    }
    if (typeof rule.factor !== 'number') out.push(`${where}: factor must be a number`)
    checkSource(where, rule.source)
    if (rule.pin && !rule.pin.reason) out.push(`${where}: pin needs a reason`)
  }
  return out
}

const pct = (v) => `${v}%`

function faceCss(family, face) {
  const lines = [`  font-family: '${family}';`, `  src: ${face.src};`]
  if (face.sizeAdjust !== undefined) lines.push(`  size-adjust: ${pct(face.sizeAdjust)};`)
  if (face.unicodeRange) lines.push(`  unicode-range: ${face.unicodeRange};`)
  if (face.ascentOverride !== undefined)
    lines.push(`  ascent-override: ${pct(face.ascentOverride)};`)
  if (face.descentOverride !== undefined)
    lines.push(`  descent-override: ${pct(face.descentOverride)};`)
  if (face.lineGapOverride !== undefined)
    lines.push(`  line-gap-override: ${pct(face.lineGapOverride)};`)
  // null omits a descriptor (CSS default); undefined emits the usual value
  if (face.weight !== null) lines.push(`  font-weight: ${face.weight ?? 'normal'};`)
  if (face.style !== null) lines.push(`  font-style: ${face.style ?? 'normal'};`)
  if (face.display !== null) lines.push(`  font-display: ${face.display ?? 'block'};`)
  return `@font-face {\n${lines.join('\n')}\n}`
}

const comment = (text) => `/* ${text.split('\n').join('\n * ')} */`

export function renderCss({ families }) {
  const parts = [
    comment(
      'Generated by tools/build-font-aliases.mjs from packages/font-metrics/measurements;\n' +
        'do not hand-edit. Every size-adjust and metric override here is a manifest entry\n' +
        'with a measurement source; fonts.css imports this file and keeps the hand-authored\n' +
        'faces (bundled subsets, PUA blank, box drawing).',
    ),
  ]
  for (const fam of families) {
    const head = [`${fam.family}: ${fam.source.kind} ${fam.source.date}`]
    if (fam.source.probe) head.push(fam.source.probe)
    if (fam.source.fonts) head.push(`fonts: ${fam.source.fonts.join('; ')}`)
    const block = [comment((fam.note ? fam.note + '\n\n' : '') + head.join('\n'))]
    for (const face of fam.faces) block.push(faceCss(fam.family, face))
    parts.push(block.join('\n'))
  }
  return parts.join('\n\n') + '\n'
}

const escapeNonAscii = (s) =>
  s.replace(/[^\x20-\x7e]/g, (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'))
const tsString = (s) => `'${escapeNonAscii(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'")}'`
const tsRegex = (m) =>
  typeof m === 'string' ? `/${escapeNonAscii(m).replace(/\//g, '\\/')}/` : m.ref

export function renderTs({ lineFactors }) {
  const refs = new Set()
  for (const r of lineFactors.rules)
    for (const m of [r.match, r.exclude]) if (m && typeof m !== 'string') refs.add(m.ref)
  const sourceText = (s) =>
    `${s.kind} ${s.date}${s.probe ? `: ${s.probe}` : ''}${s.fonts ? ` [${s.fonts.join('; ')}]` : ''}`
  const rows = lineFactors.rules.map((r) => {
    const fields = [
      `id: ${tsString(r.id)}`,
      `stage: ${tsString(r.stage)}`,
      `match: ${tsRegex(r.match)}`,
    ]
    if (r.exclude) fields.push(`exclude: ${tsRegex(r.exclude)}`)
    fields.push(`factor: ${r.factor}`, `source: ${tsString(sourceText(r.source))}`)
    if (r.pin) fields.push(`pin: ${tsRegex(r.pin.match ?? r.match)}`)
    return `  {\n    ${fields.join(',\n    ')},\n  },`
  })
  return [
    '// Generated by tools/build-font-aliases.mjs from',
    '// packages/font-metrics/measurements/line-factors.json; do not hand-edit.',
    ...(refs.size
      ? [`import { ${[...refs].sort().join(', ')} } from './font-name-patterns'`, '']
      : ['']),
    `export type LineFactorStage = ${STAGES.map(tsString).join(' | ')}`,
    '',
    'export interface LineFactorRule {',
    '  readonly id: string',
    '  /** where knownLineFactor() consults the rule relative to its behavioural branches */',
    '  readonly stage: LineFactorStage',
    '  /** tested against the lower-cased declared name */',
    '  readonly match: RegExp',
    '  readonly exclude?: RegExp',
    '  readonly factor: number',
    '  readonly source: string',
    '  /** names whose table value outranks the resolver (tested against the NFKC key) */',
    '  readonly pin?: RegExp',
    '}',
    '',
    'export const LINE_FACTOR_RULES: readonly LineFactorRule[] = [',
    ...rows,
    ']',
    '',
    'export const LINE_FACTORS: Readonly<Record<string, LineFactorRule>> = Object.fromEntries(',
    '  LINE_FACTOR_RULES.map((rule) => [rule.id, rule]),',
    ')',
    '',
    'export const LINE_FACTOR_PINS: readonly RegExp[] = LINE_FACTOR_RULES.flatMap((rule) =>',
    '  rule.pin ? [rule.pin] : [],',
    ')',
    '',
  ].join('\n')
}

// ---- Aptos refit: bucket size-adjust = mean(Aptos advance / Carlito advance) x kerning factor

function readSfnt(path, index = 0) {
  const buf = readFileSync(path)
  let off = 0
  if (buf.readUInt32BE(0) === 0x74746366) off = buf.readUInt32BE(12 + 4 * index)
  const tables = {}
  for (let i = 0, n = buf.readUInt16BE(off + 4); i < n; i++) {
    const r = off + 12 + 16 * i
    const start = buf.readUInt32BE(r + 8)
    tables[buf.toString('latin1', r, r + 4)] = buf.subarray(start, start + buf.readUInt32BE(r + 12))
  }
  const upm = tables.head.readUInt16BE(18)
  const numH = tables.hhea.readUInt16BE(34)
  const glyphOf = cmapLookup(tables.cmap)
  const names = {}
  const name = tables.name
  for (let i = 0, count = name.readUInt16BE(2), so = name.readUInt16BE(4); i < count; i++) {
    const r = 6 + 12 * i
    if (name.readUInt16BE(r) !== 3) continue
    const english = name.readUInt16BE(r + 4) === 0x409
    const id = name.readUInt16BE(r + 6)
    if (names[id] && !english) continue
    const o = so + name.readUInt16BE(r + 10)
    const raw = name.subarray(o, o + name.readUInt16BE(r + 8))
    names[id] = Array.from({ length: raw.length / 2 }, (_, k) =>
      String.fromCharCode(raw.readUInt16BE(2 * k)),
    ).join('')
  }
  return {
    fullName: names[4],
    advance(cp) {
      const g = glyphOf(cp)
      return g ? tables.hmtx.readUInt16BE(4 * Math.min(g, numH - 1)) / upm : null
    },
  }
}

function cmapLookup(b) {
  let best = null
  for (let i = 0, n = b.readUInt16BE(2); i < n; i++) {
    const r = 4 + 8 * i
    const p = b.readUInt16BE(r)
    const e = b.readUInt16BE(r + 2)
    const rank =
      p === 3 && e === 10 ? 4 : p === 0 && e >= 4 ? 3 : p === 3 && e === 1 ? 2 : p === 0 ? 1 : 0
    if (rank && (!best || rank > best.rank)) best = { off: b.readUInt32BE(r + 4), rank }
  }
  const s = b.subarray(best.off)
  const format = s.readUInt16BE(0)
  if (format === 4) {
    const segX2 = s.readUInt16BE(6)
    return (cp) => {
      if (cp > 0xffff) return 0
      for (let i = 0; i < segX2 / 2; i++) {
        if (cp > s.readUInt16BE(14 + 2 * i)) continue
        const start = s.readUInt16BE(16 + segX2 + 2 * i)
        if (cp < start) return 0
        const delta = s.readInt16BE(16 + 2 * segX2 + 2 * i)
        const roPos = 16 + 3 * segX2 + 2 * i
        const ro = s.readUInt16BE(roPos)
        if (ro === 0) return (cp + delta) & 0xffff
        const g = s.readUInt16BE(roPos + ro + 2 * (cp - start))
        return g ? (g + delta) & 0xffff : 0
      }
      return 0
    }
  }
  if (format === 12) {
    const groups = s.readUInt32BE(12)
    return (cp) => {
      for (let i = 0; i < groups; i++) {
        const r = 16 + 12 * i
        const a = s.readUInt32BE(r)
        if (cp < a) return 0
        if (cp <= s.readUInt32BE(r + 4)) return s.readUInt32BE(r + 8) + cp - a
      }
      return 0
    }
  }
  throw new Error(`cmap format ${format} unsupported`)
}

const codepoints = (range) =>
  range.split(',').flatMap((part) => {
    const m = /U\+([0-9A-F]+)(?:-([0-9A-F]+))?/i.exec(part.trim())
    const a = parseInt(m[1], 16)
    const b = m[2] ? parseInt(m[2], 16) : a
    return Array.from({ length: b - a + 1 }, (_, i) => a + i)
  })

const CUT_NAMES = {
  'normal/normal': ['Regular', ''],
  'bold/normal': ['Bold'],
  'normal/italic': ['Italic'],
  'bold/italic': ['Bold Italic'],
}

/** Word font files per family cut, located by their name-table full name */
function wordFontsFor(family, dirs) {
  const base = family.replace(/ GO$/, '')
  const files = dirs.flatMap((d) => {
    try {
      return readdirSync(d)
        .filter((f) => /\.ttf$/i.test(f))
        .map((f) => join(d, f))
    } catch {
      return []
    }
  })
  const byName = new Map()
  for (const f of files) {
    try {
      const font = readSfnt(f)
      if (font.fullName?.startsWith(base)) byName.set(font.fullName, font)
    } catch {
      /* not a plain sfnt */
    }
  }
  const out = {}
  for (const [cut, suffixes] of Object.entries(CUT_NAMES)) {
    for (const suffix of suffixes) {
      const font = byName.get(suffix ? `${base} ${suffix}` : base)
      if (font) {
        out[cut] = font
        break
      }
    }
  }
  return out
}

const carlitoFile = {
  'normal/normal': 'Carlito-Regular.ttf',
  'bold/normal': 'Carlito-Bold.ttf',
  'normal/italic': 'Carlito-Italic.ttf',
  'bold/italic': 'Carlito-BoldItalic.ttf',
}

export function refitAptos(manifest, { dfonts, cloudFonts, write }) {
  const report = []
  let changed = 0
  for (const fam of manifest.families) {
    if (fam.refit?.method !== 'aptos-buckets') continue
    const word = wordFontsFor(fam.family, [dfonts, cloudFonts].filter(Boolean))
    for (const cut of Object.keys(carlitoFile)) {
      const [weight, style] = cut.split('/')
      const faces = fam.faces.filter(
        (f) =>
          (f.weight ?? 'normal') === weight && (f.style ?? 'normal') === style && f.unicodeRange,
      )
      if (!word[cut]) {
        report.push(
          `${fam.family} ${cut}: Word font file not found, ${faces.length} buckets unverified`,
        )
        continue
      }
      const carlito = readSfnt(join(ROOT, UI_FONTS, carlitoFile[cut]))
      const raw = faces.map((face) => {
        const ratios = codepoints(face.unicodeRange)
          .map((cp) => {
            const a = word[cut].advance(cp)
            const b = carlito.advance(cp)
            return a && b ? a / b : null
          })
          .filter((r) => r !== null)
        return ratios.length ? ratios.reduce((x, y) => x + y) / ratios.length : null
      })
      const kern = fam.refit.kern?.[cut] ?? medianKern(faces, raw)
      faces.forEach((face, i) => {
        if (raw[i] === null)
          return report.push(`${fam.family} ${cut} ${face.unicodeRange}: no glyph coverage`)
        const fitted = Math.round(raw[i] * kern * 1000) / 10
        // committed values sit within the 0.1% CSS resolution of the fit
        if (Math.abs(fitted - face.sizeAdjust) <= 0.1 + 1e-9) return
        if (face.pin)
          return report.push(
            `${fam.family} ${cut} ${face.unicodeRange}: pinned ${face.sizeAdjust} (hmtx ${fitted}; ${face.pin.reason})`,
          )
        report.push(`${fam.family} ${cut} ${face.unicodeRange}: ${face.sizeAdjust} -> ${fitted}`)
        if (write) {
          face.sizeAdjust = fitted
          changed++
        }
      })
      report.push(`${fam.family} ${cut}: kern ${kern.toFixed(6)}, ${faces.length} buckets`)
    }
    if (write && changed) {
      writeFileSync(join(ROOT, fam.file), stableJson(stripFile(fam)))
      spawnSync('npx', ['prettier', '--write', fam.file], { cwd: ROOT, stdio: 'inherit' })
    }
  }
  return { report, changed }
}

// ---- generic advance-ratio buckets: one size-adjusted face per group of glyphs
// whose Word advance / stand-in advance lie within refit.span of each other
//
// refit: { method: 'hmtx-buckets', span, range, cuts: { 'weight/style': { word, base, src } } }
// word / base locate a font file: 'dfonts:<file>', 'cloud:<dir>/<full name>',
// 'system:<path>[#ttcIndex]', 'repo:<path>[#ttcIndex]'

const DFONTS_DIR = '/Applications/Microsoft Word.app/Contents/Resources/DFonts'
const CLOUD_DIR = join(
  process.env.HOME || '',
  'Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts',
)

export function locateFont(spec, { dfonts = DFONTS_DIR, cloudFonts = CLOUD_DIR } = {}) {
  const [kind, rest] = [spec.slice(0, spec.indexOf(':')), spec.slice(spec.indexOf(':') + 1)]
  const [path, index] = rest.split('#')
  if (kind === 'dfonts') return readSfnt(join(dfonts, path))
  if (kind === 'system') return readSfnt(path, Number(index ?? 0))
  if (kind === 'repo') return readSfnt(join(ROOT, path), Number(index ?? 0))
  if (kind === 'cloud') {
    const slash = path.indexOf('/')
    const dir = join(cloudFonts, path.slice(0, slash))
    const fullName = path.slice(slash + 1)
    for (const f of readdirSync(dir)) {
      try {
        const font = readSfnt(join(dir, f))
        if (font.fullName === fullName) return font
      } catch {
        /* not a plain sfnt */
      }
    }
    throw new Error(`${fullName} not found under ${dir}`)
  }
  throw new Error(`unknown font locator ${spec}`)
}

const LOWERCASE = codepoints('U+0061-007A')

const formatRange = (cps) => {
  const out = []
  for (const cp of [...cps].sort((a, b) => a - b)) {
    const last = out[out.length - 1]
    if (last && last[1] === cp - 1) last[1] = cp
    else out.push([cp, cp])
  }
  const hex = (n) => n.toString(16).toUpperCase().padStart(4, '0')
  return out.map(([a, b]) => (a === b ? `U+${hex(a)}` : `U+${hex(a)}-${hex(b)}`)).join(', ')
}

const pct1 = (ratio) => Math.round(ratio * 1000) / 10

/**
 * Faces for one cut: the rangeless face at the median lowercase ratio (glyphs
 * outside the fitted range), then one face per greedy bucket of ascending
 * ratios no wider than span, at the bucket's mean ratio
 */
export function fitBuckets(word, base, { span, range, src, weight, style }) {
  const ratios = codepoints(range)
    .map((cp) => {
      const a = word.advance(cp)
      const b = base.advance(cp)
      return a && b ? { cp, r: a / b } : null
    })
    .filter(Boolean)
    .sort((x, y) => x.r - y.r)
  if (!ratios.length) throw new Error('no shared glyphs in range')
  const lower = ratios.filter(({ cp }) => LOWERCASE.includes(cp)).map(({ r }) => r)
  const general = lower.length ? lower[Math.floor(lower.length / 2)] : ratios[0].r
  const buckets = []
  for (const g of ratios) {
    const cur = buckets[buckets.length - 1]
    if (cur && g.r / cur[0].r - 1 <= span) cur.push(g)
    else buckets.push([g])
  }
  const cut = { ...(weight ? { weight } : {}), ...(style ? { style } : {}) }
  const face = (sizeAdjust, unicodeRange) => ({
    src,
    sizeAdjust,
    ...(unicodeRange ? { unicodeRange } : {}),
    ...cut,
  })
  return [
    face(pct1(general)),
    ...buckets.map((b) =>
      face(pct1(b.reduce((s, { r }) => s + r, 0) / b.length), formatRange(b.map(({ cp }) => cp))),
    ),
  ]
}

export function refitBuckets(manifest, { write, only, dirs } = {}) {
  const report = []
  let changed = 0
  for (const fam of manifest.families) {
    if (fam.refit?.method !== 'hmtx-buckets' || (only && fam.family !== only)) continue
    const faces = []
    for (const [cut, spec] of Object.entries(fam.refit.cuts)) {
      const [weight, style] = cut.split('/')
      const fitted = fitBuckets(locateFont(spec.word, dirs), locateFont(spec.base, dirs), {
        span: fam.refit.span,
        range: fam.refit.range,
        src: spec.src,
        weight: weight === 'normal' ? undefined : weight,
        style: style === 'normal' ? undefined : style,
      })
      report.push(
        `${fam.family} ${cut}: ${fitted.length - 1} buckets, general ${fitted[0].sizeAdjust}`,
      )
      faces.push(...fitted)
    }
    const same = JSON.stringify(faces) === JSON.stringify(fam.faces)
    if (same) continue
    report.push(`${fam.family}: faces differ from the manifest`)
    changed++
    if (write) {
      fam.faces = faces
      writeFileSync(join(ROOT, fam.file), stableJson(stripFile(fam)))
      spawnSync('npx', ['prettier', '--write', fam.file], { cwd: ROOT, stdio: 'inherit' })
    }
  }
  return { report, changed }
}

/** fallback kerning factor: median committed/raw ratio of the unpinned buckets */
function medianKern(faces, raw) {
  const ks = faces
    .map((f, i) => (f.pin || raw[i] === null ? null : f.sizeAdjust / 100 / raw[i]))
    .filter((k) => k !== null)
    .sort((a, b) => a - b)
  return ks[Math.floor(ks.length / 2)]
}

const stripFile = ({ file: _file, ...rest }) => rest
const stableJson = (v) =>
  JSON.stringify(v, null, 2).replace(
    /[^\x20-\x7e\n]/g,
    (c) => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'),
  ) + '\n'

export function generate(root = ROOT) {
  const manifest = loadManifest(root)
  const problems = validateManifest(manifest)
  if (problems.length) throw new Error('manifest problems:\n  ' + problems.join('\n  '))
  return {
    [GENERATED_CSS]: renderCss(manifest),
    [GENERATED_TS]: renderTs(manifest),
  }
}

function main(argv) {
  const args = [...argv]
  const flag = (name) => {
    const i = args.indexOf(name)
    if (i === -1) return undefined
    args.splice(i, 1)
    return true
  }
  const option = (name, fallback) => {
    const i = args.indexOf(name)
    if (i === -1) return fallback
    const v = args[i + 1]
    args.splice(i, 2)
    return v
  }
  if (flag('--refit-aptos')) {
    const write = flag('--write') === true
    const dfonts = option('--dfonts', '/Applications/Microsoft Word.app/Contents/Resources/DFonts')
    const cloudFonts = option(
      '--cloud-fonts',
      join(
        process.env.HOME || '',
        'Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts/Aptos Display',
      ),
    )
    const { report, changed } = refitAptos(loadManifest(), { dfonts, cloudFonts, write })
    console.log(report.join('\n'))
    console.log(
      write ? `${changed} bucket(s) updated` : 'dry run; pass --write to update the manifest',
    )
    return 0
  }
  if (flag('--refit-buckets')) {
    const write = flag('--write') === true
    const only = option('--family')
    const { report, changed } = refitBuckets(loadManifest(), { write, only })
    console.log(report.join('\n'))
    console.log(
      write ? `${changed} family(ies) updated` : 'dry run; pass --write to update the manifest',
    )
    return 0
  }
  const check = flag('--check') === true
  if (args.length) {
    console.error(`Unexpected argument: ${args[0]}`)
    return 2
  }
  const outputs = generate()
  let stale = 0
  for (const [rel, text] of Object.entries(outputs)) {
    const path = join(ROOT, rel)
    let current = null
    try {
      current = readFileSync(path, 'utf8')
    } catch {
      /* first generation */
    }
    if (current === text) continue
    stale++
    if (check) console.error(`${rel} is stale; run node tools/build-font-aliases.mjs`)
    else {
      writeFileSync(path, text)
      console.log(`wrote ${rel}`)
    }
  }
  if (!stale) console.log('generated font files are up to date')
  return check && stale ? 1 : 0
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1])
  process.exit(main(process.argv.slice(2)))
