/**
 * Ground-truth metrics for complex scripts (Arabic/Hebrew/Thai/Devanagari): HarfBuzz (wasm)
 * shapes against the actual system fonts to compute advance widths. These scripts have ligatures
 * and contextual forms, so per-glyph accumulation (the opentype path) or heuristic estimation
 * both drift from Chromium's actual rendered width; HarfBuzz shares its shaping engine lineage
 * with Chromium, and with displayFamily forcing the draw side to use the same font family,
 * measurement and drawing match pixel-for-pixel.
 *
 * Initialization is async (wasm compilation); when not ready or fonts are missing this returns
 * null and callers fall back to the existing path.
 */
import { existsSync, readFileSync } from 'node:fs'
// Deep-path import (the package exports only expose an ESM wrapper with a top-level await,
// which the CJS main process cannot bundle)
import createHarfBuzz from '../../../../node_modules/harfbuzzjs/dist/harfbuzz.js'
import hbWasmPath from '../../../../node_modules/harfbuzzjs/dist/harfbuzz.wasm?asset'

type Script = 'arabic' | 'hebrew' | 'thai' | 'devanagari'

const SCRIPT_RES: Array<[Script, RegExp]> = [
  ['arabic', /[\u0600-\u06ff\u0750-\u077f\u08a0-\u08ff\ufb50-\ufdff\ufe70-\ufeff]/],
  ['hebrew', /[\u0590-\u05ff\ufb1d-\ufb4f]/],
  ['thai', /[\u0e00-\u0e7f]/],
  ['devanagari', /[\u0900-\u097f]/],
]

export function complexScriptOf(text: string): Script | null {
  for (const [script, re] of SCRIPT_RES) {
    if (re.test(text)) return script
  }
  return null
}

/** Per-platform script -> (font file, CSS family name, coverage-check sample char). The family name must be resolvable by Chromium by name. */
const FONT_TABLE: Record<
  string,
  Array<{ script: Script; file: string; family: string; sample: string }>
> = {
  darwin: [
    {
      script: 'arabic',
      file: '/System/Library/Fonts/GeezaPro.ttc',
      family: 'Geeza Pro',
      sample: 'ا',
    },
    {
      script: 'hebrew',
      file: '/System/Library/Fonts/Supplemental/Tahoma.ttf',
      family: 'Tahoma',
      sample: 'א',
    },
    {
      script: 'thai',
      file: '/System/Library/Fonts/Supplemental/Thonburi.ttc',
      family: 'Thonburi',
      sample: 'ก',
    },
    {
      script: 'devanagari',
      file: '/System/Library/Fonts/Supplemental/DevanagariMT.ttc',
      family: 'Devanagari MT',
      sample: 'क',
    },
  ],
  win32: [
    { script: 'arabic', file: 'C:\\Windows\\Fonts\\tahoma.ttf', family: 'Tahoma', sample: 'ا' },
    { script: 'hebrew', file: 'C:\\Windows\\Fonts\\tahoma.ttf', family: 'Tahoma', sample: 'א' },
    {
      script: 'thai',
      file: 'C:\\Windows\\Fonts\\leelawui.ttf',
      family: 'Leelawadee UI',
      sample: 'ก',
    },
    {
      script: 'devanagari',
      file: 'C:\\Windows\\Fonts\\Nirmala.ttf',
      family: 'Nirmala UI',
      sample: 'क',
    },
  ],
}

interface HbModule {
  wasmExports: Record<string, (...args: number[]) => number>
  HEAPU8: Uint8Array
  HEAPU16: Uint16Array
  HEAP32: Int32Array
  HEAPU32: Uint32Array
}

interface LoadedFont {
  fontPtr: number
  upem: number
  family: string
}

let hb: HbModule | null = null
const fonts = new Map<Script, LoadedFont>()

const SCRIPT_SAMPLES: Record<Script, string> = {
  arabic: 'ا',
  hebrew: 'א',
  thai: 'ก',
  devanagari: 'क',
}

/** A resolved face for the REQUESTED family (bytes lazily extracted): decks ship
 * with real complex-script fonts (Office CloudFonts/DFonts/embedded) — shaping
 * and drawing must use them, not the script's generic substitute. */
export interface ShapedPrefFace {
  family: string
  bytes: () => ArrayBuffer | null
}

/** On-demand faces for resolved request fonts; null = tried and unusable */
const prefFonts = new Map<string, LoadedFont | null>()

function loadHbFont(bytes: Uint8Array, family: string): LoadedFont | null {
  if (!hb) return null
  const ex = hb.wasmExports
  const ptr = ex.malloc!(bytes.length)
  hb.HEAPU8.set(bytes, ptr)
  const blob = ex.hb_blob_create!(ptr, bytes.length, 2, 0, 0)
  const face = ex.hb_face_create!(blob, 0)
  const upem = ex.hb_face_get_upem!(face)
  if (!upem) return null
  return { fontPtr: ex.hb_font_create!(face), upem, family }
}

function loadPrefFont(script: Script, pref: ShapedPrefFace, styleKey: string): LoadedFont | null {
  // Not initialized yet: do NOT cache the miss, or the face is skipped forever
  if (!hb) return null
  // Coverage is per script — one family can cover Arabic but not Thai
  const key = `${script}|${pref.family}|${styleKey}`
  const hit = prefFonts.get(key)
  if (hit !== undefined) return hit
  let loaded: LoadedFont | null = null
  try {
    const ab = pref.bytes()
    if (ab) {
      const cand = loadHbFont(new Uint8Array(ab), pref.family)
      // Coverage check: a resolved face without this script's glyphs falls back
      // to the generic substitute (e.g. a Latin-only face on Arabic text)
      if (cand && shapeUnits(cand, SCRIPT_SAMPLES[script], true) != null) loaded = cand
    }
  } catch {
    /* unusable face: remember the miss */
  }
  prefFonts.set(key, loaded)
  return loaded
}

/** `${family}|${style}|${text}` -> advance width (font units), process-level cache */
const unitsCache = new Map<string, number>()
/**
 * Renderer-process ground-truth widths (@100px, scaled linearly). HarfBuzz estimates still
 * deviate from Chromium's actual rendering (AAT/morx fonts, font matching cannot be replicated),
 * so the page's measureText is the final authority: the first layout pass records misses ->
 * refineComplexWidths measures them in batch -> the second layout pass hits the cache entirely.
 */
const gtCache = new Map<string, number>()
const gtPending = new Set<string>()
const gtKey = (text: string, family: string, bold: boolean, italic: boolean) =>
  `${bold ? 'b' : ''}${italic ? 'i' : ''}|${family}|${text}`
let started = false
let readyPromise: Promise<void> = Promise.resolve()

/**
 * Await before the first layout so complex-script metrics do not fall back to estimation due to
 * an init race (opening a file lays out immediately, while wasm compilation + font loading are
 * async).
 */
export function shapedMetricsReady(): Promise<void> {
  initShapedMetrics()
  // Init errors are already swallowed internally; the 3s cap prevents a stuck font read from blocking file open
  return Promise.race([readyPromise, new Promise<void>((r) => setTimeout(r, 3000))])
}

/** Called at startup (triggered when the fonts.ts module loads); wasm compiles in milliseconds, ready before the first file opens. */
export function initShapedMetrics(): void {
  if (started) return
  started = true
  readyPromise = (async () => {
    try {
      const wasmBinary = readFileSync(hbWasmPath)
      // Types come from the package's d.ts (0-arg + narrowed Module surface); it is really an
      // Emscripten factory: accepts wasmBinary and exposes HEAP views
      const factory = createHarfBuzz as unknown as (
        arg?: Record<string, unknown>,
      ) => Promise<unknown>
      const mod = (await factory({ wasmBinary })) as HbModule
      const ex = mod.wasmExports
      for (const entry of FONT_TABLE[process.platform] ?? []) {
        if (fonts.has(entry.script) || !existsSync(entry.file)) continue
        try {
          const bytes = readFileSync(entry.file)
          const ptr = ex.malloc!(bytes.length)
          mod.HEAPU8.set(bytes, ptr)
          // HB_MEMORY_MODE_WRITABLE, no destroy callback: fonts live for the process lifetime, never freed
          const blob = ex.hb_blob_create!(ptr, bytes.length, 2, 0, 0)
          const face = ex.hb_face_create!(blob, 0)
          const upem = ex.hb_face_get_upem!(face)
          const fontPtr = ex.hb_font_create!(face)
          const loaded: LoadedFont = { fontPtr, upem, family: entry.family }
          // Coverage check: if the sample char shapes to glyph 0 (.notdef), the font lacks this script — discard it
          hb = mod
          if (shapeUnits(loaded, entry.sample, true) != null) fonts.set(entry.script, loaded)
        } catch {
          /* One font failing does not affect other scripts */
        }
      }
      hb = mod
    } catch {
      hb = null // wasm init failed -> everything takes the original path
    }
  })()
}

/** Shape to get advance width (font units). With checkCoverage, returns null if any glyph is .notdef. */
function shapeUnits(font: LoadedFont, text: string, checkCoverage = false): number | null {
  if (!hb) return null
  const ex = hb.wasmExports
  const buf = ex.hb_buffer_create!()
  try {
    const strPtr = ex.malloc!(text.length * 2)
    for (let i = 0; i < text.length; i++) hb.HEAPU16[strPtr / 2 + i] = text.charCodeAt(i)
    ex.hb_buffer_add_utf16!(buf, strPtr, text.length, 0, text.length)
    ex.free!(strPtr)
    ex.hb_buffer_guess_segment_properties!(buf)
    ex.hb_shape!(font.fontPtr, buf, 0, 0)
    const n = ex.hb_buffer_get_length!(buf)
    if (checkCoverage) {
      const infoPtr = ex.hb_buffer_get_glyph_infos!(buf, 0) / 4
      for (let i = 0; i < n; i++) {
        if (hb.HEAPU32[infoPtr + i * 5] === 0) return null
      }
    }
    const posPtr = ex.hb_buffer_get_glyph_positions!(buf, 0) / 4
    let units = 0
    for (let i = 0; i < n; i++) units += hb.HEAP32[posPtr + i * 5]!
    return units
  } finally {
    ex.hb_buffer_destroy!(buf)
  }
}

/** Width (px) of complex-script text: prefer the renderer's measured cache; on miss, record it and return the HarfBuzz estimate; null if not applicable. */
export function shapedMeasure(
  text: string,
  fontSizePx: number,
  bold = false,
  italic = false,
  pref?: ShapedPrefFace,
): number | null {
  const script = complexScriptOf(text)
  if (!script) return null
  const styleKey = `${bold ? 'b' : ''}${italic ? 'i' : ''}`
  const font = (pref ? loadPrefFont(script, pref, styleKey) : null) ?? fonts.get(script)
  if (!font) return null
  const key = gtKey(text, font.family, bold, italic)
  const gt = gtCache.get(key)
  if (gt != null) return (gt / 100) * fontSizePx
  gtPending.add(key)
  const unitsKey = `${font.family}|${styleKey}|${text}`
  let units = unitsCache.get(unitsKey)
  if (units == null) {
    const u = shapeUnits(font, text)
    if (u == null) return null
    unitsCache.set(unitsKey, u)
    units = u
  }
  return (units / font.upem) * fontSizePx
}

/**
 * Send complex-script strings that missed during layout to the renderer's measureText (same
 * engine measures and draws) and backfill the cache. Returns true when new widths were stored,
 * meaning the caller should rebuild the layout. If the page is not ready or the call fails,
 * keep them pending and retry next time.
 */
export async function refineComplexWidths(wc: {
  executeJavaScript(code: string): Promise<unknown>
}): Promise<boolean> {
  if (!gtPending.size) return false
  const items = [...gtPending]
  gtPending.clear()
  const payload = items.map((key) => {
    const [flags = '', family = '', ...rest] = key.split('|')
    return {
      key,
      family,
      text: rest.join('|'),
      bold: flags.includes('b'),
      italic: flags.includes('i'),
    }
  })
  try {
    const widths = (await wc.executeJavaScript(`(async () => {
      const items = ${JSON.stringify(payload)}
      // Ground truth requires the REAL faces: layout may resolve to Office-private
      // fonts Chromium cannot see by name, and this refinement can run before
      // doc-fonts syncs them — measuring the fallback would bake wrong widths in
      const need = new Set(items.map((it) => it.family))
      let privateFams = new Set()
      try {
        const faces = await window.slidesApi.privateFontFaces()
        privateFams = new Set((Array.isArray(faces) ? faces : []).map((f) => f.family))
        for (const f of Array.isArray(faces) ? faces : []) {
          if (!need.has(f.family)) continue
          // Each face registers independently: one bad face must not skip the rest
          try {
            const weight = f.bold ? '700' : '400'
            const style = f.italic ? 'italic' : 'normal'
            let have = false
            document.fonts.forEach((x) => {
              if (x.family === f.family && x.weight === weight && x.style === style) have = true
            })
            if (have) continue
            const data = await window.slidesApi.privateFontData(f.id)
            if (!data) continue
            const face = new FontFace(f.family, data, { weight, style })
            document.fonts.add(face)
            await face.load()
          } catch { /* this face keeps the fallback; check() below skips it */ }
        }
      } catch { /* no private-face listing: system fonts still pass check() */ }
      // fonts.check() returns true for unknown families on modern Chromium
      // (fallback counts as renderable), so it cannot gate this. A family is
      // measurable when it is NOT a private face (main resolved it to a system
      // font Chromium sees by name) or any face of it is registered — missing
      // styles are synthesized identically by canvas drawing and measureText.
      const registered = new Set()
      document.fonts.forEach((x) => registered.add(x.family))
      const ctx = document.createElement('canvas').getContext('2d')
      const out = {}
      for (const it of items) {
        if (privateFams.has(it.family) && !registered.has(it.family)) {
          out[it.key] = null // family not available yet — stays pending
          continue
        }
        // Same font stack as the Konva renderer (generic fallback of konva-adapter displayFontFamily)
        ctx.font = (it.italic ? 'italic ' : '') + (it.bold ? 'bold ' : '') +
          '100px "' + it.family + '", \\'PingFang SC\\', \\'Microsoft YaHei\\', sans-serif'
        out[it.key] = ctx.measureText(it.text).width
      }
      return out
    })()`)) as Record<string, number>
    let added = false
    for (const [k, w] of Object.entries(widths)) {
      if (typeof w === 'number' && Number.isFinite(w)) {
        gtCache.set(k, w)
        added = true
      } else {
        gtPending.add(k) // family didn't resolve yet — retry next refine
      }
    }
    return added
  } catch {
    for (const it of items) gtPending.add(it)
    return false
  }
}

/** Draw font family for complex-script text (same font file as shaping); null if not applicable. */
/**
 * Renderer ground-truth width for non-complex text (same gtCache/refine batch as the
 * complex-script channel). Used for faces whose opentype-side advances drift from
 * Chromium's rendering (variable fonts: HVAR/instancing interpretations differ by a
 * few percent, which visually swallows word spaces). Returns null on a cache miss and
 * queues the string for refineComplexWidths; the second layout pass hits the cache.
 */
export function gtMeasure(
  text: string,
  family: string,
  fontSizePx: number,
  bold = false,
  italic = false,
): number | null {
  const key = gtKey(text, family, bold, italic)
  const gt = gtCache.get(key)
  if (gt != null) return (gt / 100) * fontSizePx
  gtPending.add(key)
  return null
}

export function shapedFamily(
  text: string,
  pref?: ShapedPrefFace,
  bold = false,
  italic = false,
): string | null {
  const script = complexScriptOf(text)
  if (!script) return null
  if (pref) {
    const f = loadPrefFont(script, pref, `${bold ? 'b' : ''}${italic ? 'i' : ''}`)
    if (f) return f.family
  }
  return fonts.get(script)?.family ?? null
}
