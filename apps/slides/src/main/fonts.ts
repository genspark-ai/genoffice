/**
 * System font metrics (main process) — parse real font files with opentype.js and inject them
 * into pptx-render's OpentypeMetrics, replacing heuristic estimation for accurate line
 * wrapping/centering.
 *
 * Strategy (pragmatic):
 *   - At startup only scan directories to build a "normalized filename -> path" index
 *     (no font parsing; ~ms cost);
 *   - Lazily parse and cache the matching file only when requested by (family, bold, italic);
 *   - .ttc collection fonts (nearly all CJK fonts: macOS Hiragino, Windows Yu Gothic, etc.)
 *     are unsupported by opentype.js; here we read the name table to pick a face by requested
 *     family, split it into a standalone sfnt via the offset table, then parse;
 *   - Filename miss -> aliases/style-less variants -> substitute by script
 *     (ja/ko/traditional-zh/serif/mono) with fonts guaranteed on this platform -> if all miss,
 *     return undefined (callers use heuristic metrics).
 */
import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { homedir, tmpdir } from 'node:os'
import { createHash } from 'node:crypto'
import * as opentype from 'opentype.js'
import type { EmbeddedFontFace } from '@genoffice/pptx-engine'
import {
  OpentypeMetrics,
  HeuristicMetrics,
  type FontMetricsProvider,
  type OpentypeFontLike,
  type RunStyle,
} from '@genoffice/pptx-render'
import { classifyCjkScript } from '../shared/cjk-script'
import {
  initShapedMetrics,
  shapedMeasure,
  shapedFamily,
  complexScriptOf,
  type ShapedPrefFace,
  gtMeasure,
} from './shaped-metrics'
import carlitoRegular from '@genoffice/ui/fonts/Carlito-Regular.ttf?asset'
import carlitoBold from '@genoffice/ui/fonts/Carlito-Bold.ttf?asset'
import carlitoItalic from '@genoffice/ui/fonts/Carlito-Italic.ttf?asset'
import carlitoBoldItalic from '@genoffice/ui/fonts/Carlito-BoldItalic.ttf?asset'

/** Fonts shipped with the app (metric substitutes for fonts most decks assume, e.g. Calibri→Carlito). */
const BUNDLED_FONTS: Record<string, string> = {
  'Carlito-Regular': carlitoRegular,
  'Carlito-Bold': carlitoBold,
  'Carlito-Italic': carlitoItalic,
  'Carlito-BoldItalic': carlitoBoldItalic,
}

/**
 * User font store (downloaded catalog fonts + fonts installed via the in-app entry).
 * Injected by slides-main at startup (fonts.ts stays free of electron imports); invisible
 * to Chromium, so faces resolved from here get the same private FontFace treatment as
 * Office DFonts.
 */
let userFontDir: string | null = null
export function setUserFontDir(dir: string): void {
  userFontDir = dir
}
export function getUserFontDir(): string | null {
  return userFontDir
}

/**
 * Document-embedded fonts (<p:embeddedFontLst> fntdata, uncompressed EOT payloads only).
 * PowerPoint renders missing families with the embedded faces, so they slot into resolution
 * between an exact system hit and the alias/substitute chain. Extracted sfnts are cached on
 * disk by content hash; the dir is private (Chromium can't see it), so drawing goes through
 * the same private FontFace channel as Office DFonts. Registrations are process-wide and
 * live until the app exits — like PowerPoint keeping embedded fonts while the deck is open.
 */
const EMBEDDED_FONT_DIR = join(tmpdir(), 'genoffice-embedded-fonts')
/** norm(typeface) -> styleKey ('<bold><italic>') -> extracted sfnt path */
const embeddedFaces = new Map<string, Map<string, string>>()

/** Register a deck's embedded faces; true when anything new was added (metrics must reset). */
export function registerEmbeddedFonts(faces: EmbeddedFontFace[]): boolean {
  let added = false
  for (const f of faces) {
    const key = norm(f.typeface)
    if (!key) continue
    const styleKey =
      (f.style === 'bold' || f.style === 'boldItalic' ? '1' : '0') +
      (f.style === 'italic' || f.style === 'boldItalic' ? '1' : '0')
    let perStyle = embeddedFaces.get(key)
    if (!perStyle) {
      perStyle = new Map()
      embeddedFaces.set(key, perStyle)
    }
    if (perStyle.has(styleKey)) continue
    const path = join(
      EMBEDDED_FONT_DIR,
      `${createHash('sha256').update(f.sfnt).digest('hex').slice(0, 32)}.ttf`,
    )
    try {
      mkdirSync(EMBEDDED_FONT_DIR, { recursive: true })
      if (!existsSync(path)) writeFileSync(path, f.sfnt)
    } catch {
      continue
    }
    perStyle.set(styleKey, path)
    added = true
  }
  return added
}

function fontDirs(): string[] {
  switch (process.platform) {
    case 'darwin':
      return [
        '/System/Library/Fonts',
        '/System/Library/Fonts/Supplemental',
        '/Library/Fonts',
        join(homedir(), 'Library/Fonts'),
      ]
    case 'win32':
      return ['C:\\Windows\\Fonts', join(homedir(), 'AppData/Local/Microsoft/Windows/Fonts')]
    default:
      return ['/usr/share/fonts', '/usr/local/share/fonts', join(homedir(), '.fonts')]
  }
}

/**
 * Office-private font dirs. PowerPoint for Mac bundles the Windows core fonts (real
 * Calibri/YaHei/Verdana…) inside the app; PowerPoint renders with them, so metrics must
 * too or every substituted family drifts from the reference. Chromium cannot resolve
 * these by name — faces resolved from here are marked private and their bytes are served
 * to the renderer for FontFace registration (same file measures and draws).
 */
function officeFontDirs(): string[] {
  if (process.platform !== 'darwin') return []
  return ['Microsoft PowerPoint', 'Microsoft Word', 'Microsoft Excel'].map((app) =>
    join('/Applications', `${app}.app`, 'Contents/Resources/DFonts'),
  )
}

const APPLE_FONT_ASSET_ROOT = '/System/Library/AssetsV2/com_apple_MobileAsset_Font7'
const APPLE_FONT_SUBSETS =
  '/System/Library/PrivateFrameworks/FontServices.framework/Versions/A/Resources/Fonts/Subsets'

/**
 * macOS on-demand font assets (CoreText downloadable fonts, e.g. NanumGothic): PowerPoint
 * renders with them but Chromium cannot resolve them by name — same private treatment as
 * Office DFonts. Materialized full downloads (AssetsV2) take precedence over the built-in
 * stub subsets; both are read-only system paths and simply absent on other platforms.
 */
function appleFontAssetDirs(): string[] {
  if (process.platform !== 'darwin') return []
  const dirs: string[] = []
  try {
    for (const d of readdirSync(APPLE_FONT_ASSET_ROOT))
      dirs.push(join(APPLE_FONT_ASSET_ROOT, d, 'AssetData'))
  } catch {
    /* asset root absent */
  }
  dirs.push(APPLE_FONT_SUBSETS)
  return dirs
}

/** Office cloud-font roots: <root>/<Family Name>/<numeric-id>.ttf — indexed by directory name. */
function cloudFontRoots(): string[] {
  const globDirs = (base: string, sub: string): string[] => {
    try {
      return readdirSync(base).map((d) => join(base, d, sub))
    } catch {
      return []
    }
  }
  switch (process.platform) {
    case 'darwin':
      return globDirs(
        join(homedir(), 'Library/Group Containers/UBF8T346G9.Office/FontCache'),
        'CloudFonts',
      )
    case 'win32':
      return globDirs(join(homedir(), 'AppData/Local/Microsoft/FontCache'), 'CloudFonts')
    default:
      return []
  }
}

/** Normalize: NFKC (full-width MS -> MS), lowercase, strip spaces/hyphens/underscores. */
function norm(s: string): string {
  return s
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s\-_]/g, '')
}

// Japanese fonts: Windows families/filenames and macOS Hiragino serve as each other's fallback
const YU_GOTHIC = ['YuGothM', 'YuGothB', 'YuGothR', 'YuGothic']
const YU_MINCHO = ['YuMin', 'YuMinDB', 'YuMincho']
const MEIRYO = ['Meiryo']
const MS_GOTHIC = ['MSGothic']
const MS_MINCHO = ['MSMincho']
// Base name first: combined with style suffixes (bold -> w6 / regular -> w3) to pick the file
// by weight; the full name with W3 is the fallback (index key = normalized filename)
const HIRAGINO_SANS = ['ヒラギノ角ゴシック', 'ヒラギノ角ゴシック W3']
const HIRAGINO_MINCHO = ['ヒラギノ明朝 ProN']

/** Common font-name aliases (localized name <-> English family <-> actual filename); keys match via norm. */
const ALIASES: Record<string, string[]> = {
  宋体: ['SimSun', 'Songti'],
  黑体: ['SimHei', 'Heiti SC'],
  微软雅黑: ['Microsoft YaHei', 'MSYH'],
  楷体: ['KaiTi', 'Kaiti SC'],
  仿宋: ['FangSong', 'STFangsong'],
  helvetica: ['Arial'],
  'helvetica neue': ['Arial'],
  calibri: ['Carlito', 'Arial'],
  'calibri light': ['Carlito', 'Arial'],
  // PowerPoint for Mac substitutes the Windows-only Lucida Sans family with Lucida Grande
  'lucida sans unicode': ['Lucida Grande', 'Arial'],
  'lucida sans': ['Lucida Grande', 'Arial'],
  // —— Japanese ——
  'yu gothic': [...YU_GOTHIC, ...HIRAGINO_SANS],
  // Yu Gothic UI faces live inside YuGothM/YuGothB.ttc (like Meiryo UI in meiryo.ttc);
  // the origKey bonus in rankFaces picks them out of the alias target's collection
  'yu gothic ui': [...YU_GOTHIC, ...HIRAGINO_SANS],
  游ゴシック: [...YU_GOTHIC, ...HIRAGINO_SANS],
  游ゴシック体: [...YU_GOTHIC, ...HIRAGINO_SANS],
  'yu mincho': [...YU_MINCHO, ...HIRAGINO_MINCHO],
  游明朝: [...YU_MINCHO, ...HIRAGINO_MINCHO],
  meiryo: [...MEIRYO, ...HIRAGINO_SANS],
  メイリオ: [...MEIRYO, ...HIRAGINO_SANS],
  // Meiryo UI is a face inside meiryo.ttc/meiryob.ttc (narrower kana than Meiryo); the
  // origKey bonus in rankFaces picks the UI face out of the alias target's collection
  'meiryo ui': [...MEIRYO, ...HIRAGINO_SANS],
  'ms gothic': [...MS_GOTHIC, ...HIRAGINO_SANS],
  'ms pgothic': [...MS_GOTHIC, ...HIRAGINO_SANS],
  'ms ui gothic': [...MS_GOTHIC, ...HIRAGINO_SANS],
  'ms ゴシック': [...MS_GOTHIC, ...HIRAGINO_SANS],
  'ms pゴシック': [...MS_GOTHIC, ...HIRAGINO_SANS],
  'ms mincho': [...MS_MINCHO, ...HIRAGINO_MINCHO],
  'ms pmincho': [...MS_MINCHO, ...HIRAGINO_MINCHO],
  'ms 明朝': [...MS_MINCHO, ...HIRAGINO_MINCHO],
  'ms p明朝': [...MS_MINCHO, ...HIRAGINO_MINCHO],
  'hiragino sans': HIRAGINO_SANS,
  'hiragino kaku gothic pron': HIRAGINO_SANS,
  'hiragino kaku gothic pro': HIRAGINO_SANS,
  ヒラギノ角ゴシック: HIRAGINO_SANS,
  'hiragino mincho pron': HIRAGINO_MINCHO,
  'hiragino mincho pro': HIRAGINO_MINCHO,
  ヒラギノ明朝: HIRAGINO_MINCHO,
  // —— Korean ——
  'malgun gothic': ['Malgun', 'MalgunBD'],
  '맑은 고딕': ['Malgun Gothic', 'Malgun'],
  바탕: ['Batang'],
  바탕체: ['Batang'],
  batangche: ['Batang'],
  gungsuh: ['Batang'],
  궁서: ['Batang'],
  굴림: ['Gulim'],
  gulimche: ['Gulim'],
  dotum: ['Gulim'],
  돋움: ['Gulim'],
  // Hangul-localized Nanum names (Google-Slides exports carry both spellings in one deck;
  // the alias keeps them on the same downloadable NanumGothic asset instead of a substitute)
  나눔고딕: ['NanumGothic'],
  나눔바른고딕: ['NanumBarunGothic', 'NanumGothic'],
  나눔명조: ['NanumMyeongjo'],
  나눔스퀘어: ['NanumSquare', 'NanumGothic'],
  // —— Traditional Chinese ——
  'microsoft jhenghei': ['MSJH'],
  微軟正黑體: ['Microsoft JhengHei', 'MSJH'],
  pmingliu: ['MingLiU'],
  新細明體: ['PMingLiU', 'MingLiU'],
  細明體: ['MingLiU'],
  'dfkai-sb': ['KaiU', 'BiauKai'],
  標楷體: ['DFKai-SB', 'KaiU', 'BiauKai'],
  'pingfang tc': ['PingFang'],
  'pingfang sc': ['PingFang'],
  'heiti tc': ['STHeiti Light', 'STHeiti Medium'],
  'heiti sc': ['STHeiti Light', 'STHeiti Medium'],
  'songti tc': ['Songti'],
  'songti sc': ['Songti'],
  宋體: ['Songti', 'PMingLiU', 'MingLiU'],
}
const ALIAS_MAP = new Map(Object.entries(ALIASES).map(([k, v]) => [norm(k), v]))
const aliasesOf = (family: string): string[] => ALIAS_MAP.get(norm(family)) ?? []

/**
 * Substitution for missing fonts (aligned with PowerPoint's font substitution): first classify
 * script by family name (ja/ko/traditional-zh) and substitute a same-script font guaranteed on
 * this platform; non-CJK substitutes by serif/mono/sans class. The key point is that the
 * substitution result is used for both measuring and drawing (displayFamily is passed through
 * to the renderer), so both sides always use the same font file — otherwise the width gap
 * between "heuristic estimate + browser-chosen fallback drawing" pushes later runs onto
 * earlier text.
 */
const SERIF_RE =
  /serif|roman|garamond|georgia|playfair|didot|bodoni|baskerville|caslon|palatino|antiqua|minion|lora|merriweather|crimson|spectral|charter|literata|song|songti|宋|mincho|明朝|ming|batang|바탕|myeongjo|명조|gungsuh|궁서|細明|標楷|儷宋/i
const MONO_RE = /mono|courier|consolas|menlo|monaco|code|typewriter/i

// PowerPoint for Mac substitutes EVERY unresolvable family with Calibri regardless of its
// apparent class — probe decks with fake serif ("Qqzgaramond") / mono ("Zxqvwt Mono Courier")
// names all export as Calibri (pdffonts on problem/weight-suffix-probe*.pptx). The class
// chains stay as fallbacks for machines without Office fonts or bundled Carlito.
const SUBSTITUTES: Record<'serif' | 'sans' | 'mono', string[]> = {
  serif: ['Calibri', 'Georgia', 'Times New Roman'],
  sans: ['Calibri', 'Arial', 'Verdana'],
  mono: ['Calibri', 'Courier New'],
}

function classifyFamily(family: string): 'serif' | 'sans' | 'mono' {
  if (MONO_RE.test(family)) return 'mono'
  if (SERIF_RE.test(family)) return 'serif'
  return 'sans'
}

/**
 * Trailing weight token of a "Family Weight" sub-family request ("Apercu Light",
 * "페이퍼로지 4 Regular", "Montserrat SemiBold"). PowerPoint resolves these through the base
 * family's faces when the full name has no file of its own (probe: "Avenir Light" renders
 * as the real Avenir-Light face). bold/italic stay out — those arrive as rPr flags.
 */
const WEIGHT_SUFFIX_RE =
  /\s+(?:\d+\s+)?(?:(?:ultra|extra|semi|demi)[- ]?(?:light|bold)|thin|hairline|extralight|ultralight|light|semilight|book|roman|regular|normal|medium|semibold|demibold|extrabold|ultrabold|black|heavy)$/i

function baseFamilyOf(family: string): string | undefined {
  const m = WEIGHT_SUFFIX_RE.exec(family.trim())
  const base = m ? family.trim().slice(0, m.index).trim() : ''
  return base || undefined
}

/** OS/2 usWeightClass the request's trailing weight token asks for (400 = regular). */
const WEIGHT_TARGETS: Array<[RegExp, number]> = [
  [/thin|hairline/i, 100],
  [/(?:ultra|extra)[- ]?light/i, 200],
  [/semi[- ]?light/i, 350],
  [/light/i, 300],
  [/book|roman|regular|normal/i, 400],
  [/medium/i, 500],
  [/(?:semi|demi)[- ]?bold/i, 600],
  [/(?:ultra|extra)[- ]?bold/i, 800],
  [/black|heavy/i, 900],
]
function weightTargetOf(family: string): number | undefined {
  const m = WEIGHT_SUFFIX_RE.exec(family.trim())
  if (!m) return undefined
  const suffix = family.trim().slice(m.index)
  for (const [re, w] of WEIGHT_TARGETS) if (re.test(suffix)) return w
  return undefined
}

// PowerPoint substitutes a missing font by the run's declared language/charset, not by
// classifying the font name (see TextRun.fontScriptHint): prod_079's JP-named font with
// charset=134 renders with Microsoft YaHei; prod_043's altLang="ko-KR" runs get Malgun.
function substitutesFor(family: string, substScript?: 'ja' | 'ko' | 'sc' | 'tc'): string[] {
  const script = substScript ?? classifyCjkScript(family)
  if (!script) return SUBSTITUTES[classifyFamily(family)]
  const serif = SERIF_RE.test(family)
  const mac = process.platform === 'darwin'
  switch (script) {
    case 'sc':
      // Office-bundled YaHei first (msyh.ttc in DFonts) to mirror PPT's pick
      return serif
        ? mac
          ? ['Songti SC', 'SimSun']
          : ['SimSun']
        : mac
          ? ['Microsoft YaHei', 'PingFang SC', 'Heiti SC']
          : ['Microsoft YaHei', 'DengXian', 'SimSun']
    case 'ja':
      return serif
        ? mac
          ? ['Hiragino Mincho ProN']
          : ['Yu Mincho', 'MS Mincho']
        : mac
          ? ['Hiragino Sans']
          : ['Yu Gothic', 'Meiryo', 'MS Gothic']
    case 'ko':
      // mac chains start with the Office-bundled faces to mirror the renderer's KO_SANS/
      // KO_SERIF draw chains — measuring Apple SD Gothic Neo while drawing the private
      // Malgun FontFace swallowed word spaces on unknown KR vendor fonts
      return serif
        ? mac
          ? ['Batang', 'AppleMyungjo', 'Apple SD Gothic Neo']
          : ['Batang', 'Malgun Gothic']
        : mac
          ? ['Malgun Gothic', 'Apple SD Gothic Neo', 'AppleGothic']
          : ['Malgun Gothic', 'Gulim']
    case 'tc':
      return serif
        ? mac
          ? ['Songti TC']
          : ['PMingLiU', 'Microsoft JhengHei']
        : mac
          ? ['PingFang TC', 'Heiti TC', 'Songti TC']
          : ['Microsoft JhengHei']
  }
}

/** Metadata for one face in a ttc/ttf (name table), used to pick a face by requested family/style. */
interface FaceInfo {
  /** Position of the offset table within the file (0 for non-ttc) */
  offset: number
  /** Family name for drawing: prefer the ASCII English name (resolvable by CSS by name) */
  display: string
  /** Normalized set of family names (name 1/16, including localized names) */
  famKeys: string[]
  /** Family + subfamily concatenation (normalized), used for style picks like bold/W6 */
  styleText: string
  /** OS/2 usWeightClass (name-independent weight evidence for ranking) */
  weight?: number
}

function readNameStrings(
  buf: Buffer,
  nameOff: number,
): { families: string[]; subfamilies: string[] } {
  const families: string[] = []
  const subfamilies: string[] = []
  const count = buf.readUInt16BE(nameOff + 2)
  const strBase = nameOff + buf.readUInt16BE(nameOff + 4)
  for (let i = 0; i < count; i++) {
    const r = nameOff + 6 + 12 * i
    const platform = buf.readUInt16BE(r)
    const encoding = buf.readUInt16BE(r + 2)
    const nameId = buf.readUInt16BE(r + 6)
    if (nameId !== 1 && nameId !== 2 && nameId !== 16 && nameId !== 17) continue
    const len = buf.readUInt16BE(r + 8)
    const off = strBase + buf.readUInt16BE(r + 10)
    if (off + len > buf.length) continue
    let s: string
    if (platform === 0 || platform === 3) {
      s = Buffer.from(buf.subarray(off, off + len))
        .swap16()
        .toString('utf16le')
    } else if (platform === 1 && encoding === 0) {
      s = buf.toString('latin1', off, off + len)
    } else {
      continue // Mac-platform non-Roman encodings (legacy Korean/Chinese codepages) cannot be decoded; skip
    }
    if (!s) continue
    const list = nameId === 1 || nameId === 16 ? families : subfamilies
    if (!list.includes(s)) list.push(s)
  }
  return { families, subfamilies }
}

function readFaceDir(buf: Buffer): FaceInfo[] {
  const offsets =
    buf.toString('ascii', 0, 4) === 'ttcf'
      ? Array.from({ length: buf.readUInt32BE(8) }, (_, i) => buf.readUInt32BE(12 + 4 * i))
      : [0]
  return offsets.map((offset) => {
    let families: string[] = []
    let subfamilies: string[] = []
    let weight: number | undefined
    try {
      const numTables = buf.readUInt16BE(offset + 4)
      for (let t = 0; t < numTables; t++) {
        const e = offset + 12 + 16 * t
        const tag = buf.toString('ascii', e, e + 4)
        if (tag === 'name') {
          ;({ families, subfamilies } = readNameStrings(buf, buf.readUInt32BE(e + 8)))
        } else if (tag === 'OS/2') {
          // usWeightClass: name records can be missing/undecodable (cloud numeric files),
          // so face ranking needs the weight straight from the table
          const w = buf.readUInt16BE(buf.readUInt32BE(e + 8) + 4)
          if (w >= 1 && w <= 1000) weight = w
        }
      }
    } catch {
      /* Even if the name table is unreadable, the first face can still be parsed */
    }
    const ascii = families.find((f) => /^[\x20-\x7e]+$/.test(f))
    return {
      offset,
      display: ascii ?? families[0] ?? '',
      famKeys: families.map(norm),
      styleText: norm([...families, ...subfamilies].join(' ')),
      ...(weight != null ? { weight } : {}),
    }
  })
}

/** OpenType layout tables — droppable for metrics-only parsing when opentype.js rejects them. */
const LAYOUT_TABLES: ReadonlySet<string> = new Set(['GSUB', 'GPOS', 'GDEF'])

/** Extract a single face from a ttc into a standalone sfnt (rewrite the table directory, copy table data by original offset). */
function extractFace(buf: Buffer, offset: number, drop?: ReadonlySet<string>): ArrayBuffer {
  const isTtc = buf.toString('ascii', 0, 4) === 'ttcf'
  if (!isTtc && !drop) {
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer
  }
  const faceOff = isTtc ? offset : 0
  const numTables = buf.readUInt16BE(faceOff + 4)
  const entries: Array<{ dirPos: number; tOff: number; tLen: number; newOff: number }> = []
  for (let t = 0; t < numTables; t++) {
    const e = faceOff + 12 + 16 * t
    if (drop?.has(buf.toString('ascii', e, e + 4))) continue
    entries.push({
      dirPos: e,
      tOff: buf.readUInt32BE(e + 8),
      tLen: buf.readUInt32BE(e + 12),
      newOff: 0,
    })
  }
  let total = 12 + 16 * entries.length
  for (const e of entries) {
    e.newOff = total
    total += (e.tLen + 3) & ~3
  }
  const out = Buffer.alloc(total)
  buf.copy(out, 0, faceOff, faceOff + 4)
  out.writeUInt16BE(entries.length, 4)
  const pow = 2 ** Math.floor(Math.log2(entries.length || 1))
  out.writeUInt16BE(pow * 16, 6)
  out.writeUInt16BE(Math.log2(pow), 8)
  out.writeUInt16BE(entries.length * 16 - pow * 16, 10)
  for (let t = 0; t < entries.length; t++) {
    const e = entries[t]!
    buf.copy(out, 12 + 16 * t, e.dirPos, e.dirPos + 8)
    out.writeUInt32BE(e.newOff, 12 + 16 * t + 8)
    out.writeUInt32BE(e.tLen, 12 + 16 * t + 12)
    buf.copy(out, e.newOff, e.tOff, e.tOff + e.tLen)
  }
  return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer
}

function rankFaces(
  faces: FaceInfo[],
  wantKey: string,
  style: RunStyle,
  origKey?: string,
  /** Weight named in the REQUEST ("Montserrat SemiBold" → 600, word-boundary parsed by the
   *  caller — the bold flag stays false on such runs); absent = the rPr bold flag decides. */
  requestWeight?: number,
): FaceInfo[] {
  const reqTarget = requestWeight ?? (style.bold ? 700 : 400)
  const styleKeys =
    style.bold && style.italic
      ? ['bolditalic']
      : style.bold
        ? ['bold', 'w6']
        : style.italic
          ? ['italic', 'oblique']
          : ['regular', 'w3', 'medium']
  // Style keywords the request did NOT ask for: without a penalty, "Bold Italic" matches the
  // 'bold' substring and ties with the true Bold face, and cloud dirs (numeric filenames) can
  // sort the italic file first — a bold request then lands on Bold Italic.
  const avoidKeys = [
    ...(style.italic ? [] : ['italic', 'oblique']),
    // 'semibold'/'extrabold' contain 'bold': a heavy weight-suffix request must not
    // penalize exactly the faces it asks for
    ...(style.bold || reqTarget >= 500 ? [] : ['bold']),
  ]
  // origKey: an alias candidate may open a ttc that also carries the exact requested face
  // (request "Meiryo UI" -> alias "Meiryo" -> meiryo.ttc, which has both) — exact match wins.
  // Sub-family faces (Poppins Light) still match exactly through their typographic family
  // (name 16 = "Poppins"), so a famKey extending origKey costs a point and the plain face wins.
  const score = (f: FaceInfo) =>
    (f.display.startsWith('.') ? -4 : 0) +
    (origKey && f.famKeys.some((k) => k === origKey) ? 4 : 0) -
    (origKey && f.famKeys.some((k) => k !== origKey && k.startsWith(origKey)) ? 1 : 0) +
    // Weight-stripped base-family lookups (wantKey "avenir" for request "Avenir Light"):
    // the sub-family face announces the full name in family+subfamily, the plain face doesn't
    (origKey && origKey !== wantKey && f.styleText.includes(origKey) ? 3 : 0) +
    (f.famKeys.some((k) => k === wantKey || k.startsWith(wantKey)) ? 2 : 0) +
    (styleKeys.some((s) => f.styleText.includes(s)) ? 1 : 0) +
    // Regular requests: an exact 'regular'/'book' face outranks 'medium'/'w3' fallbacks
    // (cloud dirs carry every weight; a tie otherwise picks whichever file sorts first)
    (reqTarget === 400 && styleKeys[0] === 'regular' && /regular|book/.test(f.styleText) ? 1 : 0) -
    (avoidKeys.some((s) => f.styleText.includes(s)) ? 2 : 0) -
    // OS/2 weight distance as a sub-integer tiebreak: numeric cloud filenames often
    // carry no readable name records, and a flat score then picks whichever file the
    // directory scan returned first (a SemiBold drew as "Montserrat" Regular)
    (f.weight != null ? Math.abs(f.weight - reqTarget) / 1000 : 0.05)
  return [...faces].sort((a, b) => score(b) - score(a))
}

class FontRegistry {
  /** Normalized file basename (no extension) -> absolute path */
  private index = new Map<string, string>()
  /** Normalized cloud family dir name -> font file paths (numeric filenames, one per style) */
  private cloud = new Map<string, string[]>()
  /** Dir prefixes invisible to Chromium (Office DFonts / cloud-font roots) */
  private privateDirs: string[] = []
  /** Path -> face directory */
  private faceDirs = new Map<string, FaceInfo[]>()
  /** `path#offset` -> parsed font (null = parse failed) */
  private parsed = new Map<string, OpentypeFontLike | null>()
  private indexed = false

  private scanFlatDir(dir: string): void {
    let names: string[]
    try {
      names = readdirSync(dir)
    } catch {
      return
    }
    for (const name of names) {
      const m = /^(.+)\.(ttf|otf|ttc|otc)$/i.exec(name)
      if (!m) continue
      // Strip the variable-font axis suffix: NotoSansSC[wght].ttf -> notosanssc
      const key = norm(m[1]!.replace(/\[[^\]]*\]$/, ''))
      const full = join(dir, name)
      try {
        if (!statSync(full).isFile()) continue
      } catch {
        continue
      }
      if (!this.index.has(key)) this.index.set(key, full)
    }
  }

  private buildIndex(): void {
    if (this.indexed) return
    this.indexed = true
    for (const [name, path] of Object.entries(BUNDLED_FONTS)) {
      this.index.set(norm(name), path)
    }
    // System dirs first so same-named Office copies (arial.ttf…) resolve non-private
    for (const dir of fontDirs()) this.scanFlatDir(dir)
    if (userFontDir) {
      this.privateDirs.push(userFontDir)
      this.scanFlatDir(userFontDir)
    }
    for (const dir of officeFontDirs()) {
      this.privateDirs.push(dir)
      this.scanFlatDir(dir)
    }
    // One prefix each covers every scanned asset dir for the isPrivate check
    this.privateDirs.push(APPLE_FONT_ASSET_ROOT, APPLE_FONT_SUBSETS, EMBEDDED_FONT_DIR)
    for (const dir of appleFontAssetDirs()) this.scanFlatDir(dir)
    for (const root of cloudFontRoots()) {
      let families: string[]
      try {
        families = readdirSync(root)
      } catch {
        continue
      }
      this.privateDirs.push(root)
      for (const fam of families) {
        const dir = join(root, fam)
        let files: string[]
        try {
          files = readdirSync(dir)
        } catch {
          continue
        }
        const paths = files.filter((f) => /\.(ttf|otf|ttc|otc)$/i.test(f)).map((f) => join(dir, f))
        if (!paths.length) continue
        const key = norm(fam)
        this.cloud.set(key, [...(this.cloud.get(key) ?? []), ...paths])
      }
    }
  }

  isPrivate(path: string): boolean {
    return this.privateDirs.some((d) => path.startsWith(d))
  }

  private facesOf(path: string): FaceInfo[] {
    const cached = this.faceDirs.get(path)
    if (cached) return cached
    let faces: FaceInfo[]
    try {
      faces = readFaceDir(readFileSync(path))
    } catch {
      faces = []
    }
    this.faceDirs.set(path, faces)
    return faces
  }

  private parseFace(path: string, offset: number): OpentypeFontLike | null {
    const key = `${path}#${offset}`
    const cached = this.parsed.get(key)
    if (cached !== undefined) return cached
    let font: OpentypeFontLike | null
    try {
      const buf = readFileSync(path)
      try {
        font = opentype.parse(extractFace(buf, offset)) as unknown as OpentypeFontLike
      } catch {
        // Legacy CJK fonts (gulim/batang) carry GSUB versions opentype.js rejects; advances
        // only need cmap/hmtx/kern, so retry without the OpenType layout tables
        font = opentype.parse(
          extractFace(buf, offset, LAYOUT_TABLES),
        ) as unknown as OpentypeFontLike
      }
    } catch {
      font = null
    }
    this.parsed.set(key, font)
    return font
  }

  private loadBest(
    path: string,
    wantKey: string,
    style: RunStyle,
    origKey?: string,
    requestWeight?: number,
  ): { font: OpentypeFontLike; family: string; path: string; offset: number } | undefined {
    for (const face of rankFaces(this.facesOf(path), wantKey, style, origKey, requestWeight)) {
      const font = this.parseFace(path, face.offset)
      if (font) return { font, family: face.display, path, offset: face.offset }
    }
    return undefined
  }

  /** Cloud families spread styles across numeric-named files: rank all faces of all files together. */
  private loadBestCloud(
    paths: string[],
    wantKey: string,
    style: RunStyle,
    origKey?: string,
    requestWeight?: number,
  ): { font: OpentypeFontLike; family: string; path: string; offset: number } | undefined {
    const all = paths.flatMap((p) => this.facesOf(p).map((face) => ({ path: p, face })))
    const ranked = rankFaces(
      all.map((x) => x.face),
      wantKey,
      style,
      origKey,
      requestWeight,
    )
    for (const face of ranked) {
      const path = all.find((x) => x.face === face)!.path
      const font = this.parseFace(path, face.offset)
      if (font) return { font, family: face.display, path, offset: face.offset }
    }
    return undefined
  }

  /** Document-embedded face for the requested family: exact style, else degrade to regular. */
  private tryEmbedded(
    origKey: string,
    style: RunStyle,
  ): { font: OpentypeFontLike; family: string; path: string; offset: number } | undefined {
    const perStyle = embeddedFaces.get(origKey)
    if (!perStyle) return undefined
    const want = `${style.bold ? 1 : 0}${style.italic ? 1 : 0}`
    for (const k of [...new Set([want, `${want[0]}0`, `0${want[1]}`, '00'])]) {
      const path = perStyle.get(k)
      if (!path) continue
      const hit = this.loadBest(path, origKey, style, origKey, weightTargetOf(style.fontFamily))
      // Subset faces may strip name records: fall back to the declared typeface for CSS
      if (hit) return { ...hit, family: hit.family || style.fontFamily }
    }
    return undefined
  }

  /**
   * Find a font by family + bold/italic; candidates in order: requested family -> aliases ->
   * script substitution (each with its own aliases). At the file level, try style variants then
   * fall back to regular; within a ttc, pick a face by the name table. Returns the matched font
   * and its "drawing family name" (the face's English family, guaranteed CSS-resolvable).
   */
  resolve(style: RunStyle):
    | {
        font: OpentypeFontLike
        family: string
        path: string
        offset: number
        /** True when the hit came from same-script/class substitution — the requested
         *  family (and its aliases) is missing. Alias hits (Calibri→Carlito) are NOT
         *  substitutions: PowerPoint has those fonts and kerns them. */
        substituted?: boolean
      }
    | undefined {
    this.buildIndex()
    const candidates: string[] = []
    const seen = new Set<string>()
    const push = (f: string) => {
      const k = norm(f)
      if (k && !seen.has(k)) {
        seen.add(k)
        candidates.push(f)
      }
    }
    push(style.fontFamily)
    for (const a of aliasesOf(style.fontFamily)) push(a)
    // Candidates after this index are same-script/class substitutes, not the requested
    // family — the sub-family cloud fallback below must run before them
    const substituteStart = candidates.length
    for (const s of substitutesFor(style.fontFamily, style.substScript)) {
      push(s)
      for (const a of aliasesOf(s)) push(a)
    }
    // wN: Japanese fonts split files by weight (Hiragino Kaku Gothic W0..W9), and Chromium
    // matches faces by CSS weight (bold=700 -> W7, normal=400 -> W4) — metrics must pick the
    // same-weight file, or the "measure W3, draw W7" width gap makes later runs overlap text
    // (measured ~15% off on Salesforce)
    const suffixes =
      style.bold && style.italic
        ? ['bolditalic', 'bi', 'bold italic', 'w7', 'w6']
        : style.bold
          ? ['bold', 'bd', 'b', 'w7', 'w6']
          : style.italic
            ? ['italic', 'it', 'i', 'oblique']
            : ['', 'regular', 'w4', 'w3']
    const origKey = norm(style.fontFamily)
    // Word-boundary parse on the ORIGINAL name: a family merely ending in a weight
    // syllable ("Highlight") must not count as a sub-family request
    const reqWeight = weightTargetOf(style.fontFamily)
    // A weight baked into the requested name (Hiragino Kaku Gothic ProN W3, Hiragino Sans W7) names an
    // exact face — it beats the bold-flag-derived file suffix (which would try W4 first and
    // measure/draw a different weight than PowerPoint)
    const wReq = /w([0-9])$/.exec(origKey)?.[1]
    if (wReq) suffixes.unshift(`w${wReq}`)
    const tryFamily = (family: string) => {
      const base = norm(family)
      // Try style-variant files first, then fall back to regular (approximate widths still far better than heuristics)
      for (const suf of [...suffixes, '', 'regular']) {
        const path = this.index.get(base + norm(suf))
        if (!path) continue
        const hit = this.loadBest(path, base, style, origKey, reqWeight)
        if (hit) return { ...hit, family: hit.family || family }
      }
      const cloudPaths = this.cloud.get(base)
      if (cloudPaths) {
        const hit = this.loadBestCloud(cloudPaths, base, style, origKey, reqWeight)
        if (hit) return { ...hit, family: hit.family || family }
      }
      return undefined
    }
    const exact = tryFamily(candidates[0]!)
    if (exact) return exact
    // Document-embedded face: beats aliases and substitutes (PowerPoint uses the embedded
    // font whenever the family is not installed), loses only to a real system hit above
    const embedded = this.tryEmbedded(origKey, style)
    if (embedded) return embedded
    for (const family of candidates.slice(1, substituteStart)) {
      const hit = tryFamily(family)
      if (hit) return hit
    }
    // Before falling to substitutes: a sub-family request lives inside the base family's
    // cloud dir (Poppins Light -> CloudFonts/Poppins). Requires a face whose family name
    // matches the request exactly, so a shorter dir can never hijack a different family
    // (Noto Sans JP must not bind to Latin-only Noto Sans).
    let bestKey = ''
    for (const k of this.cloud.keys()) {
      if (origKey.startsWith(k) && k.length > bestKey.length) bestKey = k
    }
    if (bestKey) {
      const paths = this.cloud.get(bestKey)!
      if (paths.some((p) => this.facesOf(p).some((f) => f.famKeys.includes(origKey)))) {
        const hit = this.loadBestCloud(paths, origKey, style, origKey, reqWeight)
        if (hit) return { ...hit, family: hit.family || style.fontFamily }
      }
    }
    // "Family Weight" sub-family with no file/cloud hit of its own: fall back to the base
    // family's faces (probe: PowerPoint renders "Avenir Light" with the real Avenir-Light
    // face). rankFaces' origKey bonuses pick the sub-family face out of the collection.
    const stripped = baseFamilyOf(style.fontFamily)
    if (stripped) {
      for (const family of [stripped, ...aliasesOf(stripped)]) {
        const hit = tryFamily(family)
        if (hit) return hit
      }
    }
    for (const family of candidates.slice(substituteStart)) {
      const hit = tryFamily(family)
      if (hit) return { ...hit, substituted: true }
    }
    return undefined
  }
}

/** One glyph as returned by opentype.js (only the advance is read here). */
interface OpentypeGlyph {
  advanceWidth?: number
}

/**
 * Structural view of the opentype.js runtime internals probed below (glyph lookup,
 * kern pairs, variable-font fvar/HVAR tables). Everything is feature-detected at
 * runtime; the cast just names the shape instead of erasing it with `any`.
 */
interface OpentypeRuntimeFont extends OpentypeFontLike {
  charToGlyph?(char: string): OpentypeGlyph
  getKerningValue?(left: unknown, right: unknown): number
  tables?: {
    fvar?: {
      axes?: Array<{ tag: string; minValue: number; maxValue: number; defaultValue: number }>
    }
    hvar?: unknown
  }
  variation?: {
    getTransform?(
      glyph: OpentypeGlyph,
      coords: Record<string, number>,
    ): { advanceWidth?: number } | undefined
  }
}

/**
 * Bypass opentype.js's getAdvanceWidth: it runs Bidi/GSUB text shaping and throws on
 * unsupported lookups (e.g. Inter's ccmp lookupType6/substFormat2); the caller catches and the
 * whole run falls back to heuristics — uppercase Latin gets underestimated ~18%, inter-word
 * spaces get swallowed, and runs can even overlap. Here we accumulate advance per glyph (with
 * kern pairs) without triggering the shaping path.
 */
function wrapSafeAdvance(font: OpentypeFontLike): OpentypeFontLike {
  const f = font as OpentypeRuntimeFont
  if (typeof f.charToGlyph !== 'function') return font
  return {
    unitsPerEm: font.unitsPerEm,
    ascender: font.ascender,
    descender: font.descender,
    ...(font.charToGlyphIndex
      ? { charToGlyphIndex: (ch: string) => font.charToGlyphIndex!(ch) }
      : {}),
    getAdvanceWidth(text: string, fontSize: number, options?: { kerning?: boolean }): number {
      const kern = options?.kerning !== false
      let units = 0
      let prev: unknown = null
      for (const ch of text) {
        const glyph = f.charToGlyph!(ch)
        units += glyph?.advanceWidth ?? 0
        if (kern && prev && typeof f.getKerningValue === 'function') {
          try {
            units += f.getKerningValue(prev, glyph) || 0
          } catch {
            /* Treat a failed kern lookup as 0 */
          }
        }
        prev = glyph
      }
      return (units / font.unitsPerEm) * fontSize
    },
  }
}

/**
 * Instantiate advances of variable fonts at the requested weight (opentype.js 2.x variation
 * API + HVAR).
 *
 * Background: Google Fonts variable fonts (e.g. NotoSansJP[wght].ttf) may default to
 * Thin (wght=100), and opentype's getAdvanceWidth only measures the default instance; Chromium
 * actually renders at 400/700, where digits/Latin can be ~16% wider, causing mixed-script run
 * overlap (e.g. a CJK unit suffix after "3,173" pressed onto the digits). For fonts with a
 * wght axis, apply HVAR deltas per glyph and recompute advances at bold?700:400 (no kern;
 * negligible for CJK/digit scenarios).
 */
/** Whether the parsed face is a variable font with a weight axis (fvar). */
function isVariableFont(font: OpentypeFontLike): boolean {
  const f = font as OpentypeRuntimeFont
  return !!f.tables?.fvar?.axes?.some((a) => a.tag === 'wght')
}

function instantiateWeight(font: OpentypeFontLike, bold: boolean): OpentypeFontLike {
  const f = font as OpentypeRuntimeFont
  const wghtAxis = f.tables?.fvar?.axes?.find((a) => a.tag === 'wght')
  if (
    !wghtAxis ||
    typeof f.variation?.getTransform !== 'function' ||
    !f.tables?.hvar ||
    typeof f.charToGlyph !== 'function' ||
    typeof f.charToGlyphIndex !== 'function'
  ) {
    return font
  }
  const target = Math.min(Math.max(bold ? 700 : 400, wghtAxis.minValue), wghtAxis.maxValue)
  if (target === wghtAxis.defaultValue) return font
  const coords = { wght: target }
  /** glyph index -> advance (font units, already instantiated) */
  const advCache = new Map<number, number>()
  return {
    unitsPerEm: font.unitsPerEm,
    ascender: font.ascender,
    descender: font.descender,
    ...(font.charToGlyphIndex
      ? { charToGlyphIndex: (ch: string) => font.charToGlyphIndex!(ch) }
      : {}),
    getAdvanceWidth(text: string, fontSize: number): number {
      let units = 0
      for (const ch of text) {
        const gid = f.charToGlyphIndex!(ch)
        let adv = advCache.get(gid)
        if (adv == null) {
          const glyph = f.charToGlyph!(ch)
          try {
            adv = f.variation!.getTransform!(glyph, coords)?.advanceWidth ?? glyph.advanceWidth ?? 0
          } catch {
            adv = glyph.advanceWidth ?? 0
          }
          advCache.set(gid, adv)
        }
        units += adv
      }
      return (units / font.unitsPerEm) * fontSize
    },
  }
}

/** An Office-private face the renderer must register as a FontFace (Chromium can't see the file). */
export interface PrivateFontFaceInfo {
  id: string
  family: string
  bold: boolean
  italic: boolean
}

/** id -> file location of private faces referenced by layouts so far (grows as decks open). */
const privateFaces = new Map<
  string,
  { family: string; bold: boolean; italic: boolean; path: string; offset: number }
>()

export function listPrivateFontFaces(): PrivateFontFaceInfo[] {
  return [...privateFaces.entries()].map(([id, f]) => ({
    id,
    family: f.family,
    bold: f.bold,
    italic: f.italic,
  }))
}

/** Extracted single-face sfnt bytes for a private face (ttc split out; FontFace can't take ttc). */
export function getPrivateFontData(id: string): ArrayBuffer | null {
  const f = privateFaces.get(id)
  if (!f) return null
  try {
    return extractFace(readFileSync(f.path), f.offset)
  } catch {
    return null
  }
}

/** Process-wide registry; reset after the user font store changes so new files get indexed. */
let sharedRegistry: FontRegistry | null = null
function getRegistry(): FontRegistry {
  return (sharedRegistry ??= new FontRegistry())
}
export function resetFontRegistry(): void {
  sharedRegistry = null
}

/** True when the family resolves without same-script/class substitution (alias hits count as available). */
export function familyAvailable(family: string): boolean {
  const hit = getRegistry().resolve({
    fontFamily: family,
    fontSizePx: 100,
    bold: false,
    italic: false,
  })
  return !!hit && !hit.substituted
}

/** Family names (name 1/16) declared by a local font file; [] when unreadable. */
export function fontFileFamilies(path: string): string[] {
  try {
    const faces = readFaceDir(readFileSync(path))
    return [...new Set(faces.map((f) => f.display).filter(Boolean))]
  } catch {
    return []
  }
}

/** Create a metrics provider injected with system fonts (falls back to heuristics per run when no font is found). */
export function createSystemFontMetrics(): FontMetricsProvider {
  initShapedMetrics()
  const registry = getRegistry()
  const cache = new Map<
    string,
    | {
        font: OpentypeFontLike
        family: string
        substituted?: boolean
        gtruth?: boolean
        path: string
        offset: number
      }
    | undefined
  >()
  const resolveEntry = (
    style: RunStyle,
  ):
    | {
        font: OpentypeFontLike
        family: string
        substituted?: boolean
        gtruth?: boolean
        path: string
        offset: number
      }
    | undefined => {
    const key = `${style.fontFamily}|${style.bold ? 1 : 0}${style.italic ? 1 : 0}|${style.substScript ?? ''}`
    if (cache.has(key)) return cache.get(key)
    const raw = registry.resolve(style)
    let entry:
      | {
          font: OpentypeFontLike
          family: string
          substituted?: boolean
          /** Measure via renderer ground truth: private faces (FontFace registration may
           *  land on a sibling style file) and variable fonts (instancing drift) */
          gtruth?: boolean
          path: string
          offset: number
        }
      | undefined
    // Weight-in-name requests (Hiragino Kaku Gothic ProN W3) resolve to a specific face of a system
    // collection, but CSS matches that family by weight only (normal → W4) — register the
    // exact face under a synthetic "<family> W<n>" name so drawing uses the measured face.
    const wReq = /w([0-9])$/.exec(norm(style.fontFamily))?.[1]
    if (raw && !registry.isPrivate(raw.path) && wReq) {
      if (!norm(raw.family).endsWith(`w${wReq}`)) raw.family = `${raw.family} W${wReq}`
      privateFaces.set(`${raw.family}|${style.bold ? 1 : 0}${style.italic ? 1 : 0}`, {
        family: raw.family,
        bold: style.bold,
        italic: style.italic,
        path: raw.path,
        offset: raw.offset,
      })
    }
    // "Family Weight" requests resolved from a system file (Avenir Light, Helvetica Neue
    // Light): the picked sub-family face often reports the plain family name (name 1
    // "Helvetica Neue" + subfamily "Light"), which CSS resolves to the Regular face —
    // measuring Light while drawing Regular swallows word spaces. Serve the exact face as
    // a FontFace under the requested name, same as the wN weight-in-name path above.
    // Only when the face's OS/2 weight matches the request: a base family without such a
    // face (Georgia Light → georgia.ttf Regular) keeps its own name instead of lying.
    const weightTarget = wReq ? undefined : weightTargetOf(style.fontFamily)
    if (raw && !raw.substituted && weightTarget != null) {
      const os2W = (raw.font as { tables?: { os2?: { usWeightClass?: number } } }).tables?.os2
        ?.usWeightClass
      // Renaming a plain-named face needs weight evidence; a face already carrying the
      // requested name registers unconditionally. Book/Roman requests (target 400) accept
      // an exact-400 face (harmless when it IS the Regular face, and it keeps a distinct
      // Book face measure/draw-consistent); other targets must land on a non-Regular face
      // near the request, so Georgia Light never registers georgia.ttf's Regular.
      // The rename applies to private hits too: a "Montserrat SemiBold" request landing on
      // the cloud SemiBold face (numeric filename, stripped name records) must not key the
      // registration as bare "Montserrat|00" — that overwrites the sibling Regular family
      // and every plain run then draws SemiBold, swallowing word spaces.
      const isRequestedName = norm(raw.family) === norm(style.fontFamily)
      const weightMatch =
        typeof os2W === 'number' &&
        (weightTarget === 400 ? os2W === 400 : os2W !== 400 && Math.abs(os2W - weightTarget) <= 100)
      if (!isRequestedName && weightMatch) raw.family = style.fontFamily.trim()
      if (!registry.isPrivate(raw.path) && (isRequestedName || weightMatch)) {
        privateFaces.set(`${raw.family}|${style.bold ? 1 : 0}${style.italic ? 1 : 0}`, {
          family: raw.family,
          bold: style.bold,
          italic: style.italic,
          path: raw.path,
          offset: raw.offset,
        })
      }
    }
    if (raw && registry.isPrivate(raw.path)) {
      // Register under the requested style only when this style resolved to its own face —
      // when bold/italic fell back to the same file+face as regular, skip it so Chromium
      // keeps synthesizing bold/italic from the regular face (matching PowerPoint).
      const base =
        style.bold || style.italic
          ? registry.resolve({ ...style, bold: false, italic: false })
          : undefined
      // Weight-in-name families ('Arimo Bold' + b=1) resolve both probes to the same real
      // Bold face; keying it '|00' would collide with the sibling regular family ('Arimo')
      // and draw regular glyphs — when the face's own style matches the request, keep the
      // requested key so bold/regular faces of one display family coexist.
      const os2Weight = (raw.font as { tables?: { os2?: { usWeightClass?: number } } }).tables?.os2
        ?.usWeightClass
      const italicAngle = (raw.font as { tables?: { post?: { italicAngle?: number } } }).tables
        ?.post?.italicAngle
      const faceMatchesRequest =
        typeof os2Weight === 'number' &&
        os2Weight >= 600 === style.bold &&
        (typeof italicAngle === 'number' ? Math.abs(italicAngle) > 0.01 : false) === style.italic
      if (!base || base.path !== raw.path || base.offset !== raw.offset || faceMatchesRequest) {
        privateFaces.set(`${raw.family}|${style.bold ? 1 : 0}${style.italic ? 1 : 0}`, {
          family: raw.family,
          bold: style.bold,
          italic: style.italic,
          path: raw.path,
          offset: raw.offset,
        })
      } else {
        // Same file+face as regular: register the regular face (even if no regular run
        // exists in the deck) so Chromium has a real face to synthesize bold/italic from.
        privateFaces.set(`${base.family}|00`, {
          family: base.family,
          bold: false,
          italic: false,
          path: base.path,
          offset: base.offset,
        })
      }
    }
    if (raw) {
      const inst = instantiateWeight(raw.font, style.bold)
      // Non-variable fonts (instantiateWeight returned as-is) need the safe-advance wrapper;
      // the variable-font path already accumulates per glyph and is inherently safe
      let font = inst === raw.font ? wrapSafeAdvance(raw.font) : inst
      // hhea lineGap (external leading): lives only in the table, and the wrappers above
      // rebuild the object — carry it so single spacing includes it (CoreText semantics)
      const hheaGap = (raw.font as { tables?: { hhea?: { lineGap?: number } } }).tables?.hhea
        ?.lineGap
      if (typeof hheaGap === 'number' && hheaGap > 0) font = { ...font, lineGap: hheaGap }
      // Bundled Carlito ships Linux-style hhea metrics (1.0 em) while PowerPoint spaces
      // Calibri by the OS/2 win metrics (1.22 em) — take line metrics from OS/2 win so
      // substituted decks keep PowerPoint's line pitch.
      if (raw.family.toLowerCase().startsWith('carlito')) {
        const os2 = (
          raw.font as { tables?: { os2?: { usWinAscent?: number; usWinDescent?: number } } }
        ).tables?.os2
        if (os2?.usWinAscent && os2.usWinDescent != null) {
          // Win metrics span the full 1.22em pitch (== hhea asc+desc+gap for Carlito):
          // the external leading is already inside them, adding it again double-counts
          font = {
            ...font,
            ascender: os2.usWinAscent,
            descender: -Math.abs(os2.usWinDescent),
            lineGap: 0,
          }
        }
      }
      entry = {
        font,
        family: raw.family,
        ...(registry.isPrivate(raw.path) || isVariableFont(raw.font) ? { gtruth: true } : {}),
        path: raw.path,
        offset: raw.offset,
        ...(raw.substituted ? { substituted: true } : {}),
      }
    }
    cache.set(key, entry)
    return entry
  }
  const inner = new OpentypeMetrics((style) => resolveEntry(style)?.font, new HeuristicMetrics())
  // Complex-script runs whose REQUESTED family resolved to a real (non-substituted)
  // face shape and draw with that face — decks ship real Arabic fonts via Office
  // CloudFonts/DFonts/embeds, and forcing the generic script substitute (Geeza Pro)
  // makes every run ~20% wider than PowerPoint (prod_016). Substituted resolutions
  // keep the script default so measuring and drawing stay on one font.
  const prefFaceFor = (style: RunStyle, text: string): ShapedPrefFace | undefined => {
    if (!complexScriptOf(text)) return undefined
    const e = resolveEntry(style)
    if (!e || e.substituted) return undefined
    const { family, path, offset } = e
    return {
      family,
      bytes: () => {
        try {
          return extractFace(readFileSync(path), offset)
        } catch {
          return null
        }
      },
    }
  }
  return {
    metrics: (style) => inner.metrics(style),
    // Complex scripts (ligatures/contextual forms) prefer HarfBuzz shaped metrics — opentype's
    // per-glyph accumulation measures isolated forms, drifting from actual drawing; falls back
    // to the original path when not ready or no font
    measure: (text, style) => {
      const shaped = shapedMeasure(
        text,
        style.fontSizePx,
        style.bold,
        style.italic,
        prefFaceFor(style, text),
      )
      if (shaped != null) return shaped
      // Private faces (cloud/DFonts/embeds: FontFace registration can land on a sibling
      // style file of the same family) and variable fonts (instancing drift): opentype
      // advances run a few percent off Chromium's rendering, visually swallowing word
      // spaces — take the renderer's measureText as ground truth (cached, refined in batch)
      const e = resolveEntry(style)
      if (e?.gtruth) {
        const gt = gtMeasure(text, e.family, style.fontSizePx, style.bold, style.italic)
        if (gt != null) return gt
      }
      return inner.measure(text, style)
    },
    // Substituted fonts return the substitute family; the renderer draws with it (same font file for measuring/drawing)
    displayFamily: (style, text) =>
      (text != null
        ? shapedFamily(text, prefFaceFor(style, text), style.bold, style.italic)
        : null) ??
      resolveEntry(style)?.family ??
      style.fontFamily,
    substituted: (style) => resolveEntry(style)?.substituted === true,
  }
}
