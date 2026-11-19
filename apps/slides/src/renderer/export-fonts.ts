/**
 * @font-face CSS for the PDF print window. It is a separate page: the bundled
 * substitutes (styles.css) and the Office-private FontFaces (doc-fonts.ts) the
 * canvas draws with do not exist there, so the faces the exported text uses are
 * inlined as data: URLs. System fonts need nothing — Chromium resolves them by
 * name and embeds the subsets while printing.
 */

const GENERIC = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui'])

/** Family names of every stack the page used, normalized for matching. */
export function familyNames(stacks: Iterable<string>): Set<string> {
  const out = new Set<string>()
  for (const stack of stacks) {
    for (const part of stack.split(',')) {
      const name = part.trim().replace(/^['"]|['"]$/g, '')
      if (name && !GENERIC.has(name)) out.add(name.normalize('NFKC').toLowerCase())
    }
  }
  return out
}

function toBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf)
  let bin = ''
  const CHUNK = 0x8000
  for (let i = 0; i < bytes.length; i += CHUNK)
    bin += String.fromCharCode(...bytes.subarray(i, i + CHUNK))
  return btoa(bin)
}

function faceRule(family: string, weight: string, style: string, b64: string): string {
  return `@font-face{font-family:${JSON.stringify(family)};src:url(data:font/ttf;base64,${b64});font-weight:${weight};font-style:${style};font-display:block}`
}

const bundledCache = new Map<string, Promise<string | null>>()

/** Bundled @font-face rules (styles.css) whose family the page uses, with their bytes inlined. */
async function bundledFaces(names: Set<string>): Promise<string[]> {
  const jobs: Array<Promise<string | null>> = []
  for (const sheet of Array.from(document.styleSheets)) {
    let rules: CSSRuleList
    try {
      rules = sheet.cssRules
    } catch {
      continue
    }
    for (const rule of Array.from(rules)) {
      // duck-typed: jsdom has no CSSFontFaceRule global
      if (!rule.cssText.startsWith('@font-face') || !('style' in rule)) continue
      const style = (rule as CSSFontFaceRule).style
      const family = style
        .getPropertyValue('font-family')
        .trim()
        .replace(/^['"]|['"]$/g, '')
      if (!names.has(family.normalize('NFKC').toLowerCase())) continue
      const url = /url\(\s*['"]?([^'")]+)['"]?\s*\)/.exec(style.getPropertyValue('src'))?.[1]
      if (!url) continue
      const weight = style.getPropertyValue('font-weight') || 'normal'
      const fontStyle = style.getPropertyValue('font-style') || 'normal'
      let data = bundledCache.get(url)
      if (!data) {
        data = fetch(url)
          .then((r) => (r.ok ? r.arrayBuffer() : null))
          .then((buf) => (buf ? toBase64(buf) : null))
          .catch(() => null)
        bundledCache.set(url, data)
      }
      jobs.push(data.then((b64) => (b64 ? faceRule(family, weight, fontStyle, b64) : null)))
    }
  }
  return (await Promise.all(jobs)).filter((r): r is string => !!r)
}

/** Office-private faces (embedded / Office-bundled fonts) the page uses. */
async function privateFaces(names: Set<string>): Promise<string[]> {
  let faces: Array<{ id: string; family: string; bold: boolean; italic: boolean }>
  try {
    faces = await window.slidesApi.privateFontFaces()
  } catch {
    return []
  }
  if (!Array.isArray(faces)) return []
  const rules = await Promise.all(
    faces
      .filter((f) => names.has(f.family.normalize('NFKC').toLowerCase()))
      .map(async (f) => {
        try {
          const data = await window.slidesApi.privateFontData(f.id)
          if (!data) return null
          return faceRule(
            f.family,
            f.bold ? '700' : '400',
            f.italic ? 'italic' : 'normal',
            toBase64(data),
          )
        } catch {
          return null
        }
      }),
  )
  return rules.filter((r): r is string => !!r)
}

/** All @font-face rules the print window needs for the given font stacks. */
export async function collectExportFontCss(stacks: Iterable<string>): Promise<string> {
  const names = familyNames(stacks)
  if (!names.size) return ''
  const [bundled, priv] = await Promise.all([bundledFaces(names), privateFaces(names)])
  return [...bundled, ...priv].join('\n')
}
