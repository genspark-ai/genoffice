import { loadDocxZip } from './zip-load'
import { escapeXmlAttr } from './xml-utils'

const unescapeAttr = (value: string): string =>
  value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')

const RUN_FONT_ATTRS = /(\sw:(?:ascii|hAnsi|eastAsia|cs)=")([^"]*)(")/g
const THEME_FONT_ATTR = /(\stypeface=")([^"]*)(")/

/**
 * One part's XML with the named fonts replaced wherever text is given a font:
 * the four slots of every w:rFonts, and the typefaces a theme part declares.
 * `fonts` maps a lower-case font name to the name that takes its place.
 */
export function replaceFontsXml(xml: string, fonts: ReadonlyMap<string, string>): string {
  const swap = (_all: string, head: string, value: string, tail: string) => {
    const to = fonts.get(unescapeAttr(value).toLowerCase())
    return head + (to === undefined ? value : escapeXmlAttr(to)) + tail
  }
  return xml
    .replace(/<w:rFonts\b[^>]*>/g, (tag) => tag.replace(RUN_FONT_ATTRS, swap))
    .replace(/<a:(?:latin|ea|cs|font)\b[^>]*>/g, (tag) => tag.replace(THEME_FONT_ATTR, swap))
}

const PART_WITH_FONTS = /^word\/(?:[^/]+|theme\/[^/]+|glossary\/[^/]+)\.xml$/

const byLowerName = (fonts: Readonly<Record<string, string>>): Map<string, string> =>
  new Map(
    Object.entries(fonts)
      .filter(([from, to]) => from && to)
      .map(([from, to]) => [from.toLowerCase(), to]),
  )

/** What one save's Convert Permanently wrote, kept so an undo can put the fonts back. */
export interface WrittenFontConversion {
  /** missing font -> its stand-in, as the file holds it */
  fonts: Readonly<Record<string, string>>
  /** the share still wanted: `fonts` itself until an undo takes some of it back */
  kept: Readonly<Record<string, string>>
  /** every part the conversion changed, before it and after it */
  parts: Readonly<Record<string, { before: string; after: string }>>
}

/**
 * Word's Convert Permanently at the save, and its undo. `pending` (missing
 * font -> the font that stands in) is written into every run, style, default
 * and theme font of every part that holds text. A conversion of `written`
 * that is no longer wholly kept is taken back out: a part still as that save
 * left it gets what it held before, converted by what is kept. A part changed
 * since stays converted, because a font slot that named the stand-in before
 * the conversion can no longer be told from one the conversion wrote.
 */
export async function writeFontConversions(
  bytes: Uint8Array,
  pending: Readonly<Record<string, string>>,
  written: readonly WrittenFontConversion[] = [],
): Promise<{ bytes: Uint8Array; written: WrittenFontConversion[] }> {
  const byName = byLowerName(pending)
  if (byName.size === 0 && written.every((w) => w.kept === w.fonts))
    return { bytes, written: [...written] }
  const zip = await loadDocxZip(bytes)
  const read = async (name: string) => (await zip.file(name)?.async('string')) ?? null
  let changed = false
  const next: WrittenFontConversion[] = []
  // the latest first: its parts hold what the one before it left
  for (const w of [...written].reverse()) {
    if (w.kept === w.fonts) {
      next.unshift(w)
      continue
    }
    const kept = byLowerName(w.kept)
    const parts: Record<string, { before: string; after: string }> = {}
    for (const [name, part] of Object.entries(w.parts)) {
      if ((await read(name)) !== part.after) continue
      const after = replaceFontsXml(part.before, kept)
      zip.file(name, after)
      changed = true
      parts[name] = { before: part.before, after }
    }
    if (kept.size > 0) next.unshift({ fonts: w.kept, kept: w.kept, parts })
  }
  if (byName.size > 0) {
    const parts: Record<string, { before: string; after: string }> = {}
    for (const name of Object.keys(zip.files)) {
      if (!PART_WITH_FONTS.test(name)) continue
      const before = await read(name)
      if (before === null) continue
      const after = replaceFontsXml(before, byName)
      if (after === before) continue
      zip.file(name, after)
      changed = true
      parts[name] = { before, after }
    }
    // kept also when only the editor's runs named the fonts and no part changed
    // here: an undo still has to know which fonts are back
    next.push({ fonts: pending, kept: pending, parts })
  }
  if (!changed) return { bytes, written: next }
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    written: next,
  }
}

/** Convert Permanently alone: the document with `fonts` (missing font -> stand-in) written in. */
export async function replaceFontsInDocx(
  bytes: Uint8Array,
  fonts: Readonly<Record<string, string>>,
): Promise<Uint8Array> {
  return (await writeFontConversions(bytes, fonts)).bytes
}
