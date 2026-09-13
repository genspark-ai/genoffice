import { readFile } from 'node:fs/promises'
import { extname } from 'node:path'
import {
  extractDocumentImageSources,
  resolveSafeRelativeImagePath,
  rewriteDocumentImageSources,
} from './asset-lifecycle'
import { ASSET_SNIFF_BYTES, sniffBinaryAssetMime } from './asset-mime'

/**
 * Single-file HTML export: a copy of the document with every local image
 * reference (img src and CSS url()) inlined as a data URL, so the exported
 * file opens anywhere without the sibling assets/ folder. The working
 * document is never rewritten — inlining happens only at export time (a
 * saved document deliberately keeps images in assets/: megabyte base64
 * lines make the source and AI edits fragile).
 */
export interface SingleFileHtmlResult {
  html: string
  /** references replaced with data URLs (occurrences, not unique sources) */
  inlined: number
  /** local-looking references that could not be resolved and were left as-is */
  skipped: string[]
}

/**
 * Suggested base name for the export. The working file is `<name>.html`, and
 * the save dialog anchors a bare name in the last-used folder — usually the
 * document's own — so the plain name would land on the open document and be
 * refused. A re-export of an exported copy does not stack suffixes.
 */
export function singleFileExportBaseName(name: string): string {
  return `${name.replace(/\.single$/i, '')}.single`
}

/** a reference the browser resolves elsewhere: any scheme (http:, data:, …) or scheme-relative */
function isExternalRef(source: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(source) || source.startsWith('//')
}

async function readAsDataUrl(documentPath: string, source: string): Promise<string | null> {
  const target = await resolveSafeRelativeImagePath(documentPath, source)
  if (!target) return null
  let bytes: Buffer
  try {
    bytes = await readFile(target)
  } catch {
    return null
  }
  // content decides the type; SVG is text and has no binary signature
  const mime =
    sniffBinaryAssetMime(bytes.subarray(0, ASSET_SNIFF_BYTES)) ??
    (extname(target).toLowerCase() === '.svg' ? 'image/svg+xml' : null)
  if (!mime?.startsWith('image/')) return null
  return `data:${mime};base64,${bytes.toString('base64')}`
}

interface CssUrlMatch {
  start: number
  end: number
  source: string
  quote: '"' | "'" | ''
}

/** Every CSS url(...) reference: <style> rules and inline style attributes alike. */
function scanCssUrls(text: string): CssUrlMatch[] {
  const out: CssUrlMatch[] = []
  const re = /url\(\s*(?:"([^"\n]*)"|'([^'\n]*)'|([^)"'\s]+))\s*\)/gi
  for (const match of text.matchAll(re)) {
    const source = match[1] ?? match[2] ?? match[3] ?? ''
    if (!source) continue
    out.push({
      start: match.index,
      end: match.index + match[0].length,
      source,
      quote: match[1] !== undefined ? '"' : match[2] !== undefined ? "'" : '',
    })
  }
  return out
}

/** Base64 data URLs contain no spaces, quotes or parens, so the original quoting form stays valid. */
function rewriteCssUrls(
  text: string,
  dataBySource: ReadonlyMap<string, string>,
): { html: string; rewritten: number } {
  let cursor = 0
  let output = ''
  let rewritten = 0
  for (const match of scanCssUrls(text)) {
    const data = dataBySource.get(match.source)
    if (data === undefined) continue
    output += text.slice(cursor, match.start) + `url(${match.quote}${data}${match.quote})`
    cursor = match.end
    rewritten += 1
  }
  return cursor === 0 ? { html: text, rewritten } : { html: output + text.slice(cursor), rewritten }
}

export async function inlineImagesForSingleFile(
  text: string,
  documentPath: string | null,
): Promise<SingleFileHtmlResult> {
  // an unsaved document has no assets/ to resolve against — its images are already data URLs
  if (!documentPath) return { html: text, inlined: 0, skipped: [] }
  // HTML-only scan: the Markdown-aware scanners would skip every indented <img>
  const imageSources = extractDocumentImageSources(text)
  const cssMatches = scanCssUrls(text)
  const candidates = new Set(
    [...imageSources, ...cssMatches.map((m) => m.source)].filter((s) => s && !isExternalRef(s)),
  )
  const dataBySource = new Map<string, string>()
  const skipped: string[] = []
  for (const source of candidates) {
    const data = await readAsDataUrl(documentPath, source)
    if (data) dataBySource.set(source, data)
    else skipped.push(source)
  }
  if (dataBySource.size === 0) return { html: text, inlined: 0, skipped }
  const inlinedImgs = imageSources.filter((s) => dataBySource.has(s)).length
  const afterImgs = rewriteDocumentImageSources(text, dataBySource)
  const css = rewriteCssUrls(afterImgs, dataBySource)
  return { html: css.html, inlined: inlinedImgs + css.rewritten, skipped }
}
