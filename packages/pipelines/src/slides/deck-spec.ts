/**
 * Multi-page spec → one pptx: every page is built with buildPagePptx and the
 * results are merged in order, exactly how the app lands generated pages.
 */
import {
  mergeSlideFromPptx,
  openPptx,
  promoteSlideBackground,
  savePptx,
} from '@genoffice/pptx-engine'
import {
  buildPagePptx,
  parsePageSpecObject,
  SPEC_CANVAS_H,
  SPEC_CANVAS_W,
  type BuildPageDeps,
  type PageSpec,
  type ParseSpecOptions,
} from './page-spec'

export const MAX_DECK_PAGES = 60
/** Raw LLM deck-spec budget; larger payloads are rejected before JSON.parse. */
const MAX_DECK_SPEC_RAW_CHARS = 1_000_000

export interface DeckSpec {
  pages: PageSpec[]
}

export interface DeckPageIssue {
  /** 0-based page index in the input */
  page: number
  error?: string
  warnings?: string[]
}

/**
 * Accepts `{ "pages": [...] }`, a bare array of page objects, or a single page
 * object. Pages that fail validation are dropped and reported; the deck fails
 * only when no page survives.
 */
export function parseDeckSpec(
  raw: string,
  opts: ParseSpecOptions = {},
): { ok: true; spec: DeckSpec; issues: DeckPageIssue[] } | { ok: false; error: string } {
  let parsed: unknown
  // LLM output is untrusted: refuse a megabyte dump before JSON.parse builds
  // a giant object graph from it. The CLI caller passes localImages and documents
  // data: sources up to 50 MB (image-source.ts), where a base64 payload alone runs
  // to ~67 M chars, so the cap only guards the default http(s)-only path.
  if (!opts.localImages && raw.length > MAX_DECK_SPEC_RAW_CHARS) {
    return { ok: false, error: `deck spec too large (limit ${MAX_DECK_SPEC_RAW_CHARS} chars)` }
  }
  try {
    parsed = JSON.parse(raw)
  } catch (e) {
    return { ok: false, error: `invalid JSON: ${e instanceof Error ? e.message : String(e)}` }
  }
  const rawPages = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { pages?: unknown }).pages)
      ? ((parsed as { pages: unknown[] }).pages as unknown[])
      : parsed && typeof parsed === 'object' && 'elements' in (parsed as object)
        ? [parsed]
        : null
  if (!rawPages) {
    return {
      ok: false,
      error: 'expected { "pages": [ { "background", "elements": [...] }, ... ] }',
    }
  }
  if (rawPages.length === 0) return { ok: false, error: 'the "pages" array is empty' }
  if (rawPages.length > MAX_DECK_PAGES) {
    return {
      ok: false,
      error: `too many pages (${rawPages.length}; the limit is ${MAX_DECK_PAGES})`,
    }
  }
  const pages: PageSpec[] = []
  const issues: DeckPageIssue[] = []
  for (const [i, page] of rawPages.entries()) {
    const r = parsePageSpecObject(page, SPEC_CANVAS_W, SPEC_CANVAS_H, opts)
    if (!r.ok) {
      issues.push({ page: i, error: r.error })
      continue
    }
    if (r.warnings.length) issues.push({ page: i, warnings: r.warnings })
    pages.push(r.spec)
  }
  if (pages.length === 0) {
    return {
      ok: false,
      error: `no valid pages (${issues.map((i) => `page ${i.page}: ${i.error}`).join('; ')})`,
    }
  }
  return { ok: true, spec: { pages }, issues }
}

export interface BuildDeckResult {
  bytes: Uint8Array
  /** Image sources that could not be loaded, per 0-based page */
  imageFailures: { page: number; url: string }[]
  /** Images that landed but lost most of the picture to the cover crop, per 0-based page */
  imageWarnings: { page: number; message: string }[]
}

export async function buildDeckPptx(spec: DeckSpec, deps: BuildPageDeps): Promise<BuildDeckResult> {
  const imageFailures: BuildDeckResult['imageFailures'] = []
  const imageWarnings: BuildDeckResult['imageWarnings'] = []
  const pages: Uint8Array[] = []
  for (const [i, page] of spec.pages.entries()) {
    const r = await buildPagePptx(page, deps)
    pages.push(r.bytes)
    for (const url of r.imageFailures) imageFailures.push({ page: i, url })
    for (const message of r.imageWarnings) imageWarnings.push({ page: i, message })
  }
  const base = await openPptx(pages[0]!)
  for (const one of pages.slice(1)) await mergeSlideFromPptx(base, one)
  for (const s of base.deck.slides) promoteSlideBackground(s, base.deck.size)
  return { bytes: await savePptx(base), imageFailures, imageWarnings }
}
