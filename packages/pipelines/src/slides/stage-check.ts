/**
 * Cross-checks between the deck-flow files: a page spec against its outline
 * entry, and a page's colors against the style sheet's palette. The builder
 * ignores the `title` / `type` / `layout` a page file may echo from its outline
 * entry; these checks are where they matter.
 */
import type { OutlineIssue, OutlinePage } from './outline'
import { PLACEHOLDER } from './outline'
import type { PageSpec } from './page-spec'

export function pageText(spec: PageSpec): string {
  const parts: string[] = []
  for (const el of spec.elements) {
    if (el.type === 'image') continue
    for (const p of el.paragraphs ?? []) for (const r of p.runs) parts.push(r.text)
  }
  return parts.join('\n')
}

/** Every color a page paints with, as #RRGGBB (alpha dropped), in first-use order. */
export function pageColors(spec: PageSpec): string[] {
  const seen = new Set<string>()
  const add = (c: string | undefined) => {
    if (c) seen.add(rgb(c))
  }
  add(spec.background)
  for (const el of spec.elements) {
    if (el.type === 'image') continue
    if (el.type === 'shape') {
      add(el.fill)
      add(el.stroke?.color)
    }
    for (const p of el.paragraphs ?? []) for (const r of p.runs) add(r.color)
  }
  return [...seen]
}

function rgb(color: string): string {
  return color.slice(0, 7).toUpperCase()
}

const HEX = /#([0-9a-f]{6})(?:[0-9a-f]{2})?\b/gi

/** The #RRGGBB values a style sheet (markdown) names; empty when it names none. */
export function stylePalette(markdown: string): Set<string> {
  const out = new Set<string>()
  for (const m of markdown.matchAll(HEX)) out.add(`#${m[1]!.toUpperCase()}`)
  return out
}

const NEUTRAL = new Set(['#FFFFFF', '#000000'])

/** Colors the page uses that the style sheet does not name; pure white and black are always allowed. */
export function offPaletteColors(spec: PageSpec, palette: Set<string>): string[] {
  if (palette.size === 0) return []
  return pageColors(spec).filter((c) => !palette.has(c) && !NEUTRAL.has(c))
}

function loose(s: string): string {
  return s.toLowerCase().replace(/[\s\p{P}\p{S}]+/gu, '')
}

/**
 * One page file against outline entry `page`. `raw` is the page file's parsed
 * JSON (for the echoed title / type / layout), `spec` the validated page.
 */
export function checkPageAgainstOutline(
  raw: unknown,
  spec: PageSpec,
  entry: OutlinePage,
  page: number,
): OutlineIssue[] {
  const issues: OutlineIssue[] = []
  const err = (message: string) => issues.push({ page, level: 'error', message })
  const warn = (message: string) => issues.push({ page, level: 'warning', message })
  const root = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  for (const field of ['title', 'type', 'layout'] as const) {
    const declared = typeof root[field] === 'string' ? (root[field] as string).trim() : ''
    if (declared && declared !== entry[field]) {
      err(
        `the page file's ${field} "${declared}" is not outline pages[${page}]'s "${entry[field]}"`,
      )
    }
  }
  const text = pageText(spec)
  const placeholder = PLACEHOLDER.exec(text)
  if (placeholder) {
    err(
      `placeholder "${placeholder[0]}" in the page text; every figure and name comes from the material`,
    )
  }
  if (entry.title && !loose(text).includes(loose(entry.title))) {
    warn(`the outline title "${entry.title}" does not appear in the page text`)
  }
  const photos = entry.image_queries.length
  if (photos > 0 && !spec.elements.some((el) => el.type === 'image')) {
    warn(`the outline plans ${photos} photo(s) for this page but it has no image element`)
  }
  return issues
}
