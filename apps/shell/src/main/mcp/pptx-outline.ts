import type { Op } from '@genoffice/pptx-ops'

/**
 * Headless pptx generation for the MCP server: outline -> op sequence.
 *
 * Pure functions over data — no DOM, no Electron, no filesystem. The ops use
 * the canonical `@genoffice/pptx-ops` vocabulary and are handed to the
 * `genoffice` CLI (`create --type pptx --ops`), which builds and saves the deck
 * through the same engine the app uses — so this module carries only the
 * ergonomic outline→ops adapter, not any deck-building code.
 *
 * Outline convention (markdown):
 *   `# Title`   starts a new slide (the text after `#` is its title)
 *   `## Text`   a bold body line on the current slide
 *   `- Text`    a bullet (indent with two leading spaces per level)
 *   `1. Text`   a numbered bullet
 *   plain line  a body paragraph without a bullet
 * The JSON format is `{ slides: [{ title, bullets }] }` with the same pieces:
 * `bullets` entries are bullets (a bare string gets a character bullet, and an
 * entry may name `bullet: "number"`), while `paragraphs` is plain text.
 */

export type PptxSourceFormat = 'markdown' | 'json'

export interface OutlineParagraph {
  text: string
  /** char = `•` bullet, number = `1.` bullet; absent = plain paragraph */
  bullet?: 'char' | 'number'
  /** nesting level for bullets (0-based); markdown indentation maps here */
  level?: number
  /** bold body line (markdown `##`) */
  bold?: boolean
}

export interface OutlineSlide {
  title?: string
  paragraphs: OutlineParagraph[]
}

/** standard 16:9 title/body placeholders in document-space EMU (matches createBlankPptx) */
const TITLE = { x: 914_400, y: 685_800, cx: 10_363_200, cy: 1_127_760 }
const BODY = { x: 914_400, y: 2_057_400, cx: 10_363_200, cy: 4_114_800 }
/** PowerPoint's default bullet geometry per level (marL/indent EMU) */
const BULLET_STEP = 342_900

export const MAX_OUTLINE_SLIDES = 100

/** Parse an outline in the given format into slides; throws a guided error on malformed input. */
export function parsePptxOutline(format: PptxSourceFormat, content: string): OutlineSlide[] {
  return format === 'json' ? parseJsonOutline(content) : parseMarkdownOutline(content)
}

function parseJsonOutline(raw: string): OutlineSlide[] {
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch (error) {
    throw new Error(
      `outline must be valid JSON when format is "json": ${error instanceof Error ? error.message : String(error)}`,
      { cause: error },
    )
  }
  const list = Array.isArray(parsed) ? parsed : (parsed as { slides?: unknown } | null)?.slides
  if (!Array.isArray(list)) {
    throw new Error('outline JSON must be an array of slides or { slides: [...] }')
  }
  return list.map((entry, index) => {
    if (typeof entry === 'string') return { paragraphs: [{ text: entry }] }
    if (typeof entry !== 'object' || entry === null) {
      throw new Error(`slide ${index} must be an object or a string`)
    }
    const obj = entry as { title?: unknown; bullets?: unknown; paragraphs?: unknown }
    // A `bullets` array means bullets: a bare string entry defaults to a
    // character bullet, the same as the markdown `- item` form. `paragraphs` is
    // the plain-text spelling and keeps no bullet.
    const isBullets = Array.isArray(obj.bullets)
    const rawBullets = isBullets ? obj.bullets : Array.isArray(obj.paragraphs) ? obj.paragraphs : []
    return {
      ...(typeof obj.title === 'string' && obj.title.trim() ? { title: obj.title.trim() } : {}),
      paragraphs: (rawBullets as unknown[]).map((b) => {
        if (typeof b === 'string')
          return isBullets ? { text: b, bullet: 'char' as const } : { text: b }
        const bl = b as { text?: unknown; level?: unknown; bullet?: unknown; bold?: unknown }
        if (typeof bl?.text !== 'string')
          throw new Error(`slide ${index}: each bullet needs a "text" string`)
        return {
          text: bl.text,
          ...(bl.bullet === 'number'
            ? { bullet: 'number' as const }
            : bl.bullet === 'char'
              ? { bullet: 'char' as const }
              : isBullets
                ? { bullet: 'char' as const }
                : {}),
          ...(typeof bl.level === 'number' && Number.isInteger(bl.level) && bl.level > 0
            ? { level: bl.level }
            : {}),
          ...(bl.bold === true ? { bold: true } : {}),
        }
      }),
    }
  })
}

function parseMarkdownOutline(raw: string): OutlineSlide[] {
  const slides: OutlineSlide[] = []
  let current: OutlineSlide | null = null
  const push = (): void => {
    if (current) slides.push(current)
  }
  for (const line of raw.split(/\r?\n/)) {
    const heading = /^(#{1,6})\s+(.*)$/.exec(line)
    if (heading) {
      const depth = heading[1]!.length
      const text = heading[2]!.trim()
      if (depth === 1) {
        push()
        current = { ...(text ? { title: text } : {}), paragraphs: [] }
      } else if (current && text) {
        current.paragraphs.push({ text, bold: true })
      }
      continue
    }
    const bullet = /^(\s*)[-*]\s+(.*)$/.exec(line)
    if (bullet) {
      current ??= { paragraphs: [] }
      const level = Math.min(Math.floor(bullet[1]!.length / 2), 4)
      current.paragraphs.push({ text: bullet[2]!, bullet: 'char', ...(level ? { level } : {}) })
      continue
    }
    const numbered = /^(\s*)\d+[.)]\s+(.*)$/.exec(line)
    if (numbered) {
      current ??= { paragraphs: [] }
      const level = Math.min(Math.floor(numbered[1]!.length / 2), 4)
      current.paragraphs.push({ text: numbered[2]!, bullet: 'number', ...(level ? { level } : {}) })
      continue
    }
    const text = line.trim()
    if (!text) continue
    current ??= { paragraphs: [] }
    current.paragraphs.push({ text })
  }
  push()
  return slides
}

/**
 * Map parsed slides to one flat op sequence.
 *
 * The CLI applies ops one by one to a blank one-slide deck, and later ops can
 * target slides earlier ops added, so a single array is enough: create the extra
 * pages first, then fill every page. (This used to be two `runTxn` batches here
 * because a single transaction plans against pre-transaction state — the CLI's
 * per-op application removes that constraint.)
 */
export function outlineToOps(slides: OutlineSlide[]): Op[] {
  if (slides.length === 0) {
    throw new Error('outline produced no slides — add at least one "# Slide title" (or JSON entry)')
  }
  if (slides.length > MAX_OUTLINE_SLIDES) {
    throw new Error(`outline has ${slides.length} slides; the limit is ${MAX_OUTLINE_SLIDES}`)
  }
  const createPages: Op[] = []
  for (let index = 1; index < slides.length; index++) {
    createPages.push({ op: 'addBlankSlide', target: { slide: index - 1 } })
  }
  const fillPages: Op[] = []
  slides.forEach((slide, index) => {
    if (slide.title) {
      fillPages.push({
        op: 'addElement',
        target: { slide: index },
        kind: 'textbox',
        offset: { ...TITLE },
        paragraphs: [{ runs: [{ text: slide.title, bold: true, fontSize: 36 }], align: 'left' }],
        bodyPr: { autoFit: 'shrink' },
      })
    }
    if (slide.paragraphs.length) {
      fillPages.push({
        op: 'addElement',
        target: { slide: index },
        kind: 'textbox',
        offset: { ...BODY },
        paragraphs: slide.paragraphs.map(bodyParagraph),
        bodyPr: { autoFit: 'shrink' },
      })
    }
  })
  return [...createPages, ...fillPages]
}

function bodyParagraph(p: OutlineParagraph): Record<string, unknown> {
  const runs = [{ text: p.text, fontSize: 20, ...(p.bold ? { bold: true } : {}) }]
  if (!p.bullet) {
    return { runs, ...(p.bold ? {} : { align: 'left' }) }
  }
  const level = p.level ?? 0
  return {
    runs,
    bullet:
      p.bullet === 'number'
        ? { type: 'number', numType: 'arabicPeriod' }
        : { type: 'char', char: '•' },
    marL: BULLET_STEP * (level + 1),
    indent: -BULLET_STEP,
    ...(level ? { level } : {}),
  }
}
