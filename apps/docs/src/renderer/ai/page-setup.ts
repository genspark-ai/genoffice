import {
  applyPageNumType,
  applySectionSettings,
  applyTitlePg,
  type SectionInfo,
  type SectionSettings,
} from '@genoffice/docx-engine'
import type { AgentToolDef } from '../../shared/ipc'

/**
 * Page setup as the AI sees it: one entry per section, block ranges in live
 * top-level PM indexes (what every other tool addresses), lengths in twips
 * plus a readable label. Both the app and the headless CLI build these from
 * their own section stores and hand them to the tools through
 * AiPageSetupAccess.
 */

export type SectionStartType = SectionInfo['startType']
export type SectionBreakType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'

export interface AiSectionState {
  index: number
  firstBlock: number
  lastBlock: number
  paper: string
  /** twips */
  width: number
  height: number
  orientation: 'portrait' | 'landscape'
  /** twips; left/top are the raw w:pgMar values (gutter reported separately) */
  margins: {
    top: number
    right: number
    bottom: number
    left: number
    header: number
    footer: number
    gutter: number
  }
  columns: { count: number; spacing: number }
  titlePg: boolean
  startType: SectionStartType
  pageNumberStart?: number
  pageNumberFormat?: string
}

/** What a set_page_setup call resolved to for one section; only present fields change. */
export interface ResolvedPageSetup {
  settings: SectionSettings
  titlePg?: boolean
  /** page numbering (w:pgNumType); {} removes the tag; undefined = keep */
  pgNum?: { fmt?: string; start?: number }
}

export interface AiPageSetupAccess {
  list(): AiSectionState[]
  /** the engine-level section (settings + sectPr) a set_page_setup patch merges into */
  current(index: number): SectionInfo | undefined
  /** returns an error message, or null on success */
  set(index: number, resolved: ResolvedPageSetup): string | null
  /** insert a section break after a top-level block (-1 = before the first block) */
  insertBreak(type: SectionBreakType, afterBlockIndex: number): string | null
}

export const PAPER_SIZES: ReadonlyArray<{ name: string; w: number; h: number }> = [
  { name: 'A3', w: 16838, h: 23811 },
  { name: 'A4', w: 11906, h: 16838 },
  { name: 'A5', w: 8391, h: 11906 },
  { name: 'B5', w: 10319, h: 14572 },
  { name: 'Letter', w: 12240, h: 15840 },
  { name: 'Legal', w: 12240, h: 20160 },
  { name: 'Tabloid', w: 15840, h: 24480 },
]

const TWIPS_PER: Record<string, number> = {
  twip: 1,
  twips: 1,
  pt: 20,
  px: 15,
  in: 1440,
  cm: 566.929,
  mm: 56.6929,
}

const LENGTH = /^\s*(-?\d+(?:\.\d+)?)\s*(twips?|pt|px|in|cm|mm)\s*$/i

/** "2.54cm" / "1in" / "72pt" / 1440 (twips) → twips; undefined when unparseable */
/** Largest magnitude accepted (~35in in twips): uncapped AI lengths break layout. */
const MAX_TWIPS = 50400

function boundTwips(twips: number): number | undefined {
  if (!Number.isFinite(twips)) return undefined
  const rounded = Math.round(twips)
  return Math.abs(rounded) <= MAX_TWIPS ? rounded : undefined
}

export function parseTwips(value: unknown): number | undefined {
  if (typeof value === 'number') return boundTwips(value)
  if (typeof value !== 'string') return undefined
  const m = LENGTH.exec(value)
  if (!m) return undefined
  return boundTwips(Number(m[1]) * TWIPS_PER[m[2]!.toLowerCase()]!)
}

export function twipsToCm(twips: number): string {
  return `${(twips / 566.929).toFixed(2).replace(/\.?0+$/, '')}cm`
}

/** paper name when the size matches a known sheet in either orientation, else the size in cm */
export function paperName(width: number, height: number): string {
  const [w, h] = width <= height ? [width, height] : [height, width]
  const hit = PAPER_SIZES.find((p) => Math.abs(p.w - w) <= 2 && Math.abs(p.h - h) <= 2)
  return hit ? hit.name : `${twipsToCm(width)} x ${twipsToCm(height)}`
}

export function describeSection(
  info: SectionInfo,
  index: number,
  firstBlock: number,
  lastBlock: number,
): AiSectionState {
  const s = info.settings
  const gutter = s.gutter ?? 0
  return {
    index,
    firstBlock,
    lastBlock,
    paper: paperName(s.pageWidth, s.pageHeight),
    width: s.pageWidth,
    height: s.pageHeight,
    orientation: s.orientation,
    margins: {
      top: s.marginTop - (s.gutterAtTop ? gutter : 0),
      right: s.marginRight,
      bottom: s.marginBottom,
      left: s.marginLeft - (s.gutterAtTop ? 0 : gutter),
      header: s.headerDist ?? 720,
      footer: s.footerDist ?? 720,
      gutter,
    },
    columns: { count: s.columns, spacing: s.colSpace ?? 720 },
    titlePg: info.titlePg,
    startType: info.startType,
    ...(info.pageNumberStart !== undefined ? { pageNumberStart: info.pageNumberStart } : {}),
    ...(info.pageNumberFmt !== undefined ? { pageNumberFormat: info.pageNumberFmt } : {}),
  }
}

export function sectionLine(s: AiSectionState): string {
  const m = s.margins
  const cm = twipsToCm
  const parts = [
    `${s.paper} ${s.orientation}`,
    `margins top ${cm(m.top)} right ${cm(m.right)} bottom ${cm(m.bottom)} left ${cm(m.left)}`,
  ]
  if (m.gutter) parts.push(`gutter ${cm(m.gutter)}`)
  if (s.columns.count > 1) parts.push(`${s.columns.count} columns`)
  if (s.titlePg) parts.push('different first page')
  if (s.pageNumberStart !== undefined) parts.push(`page numbers start at ${s.pageNumberStart}`)
  if (s.pageNumberFormat) parts.push(`number format ${s.pageNumberFormat}`)
  if (s.index > 0 && s.startType !== 'nextPage') parts.push(`starts ${s.startType}`)
  return parts.join(', ')
}

/** model context: one line per section (a single section stays one line) */
export function pageSetupContextLines(sections: AiSectionState[]): string[] {
  if (sections.length === 0) return []
  if (sections.length === 1)
    return [`Page setup (change with set_page_setup): ${sectionLine(sections[0]!)}`]
  return [
    `Page setup: ${sections.length} sections (set_page_setup takes section or blockIndex; insert_section_break adds one):`,
    ...sections.map(
      (s) => `- section ${s.index} (blocks ${s.firstBlock}-${s.lastBlock}): ${sectionLine(s)}`,
    ),
  ]
}

/** index of the section owning a top-level block; the last section for anything past the end */
export function sectionIndexOfBlock(sections: AiSectionState[], blockIndex: number): number {
  const i = sections.findIndex((s) => blockIndex <= s.lastBlock)
  return i >= 0 ? i : sections.length - 1
}

const PAGE_NUMBER_FORMATS = new Set([
  'decimal',
  'lowerRoman',
  'upperRoman',
  'lowerLetter',
  'upperLetter',
  'decimalEnclosedCircle',
  'decimalFullWidth',
  'chineseCounting',
  'japaneseCounting',
  'koreanCounting',
  'arabicAlpha',
  'hebrew1',
  'thaiNumbers',
])

const LENGTH_HINT = 'a number of twips or a string with a unit: "2.54cm", "1in", "72pt", "10mm"'

export const PAGE_SETUP_TOOL: AgentToolDef = {
  name: 'set_page_setup',
  description:
    'Change paper size, orientation, margins, text columns, the different-first-page flag or page numbering of the document or of one section (the message context lists the sections). ' +
    'Only the fields given change. Lengths take a unit suffix (cm, mm, in, pt) or twips. Setting the orientation alone swaps the current width and height.',
  inputSchema: {
    type: 'object',
    properties: {
      section: {
        type: 'integer',
        description:
          'section index (0-based) to change; omit with blockIndex omitted = every section',
      },
      blockIndex: {
        type: 'integer',
        description: 'change the section that contains this block index',
      },
      paper: {
        type: 'string',
        description: `named size: ${PAPER_SIZES.map((p) => p.name).join(', ')}`,
      },
      width: { type: ['string', 'number'], description: `custom page width, ${LENGTH_HINT}` },
      height: { type: ['string', 'number'], description: `custom page height, ${LENGTH_HINT}` },
      orientation: { type: 'string', enum: ['portrait', 'landscape'] },
      margins: {
        type: 'object',
        description: `any of top, right, bottom, left, header, footer, gutter; each ${LENGTH_HINT}`,
        properties: {
          top: { type: ['string', 'number'] },
          right: { type: ['string', 'number'] },
          bottom: { type: ['string', 'number'] },
          left: { type: ['string', 'number'] },
          header: { type: ['string', 'number'], description: 'header distance from the page top' },
          footer: {
            type: ['string', 'number'],
            description: 'footer distance from the page bottom',
          },
          gutter: { type: ['string', 'number'], description: 'extra binding margin' },
        },
      },
      columns: {
        type: 'object',
        description: 'text columns of the section',
        properties: {
          count: { type: 'integer', description: '1-12' },
          spacing: {
            type: ['string', 'number'],
            description: `gap between columns, ${LENGTH_HINT}`,
          },
        },
        required: ['count'],
      },
      titlePg: {
        type: 'boolean',
        description: 'different first page (own header/footer on the first page of the section)',
      },
      pageNumberStart: {
        type: ['integer', 'null'],
        description:
          'restart page numbering at this value; null = continue from the previous section',
      },
      pageNumberFormat: {
        type: 'string',
        description: 'decimal, lowerRoman, upperRoman, lowerLetter, upperLetter, ...',
      },
    },
  },
}

export const SECTION_BREAK_TOOL: AgentToolDef = {
  name: 'insert_section_break',
  description:
    'Insert a section break after a block so the blocks after it form a new section with its own page setup, headers and footers (the new section starts as a copy of the current one; change it with set_page_setup afterwards). ' +
    'type nextPage starts the new section on a new page, continuous keeps the flow (needed for column changes mid-page), evenPage / oddPage start on the next even / odd page.',
  inputSchema: {
    type: 'object',
    properties: {
      afterBlockIndex: {
        type: 'integer',
        description: 'the break goes after this block index; -1 = before the first block',
      },
      type: {
        type: 'string',
        enum: ['nextPage', 'continuous', 'evenPage', 'oddPage'],
        description: 'default nextPage',
      },
    },
    required: ['afterBlockIndex'],
  },
}

type Input = Record<string, unknown>

function lengthField(input: Input, key: string, min = 0): number | string | undefined {
  const raw = input[key]
  if (raw === undefined || raw === null) return undefined
  const v = parseTwips(raw)
  if (v === undefined) return `${key} must be ${LENGTH_HINT}`
  if (v < min) return `${key} must be at least ${min}`
  return v
}

/**
 * Merge a set_page_setup input into one section's current settings. Returns an
 * error message for anything the schema cannot express (unknown paper, margins
 * that leave no room for text, ...); an empty input is an error too.
 */
export function resolvePageSetup(
  input: Input,
  current: SectionInfo,
): ResolvedPageSetup | { error: string } {
  const cur = current.settings
  let width = cur.pageWidth
  let height = cur.pageHeight
  let touched = false

  if (input.paper !== undefined) {
    const name = String(input.paper)
      .trim()
      .toLowerCase()
      .replace(/\s*\(jis\)$/, '')
    const hit = PAPER_SIZES.find((p) => p.name.toLowerCase() === name)
    if (!hit) {
      return {
        error: `unknown paper "${input.paper}"; use ${PAPER_SIZES.map((p) => p.name).join(', ')} or width/height`,
      }
    }
    ;[width, height] = cur.orientation === 'landscape' ? [hit.h, hit.w] : [hit.w, hit.h]
    touched = true
  }
  for (const key of ['width', 'height'] as const) {
    const v = lengthField(input, key, 1)
    if (typeof v === 'string') return { error: v }
    if (v !== undefined) {
      if (key === 'width') width = v
      else height = v
      touched = true
    }
  }
  let orientation = cur.orientation
  if (input.orientation !== undefined) {
    if (input.orientation !== 'portrait' && input.orientation !== 'landscape') {
      return { error: 'orientation must be "portrait" or "landscape"' }
    }
    orientation = input.orientation
    const landscapeNow = width > height
    if ((orientation === 'landscape') !== landscapeNow && width !== height) {
      ;[width, height] = [height, width]
    }
    touched = true
  } else if (input.paper !== undefined || input.width !== undefined || input.height !== undefined) {
    orientation = width > height ? 'landscape' : 'portrait'
  }

  const gutter0 = cur.gutter ?? 0
  const margins = {
    top: cur.marginTop - (cur.gutterAtTop ? gutter0 : 0),
    right: cur.marginRight,
    bottom: cur.marginBottom,
    left: cur.marginLeft - (cur.gutterAtTop ? 0 : gutter0),
    header: cur.headerDist ?? 720,
    footer: cur.footerDist ?? 720,
    gutter: gutter0,
  }
  let marginTopFixed = cur.marginTopFixed
  let marginBottomFixed = cur.marginBottomFixed
  if (input.margins !== undefined) {
    if (!input.margins || typeof input.margins !== 'object' || Array.isArray(input.margins)) {
      return {
        error:
          'margins must be an object with any of top, right, bottom, left, header, footer, gutter',
      }
    }
    const m = input.margins as Input
    for (const key of Object.keys(m)) {
      if (!(key in margins))
        return {
          error: `margins.${key} is not a margin (top, right, bottom, left, header, footer, gutter)`,
        }
      const v = lengthField(m, key)
      if (typeof v === 'string') return { error: `margins.${v}` }
      if (v === undefined) continue
      margins[key as keyof typeof margins] = v
      if (key === 'top') marginTopFixed = undefined
      if (key === 'bottom') marginBottomFixed = undefined
      touched = true
    }
  }
  if (margins.left + margins.right + margins.gutter >= width) {
    return {
      error: `left + right margins (${twipsToCm(margins.left + margins.right + margins.gutter)}) leave no room on a ${twipsToCm(width)} wide page`,
    }
  }
  if (margins.top + margins.bottom >= height) {
    return {
      error: `top + bottom margins (${twipsToCm(margins.top + margins.bottom)}) leave no room on a ${twipsToCm(height)} tall page`,
    }
  }

  let columns = cur.columns
  let colSpace = cur.colSpace
  let colWidths = cur.colWidths
  if (input.columns !== undefined) {
    const c = input.columns as Input
    if (!c || typeof c !== 'object' || !Number.isInteger(c.count)) {
      return { error: 'columns must be { count, spacing? } with an integer count' }
    }
    const count = c.count as number
    if (count < 1 || count > 12) return { error: 'columns.count must be between 1 and 12' }
    const spacing = lengthField(c, 'spacing')
    if (typeof spacing === 'string') return { error: `columns.${spacing}` }
    if (count !== columns) colWidths = undefined
    columns = count
    if (spacing !== undefined) colSpace = spacing
    touched = true
  }

  const out: ResolvedPageSetup = {
    settings: {
      ...cur,
      pageWidth: width,
      pageHeight: height,
      orientation,
      marginTop: margins.top + (cur.gutterAtTop ? margins.gutter : 0),
      marginRight: margins.right,
      marginBottom: margins.bottom,
      marginLeft: margins.left + (cur.gutterAtTop ? 0 : margins.gutter),
      marginTopFixed,
      marginBottomFixed,
      gutter: margins.gutter,
      headerDist: margins.header,
      footerDist: margins.footer,
      columns,
      colSpace,
      colWidths,
    },
  }
  if (input.titlePg !== undefined) {
    if (typeof input.titlePg !== 'boolean') return { error: 'titlePg must be true or false' }
    out.titlePg = input.titlePg
    touched = true
  }
  if (input.pageNumberStart !== undefined || input.pageNumberFormat !== undefined) {
    let start = current.pageNumberStart
    let fmt = current.pageNumberFmt
    if (input.pageNumberStart !== undefined) {
      if (input.pageNumberStart === null) start = undefined
      else if (Number.isInteger(input.pageNumberStart) && (input.pageNumberStart as number) >= 0) {
        start = input.pageNumberStart as number
      } else return { error: 'pageNumberStart must be a non-negative integer or null' }
    }
    if (input.pageNumberFormat !== undefined) {
      if (!PAGE_NUMBER_FORMATS.has(String(input.pageNumberFormat))) {
        return { error: `pageNumberFormat must be one of ${[...PAGE_NUMBER_FORMATS].join(', ')}` }
      }
      fmt = String(input.pageNumberFormat)
    }
    out.pgNum = { ...(fmt !== undefined ? { fmt } : {}), ...(start !== undefined ? { start } : {}) }
    touched = true
  }
  if (!touched) {
    return {
      error:
        'nothing to change: give paper, width/height, orientation, margins, columns, titlePg or page numbering',
    }
  }
  return out
}

/** Rewrite one sectPr XML slice with a resolved page setup (headless store and section-break paragraphs). */
export function applyResolvedPageSetup(sectPrXml: string, resolved: ResolvedPageSetup): string {
  let xml = applySectionSettings(sectPrXml, resolved.settings)
  if (resolved.titlePg !== undefined) xml = applyTitlePg(xml, resolved.titlePg)
  if (resolved.pgNum) xml = applyPageNumType(xml, resolved.pgNum.fmt, resolved.pgNum.start)
  return xml
}

export const SECTION_BREAK_TYPES: ReadonlyArray<SectionBreakType> = [
  'nextPage',
  'continuous',
  'evenPage',
  'oddPage',
]

/** the sectPr a new section-break paragraph carries: the current section's, minus nothing (Word copies it whole) */
export function sectionBreakParagraphXml(sectPrXml: string): string {
  return `<w:p><w:pPr>${sectPrXml}</w:pPr></w:p>`
}
