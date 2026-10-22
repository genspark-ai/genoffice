import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  markerFallbackFace,
  markerTabAdvance,
  parseDocx,
  readSections,
  sectionSettingsFromXml,
  type Block,
  type ParsedDoc,
  type SectionInfo,
  type SectionSettings,
} from '@genoffice/docx-engine'
import {
  autoSpacingPt,
  autospaceMode,
  cjkDeclaredLineFactor,
  computeLineHeight,
  cssFontFamily,
  cssLeadTop,
  fontBoxCss,
  krCharsetRoute,
  krNoBreakOffsets,
  latinAutospaceGapEm,
  latinLineFactor,
  lineHeightFactor,
  missingLatinSubstitute,
  oversizedMarkerExtraPt,
  setDocFontTable,
  setEmbeddedLineMetrics,
  simsunGapLineFactor,
  symbolBulletLinePt,
  wordMetricAliasFace,
} from '../src/renderer/line-metrics'
import { listFirstLineTw, rowHeightCss, tableRowEatCss } from '../src/renderer/editor/extensions'
import { resolveTabTarget } from '../src/renderer/editor/decoration-extensions'
import {
  cellSpacingGridSharesTwips,
  distributeContentAutofit,
  expandAutofitColWidths,
} from '../src/renderer/editor/convert'
import {
  sectionCharSpacePt,
  sectionGeoms,
  singleByteCharSpacePt,
} from '../src/renderer/pagination-sections'
import {
  computeSectionedSlicesF2,
  insertParityBlanks,
  keepDemandLines,
  type BlockBox,
  type TableRowBox,
} from '../src/renderer/pagination'

/**
 * Executable Word truths: every fixture under packages/docx-engine/tests/fixtures/truths
 * with class "pure" is a minimal synthetic docx plus the values Word for Mac laid out
 * for it (pymupdf-measured, recorded in truth.json). The cases below parse the docx,
 * resolve the referenced paragraph/row/section and run the renderer's formulas on the
 * resolved props, so a formula drifting away from a probed Word value fails here
 * before it reaches the pagination model. Render-class fixtures run in e2e/docs-truths.spec.ts.
 */

export const TRUTHS_DIR = resolve(__dirname, '../../../packages/docx-engine/tests/fixtures/truths')

interface TruthCase {
  kind: string
  /** a probed Word value the current code does not reproduce: skipped, listed as a bug */
  status?: 'failing'
  note?: string
  tol?: number
  block?: number
  section?: number
  row?: number
  font?: string
  from?: number
  to?: number
  expect: unknown
  [key: string]: unknown
}

export interface Truth {
  id: string
  title: string
  section: string
  probe_date: string
  class: 'pure' | 'render'
  status?: 'failing'
  docx?: string
  cases: TruthCase[]
}

export function loadTruths(cls: 'pure' | 'render'): Array<Truth & { dir: string }> {
  return readdirSync(TRUTHS_DIR, { withFileTypes: true })
    .filter((d) => d.isDirectory() && existsSync(join(TRUTHS_DIR, d.name, 'truth.json')))
    .map((d) => {
      const truth = JSON.parse(
        readFileSync(join(TRUTHS_DIR, d.name, 'truth.json'), 'utf8'),
      ) as Truth
      return { ...truth, dir: join(TRUTHS_DIR, d.name) }
    })
    .filter((t) => t.class === cls)
    .sort((a, b) => a.id.localeCompare(b.id))
}

const PX_PER_PT = 96 / 72
const EA_TEXT_RE = /[\u1100-\u11ff\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af\uf900-\ufaff]/

interface Resolved {
  parsed: ParsedDoc
  sections: SectionSettings[]
  sectionOfBlock: (block: number) => SectionSettings
}

async function load(dir: string, docx = 'probe.docx'): Promise<Resolved> {
  const bytes = readFileSync(join(dir, docx))
  const parsed = await parseDocx(new Uint8Array(bytes))
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml')!.async('string')
  const slices = [...xml.matchAll(/<w:sectPr[^>]*\/>|<w:sectPr[\s\S]*?<\/w:sectPr>/g)]
  const sections = slices.map((m) => sectionSettingsFromXml(m[0]))
  // a block belongs to the first sectPr at or after its own body element
  const sectionOfBlock = (block: number) => {
    const el = parsed.extras.elements[block]
    const i = el ? slices.findIndex((m) => (m.index ?? 0) + m[0].length > el.start) : -1
    return sections[i === -1 ? sections.length - 1 : i]
  }
  return { parsed, sections, sectionOfBlock }
}

function blockAt(parsed: ParsedDoc, i: number): Block {
  const b = parsed.blocks.find((x) => x.docxIndex === i)
  if (!b) throw new Error(`no block with docxIndex ${i}`)
  return b
}

interface StyleFmt {
  sizeHalfPoints?: number
  /** Latin (ascii/hAnsi) slot; the East Asian slot never sets a Latin line (word-truths L376) */
  fontFamily?: string
  eaFont?: string
  basedOn?: string
  spaceBeforeTwips?: number
  spaceAfterTwips?: number
  keepNext?: boolean
  keepLines?: boolean
  widowControl?: boolean
}

/** the paragraph style chain, starting at the default paragraph style for unstyled paragraphs */
function styleChain(parsed: ParsedDoc, styleId: string | undefined): StyleFmt[] {
  const out: StyleFmt[] = []
  let id =
    styleId ??
    [...parsed.styles.values()].find(
      (st) =>
        (st as { type?: string; isDefault?: boolean }).type === 'paragraph' &&
        (st as { isDefault?: boolean }).isDefault,
    )?.styleId
  while (id) {
    const s = parsed.styles.get(id) as
      | {
          basedOn?: string
          display?: {
            sizeHalfPoints?: number
            font?: string
            fontAscii?: string
            spaceBeforeTwips?: number
            spaceAfterTwips?: number
            keepNext?: boolean
            keepLines?: boolean
            widowControl?: boolean
          }
        }
      | undefined
    if (!s) break
    const d = s.display ?? {}
    out.push({
      basedOn: s.basedOn,
      sizeHalfPoints: d.sizeHalfPoints,
      fontFamily: d.fontAscii,
      eaFont: d.font,
      spaceBeforeTwips: d.spaceBeforeTwips,
      spaceAfterTwips: d.spaceAfterTwips,
      keepNext: d.keepNext,
      keepLines: d.keepLines,
      widowControl: d.widowControl,
    })
    id = s.basedOn
  }
  return out
}

/** the face and size the paragraph's first run (or its mark) lays out with */
function runFace(parsed: ParsedDoc, b: Block): { font: string; halfPoints: number } {
  const dd = parsed.docDefaults as
    { sizeHalfPoints?: number; asciiFont?: string; eastAsiaFont?: string } | undefined
  const run = b.runs?.find((r) => r.text.trim() !== '') ?? b.runs?.[0]
  const chain = styleChain(parsed, b.styleId)
  // Word's built-in size when nothing declares one is 10pt (probe 2026-09-23)
  const halfPoints =
    run?.sizeHalfPoints ??
    (run ? undefined : b.format?.emptyRunSizeHalfPoints) ??
    chain.find((s) => s.sizeHalfPoints)?.sizeHalfPoints ??
    dd?.sizeHalfPoints ??
    20
  // Run.font is the East Asian slot when one is declared; Latin text lays out with the ascii slot
  const ea = run && EA_TEXT_RE.test(run.text)
  const font =
    (ea
      ? (run.font ?? chain.find((s) => s.eaFont)?.eaFont ?? dd?.eastAsiaFont)
      : (run?.fontAscii ?? run?.font)) ??
    (run ? undefined : b.format?.emptyRunFontFamily) ??
    chain.find((s) => s.fontFamily)?.fontFamily ??
    dd?.asciiFont ??
    'Calibri'
  return { font, halfPoints }
}

function linePt(res: Resolved, c: TruthCase): number {
  const b = blockAt(res.parsed, c.block!)
  const face = c.font ? { font: c.font, halfPoints: undefined } : undefined
  const { font, halfPoints } = {
    ...runFace(res.parsed, b),
    ...(face?.font ? { font: face.font } : {}),
  }
  const dd = res.parsed.docDefaults as
    | { lineRule?: 'auto' | 'atLeast' | 'exact'; lineRawTwips?: number; lineSpacing?: number }
    | undefined
  const f = b.format ?? {}
  const rule = f.lineRule ?? dd?.lineRule
  const twips =
    f.lineRawTwips ??
    (f.lineSpacing !== undefined ? f.lineSpacing * 240 : undefined) ??
    dd?.lineRawTwips ??
    (dd?.lineSpacing !== undefined ? dd.lineSpacing * 240 : undefined)
  const run = b.runs?.find((r) => r.text.trim() !== '') ?? b.runs?.[0]
  const eaText = !c.font && run && EA_TEXT_RE.test(run.text)
  const factor = (eaText ? cjkDeclaredLineFactor(font) : null) ?? lineHeightFactor(font)
  const naturalPx = (halfPoints / 2) * PX_PER_PT * factor
  const grid = f.snapToGrid === false ? undefined : res.sectionOfBlock(c.block!).docGrid
  return computeLineHeight(naturalPx, rule, twips, grid) / PX_PER_PT
}

/** numeric value of the px expression rowHeightCss emits (CSS vars resolved from the table) */
function evalRowCss(css: string, vars: Record<string, number>): number {
  let expr = css.replace(/^height:/, '')
  expr = expr.replace(/var\((--[a-z-]+),([^)]*)\)/g, (_m, name: string, fallback: string) =>
    name in vars ? `${vars[name]}px` : fallback.trim(),
  )
  const num = (s: string): number => {
    s = s.trim()
    let m = /^calc\((.*)\)$/.exec(s)
    if (m) return num(m[1])
    m = /^max\((.*)\)$/.exec(s)
    if (m) return Math.max(...splitTop(m[1], ',').map(num))
    const sum = splitTop(s, '+')
    if (sum.length > 1) return sum.map(num).reduce((a, b) => a + b, 0)
    const prod = splitTop(s, '*')
    if (prod.length > 1) return prod.map(num).reduce((a, b) => a * b, 1)
    if (/^-?[\d.]+px$/.test(s)) return parseFloat(s)
    if (/^-?[\d.]+$/.test(s)) return parseFloat(s)
    throw new Error(`cannot evaluate "${s}" in ${css}`)
  }
  const splitTop = (s: string, sep: string): string[] => {
    const out: string[] = []
    let depth = 0
    let cur = ''
    for (const ch of s) {
      if (ch === '(') depth++
      if (ch === ')') depth--
      if (ch === sep && depth === 0) {
        out.push(cur)
        cur = ''
      } else cur += ch
    }
    out.push(cur)
    return out
  }
  return num(expr)
}

function rowPt(res: Resolved, c: TruthCase): number {
  const b = blockAt(res.parsed, c.block!)
  const t = b.table
  if (!t) throw new Error(`block ${c.block} is not a table`)
  const r = c.row!
  const cells = t.rows[r].filter((cell) => !cell.gridGap)
  const css = rowHeightCss(
    t.rowHeightsTwips?.[r] ?? 0,
    cells.map((cell) => cell.borders ?? null),
    cells.map((cell) => cell.cellMarTwips ?? null),
    t.rowHeightRules?.[r] ?? null,
  )
  const eat = tableRowEatCss((t.borders ?? null) as never, (t.cellSpacingTwips ?? 0) > 0)
  const eatPx = parseFloat(eat.find((s) => s.startsWith('--doc-row-eat:'))?.split(':')[1] ?? '0')
  const grid = eat.includes('--doc-row-grid:0') ? 0 : 1
  const px = evalRowCss(css, {
    '--doc-row-eat': eatPx,
    '--doc-row-grid': grid,
    '--doc-cell-pad-t': (t.cellMarTwips?.top ?? 0) / 15,
    '--doc-cell-pad-b': (t.cellMarTwips?.bottom ?? 0) / 15,
  })
  return px / PX_PER_PT
}

/** vertical space Word lays between two adjacent body paragraphs (pt) */
function paraGapPt(res: Resolved, c: TruthCase): number {
  const additive = res.parsed.doNotUseHTMLParagraphAutoSpacing === true
  const dd = res.parsed.docDefaults as
    { spaceBeforeTwips?: number; spaceAfterTwips?: number } | undefined
  const a = blockAt(res.parsed, c.from!).format ?? {}
  const b = blockAt(res.parsed, c.to!).format ?? {}
  const after = a.spaceAfterAuto
    ? autoSpacingPt(additive)
    : (a.spaceAfter ?? dd?.spaceAfterTwips ?? 0) / 20
  const before = b.spaceBeforeAuto
    ? autoSpacingPt(additive)
    : (b.spaceBefore ?? dd?.spaceBeforeTwips ?? 0) / 20
  return additive ? after + before : Math.max(after, before)
}

function tableOpenSlices(res: Resolved, c: TruthCase): number[] {
  const p = c as TruthCase & {
    heading: number
    table: number
    fillerPx: number
    headingPx: number
    rowPx: number
    contentPx: number
  }
  const heading = blockAt(res.parsed, p.heading)
  const table = blockAt(res.parsed, p.table).table!
  const rows: TableRowBox[] = table.rows.map((cells, i) => ({
    height: p.rowPx,
    contentBottom: 0,
    ...(table.repeatHeaderRows?.[i] ? { isHeader: true } : {}),
    ...(table.rowCantSplit?.[i] ? { cantSplit: true } : {}),
    ...(cells.every((cell) => cell.richParas?.every((para) => para.keepNext))
      ? { keepNext: true }
      : {}),
  }))
  const blocks: BlockBox[] = [
    { top: 0, height: p.fillerPx },
    {
      top: p.fillerPx,
      height: p.headingPx,
      ...(heading.format?.keepNext ? { keepNext: true } : {}),
    },
    {
      top: p.fillerPx + p.headingPx,
      height: rows.length * p.rowPx,
      tableRows: rows,
      modernTableHeaders: (res.parsed.compatibilityMode ?? 0) >= 15,
    },
  ]
  const total = p.fillerPx + p.headingPx + rows.length * p.rowPx
  return computeSectionedSlicesF2(
    blocks,
    [{ contentHeight: p.contentPx, forceBreak: false }],
    total,
  ).map((s) => s.start)
}

function charSpace(res: Resolved, c: TruthCase): { fullWidth: number; singleByte: number } {
  const settings = res.sections[c.section ?? res.sections.length - 1]
  const d = sectionCharSpacePt({ settings } as SectionInfo)
  return {
    fullWidth: d,
    singleByte: singleByteCharSpacePt(d, res.parsed.balanceDbcsSpacing === true),
  }
}

/** tiny CSS calc() evaluator for the glyph-shift expressions: pt numbers, + - * /, parens, max() */
function evalCalc(src: string): number {
  let i = 0
  const peek = () => src[i]
  const skip = () => {
    while (src[i] === ' ') i++
  }
  const atom = (): number => {
    skip()
    if (src.startsWith('calc(', i) || src.startsWith('max(', i)) {
      const isMax = src.startsWith('max(', i)
      i += isMax ? 4 : 5
      const args = [expr()]
      skip()
      while (peek() === ',') {
        i++
        args.push(expr())
        skip()
      }
      if (src[i++] !== ')') throw new Error(`calc: ')' expected in ${src}`)
      return isMax ? Math.max(...args) : args[0]
    }
    if (peek() === '(') {
      i++
      const v = expr()
      skip()
      if (src[i++] !== ')') throw new Error(`calc: ')' expected in ${src}`)
      return v
    }
    const m = /^-?\d+(?:\.\d+)?/.exec(src.slice(i))
    if (!m) throw new Error(`calc: number expected at "${src.slice(i)}" in ${src}`)
    i += m[0].length
    if (src.startsWith('pt', i)) i += 2
    return parseFloat(m[0])
  }
  const term = (): number => {
    let v = atom()
    for (;;) {
      skip()
      if (peek() === '*') {
        i++
        v *= atom()
      } else if (peek() === '/') {
        i++
        v /= atom()
      } else return v
    }
  }
  const expr = (): number => {
    let v = term()
    for (;;) {
      skip()
      if (peek() === '+') {
        i++
        v += term()
      } else if (peek() === '-' && src[i + 1] === ' ') {
        i++
        v -= term()
      } else return v
    }
  }
  const v = expr()
  skip()
  if (i !== src.length) throw new Error(`calc: trailing "${src.slice(i)}" in ${src}`)
  return v
}

/** pt the renderer shifts a paragraph's glyphs down by (cssLeadTop) for one paragraph's rule */
function leadTopPt(c: TruthCase): number | string {
  const p = c as TruthCase & {
    lineRule?: 'auto' | 'atLeast' | 'exact'
    lineRawTwips?: number
    lineSpacing?: number
    factor: number
    sizePt: number
  }
  const css = cssLeadTop(p.lineRule, p.lineRawTwips, p.lineSpacing)
  if (css === null || css === 'initial' || css.startsWith('var(')) return css ?? 'null'
  const expr = css
    .replace(/var\(--doc-line-grid, (calc\([^)]*\)[^)]*\))\)/g, '$1')
    .replace(/var\(--doc-line-factor,[^)]*\)/g, String(p.factor))
    .replace(/var\(--doc-lead-gap, 0px\)/g, '0pt')
    .replace(/(\d+(?:\.\d+)?)em/g, (_m, n: string) => `${parseFloat(n) * p.sizePt}pt`)
    .replace(/(\d+(?:\.\d+)?)px/g, (_m, n: string) => `${(parseFloat(n) * 72) / 96}pt`)
  return Math.round(evalCalc(expr) * 1000) / 1000
}

/** px the renderer resolves a tab at (inputs in pt, grid/stops in pt) */
function tabTargetPt(c: TruthCase): number {
  const p = c as TruthCase & {
    xPt: number
    gridPt?: number
    stopsPt?: Array<{ pos: number; val?: 'left' | 'right' | 'center' | 'decimal' }>
    segWidthPt?: number
    paraWPt?: number
    minAdvPx?: number
  }
  const r = resolveTabTarget({
    x: p.xPt * PX_PER_PT,
    minAdv: p.minAdvPx ?? 4,
    segWidth: (p.segWidthPt ?? 20) * PX_PER_PT,
    restWidth: (p.segWidthPt ?? 20) * PX_PER_PT,
    paraW: (p.paraWPt ?? 468) * PX_PER_PT,
    gridPx: (p.gridPt ?? 0) * PX_PER_PT,
    stops: (p.stopsPt ?? []).map((s) => ({ x: s.pos * PX_PER_PT, val: s.val ?? 'left' })),
  })
  return Math.round((r.target / PX_PER_PT) * 1000) / 1000
}

const EM_TEXT_RE = /\S/

/** lines a paragraph lays out with: w:br / w:cr count + 1 (one br-separated line per probe line) */
function blockLines(b: Block): number {
  const text = (b.runs ?? []).map((r) => r.text).join('')
  return (text.match(/\n/g)?.length ?? 0) + 1
}

/** resolved paragraph spacing (twips) through the direct pPr, style chain and docDefaults */
function blockSpacing(parsed: ParsedDoc, b: Block): { before: number; after: number } {
  const dd = parsed.docDefaults as
    { spaceBeforeTwips?: number; spaceAfterTwips?: number } | undefined
  const chain = styleChain(parsed, b.styleId)
  const f = b.format ?? {}
  const before =
    f.spaceBefore ??
    chain.find((s) => s.spaceBeforeTwips !== undefined)?.spaceBeforeTwips ??
    dd?.spaceBeforeTwips ??
    0
  const after =
    f.spaceAfter ??
    chain.find((s) => s.spaceAfterTwips !== undefined)?.spaceAfterTwips ??
    dd?.spaceAfterTwips ??
    0
  return { before, after }
}

/**
 * Page slicing of the probe's body paragraphs from their resolved props alone:
 * each paragraph is lines x its Word line height, the collapsed inter-paragraph
 * gap rides the previous block like measureBlocks does, and the keep/widow flags
 * come from the paragraph (style flags are folded in at parse). Returns the
 * docxIndex of the block that opens each page.
 */
function pageStartBlocks(res: Resolved, c: TruthCase): number[] {
  const blocks: BlockBox[] = []
  const body = res.parsed.blocks.filter((b) => !b.hidden && b.docxIndex !== undefined)
  let top = 0
  for (let i = 0; i < body.length; i++) {
    const b = body[i]
    const lines = blockLines(b)
    const line = linePt(res, { kind: 'linePt', block: b.docxIndex!, expect: 0 }) * PX_PER_PT
    const sp = blockSpacing(res.parsed, b)
    const next = body[i + 1]
    const gapTw = next ? Math.max(sp.after, blockSpacing(res.parsed, next).before) : sp.after
    const gap = (gapTw / 20) * PX_PER_PT
    const lead = i === 0 ? (sp.before / 20) * PX_PER_PT : 0
    const f = b.format ?? {}
    const mult =
      (f.lineRule ?? 'auto') === 'auto'
        ? (f.lineSpacing ?? (f.lineRawTwips !== undefined ? f.lineRawTwips / 240 : 1))
        : 1
    const text = (b.runs ?? []).map((r) => r.text).join('')
    const chain = styleChain(res.parsed, b.styleId)
    const flag = (k: 'keepNext' | 'keepLines' | 'widowControl') =>
      f[k] ?? chain.find((st) => st[k] !== undefined)?.[k]
    const box: BlockBox = {
      top,
      height: lead + lines * line + gap,
      spaceAfterPx: gap,
      ...(lead ? { spaceBeforePx: lead, leadFoldPx: lead } : {}),
      ...(mult > 1 ? { lineLeadPx: line * (1 - 1 / mult) } : {}),
      lineBoxes: Array.from({ length: lines }, (_, k) => ({
        offsetInBlock: lead + k * line,
        height: line,
      })),
      docxIndex: b.docxIndex ?? undefined,
      ...(flag('keepNext') ? { keepNext: true } : {}),
      ...(flag('keepLines') ? { keepLines: true } : {}),
      ...(flag('widowControl') === false ? { widowControl: false } : {}),
      ...(f.pageBreakBefore ? { breakBefore: true } : {}),
      ...(EM_TEXT_RE.test(text) ? {} : { emptyPara: true }),
    }
    blocks.push(box)
    top += box.height
  }
  const set = res.sections[res.sections.length - 1]
  const contentPx =
    ((set.pageHeight - set.marginTop - set.marginBottom) / 20) * PX_PER_PT +
    ((c.contentExtraPt as number | undefined) ?? 0) * PX_PER_PT
  const slices = computeSectionedSlicesF2(
    blocks,
    [{ contentHeight: contentPx, forceBreak: false }],
    top,
  )
  return slices.map((s) => {
    const b = blocks.find((x) => x.top <= s.start + 0.01 && x.top + x.height > s.start + 0.01)
    return b?.docxIndex ?? -1
  })
}

function parityPageCount(res: Resolved): number {
  const infos = readSections(res.parsed)
  const geoms = sectionGeoms(infos, undefined, res.parsed.evenAndOddHeaders === true)
  const slices = infos.map((_s, i) => ({ start: i, end: i + 1, section: i }))
  return insertParityBlanks(slices, geoms).length
}

describe('executable Word truths (pure formulas over parsed props)', () => {
  const truths = loadTruths('pure')
  it('has fixtures', () => {
    expect(truths.length).toBeGreaterThanOrEqual(15)
  })

  for (const t of truths) {
    describe(`${t.id} (${t.section})`, () => {
      for (const [i, c] of t.cases.entries()) {
        const failing = t.status === 'failing' || c.status === 'failing'
        const run = failing ? it.skip : it
        run(`case ${i}: ${c.kind}${c.note ? ` - ${c.note}` : ''}`, async () => {
          const res = await load(t.dir, (c.docx as string | undefined) ?? t.docx)
          setDocFontTable(res.parsed.fontTable)
          try {
            const tol = c.tol ?? 0.01
            switch (c.kind) {
              case 'linePt':
                expect(linePt(res, c)).toBeCloseToPt(c.expect as number, tol)
                break
              case 'factor':
                expect(lineHeightFactor(c.font!)).toBeCloseToPt(c.expect as number, c.tol ?? 0.0001)
                break
              case 'krRoute':
                expect(krCharsetRoute(c.font!)).toBe(c.expect)
                break
              case 'latinSubstitute':
                expect(missingLatinSubstitute(c.font!)).toBe(c.expect)
                break
              case 'rowPt':
                expect(rowPt(res, c)).toBeCloseToPt(c.expect as number, tol)
                break
              case 'paraGapPt':
                expect(paraGapPt(res, c)).toBeCloseToPt(c.expect as number, tol)
                break
              case 'charSpacePt': {
                const got = charSpace(res, c)
                const want = c.expect as { fullWidth: number; singleByte: number }
                expect(got.fullWidth).toBeCloseToPt(want.fullWidth, c.tol ?? 0.001)
                expect(got.singleByte).toBeCloseToPt(want.singleByte, c.tol ?? 0.001)
                break
              }
              case 'format':
                expect(blockAt(res.parsed, c.block!).format ?? {}).toMatchObject(c.expect as object)
                break
              case 'flags':
                expect(res.parsed).toMatchObject(c.expect as object)
                break
              case 'flagsOff':
                for (const k of c.expect as string[])
                  expect((res.parsed as unknown as Record<string, unknown>)[k]).toBeFalsy()
                break
              case 'tableOpenSlices':
                expect(tableOpenSlices(res, c)).toEqual(c.expect)
                break
              case 'cjkFactor':
                expect(cjkDeclaredLineFactor(c.font!)).toBe(c.expect)
                break
              case 'simsunGap':
                expect(simsunGapLineFactor(c.font!)).toBe(c.expect)
                break
              case 'gapEm':
                expect(latinAutospaceGapEm(c.font!)).toBeCloseToPt(
                  c.expect as number,
                  c.tol ?? 0.001,
                )
                break
              case 'latinLineFactor':
                expect(latinLineFactor(c.font!)).toBeCloseToPt(c.expect as number, c.tol ?? 0.0001)
                break
              case 'aliasFace':
                expect(wordMetricAliasFace(c.font!)).toBe(c.expect)
                break
              case 'cssFamilyIncludes':
                expect(cssFontFamily(c.font!)).toContain(c.expect)
                break
              case 'krNoBreakIncludes': {
                const got = krNoBreakOffsets(c.text as string, (c.prev as string) ?? '')
                for (const off of c.expect as number[]) expect(got).toContain(off)
                break
              }
              case 'markerTab': {
                const a = c.args as [number, number, number, number?, number[]?]
                expect(markerTabAdvance(...a)).toBe(c.expect)
                break
              }
              case 'markerFallback':
                expect(markerFallbackFace(c.glyph as string, c.face as string)).toBe(c.expect)
                break
              case 'symbolBulletLinePt': {
                const p = c as TruthCase & {
                  symbolFont: string
                  markerPt: number
                  textFont: string
                  textPt: number
                }
                expect(
                  symbolBulletLinePt(p.symbolFont, p.markerPt, p.textFont, p.textPt),
                ).toBeCloseToPt(c.expect as number, tol)
                break
              }
              case 'oversizedMarkerExtraPt': {
                const p = c as TruthCase & { markerFont: string; markerPt: number; textPt: number }
                const got = oversizedMarkerExtraPt(p.markerFont, p.markerPt, p.textPt)
                if (c.expect === null) expect(got).toBeNull()
                else expect(got!).toBeCloseToPt(c.expect as number, tol)
                break
              }
              case 'keepDemand':
                expect(keepDemandLines(c.widowOn as boolean, c.nLines as number)).toBe(c.expect)
                break
              case 'distribute': {
                const p = c as TruthCase & { tableTwips: number; mins: number[]; maxes: number[] }
                const got = distributeContentAutofit(p.tableTwips, p.mins, p.maxes)
                const want = c.expect as number[]
                expect(got.length).toBe(want.length)
                got.forEach((w, k) => expect(w).toBeCloseToPt(want[k], c.tol ?? 1))
                break
              }
              case 'cellSpacingShares':
                expect(cellSpacingGridSharesTwips(c.cols as number, c.spacing as number)).toEqual(
                  c.expect,
                )
                break
              case 'tableCols':
                expect(blockAt(res.parsed, c.block!).table?.colWidthsTwips).toEqual(c.expect)
                break
              case 'tableProps':
                expect(blockAt(res.parsed, c.block!).table).toMatchObject(c.expect as object)
                break
              case 'cellProps':
                expect(
                  blockAt(res.parsed, c.block!).table!.rows[c.row!][c.cell as number],
                ).toMatchObject(c.expect as object)
                break
              case 'autofitCols': {
                const model = blockAt(res.parsed, c.block!).table!
                const avail = c.availTwips as number
                const got = expandAutofitColWidths(model, avail, avail, null).colWidthsTwips ?? []
                const want = c.expect as number[] | { sum: number }
                if (Array.isArray(want)) {
                  expect(got.length).toBe(want.length)
                  got.forEach((w, k) => expect(w).toBeCloseToPt(want[k], c.tol ?? 1))
                } else {
                  expect(got.reduce((a, b) => a + b, 0)).toBeCloseToPt(want.sum, c.tol ?? 1)
                }
                break
              }
              case 'leadTopPt': {
                const got = leadTopPt(c)
                if (typeof c.expect === 'string') expect(got).toBe(c.expect)
                else expect(got as number).toBeCloseToPt(c.expect as number, tol)
                break
              }
              case 'fontBox':
                expect(fontBoxCss(c.chain as string)).toEqual(c.expect)
                break
              case 'embeddedFactor': {
                const p = c as TruthCase & {
                  family: string
                  ascent: number
                  descent: number
                  lineGap: number
                }
                setEmbeddedLineMetrics([p])
                try {
                  expect(lineHeightFactor(p.family)).toBeCloseToPt(
                    c.expect as number,
                    c.tol ?? 0.0001,
                  )
                } finally {
                  setEmbeddedLineMetrics([])
                }
                break
              }
              case 'tabTargetPt':
                expect(tabTargetPt(c)).toBeCloseToPt(c.expect as number, tol)
                break
              case 'styleProps': {
                const st = res.parsed.styles.get(c.styleId as string) as
                  (Record<string, unknown> & { display?: Record<string, unknown> }) | undefined
                expect(st).toBeDefined()
                expect({ ...st!.display, ...st }).toMatchObject(c.expect as object)
                break
              }
              case 'docDefaults':
                expect(res.parsed.docDefaults ?? {}).toMatchObject(c.expect as object)
                break
              case 'runProps':
                expect(blockAt(res.parsed, c.block!).runs?.[c.run as number]).toMatchObject(
                  c.expect as object,
                )
                break
              case 'blockProps':
                expect(blockAt(res.parsed, c.block!)).toMatchObject(c.expect as object)
                break
              case 'blockText':
                expect((blockAt(res.parsed, c.block!).runs ?? []).map((r) => r.text).join('')).toBe(
                  c.expect,
                )
                break
              case 'sectionProps':
                expect(res.sections[c.section ?? 0]).toMatchObject(c.expect as object)
                break
              case 'numberingLevel':
                expect(
                  res.parsed.numbering.get(String(c.numId))?.levels[(c.ilvl as number) ?? 0],
                ).toMatchObject(c.expect as object)
                break
              case 'listMarkerStartTw': {
                const p = c as TruthCase & {
                  own: number | null
                  level: { indentLeft?: number; hanging?: number; firstLine?: number }
                  styleFirstLine?: number
                  left: number
                }
                const first = listFirstLineTw(p.own, p.level, p.styleFirstLine)
                expect(first === null ? null : p.left + first).toBe(c.expect)
                break
              }
              case 'blockFace':
                expect(runFace(res.parsed, blockAt(res.parsed, c.block!)).font).toBe(c.expect)
                break
              case 'autospaceMode':
                expect(autospaceMode(blockAt(res.parsed, c.block!).format?.autoSpace)).toBe(
                  c.expect,
                )
                break
              case 'pageStartBlocks':
                expect(pageStartBlocks(res, c)).toEqual(c.expect)
                break
              case 'parityPages':
                expect(parityPageCount(res)).toBe(c.expect)
                break
              default:
                throw new Error(`unknown truth case kind "${c.kind}" in ${t.id}`)
            }
          } finally {
            setDocFontTable(null)
          }
        })
      }
    })
  }
})

declare module 'vitest' {
  interface Assertion<T> {
    toBeCloseToPt(expected: number, tol: number): T
  }
}

expect.extend({
  toBeCloseToPt(received: number, expected: number, tol: number) {
    const pass = Math.abs(received - expected) <= tol + 1e-9
    return {
      pass,
      message: () =>
        `expected ${received.toFixed(4)} to be within ${tol} of ${expected} (delta ${(received - expected).toFixed(4)})`,
    }
  },
})
