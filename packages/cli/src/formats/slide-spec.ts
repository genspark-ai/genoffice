import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { basename, extname, join, resolve } from 'node:path'
import {
  auditDeck,
  buildDeckPptx,
  checkPageAgainstOutline,
  offPaletteColors,
  parseDeckSpec,
  parseOutline,
  stylePalette,
  type BuildPageDeps,
  type DeckAuditPage,
  type DeckOutline,
  type DeckPageIssue,
  type OutlineIssue,
  type PageSpec,
} from '@genoffice/pipelines/slides'
import {
  deleteSlide,
  mergeSlideFromPptx,
  moveSlide,
  promoteSlideBackground,
  type OpenedPptx,
} from '@genoffice/pptx-engine'
import { HeuristicMetrics } from '@genoffice/pptx-render'
import { extract } from '@genoffice/pdf2docx'
import { assertAllowed, type PathContext } from '../fs'
import { CliError, EXIT } from '../result'
import { imageSize } from './image-size'
import { readImageSource } from './image-source'
import { loadPdfium } from './pdf'

/**
 * The app measures text with the system fonts it indexes at startup; the CLI
 * has no font index yet, so boxes grow and the audit measures with heuristic
 * glyph widths. Close for Latin and CJK body text, looser for display fonts.
 */
const metrics = new HeuristicMetrics()

export interface BuiltDeck {
  bytes: Uint8Array
  pages: number
  issues: DeckPageIssue[]
  imageFailures: { page: number; url: string }[]
}

function specDeps(ctx: PathContext, specDir?: string): BuildPageDeps {
  return {
    fetchImage: async (url) => {
      const r = await readImageSource(url, ctx, specDir)
      return r ? { bytes: r.bytes, ext: r.ext } : null
    },
    imageDims: (bytes) => {
      const s = imageSize(bytes)
      return s ? { width: s.width, height: s.height } : null
    },
    fontMetrics: metrics,
  }
}

/**
 * Deck spec JSON → pptx bytes. Image sources are local paths (relative to the
 * current directory, then to the spec file's folder), data: URLs or http(s) URLs.
 */
export async function buildDeckFromSpec(
  text: string,
  source: string,
  ctx: PathContext,
  specDir?: string,
): Promise<BuiltDeck> {
  const parsed = parseDeckSpec(text, { localImages: true })
  if (!parsed.ok) throw new CliError(EXIT.usage, `${source}: ${parsed.error}`)
  const built = await buildDeckPptx(parsed.spec, specDeps(ctx, specDir))
  return {
    bytes: built.bytes,
    pages: parsed.spec.pages.length,
    issues: withImageWarnings(parsed.issues, inputIndexed(built.imageWarnings, parsed.issues)),
    imageFailures: inputIndexed(built.imageFailures, parsed.issues),
  }
}

/** Build-time image warnings (heavy cover crops) join the page's parse warnings so one list names everything wrong with a page. */
function withImageWarnings(
  issues: DeckPageIssue[],
  imageWarnings: { page: number; message: string }[],
): DeckPageIssue[] {
  if (imageWarnings.length === 0) return issues
  const out = issues.map((i) => ({ ...i, ...(i.warnings ? { warnings: [...i.warnings] } : {}) }))
  for (const w of imageWarnings) {
    let issue = out.find((i) => i.page === w.page)
    if (!issue) {
      issue = { page: w.page }
      out.push(issue)
    }
    issue.warnings = [...(issue.warnings ?? []), w.message]
  }
  return out.sort((a, b) => a.page - b.page)
}

/**
 * The builder numbers only the pages that survived parsing; callers reason in
 * input page numbers (the ones `issues` use), so map the failures back.
 */
function inputIndexed<T extends { page: number }>(failures: T[], issues: DeckPageIssue[]): T[] {
  const dropped = new Set(issues.filter((i) => i.error).map((i) => i.page))
  const maxPage = Math.max(-1, ...failures.map((f) => f.page))
  const kept: number[] = []
  for (let i = 0; kept.length <= maxPage; i++) if (!dropped.has(i)) kept.push(i)
  return failures.map((f) => ({ ...f, page: kept[f.page] ?? f.page }))
}

/** The `*.json` page files of a spec directory in name order: one file is one slide; an outline kept in the same folder is not a page. */
export function listPageFiles(dir: string): string[] {
  return readdirSync(dir)
    .filter(
      (f) =>
        extname(f).toLowerCase() === '.json' &&
        !f.startsWith('.') &&
        f.toLowerCase() !== 'outline.json',
    )
    .sort((a, b) => a.localeCompare(b, 'en', { numeric: true }))
    .map((f) => join(dir, f))
}

/**
 * The other files of the staged deck flow, looked up beside a page (or pages
 * folder) and one level up: `deck/pages/01.json` finds `deck/outline.json` and
 * `deck/style.md`.
 */
export interface StageFiles {
  outline?: string
  style?: string
}

export function findStageFiles(dir: string, ctx: PathContext): StageFiles {
  const find = (name: string) =>
    [join(dir, name), join(dir, '..', name)].find((p) => existsSync(p) && statSync(p).isFile())
  const out: StageFiles = {}
  const outline = find('outline.json')
  if (outline) out.outline = assertAllowed(resolve(outline), ctx.env, 'read')
  const style = find('style.md')
  if (style) out.style = assertAllowed(resolve(style), ctx.env, 'read')
  return out
}

/** What the page is checked against; the command layer resolves the files, this layer reads them. */
export interface StageContext {
  outline?: { file: string; outline: DeckOutline }
  style?: { file: string; palette: Set<string> }
  /** guidance when a stage file is missing */
  notes: string[]
}

export function stageContext(
  dir: string,
  ctx: PathContext,
  explicitOutline?: string,
): StageContext {
  const found = findStageFiles(dir, ctx)
  const outlineFile = explicitOutline ?? found.outline
  const out: StageContext = { notes: [] }
  if (outlineFile) out.outline = { file: outlineFile, outline: readOutline(outlineFile) }
  else {
    out.notes.push(
      'no outline.json beside the page files or one folder up: the staged flow writes deck/outline.json first and checks every page against it (genoffice guide slides design)',
    )
  }
  if (found.style) {
    out.style = { file: found.style, palette: stylePalette(readFileSync(found.style, 'utf-8')) }
    if (out.style.palette.size === 0) {
      out.notes.push(
        `${basename(found.style)} names no #RRGGBB colors, so the palette is not checked`,
      )
    }
  } else {
    out.notes.push(
      'no style.md beside the page files or one folder up: write the style sheet (colors, fonts, layout library) before the pages',
    )
  }
  return out
}

export interface StageFindings {
  outline: OutlineIssue[]
  offPalette: string[]
}

/** One parsed page against its outline entry and the style sheet's palette. */
export function stageFindings(
  raw: unknown,
  spec: PageSpec,
  page: number,
  stage: StageContext,
): StageFindings {
  const outline: OutlineIssue[] = []
  if (stage.outline) {
    const entry = stage.outline.outline.pages[page]
    if (entry) outline.push(...checkPageAgainstOutline(raw, spec, entry, page))
    else {
      outline.push({
        page,
        level: 'error',
        message: `no outline entry for page ${page} (the outline has ${stage.outline.outline.pages.length} pages)`,
      })
    }
  }
  return {
    outline,
    offPalette: stage.style ? offPaletteColors(spec, stage.style.palette) : [],
  }
}

export interface BuiltDeckFromDir extends BuiltDeck {
  files: string[]
  stage: StageContext
  /** outline findings (warnings only; errors abort the build) keyed by page file */
  outlineFindings: (OutlineIssue & { file: string })[]
  offPalette: { file: string; colors: string[] }[]
}

/**
 * A directory of single-page spec files → one pptx. The files are the unit an
 * agent writes and checks one at a time; issues name the file they came from.
 * Every file is checked against its outline entry; a disagreement aborts.
 */
export async function buildDeckFromDir(
  dir: string,
  ctx: PathContext,
  stage: StageContext,
): Promise<BuiltDeckFromDir> {
  const files = listPageFiles(dir)
  if (files.length === 0) throw new CliError(EXIT.usage, `${dir}: no page files (*.json)`)
  const outline = stage.outline?.outline
  if (outline && outline.pages.length !== files.length) {
    throw new CliError(
      EXIT.usage,
      `${dir}: ${files.length} page file(s) for ${outline.pages.length} outline page(s); write every page before building`,
      { files: files.map((f) => basename(f)), outline_pages: outline.pages.map((p) => p.title) },
    )
  }
  const pages: unknown[] = []
  for (const file of files) {
    try {
      pages.push(JSON.parse(readFileSync(file, 'utf-8')))
    } catch (err) {
      throw new CliError(EXIT.usage, `${file}: invalid JSON (${(err as Error).message})`)
    }
  }
  const parsed = parseDeckSpec(JSON.stringify({ pages }), { localImages: true })
  if (!parsed.ok) throw new CliError(EXIT.usage, `${dir}: ${parsed.error}`)
  const named = (i: number) => ({ file: basename(files[i]!) })

  const outlineFindings: BuiltDeckFromDir['outlineFindings'] = []
  const offPalette: BuiltDeckFromDir['offPalette'] = []
  const dropped = new Set(parsed.issues.filter((i) => i.error).map((i) => i.page))
  let kept = 0
  for (let i = 0; i < pages.length; i++) {
    if (dropped.has(i)) continue
    const f = stageFindings(pages[i], parsed.spec.pages[kept++]!, i, stage)
    outlineFindings.push(...f.outline.map((o) => ({ ...o, ...named(i) })))
    if (f.offPalette.length) offPalette.push({ ...named(i), colors: f.offPalette })
  }
  const badPages = new Set(outlineFindings.filter((o) => o.level === 'error').map((o) => o.file))
  if (badPages.size) {
    throw new CliError(
      EXIT.usage,
      `${dir}: ${badPages.size} page(s) disagree with ${basename(stage.outline!.file)}; fix the page files (or the outline) before building`,
      { outline: stage.outline!.file, findings: outlineFindings },
    )
  }

  const built = await buildDeckPptx(parsed.spec, specDeps(ctx, dir))
  const issues = withImageWarnings(parsed.issues, inputIndexed(built.imageWarnings, parsed.issues))
  return {
    bytes: built.bytes,
    pages: parsed.spec.pages.length,
    issues: issues.map((i) => ({ ...i, ...named(i.page) })),
    imageFailures: inputIndexed(built.imageFailures, parsed.issues).map((f) => ({
      ...f,
      ...named(f.page),
    })),
    files,
    stage,
    outlineFindings,
    offPalette,
  }
}

/** The position of a page file among its folder's page files: the slide it becomes. */
export function pageIndexOf(file: string): number {
  return listPageFiles(resolve(file, '..')).indexOf(resolve(file))
}

export function readOutline(path: string): DeckOutline {
  const r = parseOutline(readFileSync(path, 'utf-8'))
  if (!r.ok) throw new CliError(EXIT.usage, `${path}: ${r.error}`)
  return r.outline
}

export interface PageCheck {
  warnings: string[]
  imageFailures: string[]
  audit: string[]
  stage?: StageFindings
}

function parseOnePage(
  text: string,
  source: string,
  what: string,
): { raw: unknown; spec: PageSpec; warnings: string[] } {
  const parsed = parseDeckSpec(text, { localImages: true })
  if (!parsed.ok) throw new CliError(EXIT.usage, `${source}: ${parsed.error}`)
  if (parsed.spec.pages.length !== 1) {
    throw new CliError(
      EXIT.usage,
      `${source}: ${what} one page object, not ${parsed.spec.pages.length}`,
    )
  }
  return {
    raw: JSON.parse(text) as unknown,
    spec: parsed.spec.pages[0]!,
    warnings: parsed.issues.flatMap((i) => i.warnings ?? []),
  }
}

/**
 * One page spec built on its own and audited, so a page is fixed before the
 * next one is written; with a stage context, also checked against its outline
 * entry (`page`) and the style sheet.
 */
export async function checkPageSpec(
  text: string,
  source: string,
  ctx: PathContext,
  stage?: { context: StageContext; page: number },
): Promise<PageCheck> {
  const one = parseOnePage(text, source, 'a page file holds exactly')
  const built = await buildDeckPptx({ pages: [one.spec] }, specDeps(ctx, dirnameOf(source)))
  const audit = await auditDeckBytes(built.bytes)
  return {
    warnings: [...one.warnings, ...built.imageWarnings.map((w) => w.message)],
    imageFailures: built.imageFailures.map((f) => f.url),
    audit: audit[0]?.issues ?? [],
    ...(stage ? { stage: stageFindings(one.raw, one.spec, stage.page, stage.context) } : {}),
  }
}

function dirnameOf(source: string): string | undefined {
  return source === 'stdin' ? undefined : resolve(source, '..')
}

/** Rebuild one page from its spec and put it where slide `index` was; the other slides are untouched. */
export async function replaceSlideFromSpec(
  opened: OpenedPptx,
  index: number,
  text: string,
  source: string,
  ctx: PathContext,
  stage?: StageContext,
): Promise<{ warnings: string[]; imageFailures: string[]; stage?: StageFindings }> {
  const one = parseOnePage(text, source, 'replace takes')
  const findings = stage ? stageFindings(one.raw, one.spec, index, stage) : undefined
  const errors = findings?.outline.filter((o) => o.level === 'error') ?? []
  if (errors.length) {
    throw new CliError(
      EXIT.usage,
      `${source}: the page disagrees with ${basename(stage!.outline!.file)} pages[${index}]`,
      { outline: stage!.outline!.file, findings: findings!.outline },
    )
  }
  const built = await buildDeckPptx({ pages: [one.spec] }, specDeps(ctx, dirnameOf(source)))
  // the rebuilt page keeps the replaced slide's layout, not the last slide's (often a back cover)
  const merged = await mergeSlideFromPptx(opened, built.bytes, {
    layoutFrom: opened.deck.slides[index],
  })
  if (!merged) throw new CliError(EXIT.conversion, `${source}: the built page could not be merged`)
  promoteSlideBackground(merged, opened.deck.size)
  const last = opened.deck.slides.length - 1
  if (last !== index) moveSlide(opened, last, index)
  if (!deleteSlide(opened, index + 1)) {
    throw new CliError(EXIT.conversion, `could not remove the old slide ${index}`)
  }
  return {
    warnings: [...one.warnings, ...built.imageWarnings.map((w) => w.message)],
    imageFailures: built.imageFailures.map((f) => f.url),
    ...(findings ? { stage: findings } : {}),
  }
}

export function auditDeckBytes(bytes: Uint8Array, only?: number): Promise<DeckAuditPage[]> {
  return auditDeck(bytes, { metrics, ...(only === undefined ? {} : { only }) })
}

export interface RasterizedPage {
  index: number
  png: Uint8Array
  width: number
  height: number
}

/** Every page of a PDF as PNG at `scale` × 72 dpi (2 = 144 dpi). */
export async function rasterizePdf(
  bytes: Uint8Array,
  scale: number,
  only?: number,
  range: { flag: string; oneBased: boolean } = { flag: 'slide', oneBased: false },
): Promise<RasterizedPage[]> {
  const m = await loadPdfium()
  return extract.withPdfDocument(m, bytes, (doc) => {
    const count = m._FPDF_GetPageCount(doc)
    if (only !== undefined && (only < 0 || only >= count)) {
      const [lo, hi] = range.oneBased ? [1, count] : [0, count - 1]
      throw new CliError(EXIT.usage, `--${range.flag} out of range (${lo}-${hi})`)
    }
    const out: RasterizedPage[] = []
    for (let i = 0; i < count; i++) {
      if (only !== undefined && i !== only) continue
      const r = extract.renderPageByIndexPng(m, doc, i, scale)
      if (!r) throw new CliError(EXIT.conversion, `could not render page ${i + 1}`)
      out.push({ index: i, png: r.data, width: r.pixelWidth, height: r.pixelHeight })
    }
    return out
  })
}
