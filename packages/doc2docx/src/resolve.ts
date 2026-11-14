/**
 * Stage 3b: resolve the Word stylesheet into the formatting model.
 *
 * Resolution order per paragraph:
 *   paragraph style chain (base -> derived PAPX sprms), then the direct PAPX
 *   sprms; character formatting per run: the paragraph style chain CHPX, the
 *   default character style (istd 10) chain, the run's character style
 *   (sprmCIstd) chain, then the direct CHPX sprms.
 * Toggle operands follow Word's layering: 0/1 set false/true, 0x80 keeps the
 * inherited value and 0x81 inverts it (the base of the chain is off). Broken
 * chains (cycles, missing styles, out-of-range indices) are reported once in
 * the warnings list and the remaining layers still apply.
 */
import type {
  CharacterFormatting,
  CommentMark,
  FormattedParagraph,
  FormattedLink,
  FormattedRun,
  FormattingRead,
  FormattingSource,
  NoteMark,
  ParagraphContext,
  ParagraphFormatting,
} from './formatting'
import {
  applyCharacterGrpprl,
  applyParagraphGrpprl,
  characterFormattingOf,
  readFormattingDetailed,
} from './formatting'
import { fontNamesFor, parseFontTable } from './fonts'
import { iterateSprms, readSprmU16 } from './sprm'
import { legacyStyleId, parseStsh, type Style, type StyleSheet } from './styles'
import { createWordControlMapper } from './text'

/** the style index Word applies to paragraphs with no PAPX / no explicit style */
export const DEFAULT_PARAGRAPH_STYLE_ISTD = 0
/** the Word 97 default character style ("Default Paragraph Font") */
export const DEFAULT_CHARACTER_STYLE_ISTD = 10
const SPRM_C_ISTD = 0x4a30
const STI_HYPERLINK = 85
const STI_FOLLOWED_HYPERLINK = 87
const NO_SPRMS = new Uint8Array(0)

interface Resolver {
  styles: StyleSheet
  context: ParagraphContext
  warn: (message: string) => void
}

/** warnings are deduplicated: the same broken style is reported once, not per run */
function warnOnce(warnings: string[]): (message: string) => void {
  const seen = new Set<string>()
  return (message) => {
    if (seen.has(message)) return
    seen.add(message)
    warnings.push(message)
  }
}

/** the chain from the root base style to the given style (base first) */
function styleChain(
  styles: StyleSheet,
  istd: number,
  warn: (message: string) => void,
  what: string,
): Style[] {
  const chain: Style[] = []
  const seen = new Set<number>()
  let current: number | undefined = istd
  while (current !== undefined) {
    if (seen.has(current)) {
      warn(
        `${what} style chain contains a cycle at style index ${current}; the remaining base styles were ignored`,
      )
      break
    }
    seen.add(current)
    const style = styles.styles.get(current)
    if (!style) {
      warn(
        `${what} refers to style index ${current}, which is not in the stylesheet; it was ignored`,
      )
      break
    }
    chain.unshift(style)
    current = style.basedOn
  }
  return chain
}

/** the last sprmCIstd in a raw CHPX grpprl (a run's character style) */
function characterStyleIndex(sprms: Uint8Array): number | undefined {
  let istd: number | undefined
  for (const sprm of iterateSprms(sprms))
    if (sprm.opcode === SPRM_C_ISTD) {
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) istd = value
    }
  return istd
}

/** a stylesheet slot carries a definition worth emitting/referencing: not an
    empty cbStd == 0 placeholder */
function styleIsDefined(style: Style | undefined): boolean {
  return (
    style !== undefined &&
    (style.name !== '' || style.papxSprms.length > 0 || style.chpxSprms.length > 0)
  )
}

function resolveParagraphFormatting(
  istd: number,
  sprms: Uint8Array,
  paragraphStyles: Style[],
  resolver: Resolver,
): ParagraphFormatting {
  const formatting: ParagraphFormatting = {}
  for (const style of paragraphStyles)
    applyParagraphGrpprl(formatting, style.papxSprms, resolver.context, 'inherit')
  applyParagraphGrpprl(formatting, sprms, resolver.context, 'inherit')
  formatting.styleIndex = istd
  // the writer references the style through this id; istd 0 is Normal, which
  // the blank template already defines, so it is never emitted nor referenced
  if (istd !== DEFAULT_PARAGRAPH_STYLE_ISTD && styleIsDefined(resolver.styles.styles.get(istd)))
    formatting.styleId = legacyStyleId(istd)
  return formatting
}

/** Word draws TOC entries in the paragraph's ink: the Hyperlink style is skipped inside a TOC link */
function suppressedTocStyle(styles: StyleSheet, istd: number, inTocLink: boolean): boolean {
  if (!inTocLink) return false
  const sti = styles.styles.get(istd)?.sti
  return sti === STI_HYPERLINK || sti === STI_FOLLOWED_HYPERLINK
}

/** Word's built-in heading style ids: sti 1..9 are heading 1..9 */
const STI_HEADING_FIRST = 1
const STI_HEADING_LAST = 9

function resolveCharacterFormatting(
  sprms: Uint8Array,
  paragraphStyles: Style[],
  resolver: Resolver,
  inTocLink = false,
): CharacterFormatting {
  const formatting: CharacterFormatting = {}
  for (const style of paragraphStyles) applyCharacterGrpprl(formatting, style.chpxSprms, 'inherit')
  const defaultCharacterStyle = resolver.styles.styles.get(DEFAULT_CHARACTER_STYLE_ISTD)
  if (defaultCharacterStyle) {
    for (const style of styleChain(
      resolver.styles,
      DEFAULT_CHARACTER_STYLE_ISTD,
      resolver.warn,
      'the default character style',
    ))
      applyCharacterGrpprl(formatting, style.chpxSprms, 'inherit')
  }
  const characterIstd = characterStyleIndex(sprms)
  if (
    characterIstd !== undefined &&
    !suppressedTocStyle(resolver.styles, characterIstd, inTocLink)
  ) {
    for (const style of styleChain(
      resolver.styles,
      characterIstd,
      resolver.warn,
      'a run character style',
    ))
      applyCharacterGrpprl(formatting, style.chpxSprms, 'inherit')
    // the run carries the character style through this id (w:rStyle); istd 10
    // ("Default Paragraph Font") is the null style and is never referenced,
    // and a Hyperlink-family style is implied by the run's own link (the
    // engine convention: rStyle Hyperlink is written by the generator)
    const referencedStyle = resolver.styles.styles.get(characterIstd)
    if (
      characterIstd !== DEFAULT_CHARACTER_STYLE_ISTD &&
      styleIsDefined(referencedStyle) &&
      referencedStyle!.sti !== STI_HYPERLINK &&
      referencedStyle!.sti !== STI_FOLLOWED_HYPERLINK
    )
      formatting.styleId = legacyStyleId(characterIstd)
  }
  applyCharacterGrpprl(formatting, sprms, 'inherit')
  if (resolver.styles.fonts) {
    const names = fontNamesFor(resolver.styles.fonts, {
      ascii: formatting.ftcAscii,
      hAnsi: formatting.ftcHAnsi,
      eastAsia: formatting.ftcEastAsia,
      cs: formatting.ftcBi,
    })
    if (names) formatting.fontNames = names
  }
  return formatting
}

/** A stylesheet entry for the .docx writer's styles.xml: the style's own chain
    (base first, the same layering the direct-formatting resolution uses)
    resolved into formatting structs. The writer serializes these into
    w:style definitions whose ids come from styles.ts legacyStyleId. */
export interface LegacyStyleDef {
  istd: number
  kind: 'paragraph' | 'character'
  name: string
  basedOn?: number
  /** the chain-resolved paragraph properties (paragraph styles) */
  paragraph?: ParagraphFormatting
  /** the chain-resolved run properties (a paragraph style's CHPX layer, or a
      character style's whole chain) */
  character?: CharacterFormatting
}

/** Resolve every defined style of the stylesheet into a LegacyStyleDef.
    Normal (paragraph istd 0) is skipped — the blank .docx template already
    defines it and its chain is baked into docDefaults. */
export function legacyStyleDefs(
  styles: StyleSheet,
  context: ParagraphContext,
  warnings: string[] = [],
): LegacyStyleDef[] {
  const warn = warnOnce(warnings)
  const styleFonts = (formatting: CharacterFormatting): void => {
    if (!styles.fonts) return
    const names = fontNamesFor(styles.fonts, {
      ascii: formatting.ftcAscii,
      hAnsi: formatting.ftcHAnsi,
      eastAsia: formatting.ftcEastAsia,
      cs: formatting.ftcBi,
    })
    if (names) formatting.fontNames = names
  }
  const defs: LegacyStyleDef[] = []
  for (const style of styles.styles.values()) {
    if (!styleIsDefined(style)) continue
    // Normal is the blank template's own style, and the Hyperlink family is
    // provided by the template too (the generator implies it on link runs)
    if (style.kind === 'paragraph' && style.istd === DEFAULT_PARAGRAPH_STYLE_ISTD) continue
    if (style.sti === STI_HYPERLINK || style.sti === STI_FOLLOWED_HYPERLINK) continue
    const chain = styleChain(styles, style.istd, warn, `style ${style.istd}`)
    if (chain.length === 0) continue
    const def: LegacyStyleDef = {
      istd: style.istd,
      kind: style.kind === 'paragraph' ? 'paragraph' : 'character',
      name: style.name,
      ...(style.basedOn !== undefined ? { basedOn: style.basedOn } : {}),
    }
    if (style.kind === 'paragraph') {
      const paragraph: ParagraphFormatting = {}
      for (const entry of chain)
        applyParagraphGrpprl(paragraph, entry.papxSprms, context, 'inherit')
      delete paragraph.styleIndex
      delete paragraph.styleId
      // Word's built-in heading styles carry the outline level the sti maps
      // to; a legacy PAPX sprmSOutLvl (already resolved) wins when present
      if (
        paragraph.outlineLevel === undefined &&
        style.sti >= STI_HEADING_FIRST &&
        style.sti <= STI_HEADING_LAST
      )
        paragraph.outlineLevel = style.sti - 1
      if (Object.keys(paragraph).length > 0) def.paragraph = paragraph
    }
    const character: CharacterFormatting = {}
    for (const entry of chain) applyCharacterGrpprl(character, entry.chpxSprms, 'inherit')
    styleFonts(character)
    delete character.styleIndex
    delete character.styleId
    if (Object.keys(character).length > 0) def.character = character
    defs.push(def)
  }
  return defs
}

/** The stylesheet of a document with its font table attached (undefined without an STSH). */
export function readStyleSheet(
  source: FormattingSource,
  warnings: string[] = [],
): StyleSheet | undefined {
  if (source.fib.lcbStshf === 0) return undefined
  const styles = parseStsh(source.table, source.fib.fcStshf, source.fib.lcbStshf)
  const names = parseFontTable(source.fib, source.table, warnings)
  if (names.length > 0) styles.fonts = { names, defaults: styles.defaultFonts }
  return styles
}

function buildResolvedRuns(
  paragraph: FormattedParagraph,
  read: FormattingRead,
  paragraphStyles: Style[],
  resolver: Resolver,
  rawText: string,
): FormattedRun[] {
  const boundaries = new Set<number>([paragraph.startCp, paragraph.endCp])
  for (const run of read.characterRuns) {
    if (run.cpEnd <= paragraph.startCp || run.cpStart >= paragraph.endCp) continue
    boundaries.add(Math.max(run.cpStart, paragraph.startCp))
    boundaries.add(Math.min(run.cpEnd, paragraph.endCp))
  }
  const points = [...boundaries].sort((left, right) => left - right)
  const mapper = createWordControlMapper()
  const runs: FormattedRun[] = []
  for (let index = 0; index + 1 < points.length; index++) {
    const segmentStart = points[index]!
    const segmentEnd = points[index + 1]!
    if (segmentEnd <= segmentStart) continue
    const covering = read.characterRuns.find(
      (run) => run.cpStart <= segmentStart && run.cpEnd >= segmentEnd,
    )
    const formatting = resolveCharacterFormatting(
      covering?.sprms ?? NO_SPRMS,
      paragraphStyles,
      resolver,
    )
    const text = mapper.map(rawText.slice(segmentStart, segmentEnd))
    if (text === '') continue
    const previous = runs[runs.length - 1]
    if (previous && JSON.stringify(previous.formatting) === JSON.stringify(formatting))
      previous.text += text
    else runs.push({ text, formatting })
  }
  return runs
}

/** A hyperlink result range the resolved runs should be tagged with. */
export interface FormattedLinkRange {
  startCp: number
  endCp: number
  href: string
  tooltip?: string
  inToc?: boolean
}

/** The result range of a live field (PAGE, NUMPAGES, ...) with its instruction. */
export interface FormattedFieldRange {
  startCp: number
  endCp: number
  instruction: string
}

/** A comment anchor at a main-document CP (resolve.ts emits it as an empty marker run). */
export interface CommentMarkAt extends CommentMark {
  cp: number
}

/** A CP range whose runs should be resolved (table cell/body unit of stage 3c-2). */
export interface ParagraphRunRequest {
  startCp: number
  endCp: number
  /** paragraph style index from the unit's PAPX (0 when it has no PAPX) */
  styleIstd: number
  /** raw PAPX sprms of the unit (empty when it has no PAPX) */
  papxSprms: Uint8Array
}

function runsForRange(
  read: FormattingRead,
  startCp: number,
  endCp: number,
  formattingOf: (
    covering: FormattingRead['characterRuns'][number] | undefined,
    inTocLink?: boolean,
  ) => CharacterFormatting,
  rawText: string,
  linkRanges: FormattedLinkRange[] = [],
  fieldRanges: FormattedFieldRange[] = [],
  noteMarks: NoteMark[] = [],
  commentMarks: CommentMarkAt[] = [],
  pictureCps: number[] = [],
): FormattedRun[] {
  const boundaries = new Set<number>([startCp, endCp])
  for (const run of read.characterRuns) {
    if (run.cpEnd <= startCp || run.cpStart >= endCp) continue
    boundaries.add(Math.max(run.cpStart, startCp))
    boundaries.add(Math.min(run.cpEnd, endCp))
  }
  for (const range of [...linkRanges, ...fieldRanges]) {
    if (range.endCp <= startCp || range.startCp >= endCp) continue
    boundaries.add(Math.max(range.startCp, startCp))
    boundaries.add(Math.min(range.endCp, endCp))
  }
  // a note mark is a one-character run of its own, whatever CHPX run it sits in
  for (const mark of noteMarks) {
    if (mark.cp < startCp || mark.cp >= endCp) continue
    boundaries.add(mark.cp)
    boundaries.add(mark.cp + 1)
  }
  for (const mark of commentMarks)
    if (mark.cp >= startCp && mark.cp <= endCp) boundaries.add(mark.cp)
  for (const cp of pictureCps) {
    if (cp < startCp || cp >= endCp) continue
    boundaries.add(cp)
    boundaries.add(cp + 1)
  }
  const points = [...boundaries].sort((left, right) => left - right)
  const mapper = createWordControlMapper()
  const runs: FormattedRun[] = []
  // a live field with an empty cached result has no segment of its own: give it an empty run
  const emptyFieldAt = (cp: number): void => {
    for (const range of fieldRanges) {
      if (range.startCp !== cp || range.endCp !== cp) continue
      const covering = read.characterRuns.find((run) => run.cpStart <= cp && run.cpEnd > cp)
      runs.push({ text: '', formatting: formattingOf(covering), field: range.instruction })
    }
  }
  // comment anchors at a CP: ranges close, references follow, new ranges open
  const commentMarksAt = (cp: number): void => {
    const marks = commentMarks.filter((mark) => mark.cp === cp)
    if (marks.length === 0) return
    const order = { end: 0, reference: 1, start: 2 }
    const covering = read.characterRuns.find((run) => run.cpStart <= cp && run.cpEnd > cp)
    runs.push({
      text: '',
      formatting: formattingOf(covering),
      commentMarks: marks
        .sort((left, right) => order[left.kind] - order[right.kind] || left.id - right.id)
        .map(({ kind, id }) => ({ kind, id })),
    })
  }
  for (let index = 0; index + 1 < points.length; index++) {
    const segmentStart = points[index]!
    const segmentEnd = points[index + 1]!
    emptyFieldAt(segmentStart)
    commentMarksAt(segmentStart)
    if (segmentEnd <= segmentStart) continue
    const covering = read.characterRuns.find(
      (run) => run.cpStart <= segmentStart && run.cpEnd >= segmentEnd,
    )
    const link = linkRanges.find(
      (range) => range.startCp <= segmentStart && range.endCp >= segmentEnd,
    )
    const formatting = formattingOf(covering, link?.inToc)
    const formattedLink: FormattedLink | undefined = link
      ? {
          href: link.href,
          ...(link.tooltip ? { tooltip: link.tooltip } : {}),
          ...(link.inToc ? { plain: true } : {}),
        }
      : undefined
    const field = fieldRanges.find(
      (range) => range.startCp <= segmentStart && range.endCp >= segmentEnd,
    )
    const text = mapper.map(rawText.slice(segmentStart, segmentEnd))
    const note = noteMarks.find((mark) => mark.cp === segmentStart)
    if (note) {
      runs.push({ text: note.custom ? text : '', formatting, note })
      continue
    }
    if (pictureCps.includes(segmentStart)) {
      runs.push({ text: '', formatting, pictureCp: segmentStart })
      continue
    }
    if (text === '') continue
    const previous = runs[runs.length - 1]
    if (
      previous &&
      !previous.note &&
      previous.commentMarks === undefined &&
      previous.pictureCp === undefined &&
      JSON.stringify(previous.formatting) === JSON.stringify(formatting) &&
      JSON.stringify(previous.link ?? null) === JSON.stringify(formattedLink ?? null) &&
      previous.field === field?.instruction
    )
      previous.text += text
    else
      runs.push({
        text,
        formatting,
        ...(formattedLink ? { link: formattedLink } : {}),
        ...(field ? { field: field.instruction } : {}),
      })
  }
  emptyFieldAt(endCp)
  commentMarksAt(endCp)
  return runs
}

/** The resolved formatting of the paragraph mark at request.endCp (undefined when no CHPX covers it). */
export function resolveParagraphMarkFormatting(
  read: FormattingRead,
  source: FormattingSource,
  styles: StyleSheet | undefined,
  request: ParagraphRunRequest,
  warnings: string[] = [],
): CharacterFormatting | undefined {
  const covering = read.characterRuns.find(
    (run) => run.cpStart <= request.endCp && run.cpEnd > request.endCp,
  )
  if (!covering) return undefined
  if (!styles) return characterFormattingOf(covering.sprms)
  const resolver: Resolver = { styles, context: { source, warnings }, warn: warnOnce(warnings) }
  const paragraphStyles = styleChain(
    styles,
    request.styleIstd,
    resolver.warn,
    `the paragraph style (index ${request.styleIstd})`,
  )
  return resolveCharacterFormatting(covering.sprms, paragraphStyles, resolver)
}

/**
 * Resolve the paragraph formatting of one arbitrary CP range (a stage-3c-2
 * body unit) with the same style layering as resolveFormatting.
 */
export function resolveParagraphFormattingFor(
  source: FormattingSource,
  styles: StyleSheet | undefined,
  request: ParagraphRunRequest,
  warnings: string[] = [],
): ParagraphFormatting {
  if (!styles) {
    const formatting: ParagraphFormatting = {}
    applyParagraphGrpprl(formatting, request.papxSprms, { source, warnings }, 'inherit')
    formatting.styleIndex = request.styleIstd
    return formatting
  }
  const resolver: Resolver = { styles, context: { source, warnings }, warn: warnOnce(warnings) }
  const paragraphStyles = styleChain(
    styles,
    request.styleIstd,
    resolver.warn,
    `the paragraph style (index ${request.styleIstd})`,
  )
  const formatting: ParagraphFormatting = {}
  for (const style of paragraphStyles)
    applyParagraphGrpprl(formatting, style.papxSprms, resolver.context, 'inherit')
  applyParagraphGrpprl(formatting, request.papxSprms, resolver.context, 'inherit')
  formatting.styleIndex = request.styleIstd
  return formatting
}

/**
 * Resolve the character runs of one arbitrary CP range (a table cell paragraph
 * or a body paragraph split at cell marks) with the same style layering as
 * resolveFormatting. Without an STSH only the direct CHPX formatting applies.
 */
export function resolveRunsForRange(
  read: FormattingRead,
  source: FormattingSource,
  styles: StyleSheet | undefined,
  request: ParagraphRunRequest,
  warnings: string[] = [],
  linkRanges: FormattedLinkRange[] = [],
  fieldRanges: FormattedFieldRange[] = [],
  noteMarks: NoteMark[] = [],
  commentMarks: CommentMarkAt[] = [],
  pictureCps: number[] = [],
): FormattedRun[] {
  if (!styles) {
    return runsForRange(
      read,
      request.startCp,
      request.endCp,
      (covering) => characterFormattingOf(covering?.sprms ?? NO_SPRMS),
      source.rawText,
      linkRanges,
      fieldRanges,
      noteMarks,
      commentMarks,
      pictureCps,
    )
  }
  const resolver: Resolver = { styles, context: { source, warnings }, warn: warnOnce(warnings) }
  const paragraphStyles = styleChain(
    styles,
    request.styleIstd,
    resolver.warn,
    `the paragraph style (index ${request.styleIstd})`,
  )
  return runsForRange(
    read,
    request.startCp,
    request.endCp,
    (covering, inTocLink) =>
      resolveCharacterFormatting(
        covering?.sprms ?? NO_SPRMS,
        paragraphStyles,
        resolver,
        inTocLink === true,
      ),
    source.rawText,
    linkRanges,
    fieldRanges,
    noteMarks,
    commentMarks,
    pictureCps,
  )
}

/** the note marks that fall inside a paragraph range */
export function noteMarksWithin(marks: NoteMark[], startCp: number, endCp: number): NoteMark[] {
  return marks.filter((mark) => mark.cp >= startCp && mark.cp < endCp)
}

/** the comment anchors that fall inside a paragraph range (both ends inclusive) */
export function commentMarksWithin(
  marks: CommentMarkAt[],
  startCp: number,
  endCp: number,
): CommentMarkAt[] {
  return marks.filter((mark) => mark.cp >= startCp && mark.cp <= endCp)
}

/** clip live-field result ranges to a paragraph range */
export function fieldRangesWithin(
  fields: Array<{ startCp: number; endCp: number; instruction: string }>,
  startCp: number,
  endCp: number,
): FormattedFieldRange[] {
  return fields
    .filter(
      (field) =>
        (field.endCp > startCp && field.startCp < endCp) ||
        (field.startCp === field.endCp && field.startCp >= startCp && field.startCp <= endCp),
    )
    .map((field) => ({
      startCp: Math.max(field.startCp, startCp),
      endCp: Math.min(field.endCp, endCp),
      instruction: field.instruction,
    }))
}

/**
 * Resolve a stage-2 formatting read through a stylesheet: every paragraph and
 * run carries the merged style + direct formatting (styleIndex stays as a
 * diagnostic). Exposes the writer's input model shape.
 */
export function resolveFormatting(
  read: FormattingRead,
  source: FormattingSource,
  styles: StyleSheet,
  warnings: string[] = [],
): FormattedParagraph[] {
  const resolver: Resolver = { styles, context: { source, warnings }, warn: warnOnce(warnings) }
  return read.paragraphs.map((paragraph, index) => {
    const raw = read.paragraphSprms[index] ?? {
      istd: DEFAULT_PARAGRAPH_STYLE_ISTD,
      sprms: NO_SPRMS,
    }
    const paragraphStyles = styleChain(
      styles,
      raw.istd,
      resolver.warn,
      `the paragraph style (index ${raw.istd})`,
    )
    return {
      startCp: paragraph.startCp,
      endCp: paragraph.endCp,
      text: paragraph.text,
      formatting: resolveParagraphFormatting(raw.istd, raw.sprms, paragraphStyles, resolver),
      runs: buildResolvedRuns(paragraph, read, paragraphStyles, resolver, source.rawText),
      ...markFormattingEntry(
        resolveParagraphMarkFormatting(read, source, styles, {
          startCp: paragraph.startCp,
          endCp: paragraph.endCp,
          styleIstd: raw.istd,
          papxSprms: raw.sprms,
        }),
      ),
    }
  })
}

/** spread helper: a markFormatting entry only when the mark has one */
export function markFormattingEntry(mark: CharacterFormatting | undefined): {
  markFormatting?: CharacterFormatting
} {
  return mark ? { markFormatting: mark } : {}
}

/** Read the body model and resolve its stylesheet (documents without an STSH keep direct formatting). */
export function readResolvedFormatting(
  source: FormattingSource,
  warnings: string[] = [],
): FormattedParagraph[] {
  const read = readFormattingDetailed(source, warnings)
  // a document without an STSH has nothing to resolve: keep the direct model
  const styles = readStyleSheet(source, warnings)
  if (!styles) return read.paragraphs
  return resolveFormatting(read, source, styles, warnings)
}
