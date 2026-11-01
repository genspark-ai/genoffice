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
  FormattedParagraph,
  FormattedRun,
  FormattingRead,
  FormattingSource,
  ParagraphContext,
  ParagraphFormatting,
} from './formatting'
import { applyCharacterGrpprl, applyParagraphGrpprl, readFormattingDetailed } from './formatting'
import { iterateSprms, readSprmU16 } from './sprm'
import { parseStsh, type Style, type StyleSheet } from './styles'
import { createWordControlMapper } from './text'

/** the style index Word applies to paragraphs with no PAPX / no explicit style */
export const DEFAULT_PARAGRAPH_STYLE_ISTD = 0
/** the Word 97 default character style ("Default Paragraph Font") */
export const DEFAULT_CHARACTER_STYLE_ISTD = 10
const SPRM_C_ISTD = 0x4a30
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
  return formatting
}

function resolveCharacterFormatting(
  sprms: Uint8Array,
  paragraphStyles: Style[],
  resolver: Resolver,
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
  if (characterIstd !== undefined) {
    for (const style of styleChain(
      resolver.styles,
      characterIstd,
      resolver.warn,
      'a run character style',
    ))
      applyCharacterGrpprl(formatting, style.chpxSprms, 'inherit')
  }
  applyCharacterGrpprl(formatting, sprms, 'inherit')
  return formatting
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
    }
  })
}

/** Read the body model and resolve its stylesheet (documents without an STSH keep direct formatting). */
export function readResolvedFormatting(
  source: FormattingSource,
  warnings: string[] = [],
): FormattedParagraph[] {
  const read = readFormattingDetailed(source, warnings)
  // a document without an STSH has nothing to resolve: keep the direct model
  if (source.fib.lcbStshf === 0) return read.paragraphs
  const styles = parseStsh(source.table, source.fib.fcStshf, source.fib.lcbStshf)
  return resolveFormatting(read, source, styles, warnings)
}
