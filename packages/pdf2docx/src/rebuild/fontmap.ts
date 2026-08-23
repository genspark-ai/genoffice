/**
 * Output-font existence check (P21 A): an rFonts family that is not installed
 * leaves the substitution to Word/LO, which picks by name similarity — often a
 * face with very different advances (a whole page set in an embedded URW clone
 * like 'Nimbus Roman No9 L' reflows and spills onto extra pages). Families we
 * cannot locate map to a metric-compatible installed classic; families with no
 * known stand-in stay untouched (a wrong "compatible" face is worse than the
 * renderer's own fallback). CJK families are deliberately absent from the
 * table: their substitutes are not metric-compatible and the regional-artifact
 * normalization (analyze/chars.ts) already handles the known cases.
 */
import { isFamilyInstalled } from '../../../font-metrics/src/index'
import type { FurnitureHf } from '../analyze/furniture'
import { stripTrailingStyleWords } from '../extract/fontname'
import type { IrPage, TextBlock } from '../ir'

const norm = (s: string): string =>
  s
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s\-_]/g, '')

// candidate lists are ordered: metric-equal clones first, same-metrics
// classics next, wider lookalikes last; the first INSTALLED candidate wins
const TIMES = [
  'Times New Roman',
  'Liberation Serif',
  'Tinos',
  'Nimbus Roman',
  'Nimbus Roman No9 L',
  'Times',
  'FreeSerif',
  'DejaVu Serif',
]
const HELVETICA = [
  'Helvetica',
  'Arial',
  'Liberation Sans',
  'Arimo',
  'Nimbus Sans',
  'Nimbus Sans L',
  'Helvetica Neue',
  'FreeSans',
  'DejaVu Sans',
]
const ARIAL = ['Arial', 'Liberation Sans', 'Arimo', 'Helvetica', 'Nimbus Sans', 'FreeSans']
const HELVETICA_NEUE = ['Helvetica Neue', 'Helvetica', 'Arial', 'Liberation Sans', 'Arimo']
const COURIER = [
  'Courier New',
  'Liberation Mono',
  'Cousine',
  'Nimbus Mono PS',
  'Nimbus Mono L',
  'Courier',
  'DejaVu Sans Mono',
]

/** normalized source family → metric-compatible stand-in candidates */
const CLASSIC_MAP: Record<string, string[]> = {
  // Times metrics (URW Nimbus Roman clones, Liberation/Google twins, PS aliases)
  timesroman: TIMES,
  times: TIMES,
  timesnewroman: TIMES,
  timesnewromanps: TIMES,
  timesnewromanpsmt: TIMES,
  nimbusromno9l: TIMES,
  nimbusromanno9l: TIMES,
  nimbusroman: TIMES,
  liberationserif: TIMES,
  tinos: TIMES,
  thorndale: TIMES,
  cgtimes: TIMES, // HP/Monotype CG Times, metric twin of Times New Roman

  // Helvetica metrics
  helvetica: HELVETICA,
  helveticaworld: HELVETICA,
  helveticalt: HELVETICA,
  helveticaltstd: HELVETICA,
  nimbussansl: HELVETICA,
  nimbussans: HELVETICA,
  // Arial metrics (equal to Helvetica's, but prefer the exact twin)
  arial: ARIAL,
  arialmt: ARIAL,
  liberationsans: ARIAL,
  arimo: ARIAL,
  albany: ARIAL,
  helveticaneue: HELVETICA_NEUE,
  helveticaneuelt: HELVETICA_NEUE,
  helveticaneueltstd: HELVETICA_NEUE,
  // Courier metrics
  courier: COURIER,
  couriernew: COURIER,
  courierstd: COURIER,
  nimbusmonl: COURIER,
  nimbusmonol: COURIER,
  nimbusmono: COURIER,
  nimbusmonops: COURIER,
  liberationmono: COURIER,
  cousine: COURIER,
  cumberland: COURIER,
}

export type InstalledCheck = (family: string) => boolean

/** tracking below this (ems) is metric reconciliation, dropped on substitution */
const TINY_TRACKING_EMS = 0.01

/**
 * The family name to write into rFonts: the source family when installed,
 * else the first installed metric-compatible stand-in, else the source
 * family unchanged. Styled aliases resolve through their base name
 * ('Helvetica Neue LTStd It' → 'Helvetica Neue LTStd' → 'Helvetica Neue').
 */
export function resolveOutputFamily(family: string, installed: InstalledCheck): string {
  if (!family || installed(family)) return family
  const stripped = stripTrailingStyleWords(family)
  if (stripped !== family && installed(stripped)) return stripped
  const candidates = CLASSIC_MAP[norm(family)] ?? CLASSIC_MAP[norm(stripped)]
  if (candidates) {
    for (const cand of candidates) if (installed(cand)) return cand
  }
  return family
}

const textBlocksOf = (page: IrPage): TextBlock[] => [
  ...page.blocks.flatMap((b) => {
    if (b.kind === 'text') return [b]
    if (b.kind === 'table') return b.rows.flat().flatMap((cell) => cell.blocks)
    return []
  }),
  ...(page.footnotes ?? []).flatMap((f) => f.blocks),
]

/**
 * Rewrite every span's/HF band's fontFamily to its resolvable output family,
 * in place. Resolutions are memoized per family — the installed check walks
 * the system font index.
 */
export function applyOutputFontSubstitutions(
  pages: IrPage[],
  furnitureHf: readonly FurnitureHf[] = [],
  installed: InstalledCheck = isFamilyInstalled,
): void {
  const memo = new Map<string, string>()
  const resolve = (family: string): string => {
    let out = memo.get(family)
    if (out === undefined) {
      out = resolveOutputFamily(family, installed)
      memo.set(family, out)
    }
    return out
  }
  for (const page of pages) {
    for (const block of textBlocksOf(page)) {
      for (const line of block.lines) {
        for (const span of line.spans) {
          const out = resolve(span.fontFamily)
          if (out !== span.fontFamily) {
            span.fontFamily = out
            // hair-thin POSITIVE tracking is the analyser reconciling the
            // ORIGINAL font's /Widths with its metrics — meaningless on a
            // substituted face, and one twip per char is enough to re-wrap a
            // full body line (P22 A). Deliberate display tracking is far
            // larger. Tiny NEGATIVE tracking stays: dropping it widens every
            // line and grows pages (testPDF_protected 41→42).
            if (
              span.charSpacingPt !== undefined &&
              span.charSpacingPt > 0 &&
              span.charSpacingPt <= TINY_TRACKING_EMS * span.fontSize
            ) {
              delete span.charSpacingPt
            }
          }
        }
      }
    }
  }
  for (const hf of furnitureHf) hf.fontFamily = resolve(hf.fontFamily)
}
