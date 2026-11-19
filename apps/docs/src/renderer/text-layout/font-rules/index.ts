/**
 * Every face / line-factor / CSS-chain decision of the Docs renderer, behind
 * one entry point. `resolveFace` returns pure data; the emitters (run-faces,
 * para-line-box, doc-style-css, strut-face) map it and decide nothing. New
 * rules enter here (or the resolver / the measurement manifests they read),
 * never line-metrics.ts: tools/check-font-rules.mjs enforces it.
 */
import { faceKey, type SubstituteRule } from '@genoffice/font-metrics/resolve'
import { resolveDeclaredFace } from '../../font-resolve'
import { declaredFaceMissing, isBundledFont, isFontAvailable } from './availability'
import { cssFontFamily } from './css-chains'
import { isEaLayoutFace, krCharsetRoute, missingLatinSubstitute } from './font-table'
import { latinSlotFace } from './latin-slot'
import { type FactorSource, lineFactorDecision, tableLineHeightFactor } from './line-factor'

export * from './availability'
export * from './css-chains'
export * from './font-table'
export * from './latin-slot'
export * from './line-factor'
export * from './slot-faces'

export type FaceSource =
  | 'embedded'
  | 'altName'
  | 'charset'
  | 'prefix'
  | 'script'
  | 'installed'
  | 'shipped'
  | 'fallback'
  | 'unknown'

export interface FaceDecision {
  /** NFKC-lowercased chain head */
  declared: string
  /** the family Word lays out with (resolver), or the declared name when the store has no answer */
  face: string
  source: FaceSource
  substituted: boolean
  /** resolver path: declared key, substitutes, final face */
  path: string[]
  lineFactor: number
  factorSource: FactorSource
  /** cssFontFamily of the declared name */
  chain: string
  /** first bundled / metric alias family of the chain, else null */
  standIn: string | null
  latinSlot: string | null
  kr: 'malgun' | 'batang' | null
  missingLatin: 'calibri' | 'cambria' | null
  eastAsianLayout: boolean
  missing: boolean
}

/** first metric alias / bundled family of the chain ahead of its generic tail, else null */
function standInOf(chain: string): string | null {
  for (const raw of chain.split(',')) {
    const fam = raw.trim().replace(/^['"]|['"]$/g, '')
    if (!fam || /^(serif|sans-serif|monospace)$/.test(fam)) return null
    if (/ GO$/.test(fam) || /^GenOffice /.test(fam) || isBundledFont(fam)) return fam
  }
  return null
}

const RULE_SOURCE: Record<Exclude<SubstituteRule, 'styleSuffix'>, FaceSource> = {
  charset81: 'charset',
  charset80: 'charset',
  altName: 'altName',
  arabicCodePage: 'prefix',
  table: 'prefix',
  hanKr: 'script',
  han: 'script',
  hangul: 'script',
  fallback: 'fallback',
}

function sourceOf(face: ReturnType<typeof resolveDeclaredFace>): FaceSource {
  if (!face) return 'unknown'
  if (face.source === 'embedded') return 'embedded'
  const rule = face.rules[0]
  // a style word dropped from the name is still the base family laid out itself
  if (!rule || rule === 'styleSuffix')
    return face.source === 'installed' || face.source === 'system' ? 'installed' : 'shipped'
  return RULE_SOURCE[rule]
}

/** The whole decision for one declared font name, as data. */
export function resolveFace(name: string): FaceDecision {
  const key = faceKey(name)
  const face = resolveDeclaredFace(name)
  const { factor: lineFactor, source: factorSource } = lineFactorDecision(name)
  const chain = cssFontFamily(name)
  const head = name.split(',')[0].replace(/['"]/g, '').trim()
  return {
    declared: key,
    face: face?.family ?? head,
    source: sourceOf(face),
    substituted: !!face?.substituted,
    path: face?.path ?? [key],
    lineFactor,
    factorSource,
    chain,
    standIn: standInOf(chain),
    latinSlot: latinSlotFace(name),
    kr: krCharsetRoute(name),
    missingLatin: missingLatinSubstitute(name),
    eastAsianLayout: isEaLayoutFace(name),
    missing: declaredFaceMissing(head) || !isFontAvailable(head),
  }
}

export { tableLineHeightFactor }
