/**
 * KB / dictionary terminology rules for the live translate path.
 *
 * `dictionary.ts` owns the *dictionary builder* (file segments -> a
 * `--dictionary` JSON) and applies the KB while it builds. This module owns the
 * same decisions for the editor / snippet path, where the input is one
 * selection rather than a file:
 *
 *   - which mandatory terms the source text actually touched, so callers can
 *     render the "KB · N" badge (`TranslateResponse.matchedTerms`)
 *   - deterministic enforcement of a mandatory term the model left untranslated
 *
 * It is deliberately standalone — it imports only `knowledge-base.ts` — so
 * `provider.ts` can use it without pulling in `dictionary.ts` (which re-exports
 * the package index) and creating an import cycle.
 */
import type { KnowledgeBase, ResolvedRules } from './knowledge-base'

/** One mandatory `source -> target` mapping, from the KB or a generated dictionary. */
export interface TerminologyPair {
  source: string
  target: string
}

export interface ResolveKbOptions {
  sourceLang: string
  targetLang: string
  /** Glossary bucket — forwarded to the KB term filter. */
  category?: string | undefined
  /** Customer name — forwarded to the KB customer-preference filter. */
  customerName?: string | undefined
}

/** Resolve the KB down to the rule set that applies to a single call. */
export function resolveKbForCall(kb: KnowledgeBase, opts: ResolveKbOptions): ResolvedRules {
  return kb.resolve({
    sourceLang: opts.sourceLang,
    targetLang: opts.targetLang,
    ...(opts.category !== undefined ? { category: opts.category } : {}),
    ...(opts.customerName !== undefined ? { customerName: opts.customerName } : {}),
  })
}

/**
 * Mandatory pairs from resolved KB rules, longest source first.
 *
 * Ordering matters: `applyTerminology` rewrites with `split/join`, so a
 * multi-word term has to run before the single word nested inside it
 * ("fabric weight" before "fabric").
 */
export function terminologyPairs(resolved: ResolvedRules): TerminologyPair[] {
  return resolved.terms
    .filter((term) => Boolean(term.sourceTerm) && Boolean(term.targetTerm))
    .map((term) => ({ source: term.sourceTerm, target: term.targetTerm }))
    .sort((a, b) => b.source.length - a.source.length)
}

/**
 * Which of `pairs` the source text actually contains — the terms the model was
 * told to honour for this call. Deduplicated, order preserved.
 */
export function matchTermsInSource(
  sourceText: string,
  pairs: readonly TerminologyPair[],
): string[] {
  const out: string[] = []
  const seen = new Set<string>()
  for (const pair of pairs) {
    if (seen.has(pair.source)) continue
    if (!sourceText.includes(pair.source)) continue
    seen.add(pair.source)
    out.push(pair.source)
  }
  return out
}

/**
 * Force each mandatory target term wherever the model left the source term
 * verbatim. Only segments the model failed to translate are rewritten, so a
 * faithful translation passes through byte-for-byte.
 *
 * `pairs` must already be sorted longest-source-first; see
 * {@link terminologyPairs}.
 */
export function applyTerminology(text: string, pairs: readonly TerminologyPair[]): string {
  let out = text
  for (const pair of pairs) {
    if (pair.source === pair.target) continue
    if (!out.includes(pair.source)) continue
    out = out.split(pair.source).join(pair.target)
  }
  return out
}
