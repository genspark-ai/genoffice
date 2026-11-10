import { placeholderSource, type RedactionProjection } from './redact'
import type { ParseMap } from '../document/parse-map'
import type { ElementEntry } from '../document/parse-map'

/**
 * Stop a model edit from damaging something the reader withheld.
 *
 * ## Damage, not disclosure
 *
 * The read path is what keeps the words from the model (see `redact.ts`). What is
 * left is the model **rewriting or removing** them, which is silent: the file
 * still opens, it just no longer says what the reader wrote. That is the failure
 * this guards.
 *
 * ## Why it is cheap here
 *
 * Every op except `str_replace` addresses one element by `sid`, and a withheld
 * region is a range in the same source string — so "does this op touch it" is an
 * interval intersection, with no position arithmetic and no before/after text
 * comparison. `str_replace` is an exact string match over the source, so the
 * same question is "does any occurrence of `old` land inside a withheld
 * region".
 *
 * ## What is deliberately still allowed
 *
 * A `str_replace` whose `old` merely *contains* a marker is how a model rewrites
 * the sentence around a placeholder, and that is the whole point of the feature.
 * Only a needle that lands **on** withheld characters is refused. Moving,
 * restyling and reordering an element that happens to sit next to a withheld
 * region are fine — the withheld characters themselves are untouched.
 */

/** op fields that name an element the op will modify */
const SID_FIELDS = ['sid', 'ref_sid'] as const

export interface RedactionRefusal {
  reason: string
  label: string
  where: string
}

function elementOf(map: ParseMap, sid: unknown): ElementEntry | null {
  return typeof sid === 'number' ? (map.bySid.get(sid) ?? null) : null
}

/** Does this element's range contain or sit inside a withheld region? */
function hit(proj: RedactionProjection, e: ElementEntry | null): RedactionRefusal | null {
  if (!e) return null
  for (const s of proj.spans) {
    // intersection, in either direction: a span inside the element is as
    // disqualifying as the element sitting inside a span
    if (s.rawFrom < e.range[1] && s.rawTo > e.range[0]) {
      return {
        label: s.label,
        where: s.where,
        reason:
          `this edit would reach <${e.tag}>, which holds a span withheld from the model ` +
          `(${placeholderSource(s.label)}). The model was never shown those words, so it ` +
          `cannot have meant to change them — rephrase around the placeholder instead.`,
      }
    }
  }
  return null
}

/**
 * Check a batch. Returns null when it may proceed.
 *
 * The whole batch is checked before any of it runs, so a refusal leaves the
 * document exactly as it was.
 */
export function redactGuardForOps(
  ops: readonly unknown[],
  proj: RedactionProjection,
  map: ParseMap,
): RedactionRefusal | null {
  if (proj.empty) return null
  for (const [i, raw] of ops.entries()) {
    if (typeof raw !== 'object' || raw === null) continue
    const op = raw as Record<string, unknown>
    const name = String(op.op ?? '')

    if (name === 'str_replace') {
      const old = op.old
      if (typeof old !== 'string' || old === '') continue
      // exact-string matching, same as the op itself
      for (const s of proj.spans) {
        for (
          let at = proj.raw.indexOf(old, Math.max(0, s.rawFrom - old.length + 1));
          at !== -1;
          at = proj.raw.indexOf(old, at + 1)
        ) {
          if (at >= s.rawFrom && at + old.length <= s.rawTo) {
            return {
              label: s.label,
              where: s.where,
              reason:
                `this replacement would rewrite text withheld from the model ` +
                `(${placeholderSource(s.label)}). It cannot have meant to change words it ` +
                `never saw — rephrase around the placeholder instead.`,
            }
          }
        }
      }
      continue
    }

    for (const field of SID_FIELDS) {
      const refusal = hit(proj, elementOf(map, op[field]))
      if (refusal) return { ...refusal, reason: `op #${i} (${name}): ${refusal.reason}` }
    }
  }
  return null
}
