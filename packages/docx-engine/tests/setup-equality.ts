import { expect } from 'vitest'

/**
 * `Run.effective` / `ParaFormat.effective` are display-only views the parser
 * attaches to every run and paragraph. An expectation that does not mention
 * the key is comparing the saved model, so the key is left out of toEqual on
 * that side only; an expectation that spells it out is compared in full.
 */
const isRecord = (v: unknown): v is Record<string, unknown> =>
  typeof v === 'object' && v !== null && !Array.isArray(v)

expect.addEqualityTesters([
  function effectiveAgnostic(a, b) {
    if (!isRecord(a) || !isRecord(b)) return undefined
    const inA = 'effective' in a
    const inB = 'effective' in b
    if (inA === inB) return undefined
    const strip = ({ effective: _effective, ...rest }: Record<string, unknown>) => rest
    // equal without the key: the key was the only difference. Otherwise defer, so
    // toMatchObject keeps its subset semantics and toEqual reports the real diff
    return this.equals(inA ? strip(a) : a, inB ? strip(b) : b) || undefined
  },
])
