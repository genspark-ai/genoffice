import { afterEach, describe, expect, it } from 'vitest'

import { canonicalFunctionName, extractFunctionNames } from '../src/renderer/formula-functions'
import { installSupportedFunctionProbe } from '../src/renderer/function-registry-probe'
import {
  formulaKeepsCache,
  hasUsableCachedValue,
  setSupportedFunctionProbe,
} from '../src/renderer/univer-sync'
import type { UniverRuntime } from '../src/renderer/univer-state'

// A stand-in for the live Univer registry: the functions the engine ships.
const ENGINE_FUNCTIONS = new Set([
  'SUM',
  'IF',
  'IFERROR',
  'ISERROR',
  'IFNA',
  'IFS',
  'CHOOSE',
  'LOG10',
  'CONCAT',
  'FILTER',
  'INDEX',
  'MATCH',
  'TRUE',
])

afterEach(() => {
  setSupportedFunctionProbe(null)
})

describe('extractFunctionNames', () => {
  it('lists every called function, canonical and deduplicated', () => {
    expect(extractFunctionNames('IFERROR(FOOBARFN(A1),0)')).toEqual(['IFERROR', 'FOOBARFN'])
    expect(extractFunctionNames('IF(ISERROR(X(A1)),"",X(A1))')).toEqual(['IF', 'ISERROR', 'X'])
    expect(extractFunctionNames('=sum(a1:a2)+Sum(B1)')).toEqual(['SUM'])
  })

  it('strips Excel storage markers before lookup', () => {
    expect(extractFunctionNames('_xlfn.CONCAT(A1,_xlfn._xlws.FILTER(B:B,C:C))')).toEqual([
      'CONCAT',
      'FILTER',
    ])
    expect(canonicalFunctionName('_xlfn._xlws.sort')).toBe('SORT')
    expect(canonicalFunctionName('__xludf.DUMMYFUNCTION')).toBe('__XLUDF.DUMMYFUNCTION')
  })

  it('ignores names inside string literals and quoted sheet names', () => {
    expect(extractFunctionNames('CONCAT("SUM(",A1,")")')).toEqual(['CONCAT'])
    expect(extractFunctionNames('CONCAT("say ""FOO(""",A1)')).toEqual(['CONCAT'])
    expect(extractFunctionNames("SUM('Data (2024)'!A1:A3)")).toEqual(['SUM'])
    expect(extractFunctionNames("SUM('It''s (raw)'!A1)")).toEqual(['SUM'])
  })

  it('does not mistake references and names for calls', () => {
    expect(extractFunctionNames('Sheet1!A1+LOG10(A1)')).toEqual(['LOG10'])
    expect(extractFunctionNames('$A$1*Rate+TaxTable')).toEqual([])
    expect(extractFunctionNames('SUM(Table1[[#Headers],[Total (net)]])')).toEqual(['SUM'])
    expect(extractFunctionNames("[1]Sheet1!A1+'[2]Other'!B2")).toEqual([])
    expect(extractFunctionNames('1E5+2.5E-3*SUM(A1)')).toEqual(['SUM'])
    expect(extractFunctionNames('IF(B18>0,B18*الافتراضات!$B$30,0)')).toEqual(['IF'])
  })

  it('accepts blanks between the name and its parenthesis', () => {
    expect(extractFunctionNames('SUM (A1) + FOO\t(B1)')).toEqual(['SUM', 'FOO'])
  })
})

describe('formulaKeepsCache with the engine registry probe', () => {
  const install = (): void => {
    setSupportedFunctionProbe({ ready: () => true, supports: (name) => ENGINE_FUNCTIONS.has(name) })
  }

  it('keeps the cache for a bare unsupported call', () => {
    install()
    expect(formulaKeepsCache('FOOBARFN(A1)')).toBe(true)
  })

  it('keeps the cache when the unsupported call is wrapped in an error guard', () => {
    install()
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(true)
    expect(formulaKeepsCache('IF(ISERROR(FOOBARFN(A1)),"n/a",FOOBARFN(A1))')).toBe(true)
    expect(formulaKeepsCache('IFNA(FOOBARFN(A1),"")')).toBe(true)
    expect(formulaKeepsCache('IFS(ISERROR(FOOBARFN(A1)),0,TRUE(),FOOBARFN(A1))')).toBe(true)
    expect(formulaKeepsCache('CHOOSE(1,FOOBARFN(A1),0)')).toBe(true)
  })

  it('looks the function up by its plain name behind the _xlfn. marker', () => {
    install()
    expect(formulaKeepsCache('IFERROR(_xlfn.FOOBARFN(A1),0)')).toBe(true)
    expect(formulaKeepsCache('_xlfn.CONCAT(A1,_xlfn._xlws.FILTER(B:B,C:C))')).toBe(false)
  })

  it('does not count function names inside string literals', () => {
    install()
    expect(formulaKeepsCache('CONCAT("FOOBARFN(",A1,")")')).toBe(false)
    expect(formulaKeepsCache('IF(A1="FOOBARFN(x)",1,0)')).toBe(false)
  })

  it('recalculates formulas built only of supported functions', () => {
    install()
    expect(formulaKeepsCache('IFERROR(SUM(A1:A2),0)')).toBe(false)
    expect(formulaKeepsCache('IF(ISERROR(INDEX(A:A,MATCH(B1,C:C,0))),"",1)')).toBe(false)
    expect(formulaKeepsCache('sum(A1:A5)')).toBe(false)
  })

  it('treats sheet and cell references as references, not calls', () => {
    install()
    expect(formulaKeepsCache('Sheet1!A1+LOG10(A1)')).toBe(false)
    expect(formulaKeepsCache("SUM('Data (2024)'!A1:A3)")).toBe(false)
  })

  it('has nothing to keep without a cached value, so the engine gets the formula', () => {
    install()
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)', false)).toBe(false)
    expect(formulaKeepsCache('FOOBARFN(A1)', false)).toBe(false)
    // The pre-existing classes keep the cell cache-only regardless.
    expect(formulaKeepsCache('IFERROR(__xludf.DUMMYFUNCTION("X"),0)', false)).toBe(true)
    expect(formulaKeepsCache('DOLLAR(A1)', false)).toBe(true)
    expect(formulaKeepsCache('A1*TaxRate', false)).toBe(true)
  })

  it('assumes every function is supported until a probe is installed', () => {
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(false)
  })

  it('drops memoised verdicts when the probe changes', () => {
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(false)
    install()
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(true)
    setSupportedFunctionProbe({ ready: () => true, supports: () => true })
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(false)
  })

  it('keeps the pre-existing keep-cache classes', () => {
    install()
    expect(formulaKeepsCache('IFERROR(__xludf.DUMMYFUNCTION("""X"""),46235.0)')).toBe(true)
    expect(formulaKeepsCache('NUMBERSTRING(B2,1)')).toBe(true)
    expect(formulaKeepsCache('SUM(Revenue)')).toBe(true)
    expect(formulaKeepsCache('[1]Sheet1!A1*2')).toBe(true)
  })
})

describe('installSupportedFunctionProbe', () => {
  // A stand-in IFunctionService over a mutable executor set.
  const runtimeOver = (executors: Set<string>): UniverRuntime => {
    const functionService = {
      hasExecutor: (name: string) => executors.has(name),
      getExecutors: () => new Map([...executors].map((name) => [name, {}])),
    }
    return {
      univer: { __getInjector: () => ({ get: () => functionService }) },
    } as unknown as UniverRuntime
  }

  it('answers from the live registry', () => {
    const disposable = installSupportedFunctionProbe(
      runtimeOver(new Set(['SUM', 'IFERROR', 'CELL'])),
    )
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(true)
    expect(formulaKeepsCache('IFERROR(SUM(A1:A2),0)')).toBe(false)
    expect(formulaKeepsCache('CELL("filename",A1)')).toBe(false)
    disposable.dispose()
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(false)
  })

  it('assumes support until the builtin batch lands, without memoising', () => {
    // Before the engine plugin's onReady only the app's own executors exist:
    // a non-empty registry that still knows nothing about SUM or IFERROR.
    const executors = new Set(['CELL', 'RATE', 'MINIFS'])
    const disposable = installSupportedFunctionProbe(runtimeOver(executors))
    expect(formulaKeepsCache('IFERROR(SUM(A1:A2),0)')).toBe(false)
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(false)
    for (const name of ['SUM', 'IFERROR', 'IF']) executors.add(name)
    // The same formulas, asked again once the batch has landed: a verdict
    // memoised in the blind window would still say "recompute".
    expect(formulaKeepsCache('IFERROR(FOOBARFN(A1),0)')).toBe(true)
    expect(formulaKeepsCache('IFERROR(SUM(A1:A2),0)')).toBe(false)
    disposable.dispose()
  })
})

describe('hasUsableCachedValue', () => {
  it('accepts numbers, text and booleans, rejects blanks and error literals', () => {
    expect(hasUsableCachedValue(42)).toBe(true)
    expect(hasUsableCachedValue(0)).toBe(true)
    expect(hasUsableCachedValue('hello')).toBe(true)
    expect(hasUsableCachedValue('')).toBe(true)
    expect(hasUsableCachedValue(false)).toBe(true)
    expect(hasUsableCachedValue(null)).toBe(false)
    expect(hasUsableCachedValue(undefined)).toBe(false)
    expect(hasUsableCachedValue('#NAME?')).toBe(false)
    expect(hasUsableCachedValue('#N/A')).toBe(false)
  })
})
