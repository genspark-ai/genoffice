import { describe, expect, it } from 'vitest'

import {
  spillsDynamicArray,
  withFutureFunctionMarkers,
} from '@genoffice/xlsx-gateway/gateway/future-functions'
import { dynamicArrayCellMetaIndex } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

describe('withFutureFunctionMarkers', () => {
  it('prefixes future functions for storage', () => {
    expect(withFutureFunctionMarkers('MINIFS(C7:C10,C7:C10,">0")')).toBe(
      '_xlfn.MINIFS(C7:C10,C7:C10,">0")',
    )
    expect(withFutureFunctionMarkers('IFS(A1>0,STDEV.P(B:B),TRUE,0)')).toBe(
      '_xlfn.IFS(A1>0,_xlfn.STDEV.P(B:B),TRUE,0)',
    )
  })

  it('marks worksheet-scope dynamic-array functions with _xlws', () => {
    expect(withFutureFunctionMarkers('SORT(FILTER(A:A,B:B>0))')).toBe(
      '_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A:A,B:B>0))',
    )
  })

  it('canonicalizes lowercase and mixed-case calls', () => {
    expect(withFutureFunctionMarkers('minifs(A:A,A:A,">0")')).toBe('_xlfn.MINIFS(A:A,A:A,">0")')
    expect(withFutureFunctionMarkers('XLookup(1,A:A,B:B)')).toBe('_xlfn.XLOOKUP(1,A:A,B:B)')
    expect(withFutureFunctionMarkers('sum(A1:A3)')).toBe('sum(A1:A3)')
  })

  it('restores the implicit-intersection and spill storage forms', () => {
    expect(withFutureFunctionMarkers('SINGLE(A:A)+SUM(ANCHORARRAY(C3))')).toBe(
      '_xlfn.SINGLE(A:A)+SUM(_xlfn.ANCHORARRAY(C3))',
    )
  })

  it('leaves classic functions, strings, and existing markers alone', () => {
    expect(withFutureFunctionMarkers('SUM(A1:A3)+IF(B1,1,0)')).toBe('SUM(A1:A3)+IF(B1,1,0)')
    expect(withFutureFunctionMarkers('CONCATENATE("MINIFS(",A1,")")')).toBe(
      'CONCATENATE("MINIFS(",A1,")")',
    )
    expect(withFutureFunctionMarkers('_xlfn.MINIFS(A:A,A:A,">0")')).toBe(
      '_xlfn.MINIFS(A:A,A:A,">0")',
    )
  })
})

describe('spillsDynamicArray', () => {
  it('recognises spill functions at any depth, in any case, with or without markers', () => {
    expect(spillsDynamicArray('FILTER(A:A,B:B>0)')).toBe(true)
    expect(spillsDynamicArray('IFERROR(sort(unique(A1:A9)),"")')).toBe(true)
    expect(spillsDynamicArray('_xlfn._xlws.FILTER(A:A,B:B)')).toBe(true)
    expect(spillsDynamicArray('SUM(FILTER(A:A,B:B>0))')).toBe(true)
  })

  it('leaves scalar formulas and string literals alone', () => {
    expect(spillsDynamicArray('SUM(A1:A3)+XLOOKUP(1,A:A,B:B)')).toBe(false)
    expect(spillsDynamicArray('CONCAT("FILTER(",A1,")")')).toBe(false)
  })
})

describe('dynamicArrayCellMetaIndex', () => {
  const part = (types: string[], records: number[]) =>
    `<metadata><metadataTypes count="${types.length}">${types
      .map((n) => `<metadataType name="${n}" minSupportedVersion="120000"/>`)
      .join('')}</metadataTypes><cellMetadata count="${records.length}">${records
      .map((t) => `<bk><rc t="${t}" v="0"/></bk>`)
      .join('')}</cellMetadata></metadata>`

  it('finds the record that points at XLDAPR wherever the type sits', () => {
    expect(dynamicArrayCellMetaIndex(part(['XLDAPR'], [1]))).toBe(1)
    expect(dynamicArrayCellMetaIndex(part(['XLRICHVALUE', 'XLDAPR'], [1, 2]))).toBe(2)
    expect(dynamicArrayCellMetaIndex(part(['XLRICHVALUE', 'XLDAPR'], [2]))).toBe(1)
  })

  it('refuses when XLDAPR is absent or has no cell record', () => {
    expect(dynamicArrayCellMetaIndex(part(['XLRICHVALUE'], [1]))).toBeNull()
    expect(dynamicArrayCellMetaIndex(part(['XLRICHVALUE', 'XLDAPR'], [1]))).toBeNull()
  })
})
