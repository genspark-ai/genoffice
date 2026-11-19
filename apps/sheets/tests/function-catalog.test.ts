import { FunctionType, type IFunctionInfo } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { buildFunctionCatalog, functionSyntax } from '../src/renderer/function-catalog'

function info(
  name: string,
  type: FunctionType,
  params: { name: string; require?: boolean; repeat?: boolean }[] = [],
): IFunctionInfo {
  return {
    functionName: name,
    functionType: type,
    description: `${name} long`,
    abstract: `${name} short`,
    functionParameter: params.map((param) => ({
      name: param.name,
      detail: '',
      example: '',
      require: param.require === false ? 0 : 1,
      repeat: param.repeat ? 1 : 0,
    })),
  }
}

describe('functionSyntax', () => {
  it('brackets optional params and appends an ellipsis for repeating ones', () => {
    const xlookup = info('XLOOKUP', FunctionType.Lookup, [
      { name: 'lookup_value' },
      { name: 'lookup_array' },
      { name: 'return_array' },
      { name: 'if_not_found', require: false },
    ])
    expect(functionSyntax(xlookup)).toBe(
      'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])',
    )
    const sum = info('SUM', FunctionType.Math, [
      { name: 'number1' },
      { name: 'number2', require: false, repeat: true },
    ])
    expect(functionSyntax(sum)).toBe('SUM(number1, [number2], …)')
    expect(functionSyntax(info('NOW', FunctionType.Date))).toBe('NOW()')
  })
})

describe('buildFunctionCatalog', () => {
  const live = [
    info('XLOOKUP', FunctionType.Lookup, [{ name: 'v' }]),
    info('LET', FunctionType.Logical),
    info('FILTER', FunctionType.Lookup),
    info('MyRange', FunctionType.DefinedName),
    info('Table1', FunctionType.Table),
    info('CUSTOMFN', FunctionType.User),
  ]

  it('lists every engine-described function sorted by name, skipping names and tables', () => {
    const names = buildFunctionCatalog(live).map((spec) => spec.name)
    expect(names).toEqual(['CUSTOMFN', 'FILTER', 'LET', 'XLOOKUP'])
  })

  it('maps engine types to Excel categories and carries both texts', () => {
    const catalog = buildFunctionCatalog(live)
    const xlookup = catalog.find((spec) => spec.name === 'XLOOKUP')
    expect(xlookup).toMatchObject({
      category: 'Lookup',
      syntax: 'XLOOKUP(v)',
      abstract: 'XLOOKUP short',
      description: 'XLOOKUP long',
    })
    expect(catalog.find((spec) => spec.name === 'CUSTOMFN')?.category).toBe('Other')
  })

  it('uses fallback entries only for names the engine does not describe', () => {
    const fallback = [
      {
        name: 'CELL',
        category: 'Information',
        syntax: 'CELL(info_type, [reference])',
        abstract: 'c',
        description: 'c',
      },
      {
        name: 'XLOOKUP',
        category: 'Lookup',
        syntax: 'stale',
        abstract: 'stale',
        description: 'stale',
      },
    ]
    const catalog = buildFunctionCatalog(live, fallback)
    expect(catalog.find((spec) => spec.name === 'CELL')?.syntax).toBe(
      'CELL(info_type, [reference])',
    )
    expect(catalog.find((spec) => spec.name === 'XLOOKUP')?.syntax).toBe('XLOOKUP(v)')
  })
})
