import { FunctionType, type IFunctionInfo } from '@univerjs/engine-formula'
import { IDescriptionService } from '@univerjs/sheets-formula'

import type { UniverRuntime } from './univer-state'

/// One Insert Function row. `category` is a stable English id, displayed
/// through the dialog's label table; `abstract` is the one-line list text,
/// `description` the longer text shown once a function is picked.
export interface FunctionSpec {
  readonly name: string
  readonly category: string
  readonly syntax: string
  readonly abstract: string
  readonly description: string
}

/// Excel's Insert Function category order.
export const FUNCTION_CATEGORIES = [
  'Financial',
  'Date & Time',
  'Math',
  'Statistical',
  'Lookup',
  'Database',
  'Text',
  'Logical',
  'Information',
  'Engineering',
  'Cube',
  'Compatibility',
  'Web',
  'Array',
  'Other',
] as const

export type FunctionCategory = (typeof FUNCTION_CATEGORIES)[number]

const CATEGORY_BY_TYPE: Partial<Record<FunctionType, FunctionCategory>> = {
  [FunctionType.Financial]: 'Financial',
  [FunctionType.Date]: 'Date & Time',
  [FunctionType.Math]: 'Math',
  [FunctionType.Statistical]: 'Statistical',
  [FunctionType.Lookup]: 'Lookup',
  [FunctionType.Database]: 'Database',
  [FunctionType.Text]: 'Text',
  [FunctionType.Logical]: 'Logical',
  [FunctionType.Information]: 'Information',
  [FunctionType.Engineering]: 'Engineering',
  [FunctionType.Cube]: 'Cube',
  [FunctionType.Compatibility]: 'Compatibility',
  [FunctionType.Web]: 'Web',
  [FunctionType.Array]: 'Array',
  [FunctionType.User]: 'Other',
}

/// Excel syntax line: required params bare, optional in brackets, a
/// repeating tail as an ellipsis.
export function functionSyntax(
  info: Pick<IFunctionInfo, 'functionName' | 'functionParameter'>,
): string {
  const parts = info.functionParameter.map((param) =>
    param.require ? param.name : `[${param.name}]`,
  )
  if (info.functionParameter.some((param) => param.repeat)) parts.push('…')
  return `${info.functionName}(${parts.join(', ')})`
}

/// Descriptions for the active Univer locale. The description service is
/// what resolves the locale keys into the function registry (lazily, on first
/// resolve), so it is asked rather than the registry directly.
export function readLiveFunctionInfos(runtime: UniverRuntime): IFunctionInfo[] {
  return [...runtime.univer.__getInjector().get(IDescriptionService).getDescriptions().values()]
}

/**
 * Catalog for the Insert Function dialog: every function the running engine
 * describes, plus `fallback` entries for names the engine has no description
 * for (executors the app registers itself). Sorted by name.
 */
export function buildFunctionCatalog(
  live: readonly IFunctionInfo[],
  fallback: readonly FunctionSpec[] = [],
): FunctionSpec[] {
  const byName = new Map<string, FunctionSpec>()
  for (const info of live) {
    const category = CATEGORY_BY_TYPE[info.functionType]
    if (!category) continue
    const name = info.functionName.toUpperCase()
    byName.set(name, {
      name,
      category,
      syntax: functionSyntax({ ...info, functionName: name }),
      abstract: info.abstract,
      description: info.description,
    })
  }
  for (const spec of fallback) {
    if (!byName.has(spec.name)) byName.set(spec.name, spec)
  }
  return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name))
}
