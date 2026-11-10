/**
 * MINIFS/MAXIFS over an empty match set must return 0 (Excel semantics).
 * Univer's empty-set branch builds `ArrayValueObject.create('0')`, but the
 * string form is parsed as an array literal whose brace-stripping slice turns
 * '0' into '' — the cell renders blank. Wrap both executors and rewrite that
 * 1x1 empty-string artifact into a real 0.
 */
import {
  ArrayValueObject,
  BaseFunction,
  FUNCTION_NAMES_STATISTICAL,
  IFunctionService,
  NumberValueObject,
  functionStatistical,
} from '@univerjs/engine-formula'
import type { BaseValueObject } from '@univerjs/engine-formula'

import type { UniverRuntime } from './univer-state'

interface ValueLike {
  isArray?(): boolean
  isString?(): boolean
  getValue(): unknown
  getRowCount?(): number
  getColumnCount?(): number
  getFirstCell?(): ValueLike
  getArrayValue?(): ValueLike[][]
}

function isEmptySetArtifact(value: ValueLike | undefined): boolean {
  if (!value?.isArray?.() || value.getRowCount?.() !== 1 || value.getColumnCount?.() !== 1) {
    return false
  }
  const cell = value.getFirstCell?.()
  return cell?.isString?.() === true && cell.getValue() === ''
}

export function fixIfsEmptySetResult(value: ValueLike): ValueLike {
  if (isEmptySetArtifact(value)) return NumberValueObject.create(0) as unknown as ValueLike
  if (!value?.isArray?.() || !value.getArrayValue) return value
  // Array-criteria expansion: each zero-match slot carries its own artifact.
  const rows = value.getArrayValue()
  if (!rows.some((row) => row.some(isEmptySetArtifact))) return value
  const mapped = rows.map((row) =>
    row.map((cell) =>
      isEmptySetArtifact(cell) ? (NumberValueObject.create(0) as unknown as ValueLike) : cell,
    ),
  )
  return ArrayValueObject.create({
    calculateValueList: mapped,
    rowCount: mapped.length,
    columnCount: mapped[0]?.length ?? 0,
    unitId: '',
    sheetId: '',
    row: 0,
    column: 0,
  } as unknown as Parameters<typeof ArrayValueObject.create>[0]) as unknown as ValueLike
}

function createExecutor(name: string): BaseFunction | null {
  const Builtin = functionStatistical.find(([, builtinName]) => builtinName === name)?.[0] as
    (new (functionName: string) => BaseFunction) | undefined
  if (!Builtin) return null
  class IfsWithEmptySetZero extends Builtin {
    override calculate(...variants: BaseValueObject[]): ReturnType<BaseFunction['calculate']> {
      return fixIfsEmptySetResult(super.calculate(...variants) as ValueLike) as ReturnType<
        BaseFunction['calculate']
      >
    }
  }
  return new IfsWithEmptySetZero(name)
}

export function createIfsEmptySetExecutors(): BaseFunction[] {
  return [FUNCTION_NAMES_STATISTICAL.MINIFS, FUNCTION_NAMES_STATISTICAL.MAXIFS]
    .map(createExecutor)
    .filter((executor): executor is BaseFunction => executor !== null)
}

export function installIfsEmptySetFix(runtime: UniverRuntime): { dispose(): void } {
  const functionService = runtime.univer.__getInjector().get(IFunctionService)
  const executors = createIfsEmptySetExecutors()
  if (executors.length === 0) return { dispose() {} }
  const originals = executors.map(
    (executor) => functionService.getExecutor(executor.name) ?? undefined,
  )
  functionService.registerExecutors(...executors)
  const timer = setTimeout(() => {
    const stale = executors.filter(
      (executor) => functionService.getExecutor(executor.name) !== executor,
    )
    if (stale.length > 0) functionService.registerExecutors(...stale)
  }, 0)
  return {
    dispose() {
      clearTimeout(timer)
      for (const original of originals) {
        if (original) functionService.registerExecutors(original)
      }
    },
  }
}
