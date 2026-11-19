/**
 * CELL("filename") support. Univer's builtin CELL demands exactly two
 * arguments and hard-codes the "filename" info_type to #VALUE!, which breaks
 * the ubiquitous `=MID(CELL("filename",B1),FIND("]",...)+1,31)` sheet-name
 * idiom. This executor replaces it: "filename" answers Excel's
 * `dir/[Book.xlsx]Sheet` (empty string while the workbook has never been
 * saved), everything else delegates to the builtin.
 */
import {
  BaseFunction,
  ErrorType,
  ErrorValueObject,
  FUNCTION_NAMES_INFORMATION,
  IFunctionService,
  StringValueObject,
  type BaseValueObject,
} from '@univerjs/engine-formula'

type CalcResult = ReturnType<BaseFunction['calculate']>

import type { UniverRuntime } from './univer-state'

export type WorkbookPathProvider = () => string | null | undefined

interface ValueLike {
  isError(): boolean
  isReferenceObject?(): boolean
  isArray?(): boolean
  getValue?(): unknown
}

export class CellWithFilename extends BaseFunction {
  override needsReferenceObject = true
  override needsSheetsInfo = true
  override minParams = 1
  override maxParams = 2

  constructor(
    name: string,
    private _original: BaseFunction | undefined,
    private readonly _getPath: WorkbookPathProvider,
  ) {
    super(name)
  }

  setOriginal(original: BaseFunction | undefined): void {
    this._original = original
  }

  getOriginal(): BaseFunction | undefined {
    return this._original
  }

  override calculate(rawInfoType: BaseValueObject, rawReference?: BaseValueObject): CalcResult {
    const infoType = rawInfoType as ValueLike
    const reference = rawReference as ValueLike | undefined
    if (infoType.isError()) return rawInfoType
    let single: ValueLike & { toArrayValueObject?(): ValueLike } = infoType
    if (single.isReferenceObject?.()) single = single.toArrayValueObject!()
    if (single.isArray?.()) {
      const arr = single as ValueLike & {
        getRowCount(): number
        getColumnCount(): number
        get(r: number, c: number): ValueLike
      }
      if (arr.getRowCount() !== 1 || arr.getColumnCount() !== 1) {
        return this._delegate(infoType, reference)
      }
      single = arr.get(0, 0)
      if (single.isError()) return single as CalcResult
    }
    if (`${single.getValue?.()}`.toLowerCase() === 'filename') return this._filename(reference)
    return this._delegate(infoType, reference)
  }

  private _delegate(infoType: ValueLike, reference?: ValueLike): CalcResult {
    // The builtin dereferences its second argument unconditionally; Excel's
    // ref-less form ("the last changed cell") is unknowable here, so only
    // "filename" supports it.
    if (!reference || !this._original) return ErrorValueObject.create(ErrorType.NA)
    return (this._original.calculate as (a: ValueLike, b: ValueLike) => CalcResult)(
      infoType,
      reference,
    )
  }

  private _filename(reference?: ValueLike): CalcResult {
    if (reference?.isError()) return reference as CalcResult
    const path = this._getPath()
    if (!path) return StringValueObject.create('')
    let sheetId = this.subUnitId
    if (reference?.isReferenceObject?.()) {
      const ref = reference as ValueLike & {
        getForcedSheetId(): string | undefined
        getDefaultSheetId(): string | undefined
      }
      sheetId = ref.getForcedSheetId() ?? ref.getDefaultSheetId() ?? this.subUnitId
    }
    const { sheetNameMap } = this.getSheetsInfo() as {
      sheetNameMap: Record<string, string> | undefined
    }
    const sheetName = (sheetId && sheetNameMap?.[sheetId]) || ''
    const cut = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\')) + 1
    return StringValueObject.create(`${path.slice(0, cut)}[${path.slice(cut)}]${sheetName}`)
  }
}

export function installCellFilenameFunction(
  runtime: UniverRuntime,
  getPath: WorkbookPathProvider,
): { dispose(): void } {
  const functionService = runtime.univer.__getInjector().get(IFunctionService)
  const name = FUNCTION_NAMES_INFORMATION.CELL
  const executor = new CellWithFilename(
    name,
    functionService.getExecutor(name) ?? undefined,
    getPath,
  )
  functionService.registerExecutors(executor)
  // Registration order vs. the builtin set is not guaranteed; re-assert once
  // (adopting a builtin that landed on top as the delegate).
  const timer = setTimeout(() => {
    const current = functionService.getExecutor(name)
    if (current !== executor) {
      executor.setOriginal(current ?? undefined)
      functionService.registerExecutors(executor)
    }
  }, 0)
  return {
    dispose() {
      clearTimeout(timer)
      const original = executor.getOriginal()
      if (original) functionService.registerExecutors(original)
    },
  }
}
