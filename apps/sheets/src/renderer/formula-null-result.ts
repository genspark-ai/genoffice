/**
 * Empty-value formula results display as 0, like Excel.
 *
 * A bare `=H19` on an empty cell already shows 0: the root reference goes
 * through BaseReferenceObject.getFirstCell(), which maps a missing cell to
 * NumberValueObject(0). But an empty reference passed THROUGH a function
 * (IFERROR, IF, CHOOSE, SWITCH, …) arrives as a NullValueObject, is returned
 * as-is, and the runtime stores `v: null` — a blank cell where Excel shows 0.
 *
 * Coercing at the root keeps intermediate semantics intact: `=IFERROR(H19,
 * "NA")&""` still concatenates the emptiness to "" before the root ever sees
 * it, matching Excel.
 */
import { IFormulaRuntimeService, NumberValueObject } from '@univerjs/engine-formula'

import type { UniverRuntime } from './univer-state'

interface VariantLike {
  isValueObject?(): boolean
  isArray?(): boolean
  isNull?(): boolean
}

export function coerceNullResult<T>(variant: T): T | NumberValueObject {
  const value = variant as VariantLike | null | undefined
  if (value?.isValueObject?.() && !value.isArray?.() && value.isNull?.()) {
    return NumberValueObject.create(0)
  }
  return variant
}

export function installFormulaNullResultFix(runtime: UniverRuntime): { dispose(): void } {
  const runtimeService = runtime.univer.__getInjector().get(IFormulaRuntimeService)
  const original = runtimeService.setRuntimeData.bind(runtimeService)
  runtimeService.setRuntimeData = (variant) => original(coerceNullResult(variant))
  return {
    dispose() {
      runtimeService.setRuntimeData = original
    },
  }
}
