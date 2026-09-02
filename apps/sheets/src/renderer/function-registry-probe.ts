/**
 * Wires the keep-cached-value decision to the live Univer function registry.
 *
 * A file formula calling a function the engine lacks recalculates to #NAME?,
 * and when the author wrapped the call (IFERROR/ISERROR/IFNA/IFS/CHOOSE) the
 * error is swallowed into the fallback literal — a wrong number where the
 * file's cached <v> holds Excel's true result. The registry is the source of
 * truth for "implemented": builtins, plus every executor the app registers
 * itself (CELL, RATE, MINIFS/MAXIFS overrides), all land there.
 */
import { FUNCTION_NAMES_MATH, IFunctionService } from '@univerjs/engine-formula'

import { setSupportedFunctionProbe } from './univer-sync'
import type { UniverRuntime } from './univer-state'

/// The engine registers every builtin in one batch at its plugin's onReady;
/// SUM being visible means the batch has landed. onReady may run after this
/// installer — and after the app's own CELL / RATE / MINIFS executors, so a
/// non-empty registry is not a complete one.
const BUILTIN_SENTINEL = FUNCTION_NAMES_MATH.SUM

export function installSupportedFunctionProbe(runtime: UniverRuntime): { dispose(): void } {
  const functionService = runtime.univer.__getInjector().get(IFunctionService)
  setSupportedFunctionProbe({
    ready: () => functionService.hasExecutor(BUILTIN_SENTINEL),
    supports: (name) => functionService.hasExecutor(name),
  })
  return {
    dispose() {
      setSupportedFunctionProbe(null)
    },
  }
}
