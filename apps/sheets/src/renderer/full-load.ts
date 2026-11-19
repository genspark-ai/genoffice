import { preloadEntireWorkbook } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/// The preloads in flight, one per workbook: `preloadEntireWorkbook` has no
/// guard of its own, and a second concurrent run would re-read and re-install
/// every block while the first is still filling the model.
const inFlight = new WeakMap<LazyWorkbookState, Promise<void>>()

/// The one way to fully load the open workbook. Every trigger (the Load-all
/// prompt, formula-mode open, headless export, Print / PDF export) goes
/// through here, so a second request while a load runs joins the running one
/// instead of starting another. Progress messages go to the caller that
/// started the load. Resolves without a load when the workbook is already
/// fully loaded or none is open.
export function startFullLoad(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  setMessage: (message: string) => void,
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state || state.flags.preloadComplete) return Promise.resolve()
  const running = inFlight.get(state)
  if (running) return running
  const load = preloadEntireWorkbook(runtime, lazyWorkbookRef, setMessage).finally(() => {
    if (inFlight.get(state) === load) inFlight.delete(state)
  })
  inFlight.set(state, load)
  return load
}
