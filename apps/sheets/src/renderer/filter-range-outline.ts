/**
 * Excel draws no outline around an auto-filtered range, but Univer's
 * SheetsFilterRenderController paints a selection-style border over the whole
 * filter range whenever a sheet has a filter (alpha feedback: "a mysterious
 * outer border appeared"). The controller class isn't exported, so it is
 * caught at render-module registration — recognized by its distinctive
 * prototype — and its range painter is stubbed out. The funnel buttons keep
 * rendering; only the range outline goes.
 */
import { IRenderManagerService } from '@univerjs/engine-render'

import type { UniverRuntime } from './univer-state'

export function installFilterRangeOutlineSuppression(runtime: UniverRuntime): void {
  const service = runtime.univer.__getInjector().get(IRenderManagerService)
  const original = service.registerRenderModule.bind(service)
  const patched: typeof service.registerRenderModule = (type, dep) => {
    const ctor = Array.isArray(dep) ? dep[0] : dep
    const prototype = (ctor as { prototype?: Record<string, unknown> } | undefined)?.prototype
    if (
      prototype &&
      typeof prototype._renderRange === 'function' &&
      typeof prototype._renderButtons === 'function'
    ) {
      prototype._renderRange = () => {}
    }
    return original(type, dep)
  }
  ;(service as { registerRenderModule: typeof service.registerRenderModule }).registerRenderModule =
    patched
}
