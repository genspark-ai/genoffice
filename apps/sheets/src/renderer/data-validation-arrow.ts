import { DataValidationRenderMode, getOriginCellValue } from '@univerjs/core'
import { DataValidatorRegistryService } from '@univerjs/data-validation'

import type { UniverRuntime } from './univer-state'

interface ValidationCanvasRenderer {
  drawWith: ((...args: unknown[]) => unknown) | undefined
  isHit: ((...args: unknown[]) => unknown) | undefined
  _dataValidationModel?: {
    getRuleByLocation(
      unitId: string,
      subUnitId: string,
      row: number,
      col: number,
    ): { renderMode: DataValidationRenderMode } | undefined
  }
}

interface ListValidator {
  canvasRender?: ValidationCanvasRenderer | null
}

function hasValue(value: unknown): boolean {
  return value !== undefined && value !== null && value !== ''
}

function rendererHasValue(info: unknown): boolean {
  if (!info || typeof info !== 'object') return false
  const data = (info as { data?: unknown }).data
  return hasValue(getOriginCellValue(data as never))
}

/**
 * Keep validated cells in TEXT mode so their workbook styling and text layout
 * stay untouched. For populated cells only, invoke Univer's arrow renderer
 * with an empty display value: this adds the arrow and hit target without
 * repainting the text or background.
 */
export function installPopulatedDataValidationArrow(runtime: UniverRuntime): { dispose(): void } {
  const registry = runtime.univer.__getInjector().get(DataValidatorRegistryService)
  const wrappedRenderers = new WeakSet<ValidationCanvasRenderer>()
  const restoreRenderers: Array<() => void> = []
  let animationFrame: number | undefined
  let attempts = 0
  let disposed = false

  const install = () => {
    if (disposed) return
    const validator = registry.getValidatorItem('list') as ListValidator | undefined
    const renderer = validator?.canvasRender

    // Univer assigns its final canvas renderer during the Rendered lifecycle.
    // Check for a replacement for the first second instead of wrapping a
    // short-lived bootstrap instance.
    if (renderer && !wrappedRenderers.has(renderer)) {
      wrappedRenderers.add(renderer)
      const originalDraw = renderer.drawWith
      const originalIsHit = renderer.isHit

      if (originalDraw) {
        renderer.drawWith = function (...args: unknown[]) {
          const info = args[1] as
            | {
                data?: Record<string, unknown>
                unitId?: string
                subUnitId?: string
                row?: number
                col?: number
              }
            | undefined
          if (
            !rendererHasValue(info) ||
            !info?.data ||
            info.unitId === undefined ||
            info.subUnitId === undefined ||
            info.row === undefined ||
            info.col === undefined
          ) {
            return
          }
          const rule = renderer._dataValidationModel?.getRuleByLocation(
            info.unitId,
            info.subUnitId,
            info.row,
            info.col,
          )
          if (!rule) return
          const previousMode = rule.renderMode
          rule.renderMode = DataValidationRenderMode.ARROW
          try {
            const arrowInfo = { ...info, data: { ...info.data, v: '' } }
            return originalDraw.apply(this, [args[0], arrowInfo, ...args.slice(2)])
          } finally {
            rule.renderMode = previousMode
          }
        }
      }
      if (originalIsHit) {
        renderer.isHit = function (...args: unknown[]) {
          const info = args[1] as
            | {
                data?: Record<string, unknown>
                unitId?: string
                subUnitId?: string
                row?: number
                col?: number
              }
            | undefined
          if (
            !rendererHasValue(info) ||
            info?.unitId === undefined ||
            info.subUnitId === undefined ||
            info.row === undefined ||
            info.col === undefined
          ) {
            return false
          }
          const rule = renderer._dataValidationModel?.getRuleByLocation(
            info.unitId,
            info.subUnitId,
            info.row,
            info.col,
          )
          if (!rule) return false
          const previousMode = rule.renderMode
          rule.renderMode = DataValidationRenderMode.ARROW
          try {
            return originalIsHit.apply(this, args)
          } finally {
            rule.renderMode = previousMode
          }
        }
      }
      restoreRenderers.push(() => {
        renderer.drawWith = originalDraw
        renderer.isHit = originalIsHit
      })
    }

    attempts += 1
    if (attempts < 60 && typeof globalThis.requestAnimationFrame === 'function') {
      animationFrame = globalThis.requestAnimationFrame(install)
    }
  }

  install()

  return {
    dispose: () => {
      disposed = true
      if (animationFrame !== undefined) globalThis.cancelAnimationFrame(animationFrame)
      for (const restore of restoreRenderers) restore()
    },
  }
}
