/** Word's Preferences ▸ General ▸ "Show measurements in units of": one
 * preference every length field (ruler, margins, indents, sizes) follows.
 * Storage stays in twips; only presentation changes. Font size, line and
 * paragraph spacing stay in points, as in Word.
 *
 * The preference is app-wide (Settings ▸ General, held by the shell, see
 * @genoffice/ui/measurement); this module keeps the editor's copy of it and
 * re-exports the shared conversions. */
import {
  defaultMeasurementUnit,
  isMeasurementUnit,
  isMeasurementUnitState,
  type MeasurementUnit,
  type MeasurementUnitState,
} from '@genoffice/ui/measurement'

export {
  MEASUREMENT_UNITS,
  UNIT_STEP,
  UNIT_SUFFIX,
  defaultMeasurementUnit,
  formatLength,
  fromUnit,
  isMeasurementUnit,
  parseLength,
  pxToTwips,
  toUnit,
  twipsPerUnit,
  twipsToPx,
} from '@genoffice/ui/measurement'
export type { MeasurementUnit } from '@genoffice/ui/measurement'

/** where Docs kept its own unit before the setting became app-wide (migrated once) */
export const MEASUREMENT_UNIT_KEY = 'aidocs.measurementUnit'
const CHANGE_EVENT = 'aidocs:measurement-unit'

/** the shell's answer once it arrives; null standalone or before it answers */
let appWide: MeasurementUnitState | null = null
let appWideSetter: ((unit: MeasurementUnit | null) => void) | null = null

export function storedMeasurementUnit(): MeasurementUnit | null {
  try {
    const raw = globalThis.localStorage?.getItem(MEASUREMENT_UNIT_KEY)
    return isMeasurementUnit(raw) ? raw : null
  } catch {
    return null
  }
}

export function measurementUnit(): MeasurementUnit {
  return (
    appWide?.unit ??
    storedMeasurementUnit() ??
    defaultMeasurementUnit(globalThis.navigator?.language ?? 'en-US')
  )
}

const notify = (): void => {
  globalThis.dispatchEvent?.(new Event(CHANGE_EVENT))
}

/** pick the unit (null = follow the system region); app-wide when the shell is there */
export function setMeasurementUnit(unit: MeasurementUnit | null): void {
  if (appWideSetter) {
    // the shell writes it and broadcasts the result back to every editor
    appWideSetter(unit)
    return
  }
  try {
    if (unit) globalThis.localStorage?.setItem(MEASUREMENT_UNIT_KEY, unit)
    else globalThis.localStorage?.removeItem(MEASUREMENT_UNIT_KEY)
  } catch {
    /* storage unavailable: the choice just doesn't persist */
  }
  notify()
}

export function subscribeMeasurementUnit(listener: () => void): () => void {
  window.addEventListener(CHANGE_EVENT, listener)
  window.addEventListener('storage', listener)
  return () => {
    window.removeEventListener(CHANGE_EVENT, listener)
    window.removeEventListener('storage', listener)
  }
}

export interface AppWideMeasurementSource {
  get(): Promise<unknown>
  set(unit: MeasurementUnit | null): Promise<unknown> | void
  onChanged(handler: (state: unknown) => void): () => void
}

/**
 * Follow the shell's app-wide unit from now on. A unit picked in Docs before
 * the setting moved to the shell is carried over once, when the shell still
 * follows the system region, and then forgotten locally.
 */
export function connectAppWideMeasurementUnit(source: AppWideMeasurementSource): () => void {
  const apply = (state: unknown): void => {
    if (!isMeasurementUnitState(state)) return
    appWide = state
    notify()
  }
  appWideSetter = (unit) => {
    void Promise.resolve(source.set(unit)).catch(() => undefined)
  }
  const off = source.onChanged(apply)
  void source
    .get()
    .then((state) => {
      const legacy = storedMeasurementUnit()
      if (legacy && isMeasurementUnitState(state) && !state.explicit) {
        void Promise.resolve(source.set(legacy)).catch(() => undefined)
      }
      if (legacy) {
        try {
          globalThis.localStorage?.removeItem(MEASUREMENT_UNIT_KEY)
        } catch {
          /* nothing to forget */
        }
      }
      apply(state)
    })
    .catch(() => undefined)
  return () => {
    off()
    appWideSetter = null
  }
}
