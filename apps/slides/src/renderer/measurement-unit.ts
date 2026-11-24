/**
 * The app-wide measurement unit (Settings ▸ General ▸ "Show measurements in
 * units of") as Slides sees it: the ruler, guide readouts, Format pane sizes,
 * positions and text insets, bullet indents and the Slide Size dialog all
 * follow it. The store lives in @genoffice/ui; this adds Slides' unit labels.
 */
import { MEASUREMENT_UNITS, type MeasurementUnit } from '@genoffice/ui/measurement'

import type { TFunc } from './i18n/locale'

export {
  connectMeasurementUnit,
  measurementUnit,
  resetMeasurementUnitForTests,
  useMeasurementUnit,
} from '@genoffice/ui/use-measurement-unit'

/** short unit labels shown after a length (pt is the Format pane's point label) */
const UNIT_ABBR_KEYS = {
  in: 'paneSizeIn',
  cm: 'paneSizeCm',
  mm: 'paneSizeMm',
  pt: 'paneFormatPt',
  pi: 'paneSizePi',
} as const satisfies Record<MeasurementUnit, string>

/** localized unit labels, for formatLength's suffix and parseLength's local suffixes */
export function unitSuffixes(t: TFunc): Record<MeasurementUnit, string> {
  return Object.fromEntries(MEASUREMENT_UNITS.map((u) => [u, t(UNIT_ABBR_KEYS[u])])) as Record<
    MeasurementUnit,
    string
  >
}
