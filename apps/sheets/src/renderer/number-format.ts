import { DEFAULT_NUMFMT_OPTIONS, numfmtPattern, type NumfmtOptions } from './numfmt-dialog'
import { numberFormatLabel } from './selection-format'

const pattern = (overrides: Partial<NumfmtOptions>): string =>
  numfmtPattern({ ...DEFAULT_NUMFMT_OPTIONS, ...overrides })

/// Excel's Number-group category dropdown: label → representative pattern,
/// built by the same builders as the Format Cells dialog (numfmt-dialog.ts).
export const NUMBER_FORMAT_CATEGORIES: ReadonlyArray<{
  readonly label: string
  readonly pattern: string
}> = [
  { label: 'General', pattern: pattern({ category: 'general' }) },
  { label: 'Number', pattern: pattern({ category: 'number' }) },
  { label: 'Currency', pattern: pattern({ category: 'currency' }) },
  { label: 'Accounting', pattern: pattern({ category: 'accounting' }) },
  { label: 'Short Date', pattern: pattern({ category: 'date', datePattern: 'm/d/yyyy' }) },
  {
    label: 'Long Date',
    pattern: pattern({ category: 'date', datePattern: 'dddd, mmmm dd, yyyy' }),
  },
  { label: 'Time', pattern: pattern({ category: 'time', timePattern: 'h:mm:ss AM/PM' }) },
  { label: 'Percentage', pattern: pattern({ category: 'percentage' }) },
  { label: 'Fraction', pattern: pattern({ category: 'fraction' }) },
  { label: 'Scientific', pattern: pattern({ category: 'scientific' }) },
  { label: 'Text', pattern: pattern({ category: 'text' }) },
]

/// The dropdown echo for the current selection. The generic 'Date' category
/// splits into Short/Long by whether the pattern spells out day or month names.
export function categoryOptionForPattern(pattern: string): string {
  const label = numberFormatLabel(pattern)
  if (label !== 'Date') return label
  return /dddd|mmmm/i.test(pattern) ? 'Long Date' : 'Short Date'
}
