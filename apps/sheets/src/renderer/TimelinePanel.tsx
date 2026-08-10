import { monthKeyParts, type MonthKey, type TimelineMember } from '../domain/pivot-timeline'
import { useI18n } from './i18n/locale'

/// Timeline: a month-granularity date-range filter bound to one pivot date
/// field. Like slicers it drives the pivot's hidden entries (applyPivotSlicer)
/// and exists only within the session; the filtered output area persists
/// through the pivot refresh write-back path.

export interface TimelineUiState {
  readonly id: string
  readonly sheetId: string
  /// Bound pivot part path (fileMeta.pivotTables[].path).
  readonly pivotPath: string
  /// Cache-field index and display name of the filtered date field.
  readonly field: number
  readonly fieldName: string
  readonly members: readonly TimelineMember[]
  readonly minMonth: MonthKey
  readonly maxMonth: MonthKey
  /// Selected inclusive month range; null = unfiltered.
  readonly range: { readonly start: MonthKey; readonly end: MonthKey } | null
}

function monthLabel(key: MonthKey): string {
  const { year, month } = monthKeyParts(key)
  return `${year}-${String(month).padStart(2, '0')}`
}

/// Field picker when creating a timeline: lists the pivot's date fields.
export function TimelineFieldPicker({
  fields,
  onPick,
  onClose,
}: {
  readonly fields: readonly { readonly field: number; readonly name: string }[]
  /// Returns an error message; null = created successfully (the caller closes
  /// the dialog).
  readonly onPick: (field: number) => string | null
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog slicer-picker"
        role="dialog"
        aria-label={t('dlgTimelineInsertTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgTimelineInsertTitle')}</header>
        <p className="dialog-note">{t('dlgTimelinePickNote')}</p>
        <div className="slicer-picker-fields">
          {fields.map((entry) => (
            <button
              key={entry.field}
              className="slicer-picker-field"
              onClick={() => {
                if (onPick(entry.field) === null) onClose()
              }}
            >
              {entry.name}
            </button>
          ))}
        </div>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
        </div>
      </div>
    </div>
  )
}

/// Timeline overlay on the worksheet: one panel per timeline with a month
/// strip. Click selects one month; Shift-click extends the range.
export function TimelinePanels({
  timelines,
  onRange,
  onClear,
  onRemove,
}: {
  readonly timelines: readonly TimelineUiState[]
  readonly onRange: (timelineId: string, start: MonthKey, end: MonthKey) => void
  readonly onClear: (timelineId: string) => void
  readonly onRemove: (timelineId: string) => void
}): React.JSX.Element | null {
  const { t } = useI18n()
  if (timelines.length === 0) return null
  return (
    <div className="slicer-stack">
      {timelines.map((timeline) => {
        const months: MonthKey[] = []
        for (let key = timeline.minMonth; key <= timeline.maxMonth; key++) months.push(key)
        const range = timeline.range
        const label =
          range === null
            ? `${monthLabel(timeline.minMonth)} – ${monthLabel(timeline.maxMonth)}`
            : range.start === range.end
              ? monthLabel(range.start)
              : `${monthLabel(range.start)} – ${monthLabel(range.end)}`
        return (
          <section
            key={timeline.id}
            className="slicer-panel timeline-panel"
            aria-label={t('dlgTimelineAria', { name: timeline.fieldName })}
          >
            <header>
              <span className="slicer-title" data-tip={timeline.pivotPath}>
                {timeline.fieldName}
              </span>
              <button
                className="slicer-clear"
                data-tip={t('dlgTimelineClear')}
                aria-label={t('dlgTimelineClear')}
                disabled={range === null}
                onClick={() => onClear(timeline.id)}
              >
                ⧉
              </button>
              <button
                className="slicer-close"
                data-tip={t('dlgTimelineRemove')}
                aria-label={t('dlgTimelineRemove')}
                onClick={() => onRemove(timeline.id)}
              >
                ×
              </button>
            </header>
            <div className="timeline-range">{label}</div>
            <div className="timeline-months">
              {months.map((key) => {
                const selected = range !== null && key >= range.start && key <= range.end
                const { month } = monthKeyParts(key)
                return (
                  <button
                    key={key}
                    className={selected ? 'timeline-month selected' : 'timeline-month'}
                    data-tip={monthLabel(key)}
                    onClick={(event) => {
                      if (event.shiftKey && range !== null) {
                        onRange(timeline.id, Math.min(range.start, key), Math.max(range.end, key))
                      } else {
                        onRange(timeline.id, key, key)
                      }
                    }}
                  >
                    {month === 1 ? monthLabel(key) : String(month)}
                  </button>
                )
              })}
            </div>
          </section>
        )
      })}
    </div>
  )
}
