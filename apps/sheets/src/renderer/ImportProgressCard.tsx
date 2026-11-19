/**
 * The progress bar and the cancel button a delimited import (a .csv/.tsv
 * opened as a workbook) shows while it runs.
 *
 * The card is its own component so the progress can be checked by rendering
 * it: the import is the one place where the user waits on something long
 * enough to want a way out, and a bar that never moves is worse than no bar.
 */
import type { ImportProgress } from './import-progress'
import { importPercent } from './import-progress'

export interface ImportProgressCardProps {
  /** how far the import is, or null when nothing is importing */
  readonly progress: ImportProgress | null
  /** the message above the bar (the app's own "opening workbook" line) */
  readonly label: string
  /** the label on the cancel button */
  readonly cancelLabel: string
  /** the user asked to stop */
  readonly onCancel: () => void
}

export function ImportProgressCard({
  progress,
  label,
  cancelLabel,
  onCancel,
}: ImportProgressCardProps): React.JSX.Element {
  const percent = progress ? Math.round(importPercent(progress) * 100) : 0
  return (
    <div className="workbook-opening-card">
      <div>{label}</div>
      {progress && (
        <>
          <div
            className="workbook-opening-bar"
            role="progressbar"
            aria-valuemin={0}
            aria-valuemax={100}
            aria-valuenow={percent}
          >
            <div className="workbook-opening-bar-fill" style={{ width: `${percent}%` }} />
          </div>
          <div className="workbook-opening-percent">{`${percent}%`}</div>
          <button type="button" className="workbook-opening-cancel" onClick={onCancel}>
            {cancelLabel}
          </button>
        </>
      )}
    </div>
  )
}
