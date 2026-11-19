/**
 * How far along a delimited import (a .csv/.tsv opened as a workbook) is, as
 * one fraction for the bar.
 *
 * The import has two phases: the text is read (a pass over the file that
 * measures the grid), then the workbook is written (a second pass plus the
 * compression of the archive — where the time goes: 0.7 s against 12 s for a
 * 39 MB file). The read phase is a small slice of the work, so it takes a
 * small slice of the bar; without the split the bar would sit at zero through
 * the whole read and then jump.
 */
import { useEffect, useState } from 'react'

export interface ImportProgress {
  readonly phase: 'read' | 'write'
  readonly fraction: number
}

/** the share of the bar the read phase takes */
export const IMPORT_READ_SHARE = 0.05

export function importPercent(progress: ImportProgress): number {
  const fraction = Math.max(0, Math.min(progress.fraction, 1))
  const percent =
    progress.phase === 'read'
      ? fraction * IMPORT_READ_SHARE
      : IMPORT_READ_SHARE + fraction * (1 - IMPORT_READ_SHARE)
  return Math.max(0, Math.min(percent, 1))
}

/**
 * The progress of the import running in this tab, or null when none is.
 *
 * The main process owns the lifetime: it reports the position while the
 * import runs and null when it is over. The bar therefore belongs to one
 * import — a bar left standing would show the next workbook opening at the
 * previous file's percentage, with a cancel button that has nothing to
 * cancel.
 */
export function useImportProgress(): ImportProgress | null {
  const [progress, setProgress] = useState<ImportProgress | null>(null)
  useEffect(() => {
    const stop = window.desktopApi?.onImportProgress?.((next) => {
      setProgress(next)
    })
    return () => stop?.()
  }, [])
  return progress
}
