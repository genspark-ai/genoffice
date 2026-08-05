/**
 * CSV export: serializes the active sheet's used range as comma-separated
 * text (Excel CSV conventions: TRUE/FALSE for booleans, quoting for fields
 * carrying commas/quotes/newlines) and asks the main process to run the save
 * dialog and write the file. Mirrors handleExportPdf's flow.
 */
import { t } from './i18n/locale'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

export interface CsvExportContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
}

/// One CSV field: booleans as TRUE/FALSE, embedded newlines normalized to
/// \n, CSV-style quoting when the text carries commas, quotes or newlines.
function csvField(text: string): string {
  const normalized = text.replace(/\r\n|\r+/g, '\n').replace(/\n+$/, '')
  return /[",\n]/.test(normalized) ? `"${normalized.replace(/"/g, '""')}"` : normalized
}

/// Exports the active sheet's used range; the main process shows the save
/// dialog and writes the file. Resolves to the written path ('' if canceled).
export async function handleExportCsv(ctx: CsvExportContext): Promise<void> {
  const runtime = ctx.univerRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!runtime || !worksheet) return
  const state = ctx.lazyWorkbookRef.current
  if (state && !state.flags.preloadComplete) {
    ctx.setMessage(t('appPdfNeedsFullLoad'))
    return
  }
  try {
    const baseName = (state?.file.name ?? 'Book1').replace(/\.[^.]+$/, '')
    const lastRow = worksheet.getLastRow()
    const lastColumn = worksheet.getLastColumn()
    if (lastRow < 0 || lastColumn < 0) {
      ctx.setMessage(t('appExportCsvEmpty'))
      return
    }
    const grid = worksheet.getRange(0, 0, lastRow + 1, lastColumn + 1)
    const display = grid.getDisplayValues()
    const lines = display.map((row) => row.map((value) => csvField(String(value ?? ''))).join(','))
    const csv = `${lines.join('\n')}\n`
    ctx.setMessage(t('appExportCsvExporting'))
    const result = await window.desktopApi.exportCsv({ fileName: `${baseName}.csv`, csv })
    ctx.setMessage(
      result.canceled ? t('appExportCsvCanceled') : t('appExportCsvExported', { path: result.path }),
    )
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appExportCsvFailed'))
  }
}
