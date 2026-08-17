/**
 * Page Layout commands, header/footer, freeze journaling and PDF export.
 * Extracted from App.tsx; the App component passes a PageLayoutContext built
 * fresh per call so refs and state never go stale. Page-setup edits journal
 * per-sheet print settings; nothing renders in the grid (Univer has no
 * page-layout view), everything lands in the saved file.
 */
import { columnLabel } from '../domain/cell-address'
import {
  isSheetRemoved,
  journalSize,
  recordPageSetup,
  recordThemeColors,
  recordThemeFonts,
  type PageSetupJournalState,
} from './edit-journal'
import type { HeaderFooterResult } from './HeaderFooterDialog'
import { t } from './i18n/locale'
import { effectivePageBreaks } from './page-break-preview'
import { COLOR_SCHEMES, FONT_SCHEMES, rethemeStyles, THEME_PRESETS } from './themes'
import { loadVisibleRange } from './univer-sync'
import { buildSheetPrintPayload, type PrintWorksheet } from './print-html'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

const PAPER_NAMES: Record<string, string> = {
  1: 'Letter',
  3: 'Tabloid',
  5: 'Legal',
  7: 'Executive',
  8: 'A3',
  9: 'A4',
  11: 'A5',
}

/** The App refs/state the page-layout actions need; built fresh per call. */
export interface PageLayoutContext {
  univerRef: { readonly current: UniverRuntime | null }
  /// The App's live ref (not a snapshot): loadVisibleRange's staleness
  /// guards compare against `.current` after awaits.
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  /// Re-renders the Page Break Preview overlay when page geometry changed.
  refreshPageBreakPreview?: () => void
}

export function handlePageLayoutCommand(ctx: PageLayoutContext, rest: string): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime) return
  if (!state) {
    ctx.setMessage(t('appPageSetupNeedsFile'))
    return
  }
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  const sheetId = worksheet?.getSheetId()
  if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) return
  const record = (patch: PageSetupJournalState, note: string): void => {
    recordPageSetup(state.editJournal, sheetId, patch)
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appPageSetupRecorded', { note }))
    ctx.refreshPageBreakPreview?.()
  }
  const separator = rest.indexOf(':')
  const key = separator === -1 ? rest : rest.slice(0, separator)
  const value = separator === -1 ? '' : rest.slice(separator + 1)
  const prior = state.editJournal.pageSetup.get(sheetId) ?? {}
  switch (key) {
    case 'orientation':
      if (value !== 'portrait' && value !== 'landscape') return
      record(
        { orientation: value },
        value === 'portrait' ? t('appOrientationPortrait') : t('appOrientationLandscape'),
      )
      return
    case 'margins':
      if (value !== 'normal' && value !== 'wide' && value !== 'narrow') return
      record(
        { margins: value },
        t(
          value === 'normal'
            ? 'appMarginsNormal'
            : value === 'wide'
              ? 'appMarginsWide'
              : 'appMarginsNarrow',
        ),
      )
      return
    case 'paper': {
      const code = Number(value)
      if (!Number.isInteger(code) || code < 1 || code > 118) return
      record({ paperSize: code }, t('appPaperSizeNote', { name: PAPER_NAMES[value] ?? value }))
      return
    }
    case 'scale': {
      const scale = Number(value)
      if (!Number.isInteger(scale) || scale < 10 || scale > 400) return
      record({ scale, fitToPage: false }, t('appPrintScaleNote', { scale }))
      return
    }
    case 'fit-width':
    case 'fit-height': {
      const pages = Number(value)
      if (!Number.isInteger(pages) || pages < 0 || pages > 1_000) return
      // Excel treats the untouched other axis as Automatic (0) once
      // fit-to-page engages.
      const fitToWidth = key === 'fit-width' ? pages : (prior.fitToWidth ?? 0)
      const fitToHeight = key === 'fit-height' ? pages : (prior.fitToHeight ?? 0)
      const fitValue = pages === 0 ? t('appFitAutomatic') : t('appFitPages', { count: pages })
      record(
        { fitToWidth, fitToHeight, fitToPage: fitToWidth > 0 || fitToHeight > 0 },
        key === 'fit-width'
          ? t('appFitWidthNote', { value: fitValue })
          : t('appFitHeightNote', { value: fitValue }),
      )
      return
    }
    case 'print-gridlines':
      record(
        { printGridlines: value === '1' },
        value === '1' ? t('appGridlinesWillPrint') : t('appGridlinesWontPrint'),
      )
      return
    case 'print-headings':
      record(
        { printHeadings: value === '1' },
        value === '1' ? t('appHeadingsWillPrint') : t('appHeadingsWontPrint'),
      )
      return
    case 'print-area': {
      if (value === 'clear') {
        record({ printArea: null }, t('appPrintAreaCleared'))
        return
      }
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range) {
        ctx.setMessage(t('appSelectPrintRange'))
        return
      }
      const startColumn = range.getColumn()
      const startRow = range.getRow()
      const area =
        `${columnLabel(startColumn)}${startRow + 1}` +
        `:${columnLabel(startColumn + range.getWidth() - 1)}${startRow + range.getHeight()}`
      record({ printArea: area }, t('appPrintAreaNote', { area }))
      return
    }
    case 'theme':
    case 'theme-colors':
    case 'theme-fonts': {
      if (!state.file.themeColors || (key === 'theme-fonts' && !state.file.themeFonts)) {
        ctx.setMessage(t('appThemeNeedsThemePart'))
        return
      }
      const colors =
        key === 'theme-fonts' ? undefined : COLOR_SCHEMES.find((entry) => entry.id === value)
      // A theme without a rewritable fontScheme keeps its fonts: journaling
      // one would poison every subsequent save.
      const fonts =
        key === 'theme-colors' || !state.file.themeFonts
          ? undefined
          : key === 'theme'
            ? THEME_PRESETS.find((entry) => entry.id === value)?.fonts
            : FONT_SCHEMES.find((entry) => entry.id === value)
      if (colors === undefined && fonts === undefined) return
      if (colors !== undefined) {
        recordThemeColors(state.editJournal, colors.name, colors.values)
        state.file.themeColors = [...colors.values]
      }
      if (fonts !== undefined) {
        recordThemeFonts(state.editJournal, fonts.name, fonts.major, fonts.minor)
        state.file.themeFonts = { major: fonts.major, minor: fonts.minor }
      }
      // Live re-resolution: styles carrying theme provenance recolor now;
      // charts/CF colors follow on the save's reopen (which reparses the
      // rewritten theme part).
      state.file.styles = rethemeStyles(
        state.file.styles,
        state.file.themeColors,
        state.file.themeFonts ?? null,
      )
      for (const loadedSheetId of [...state.loadedRanges.keys()]) {
        state.loadedRanges.delete(loadedSheetId)
        state.appliedRowKeys.delete(loadedSheetId)
      }
      ctx.setPendingEdits(journalSize(state.editJournal))
      const active = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (active) {
        // Silent repaint: this reload only re-applies the re-themed styles;
        // its async streaming progress would otherwise land after (and wipe)
        // the theme confirmation below.
        void loadVisibleRange(runtime, ctx.lazyWorkbookRef, active, () => undefined)
      }
      ctx.setMessage(t('appThemeApplied', { name: (colors ?? fonts)?.name ?? value }))
      return
    }
    case 'breaks': {
      if (value !== 'insert' && value !== 'remove' && value !== 'reset') return
      const current = effectivePageBreaks(state, sheetId)
      if (current === null) {
        ctx.setMessage(t('appBreaksNeedIndexed'))
        return
      }
      if (value === 'reset') {
        record({ rowBreaks: [], colBreaks: [] }, t('appBreaksReset'))
        return
      }
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range || !worksheet) {
        ctx.setMessage(t('appBreaksNeedCell'))
        return
      }
      const row = range.getRow()
      const column = range.getColumn()
      // Excel: a full-row selection places only the horizontal break, a
      // full-column selection only the vertical one; a cell places both.
      const fullRow = range.getWidth() >= worksheet.getMaxColumns()
      const fullColumn = range.getHeight() >= worksheet.getMaxRows()
      const wantRow = !fullColumn && row > 0
      const wantColumn = !fullRow && column > 0
      if (value === 'insert') {
        if (!wantRow && !wantColumn) {
          ctx.setMessage(t('appBreaksNeedCell'))
          return
        }
        const rowBreaks = wantRow
          ? [...new Set([...current.rowBreaks, row])].sort((a, b) => a - b)
          : current.rowBreaks
        const colBreaks = wantColumn
          ? [...new Set([...current.colBreaks, column])].sort((a, b) => a - b)
          : current.colBreaks
        record({ rowBreaks, colBreaks }, t('appBreakInserted'))
        return
      }
      const rowBreaks = current.rowBreaks.filter((id) => !(wantRow && id === row))
      const colBreaks = current.colBreaks.filter((id) => !(wantColumn && id === column))
      if (
        rowBreaks.length === current.rowBreaks.length &&
        colBreaks.length === current.colBreaks.length
      ) {
        ctx.setMessage(t('appBreakNoneHere'))
        return
      }
      record({ rowBreaks, colBreaks }, t('appBreakRemoved'))
      return
    }
    case 'print-titles': {
      if (value === 'clear') {
        record({ printTitles: null }, t('appPrintTitlesCleared'))
        return
      }
      if (value === 'first-row') {
        record({ printTitles: '1:1' }, t('appRow1Repeats'))
        return
      }
      const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!range) {
        ctx.setMessage(t('appSelectRepeatRows'))
        return
      }
      const rows = `${range.getRow() + 1}:${range.getRow() + range.getHeight()}`
      record({ printTitles: rows }, t('appRowsRepeat', { rows }))
      return
    }
    default:
      return
  }
}

/// Insert → Header & Footer OK: journals the printed header/footer of the
/// active sheet; the save writes the worksheet's <headerFooter> element.
export function handleApplyHeaderFooter(
  ctx: PageLayoutContext,
  result: HeaderFooterResult,
): string | null {
  const state = ctx.lazyWorkbookRef.current
  if (!state) return t('appHfNeedsFile')
  const sheetId = ctx.univerRef.current?.univerAPI
    .getActiveWorkbook()
    ?.getActiveSheet()
    ?.getSheetId()
  if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) {
    return t('appActiveSheetUnavailable')
  }
  recordPageSetup(state.editJournal, sheetId, { header: result.header, footer: result.footer })
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appHfUpdated'))
  return null
}

/// Lays the active sheet out as HTML with its Page Layout settings and asks
/// the main process to render the PDF (hidden window + save dialog).
export async function handleExportPdf(ctx: PageLayoutContext): Promise<void> {
  const runtime = ctx.univerRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!runtime || !worksheet) return
  const state = ctx.lazyWorkbookRef.current
  if (state && !state.flags.preloadComplete) {
    ctx.setMessage(t('appPdfNeedsFullLoad'))
    return
  }
  try {
    const pageSetup = state?.editJournal.pageSetup.get(worksheet.getSheetId()) ?? {}
    const baseName = (state?.file.name ?? 'Book1').replace(/\.[^.]+$/, '')
    const payload = buildSheetPrintPayload(
      worksheet as unknown as PrintWorksheet,
      pageSetup,
      `${baseName}.pdf`,
      worksheet.getSheetName(),
    )
    ctx.setMessage(t('appPdfRendering'))
    const result = await window.desktopApi.exportPdf(payload)
    ctx.setMessage(
      result.canceled ? t('appPdfCanceled') : t('appPdfExported', { path: result.path }),
    )
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appPdfExportFailed'))
  }
}
