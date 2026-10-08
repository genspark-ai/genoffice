/// Excel's Print with Preview: settings on the left, the paginated pages on
/// the right. The preview is the very HTML the printer and the PDF export
/// receive (one paper-sized box per page), shown one page at a time.
import { useEffect, useMemo, useRef, useState } from 'react'

import type { PrinterInfo } from '../shared/desktop-api'
import type { PageSetupJournalState } from './edit-journal'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
import { type PrintJob, type PrintPreviewHost, type PrintRange } from './page-layout-actions'
import { assemblePrintHtml, type PaperGeometry } from './print-html'

type MarginsPreset = 'normal' | 'wide' | 'narrow' | 'custom'
type Scaling = 'none' | 'sheet' | 'columns' | 'rows' | 'fit-file' | 'custom'

interface Settings {
  readonly orientation: 'portrait' | 'landscape'
  readonly paperSize: number
  readonly margins: MarginsPreset
  readonly scaling: Scaling
  readonly scale: number
  readonly printGridlines: boolean
  readonly printHeadings: boolean
}

const PAPERS: readonly { readonly code: number; readonly label: string }[] = [
  { code: 1, label: 'Letter' },
  { code: 5, label: 'Legal' },
  { code: 3, label: 'Tabloid' },
  { code: 7, label: 'Executive' },
  { code: 8, label: 'A3' },
  { code: 9, label: 'A4' },
  { code: 11, label: 'A5' },
  { code: 12, label: 'B4' },
  { code: 13, label: 'B5' },
]

const PREVIEW_PADDING_PX = 24
/// Option values that are not printer names (Electron: '' = system default).
const DEFAULT_PRINTER = ''
const SYSTEM_DIALOG = '\u0000system-dialog'

function initialSettings(host: PrintPreviewHost): Settings {
  const { setup } = host
  const scaling: Scaling = !setup.fitToPage
    ? setup.scale === 100
      ? 'none'
      : 'custom'
    : setup.fitToWidth === 1 && setup.fitToHeight === 1
      ? 'sheet'
      : setup.fitToWidth === 1 && setup.fitToHeight === 0
        ? 'columns'
        : setup.fitToWidth === 0 && setup.fitToHeight === 1
          ? 'rows'
          : 'fit-file'
  return {
    orientation: setup.orientation,
    paperSize: setup.paperSize,
    margins: host.marginsPreset,
    scaling,
    scale: setup.fitToPage ? 100 : setup.scale,
    printGridlines: setup.printGridlines,
    printHeadings: setup.printHeadings,
  }
}

function fitFor(scaling: Scaling): { fitToWidth: number; fitToHeight: number } | null {
  switch (scaling) {
    case 'sheet':
      return { fitToWidth: 1, fitToHeight: 1 }
    case 'columns':
      return { fitToWidth: 1, fitToHeight: 0 }
    case 'rows':
      return { fitToWidth: 0, fitToHeight: 1 }
    default:
      return null
  }
}

/// The sheet's own fit-to-page page counts stay untouched under 'fit-file'.
function overridesFor(settings: Settings): PageSetupJournalState {
  const fit = fitFor(settings.scaling)
  return {
    orientation: settings.orientation,
    paperSize: settings.paperSize,
    ...(settings.margins === 'custom' ? {} : { margins: settings.margins }),
    ...(fit
      ? { ...fit, fitToPage: true }
      : settings.scaling === 'fit-file'
        ? {}
        : { scale: settings.scaling === 'none' ? 100 : settings.scale, fitToPage: false }),
    printGridlines: settings.printGridlines,
    printHeadings: settings.printHeadings,
  }
}

export function PrintDialog({
  host,
  onClose,
  onPageSetup,
  revision,
}: {
  readonly host: PrintPreviewHost
  readonly onClose: () => void
  readonly onPageSetup: () => void
  /// Bumped when Page Setup (Header & Footer) closes so the preview re-lays out.
  readonly revision: number
}): React.JSX.Element {
  const { t } = useI18n()
  const [settings, setSettings] = useState<Settings>(() => initialSettings(host))
  const [range, setRange] = useState<PrintRange>('active')
  const [printers, setPrinters] = useState<PrinterInfo[]>([])
  const [deviceName, setDeviceName] = useState(DEFAULT_PRINTER)
  const [copies, setCopies] = useState(1)
  const [collate, setCollate] = useState(true)
  const [pagesMode, setPagesMode] = useState<'all' | 'range'>('all')
  const [from, setFrom] = useState(1)
  const [to, setTo] = useState(1)
  const [job, setJob] = useState<PrintJob | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [page, setPage] = useState(1)
  const [fit, setFit] = useState(true)
  const [paneSize, setPaneSize] = useState({ width: 0, height: 0 })
  const paneRef = useRef<HTMLDivElement | null>(null)
  const frameRef = useRef<HTMLIFrameElement | null>(null)
  const modal = useModalDialog(onClose)

  useEffect(() => {
    let alive = true
    host
      .listPrinters()
      .then((list) => {
        if (alive) setPrinters(list)
      })
      .catch(() => undefined)
    return () => {
      alive = false
    }
  }, [host])

  // The job rebuilds a beat after the last change so a slider drag does not
  // lay the sheet out on every tick.
  useEffect(() => {
    let alive = true
    const timer = window.setTimeout(() => {
      host
        .buildJob(range, overridesFor(settings))
        .then((built) => {
          if (!alive) return
          setJob(built)
          setError(null)
          setPage((current) => Math.min(Math.max(current, 1), built.pageCount))
          setTo((current) => (current > built.pageCount ? built.pageCount : current))
        })
        .catch((reason: unknown) => {
          if (!alive) return
          setJob(null)
          setError(reason instanceof Error ? reason.message : String(reason))
        })
    }, 120)
    return () => {
      alive = false
      window.clearTimeout(timer)
    }
  }, [host, range, settings, revision])

  useEffect(() => {
    const pane = paneRef.current
    if (!pane) return
    const measure = () => setPaneSize({ width: pane.clientWidth, height: pane.clientHeight })
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(pane)
    return () => observer.disconnect()
  }, [])

  const assembled = useMemo(() => (job ? assemblePrintHtml(job.documents) : null), [job])
  const pageCount = job?.pageCount ?? 0
  const current = Math.min(Math.max(page, 1), Math.max(pageCount, 1))
  const paper: PaperGeometry | null = (() => {
    if (!job || !assembled) return null
    const entry = assembled.pages[current - 1]
    return entry ? (job.documents[entry.document]?.paper ?? null) : null
  })()

  // Only the current page shows; the others stay laid out so page numbers
  // and the printed HTML never diverge from what is on screen.
  useEffect(() => {
    const frame = frameRef.current
    const doc = frame?.contentDocument
    if (!frame || !doc) return
    const apply = () => {
      const target = frame.contentDocument
      if (!target) return
      let style = target.getElementById('pv-page')
      if (!style) {
        style = target.createElement('style')
        style.id = 'pv-page'
        target.head?.appendChild(style)
      }
      style.textContent = `.page{display:none}.page:nth-child(${current}){display:block}`
    }
    apply()
    frame.addEventListener('load', apply)
    return () => frame.removeEventListener('load', apply)
  }, [assembled, current])

  const pageWidthPx = (paper?.widthIn ?? 8.27) * 96
  const pageHeightPx = (paper?.heightIn ?? 11.69) * 96
  const zoom = fit
    ? Math.max(
        0.05,
        Math.min(
          (paneSize.width - PREVIEW_PADDING_PX) / pageWidthPx,
          (paneSize.height - PREVIEW_PADDING_PX) / pageHeightPx,
        ),
      )
    : 1

  const update = (patch: Partial<Settings>, command: string | null): void => {
    setSettings((prev) => ({ ...prev, ...patch }))
    if (command !== null) host.applySetting(command)
  }
  const setScaling = (scaling: Scaling, scale = settings.scale): void => {
    const fitPages = fitFor(scaling)
    update(
      { scaling, scale },
      fitPages
        ? `fit:${fitPages.fitToWidth},${fitPages.fitToHeight}`
        : scaling === 'fit-file'
          ? `fit:${host.setup.fitToWidth},${host.setup.fitToHeight}`
          : `scale:${scaling === 'none' ? 100 : scale}`,
    )
  }

  const pages = pagesMode === 'all' ? null : { from, to }
  const rangeInvalid =
    pagesMode === 'range' && (from < 1 || to < from || (pageCount > 0 && to > pageCount))
  const canOutput = job !== null && !busy && !rangeInvalid

  const run = async (action: (built: PrintJob) => Promise<string>): Promise<void> => {
    if (!job) return
    setBusy(true)
    try {
      const message = await action(job)
      host.setMessage(message)
      onClose()
    } catch (reason: unknown) {
      setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog print-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgPrintTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgPrintTitle')}</header>
        <div className="print-dialog-body">
          <div className="print-settings dialog-grid">
            <label className="dialog-span">
              {t('dlgPrintPrinter')}
              <select value={deviceName} onChange={(event) => setDeviceName(event.target.value)}>
                <option value={DEFAULT_PRINTER}>{t('dlgPrintDefaultPrinter')}</option>
                {printers.map((printer) => (
                  <option key={printer.name} value={printer.name}>
                    {printer.displayName}
                  </option>
                ))}
                <option value={SYSTEM_DIALOG}>{t('dlgPrintSystemDialog')}</option>
              </select>
            </label>
            <label>
              {t('dlgPrintCopies')}
              <input
                type="number"
                min={1}
                max={999}
                value={copies}
                onChange={(event) =>
                  setCopies(Math.min(999, Math.max(1, Number(event.target.value) || 1)))
                }
              />
            </label>
            <label className="dialog-check">
              <input
                type="checkbox"
                checked={collate}
                onChange={(event) => setCollate(event.target.checked)}
              />
              {t('dlgPrintCollate')}
            </label>
            <label className="dialog-span">
              {t('dlgPrintRange')}
              <select
                value={range}
                onChange={(event) => setRange(event.target.value as PrintRange)}
              >
                <option value="active">{t('dlgPrintActiveSheets')}</option>
                <option value="workbook">{t('dlgPrintEntireWorkbook')}</option>
                <option value="selection" disabled={!host.hasSelection}>
                  {t('dlgPrintSelection')}
                </option>
              </select>
            </label>
            <label className="dialog-span">
              {t('dlgPrintPages')}
              <span className="print-pages-row">
                <select
                  value={pagesMode}
                  onChange={(event) => setPagesMode(event.target.value as 'all' | 'range')}
                >
                  <option value="all">{t('dlgPrintAllPages')}</option>
                  <option value="range">{t('dlgPrintPageRange')}</option>
                </select>
                <input
                  type="number"
                  min={1}
                  max={Math.max(pageCount, 1)}
                  value={from}
                  disabled={pagesMode === 'all'}
                  aria-label={t('dlgPrintPagesFrom')}
                  onChange={(event) => setFrom(Number(event.target.value) || 1)}
                />
                <input
                  type="number"
                  min={1}
                  max={Math.max(pageCount, 1)}
                  value={to}
                  disabled={pagesMode === 'all'}
                  aria-label={t('dlgPrintPagesTo')}
                  onChange={(event) => setTo(Number(event.target.value) || 1)}
                />
              </span>
            </label>
            <label>
              {t('dlgPrintOrientation')}
              <select
                value={settings.orientation}
                onChange={(event) => {
                  const orientation = event.target.value as Settings['orientation']
                  update({ orientation }, `orientation:${orientation}`)
                }}
              >
                <option value="portrait">{t('dlgPrintPortrait')}</option>
                <option value="landscape">{t('dlgPrintLandscape')}</option>
              </select>
            </label>
            <label>
              {t('dlgPrintPaper')}
              <select
                value={settings.paperSize}
                onChange={(event) => {
                  const paperSize = Number(event.target.value)
                  update({ paperSize }, `paper:${paperSize}`)
                }}
              >
                {PAPERS.map((entry) => (
                  <option key={entry.code} value={entry.code}>
                    {entry.label}
                  </option>
                ))}
                {PAPERS.some((entry) => entry.code === settings.paperSize) ? null : (
                  <option value={settings.paperSize}>{t('dlgPrintPaperFromFile')}</option>
                )}
              </select>
            </label>
            <label>
              {t('dlgPrintMargins')}
              <select
                value={settings.margins}
                onChange={(event) => {
                  const margins = event.target.value as MarginsPreset
                  update({ margins }, margins === 'custom' ? null : `margins:${margins}`)
                }}
              >
                <option value="normal">{t('dlgPrintMarginsNormal')}</option>
                <option value="wide">{t('dlgPrintMarginsWide')}</option>
                <option value="narrow">{t('dlgPrintMarginsNarrow')}</option>
                {host.marginsPreset === 'custom' && (
                  <option value="custom">{t('dlgPrintMarginsCustom')}</option>
                )}
              </select>
            </label>
            <label>
              {t('dlgPrintScaling')}
              <select
                value={settings.scaling}
                onChange={(event) => setScaling(event.target.value as Scaling)}
              >
                <option value="none">{t('dlgPrintNoScaling')}</option>
                <option value="sheet">{t('dlgPrintFitSheet')}</option>
                <option value="columns">{t('dlgPrintFitColumns')}</option>
                <option value="rows">{t('dlgPrintFitRows')}</option>
                {host.setup.fitToPage && (
                  <option value="fit-file">
                    {t('dlgPrintFitFile', {
                      width: host.setup.fitToWidth || '*',
                      height: host.setup.fitToHeight || '*',
                    })}
                  </option>
                )}
                <option value="custom">{t('dlgPrintCustomScale')}</option>
              </select>
            </label>
            {settings.scaling === 'custom' && (
              <label className="dialog-span">
                {t('dlgPrintCustomScalePercent')}
                <input
                  type="number"
                  min={10}
                  max={400}
                  value={settings.scale}
                  onChange={(event) => {
                    const scale = Math.min(400, Math.max(10, Number(event.target.value) || 100))
                    setScaling('custom', scale)
                  }}
                />
              </label>
            )}
            <label className="dialog-check">
              <input
                type="checkbox"
                checked={settings.printGridlines}
                onChange={(event) =>
                  update(
                    { printGridlines: event.target.checked },
                    `print-gridlines:${event.target.checked ? '1' : '0'}`,
                  )
                }
              />
              {t('dlgPrintGridlines')}
            </label>
            <label className="dialog-check">
              <input
                type="checkbox"
                checked={settings.printHeadings}
                onChange={(event) =>
                  update(
                    { printHeadings: event.target.checked },
                    `print-headings:${event.target.checked ? '1' : '0'}`,
                  )
                }
              />
              {t('dlgPrintHeadings')}
            </label>
          </div>
          <div className="print-preview">
            <div className="print-preview-bar">
              <button
                className="secondary"
                disabled={current <= 1}
                aria-label={t('dlgPrintPrevPage')}
                onClick={() => setPage(current - 1)}
              >
                ‹
              </button>
              <span>
                {pageCount > 0
                  ? t('dlgPrintPageOf', { page: current, total: pageCount })
                  : error
                    ? t('dlgPrintNoPages')
                    : t('dlgPrintRendering')}
              </span>
              <button
                className="secondary"
                disabled={current >= pageCount}
                aria-label={t('dlgPrintNextPage')}
                onClick={() => setPage(current + 1)}
              >
                ›
              </button>
              <span className="print-preview-spacer" />
              <button
                className="secondary"
                aria-pressed={fit}
                onClick={() => setFit((value) => !value)}
              >
                {fit ? t('dlgPrintZoomActual') : t('dlgPrintZoomFit')}
              </button>
            </div>
            <div className="print-preview-pane" ref={paneRef}>
              {assembled && (
                <div
                  className="print-preview-page"
                  style={{
                    width: pageWidthPx * zoom,
                    height: pageHeightPx * zoom,
                  }}
                >
                  <iframe
                    ref={frameRef}
                    title={t('dlgPrintTitle')}
                    sandbox="allow-same-origin"
                    srcDoc={assembled.html}
                    style={{
                      width: pageWidthPx,
                      height: pageHeightPx,
                      transform: `scale(${zoom})`,
                    }}
                  />
                </div>
              )}
            </div>
          </div>
        </div>
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={onPageSetup}>
            {t('dlgPrintPageSetup')}
          </button>
          <span className="print-preview-spacer" />
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="secondary"
            disabled={!canOutput}
            onClick={() => void run((built) => host.exportPdf(built, pages))}
          >
            {t('dlgPrintToPdf')}
          </button>
          <button
            className="primary-action"
            disabled={!canOutput}
            onClick={() =>
              void run((built) =>
                host.print(built, {
                  deviceName: deviceName === SYSTEM_DIALOG ? null : deviceName,
                  copies,
                  collate,
                  pages,
                }),
              )
            }
          >
            {t('dlgPrintButton')}
          </button>
        </div>
      </div>
    </div>
  )
}
