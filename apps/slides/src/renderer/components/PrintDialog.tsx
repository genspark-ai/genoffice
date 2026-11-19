/**
 * Print dialog modeled on PowerPoint's print sheet: live page preview on the left
 * (the exact pages that will print, built by the shared print-html assembly),
 * range / layout / orientation / framing options on the right. "Print" hands the
 * assembled pages to the system print dialog.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { useEscOverlay } from '../esc-overlay'
import type { RenderSlide } from '@genoffice/pptx-render'
import { useI18n } from '../i18n/locale'
import {
  buildPrintDocumentHtml,
  currentRangeIndices,
  parsePrintRange,
  printPageCount,
  type PrintLayout,
  type PrintOrder,
  type PrintOrientation,
} from '../../shared/print-html'
import { renderSlidesToPngBase64 } from '../export-render'
import type { PrintDuplex, PrinterInfo, PrinterOption } from '../../shared/ipc'

type RangeMode = 'all' | 'current' | 'custom'

const IS_MAC = navigator.platform.toLowerCase().includes('mac')
/** the escape hatch: hand the job to the OS print dialog instead of printing here */
const PRINT_SYSTEM_DIALOG = '\u0000system-dialog'
const PRINT_DUPLEX_MODES: PrintDuplex[] = ['simplex', 'longEdge', 'shortEdge']

export function PrintDialog({
  slides,
  images,
  current,
  onClose,
  setStatus,
}: {
  /** The whole deck, hidden pages included (the hidden toggle filters at print time) */
  slides: RenderSlide[]
  images: Map<string, HTMLImageElement>
  /** Current page index, for the "current slide" range */
  current: number
  onClose: () => void
  setStatus: (s: string) => void
}) {
  useEscOverlay(true)
  const { t } = useI18n()
  const [pngs, setPngs] = useState<string[] | null>(null)
  const [notes, setNotes] = useState<string[]>([])
  const [layout, setLayout] = useState<PrintLayout>('full')
  const [rangeMode, setRangeMode] = useState<RangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [orientation, setOrientation] = useState<PrintOrientation>('portrait')
  const [order, setOrder] = useState<PrintOrder>('horizontal')
  const [frame, setFrame] = useState(false)
  const [includeHidden, setIncludeHidden] = useState(false)
  const [printing, setPrinting] = useState(false)
  // the printer section: '' = the OS default, a name = that queue, SYSTEM = the OS dialog
  const [printers, setPrinters] = useState<PrinterInfo[]>([])
  const [printDevice, setPrintDevice] = useState(IS_MAC ? PRINT_SYSTEM_DIALOG : '')
  const [printCopies, setPrintCopies] = useState(1)
  const [printCollate, setPrintCollate] = useState(true)
  const [printDuplex, setPrintDuplex] = useState<PrintDuplex>('simplex')
  const [driverOptions, setDriverOptions] = useState<PrinterOption[]>([])
  const [driverChoices, setDriverChoices] = useState<Record<string, string>>({})
  const blobUrlsRef = useRef<string[]>([])
  // Only the newest printer-options request may fill the dropdowns: a slow
  // answer for an earlier queue must not land after a later selection's.
  const driverOptionsRequest = useRef(0)
  const frameRef = useRef<HTMLIFrameElement | null>(null)
  const paneRef = useRef<HTMLDivElement | null>(null)
  const dialogRef = useRef<HTMLDivElement>(null)
  const previouslyFocused = useRef<HTMLElement | null>(null)

  // Focus the first field on mount; return focus to the opener on unmount
  useEffect(() => {
    previouslyFocused.current = document.activeElement as HTMLElement | null
    const root = dialogRef.current
    if (root && !root.contains(document.activeElement)) {
      root.querySelector<HTMLElement>('input, textarea, select, button')?.focus()
    }
    return () => {
      previouslyFocused.current?.focus?.()
    }
  }, [])

  // One offscreen render of the whole deck on open (print quality, 2x); the same
  // bitmaps feed the preview (as blob URLs) and the print job (as base64).
  useEffect(() => {
    let alive = true
    void (async () => {
      try {
        const rendered = await renderSlidesToPngBase64(slides, images)
        const noteTexts = await Promise.all(slides.map((_s, i) => window.slidesApi.getNotes(i)))
        if (!alive) return
        blobUrlsRef.current = rendered.map((b64) =>
          URL.createObjectURL(
            new Blob([Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))], { type: 'image/png' }),
          ),
        )
        setNotes(noteTexts)
        setPngs(rendered)
      } catch (err) {
        if (alive) setStatus(t('appPrintFailed', { error: String(err) }))
      }
    })()
    return () => {
      alive = false
      for (const url of blobUrlsRef.current) URL.revokeObjectURL(url)
      blobUrlsRef.current = []
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onClose()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose])

  useEffect(() => {
    let alive = true
    window.slidesApi
      .listPrinters()
      .then((list) => {
        if (alive) setPrinters(list)
      })
      .catch(() => undefined)
    return () => {
      alive = false
    }
  }, [])

  /** 0-based deck indices selected by the range options (before the hidden filter) */
  const rangeIndices = useMemo<number[] | null>(() => {
    if (rangeMode === 'all') return slides.map((_s, i) => i)
    if (rangeMode === 'current') return currentRangeIndices(current, slides.length)
    return parsePrintRange(customRange, slides.length)
  }, [rangeMode, customRange, slides, current])

  const selected = useMemo(
    () => (rangeIndices ?? []).filter((i) => includeHidden || !slides[i].hidden),
    [rangeIndices, includeHidden, slides],
  )
  const rangeInvalid = rangeMode === 'custom' && rangeIndices === null
  const pageCount = printPageCount(selected.length, layout)

  const previewHtml = useMemo(() => {
    if (!pngs || selected.length === 0) return null
    const first = slides[selected[0]]
    return buildPrintDocumentHtml({
      srcs: selected.map((i) => blobUrlsRef.current[i]),
      ratio: first.widthPx / first.heightPx,
      layout,
      notes: selected.map((i) => notes[i] ?? ''),
      orientation,
      order,
      frame,
      preview: true,
    })
  }, [pngs, selected, slides, layout, notes, orientation, order, frame])

  /** Scale the preview document so one page fits the pane width (iframe body zoom) */
  const applyZoom = useCallback(() => {
    const doc = frameRef.current?.contentDocument
    const pane = paneRef.current
    const page = doc?.querySelector('.page') as HTMLElement | null
    if (!doc || !pane || !page) return
    doc.body.style.zoom = ''
    const pageW = page.getBoundingClientRect().width
    if (pageW > 0) doc.body.style.zoom = String(Math.min((pane.clientWidth - 44) / pageW, 1))
  }, [])

  useEffect(() => {
    const pane = paneRef.current
    if (!pane) return
    const ro = new ResizeObserver(applyZoom)
    ro.observe(pane)
    return () => ro.disconnect()
  }, [applyZoom])

  const doPrint = async () => {
    if (!pngs || selected.length === 0 || printing) return
    setPrinting(true)
    try {
      const first = slides[selected[0]]
      const job =
        printDevice === PRINT_SYSTEM_DIALOG
          ? undefined
          : {
              deviceName: printDevice,
              copies: printCopies,
              collate: printCollate,
              duplex: printDuplex,
              driverOptions: driverOptions
                .filter(
                  (g) =>
                    driverChoices[g.name] !== undefined && driverChoices[g.name] !== g.defaultValue,
                )
                .map((g) => ({ name: g.name, value: driverChoices[g.name]! })),
            }
      const r = await window.slidesApi.printSlides({
        pngsBase64: selected.map((i) => pngs[i]),
        widthPx: first.widthPx,
        heightPx: first.heightPx,
        layout,
        ...(layout === 'notes' ? { notes: selected.map((i) => notes[i] ?? '') } : {}),
        ...(layout !== 'full' ? { orientation } : {}),
        ...(layout !== 'full' && layout !== 'notes' ? { order } : {}),
        ...(frame && layout === 'full' ? { frame: true } : {}),
        ...(job ? { job } : {}),
      })
      if (r.ok) onClose()
      else if (r.error) setStatus(t('appPrintFailed', { error: r.error }))
      // ok=false without an error = canceled in the system dialog: keep the dialog open
    } finally {
      setPrinting(false)
    }
  }

  const layouts: Array<[PrintLayout, string]> = [
    ['full', t('appPrintLayoutFull')],
    ['handout2', t('appPrintLayoutHandout2')],
    ['handout3', t('appPrintLayoutHandout3')],
    ['handout4', t('appPrintLayoutHandout4')],
    ['handout6', t('appPrintLayoutHandout6')],
    ['handout9', t('appPrintLayoutHandout9')],
    ['notes', t('appPrintLayoutNotes')],
  ]
  // A single column (or no grid at all) leaves the order option nothing to swap
  const orderMatters = layout !== 'full' && layout !== 'notes' && layout !== 'handout3'

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        ref={dialogRef}
        className="modal print-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('appPrintTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h2>{t('appPrintTitle')}</h2>
        <div className="print-dialog-body">
          <div className="print-preview-pane" ref={paneRef}>
            {previewHtml ? (
              <iframe
                ref={frameRef}
                title={t('appPrintTitle')}
                sandbox="allow-same-origin"
                srcDoc={previewHtml}
                onLoad={applyZoom}
              />
            ) : (
              <div className="print-preview-empty">
                {pngs ? t('appPrintNoPages') : t('appPrintRendering')}
              </div>
            )}
          </div>
          <div className="print-options">
            <fieldset>
              <legend>{t('appPrintRange')}</legend>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-range"
                  checked={rangeMode === 'all'}
                  onChange={() => setRangeMode('all')}
                />
                {t('appPrintRangeAll')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-range"
                  checked={rangeMode === 'current'}
                  onChange={() => setRangeMode('current')}
                />
                {t('appPrintRangeCurrent')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-range"
                  checked={rangeMode === 'custom'}
                  onChange={() => setRangeMode('custom')}
                />
                {t('appPrintRangeCustom')}
              </label>
              <input
                type="text"
                className={`print-range-input ${rangeInvalid ? 'invalid' : ''}`}
                placeholder={t('appPrintRangeHint')}
                value={customRange}
                disabled={rangeMode !== 'custom'}
                onFocus={() => setRangeMode('custom')}
                onChange={(e) => setCustomRange(e.target.value)}
              />
            </fieldset>
            <fieldset>
              <legend>{t('appPrintLayoutGroup')}</legend>
              {layouts.map(([k, label]) => (
                <label key={k} className="print-radio">
                  <input
                    type="radio"
                    name="print-layout"
                    checked={layout === k}
                    onChange={() => setLayout(k)}
                  />
                  {label}
                </label>
              ))}
            </fieldset>
            <fieldset>
              <legend>{t('appPrintOrderGroup')}</legend>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-order"
                  disabled={!orderMatters}
                  checked={order === 'horizontal'}
                  onChange={() => setOrder('horizontal')}
                />
                {t('appPrintOrderHorizontal')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-order"
                  disabled={!orderMatters}
                  checked={order === 'vertical'}
                  onChange={() => setOrder('vertical')}
                />
                {t('appPrintOrderVertical')}
              </label>
            </fieldset>
            <fieldset>
              <legend>{t('appPrintOrientation')}</legend>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-orientation"
                  disabled={layout === 'full'}
                  checked={orientation === 'portrait'}
                  onChange={() => setOrientation('portrait')}
                />
                {t('appPrintPortrait')}
              </label>
              <label className="print-radio">
                <input
                  type="radio"
                  name="print-orientation"
                  disabled={layout === 'full'}
                  checked={orientation === 'landscape'}
                  onChange={() => setOrientation('landscape')}
                />
                {t('appPrintLandscape')}
              </label>
            </fieldset>
            <fieldset>
              <label className="print-radio">
                <input
                  type="checkbox"
                  disabled={layout !== 'full'}
                  checked={frame}
                  onChange={(e) => setFrame(e.target.checked)}
                />
                {t('appPrintFrame')}
              </label>
              <label className="print-radio">
                <input
                  type="checkbox"
                  checked={includeHidden}
                  onChange={(e) => setIncludeHidden(e.target.checked)}
                />
                {t('appPrintHidden')}
              </label>
            </fieldset>
            <fieldset>
              <legend>{t('appPrintPrinter')}</legend>
              <select
                className="print-select"
                value={printDevice}
                onChange={(e) => {
                  const device = e.target.value
                  setPrintDevice(device)
                  setDriverChoices({})
                  setDriverOptions([])
                  // Bumped on every switch (including the system-dialog entry),
                  // so a pending request for an earlier queue can never fill the
                  // groups of the queue selected now.
                  const request = ++driverOptionsRequest.current
                  if (device && device !== PRINT_SYSTEM_DIALOG && IS_MAC) {
                    void window.slidesApi
                      .printerOptions(device)
                      .then((groups) => {
                        if (request === driverOptionsRequest.current) setDriverOptions(groups)
                      })
                      .catch(() => {
                        if (request === driverOptionsRequest.current) setDriverOptions([])
                      })
                  }
                }}
              >
                <option value="">{t('appPrintDefaultPrinter')}</option>
                {printers.map((p) => (
                  <option key={p.name} value={p.name}>
                    {p.displayName}
                  </option>
                ))}
                <option value={PRINT_SYSTEM_DIALOG}>{t('appPrintSystemDialog')}</option>
              </select>
              {printDevice !== PRINT_SYSTEM_DIALOG && (
                <>
                  <div className="print-copies-row">
                    <label>
                      {t('appPrintCopies')}
                      <input
                        type="number"
                        className="print-range-input"
                        min={1}
                        max={999}
                        value={printCopies}
                        onChange={(e) =>
                          setPrintCopies(
                            Math.min(999, Math.max(1, Math.round(Number(e.target.value)) || 1)),
                          )
                        }
                      />
                    </label>
                    <label className="print-radio">
                      <input
                        type="checkbox"
                        checked={printCollate}
                        disabled={printCopies < 2}
                        onChange={(e) => setPrintCollate(e.target.checked)}
                      />
                      {t('appPrintCollate')}
                    </label>
                  </div>
                  <label className="print-radio">
                    {t('appPrintDuplex')}
                    <select
                      className="print-select"
                      value={printDuplex}
                      onChange={(e) => setPrintDuplex(e.target.value as PrintDuplex)}
                    >
                      {PRINT_DUPLEX_MODES.map((m) => (
                        <option key={m} value={m}>
                          {t(
                            m === 'simplex'
                              ? 'appPrintDuplexOff'
                              : m === 'longEdge'
                                ? 'appPrintDuplexLong'
                                : 'appPrintDuplexShort',
                          )}
                        </option>
                      ))}
                    </select>
                  </label>
                  {driverOptions.map((group) => (
                    <label key={group.name} className="print-radio">
                      {group.translation}
                      <select
                        className="print-select"
                        value={driverChoices[group.name] ?? group.defaultValue ?? group.values[0]}
                        onChange={(e) =>
                          setDriverChoices((c) => ({ ...c, [group.name]: e.target.value }))
                        }
                      >
                        {group.values.map((v) => (
                          <option key={v} value={v}>
                            {v}
                          </option>
                        ))}
                      </select>
                    </label>
                  ))}
                </>
              )}
            </fieldset>
            <div className="print-page-count">
              {selected.length > 0 ? t('appPrintPageCount', { n: pageCount }) : ''}
            </div>
          </div>
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('appSettingsCancel')}</button>
          <button
            className="primary"
            disabled={!pngs || selected.length === 0 || printing}
            onClick={() => void doPrint()}
          >
            {printing ? t('appPrintProgress') : t('appPrintTitle')}
          </button>
        </div>
      </div>
    </div>
  )
}
