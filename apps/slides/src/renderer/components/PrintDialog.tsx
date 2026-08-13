/**
 * Print dialog modeled on PowerPoint's print sheet: live page preview on the left
 * (the exact pages that will print, built by the shared print-html assembly),
 * range / layout / orientation / framing options on the right. "Print" hands the
 * assembled pages to the system print dialog.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import { useI18n } from '../i18n/locale'
import {
  buildPrintDocumentHtml,
  parsePrintRange,
  printPageCount,
  type PrintLayout,
  type PrintOrientation,
} from '../../shared/print-html'
import { renderSlidesToPngBase64 } from '../export-render'

type RangeMode = 'all' | 'current' | 'custom'

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
  const { t } = useI18n()
  const [pngs, setPngs] = useState<string[] | null>(null)
  const [notes, setNotes] = useState<string[]>([])
  const [layout, setLayout] = useState<PrintLayout>('full')
  const [rangeMode, setRangeMode] = useState<RangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [orientation, setOrientation] = useState<PrintOrientation>('portrait')
  const [frame, setFrame] = useState(false)
  const [includeHidden, setIncludeHidden] = useState(false)
  const [printing, setPrinting] = useState(false)
  const blobUrlsRef = useRef<string[]>([])
  const frameRef = useRef<HTMLIFrameElement | null>(null)
  const paneRef = useRef<HTMLDivElement | null>(null)

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

  /** 0-based deck indices selected by the range options (before the hidden filter) */
  const rangeIndices = useMemo<number[] | null>(() => {
    if (rangeMode === 'all') return slides.map((_s, i) => i)
    if (rangeMode === 'current') return [current]
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
      frame,
      preview: true,
    })
  }, [pngs, selected, slides, layout, notes, orientation, frame])

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
      const r = await window.slidesApi.printSlides({
        pngsBase64: selected.map((i) => pngs[i]),
        widthPx: first.widthPx,
        heightPx: first.heightPx,
        layout,
        ...(layout === 'notes' ? { notes: selected.map((i) => notes[i] ?? '') } : {}),
        ...(layout !== 'full' ? { orientation } : {}),
        ...(frame && layout === 'full' ? { frame: true } : {}),
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
    ['handout6', t('appPrintLayoutHandout6')],
    ['notes', t('appPrintLayoutNotes')],
  ]

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal print-dialog" onClick={(e) => e.stopPropagation()}>
        <h2>{t('appPrintTitle')}</h2>
        <div className="print-dialog-body">
          <div className="print-preview-pane" ref={paneRef}>
            {previewHtml ? (
              <iframe
                ref={frameRef}
                title="print-preview"
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
