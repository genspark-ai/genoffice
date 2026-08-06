/**
 * Insert-tab actions extracted from App.tsx: shapes/text boxes,
 * images, tables, icons, charts, SmartArt, WordArt, fields, links, Zoom,
 * header/footer, equations, media, 3D models, and screen recording.
 * Functions take the ActionCtx built fresh per call.
 */
import type { InsertKind, LinkTargetOp } from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { applySelectionLink, saveEditSelection, selectionLink } from './TextEditOverlay'
import { FIT_WIDTH } from './app-constants'
import type { WordArtPreset } from '@genoffice/ui'
import {
  chartSampleData,
  iconSvg,
  type ChartPresetDef,
  type IconDef,
  type SmartArtDef,
} from './insert-presets'
import { t } from './i18n/locale'
import { isLineDrawKind, type DrawRect } from './draw-shape'

/** Draw-mode commit: insert a gallery shape at the drawn box (PowerPoint click-or-drag sizing). */
export async function insertShapeAt(
  ctx: ActionCtx,
  kind: InsertKind,
  rect: DrawRect,
): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const isLine = isLineDrawKind(kind)
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind,
    xPx: Math.round(rect.x),
    yPx: Math.round(rect.y),
    wPx: Math.round(rect.w),
    hPx: Math.round(rect.h),
    fitWidthPx: FIT_WIDTH,
    ...(isLine ? { stroke: { color: '#000000', widthPt: 1 } } : { fillColor: '#C43E1C' }),
  })
  if (!r) return
  let updated = r.slide
  // Connectors render top-left → bottom-right inside their box; leftward/upward drags are restored by mirroring
  for (const axis of [
    ...(rect.flipH ? (['h'] as const) : []),
    ...(rect.flipV ? (['v'] as const) : []),
  ]) {
    const f = await window.slidesApi.flipElements({
      slideIndex: current,
      sourceIds: [r.sourceId],
      axis,
    })
    if (f) updated = f
  }
  ctx.applySlide(current, updated)
  ctx.setSelectedIds([r.sourceId])
}

export async function insertElement(ctx: ActionCtx, kind: InsertKind): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  // Lines insert as a horizontal stroke-only connector (no fill, no text);
  // bent/curved connectors need a real box for their elbow/curve geometry
  const isStraightLine = kind === 'line' || kind === 'lineArrow' || kind === 'lineArrowDouble'
  const isLine = isStraightLine || kind === 'lineBent' || kind === 'lineCurved'
  const w = kind === 'textbox' ? 360 : 240
  const h = kind === 'textbox' ? 60 : isStraightLine ? 0 : 160
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind,
    xPx: Math.round((slide.widthPx - w) / 2),
    yPx: Math.round((slide.heightPx - h) / 2),
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
    ...(kind === 'textbox'
      ? { text: '' }
      : isLine
        ? { stroke: { color: '#000000', widthPt: 1 } }
        : { fillColor: '#C43E1C' }),
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    if (kind === 'textbox') ctx.setEditing({ sourceId: r.sourceId })
  }
}

export async function insertImage(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide) return
  const r = await window.slidesApi.insertImage(ctx.current, FIT_WIDTH)
  if (!r) return
  if ('error' in r) {
    ctx.setSelectedIds([])
    ctx.setStatus(t('appStatusImageUnsupported', { ext: r.ext.toUpperCase() }))
    return
  }
  ctx.applySlide(ctx.current, r.slide)
  ctx.setSelectedIds([r.sourceId])
}

export async function insertTable(ctx: ActionCtx, rows: number, cols: number): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const w = Math.min(Math.round(slide.widthPx * 0.7), 760)
  const h = Math.min(Math.round(slide.heightPx * 0.6), rows * 44)
  const r = await window.slidesApi.addTable({
    slideIndex: current,
    rows,
    cols,
    xPx: Math.round((slide.widthPx - w) / 2),
    yPx: Math.round((slide.heightPx - h) / 2),
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(t('appStatusTableInserted', { rows, cols }))
  }
}

/** Icons: SVG rasterized to PNG then inserted as an image (embedded PNG has the best pptx compatibility). */
export async function insertIcon(ctx: ActionCtx, def: IconDef, color: string): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  try {
    const svg = iconSvg(def, color)
    const img = new Image()
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve()
      img.onerror = () => reject(new Error('svg decode failed'))
      img.src = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`
    })
    const canvas = document.createElement('canvas')
    canvas.width = 512
    canvas.height = 512
    canvas.getContext('2d')!.drawImage(img, 0, 0, 512, 512)
    const base64 = canvas.toDataURL('image/png').split(',')[1]!
    const size = 96
    const r = await window.slidesApi.addImageBytes({
      slideIndex: current,
      base64,
      ext: 'png',
      xPx: Math.round((slide.widthPx - size) / 2),
      yPx: Math.round((slide.heightPx - size) / 2),
      wPx: size,
      hPx: size,
      fitWidthPx: FIT_WIDTH,
      name: `Icon ${def.name}`,
    })
    if (r && !('error' in r)) {
      ctx.applySlide(current, r.slide)
      ctx.setSelectedIds([r.sourceId])
      ctx.setStatus(t('appStatusIconInserted', { name: def.name }))
    }
  } catch {
    ctx.setStatus(t('appStatusIconInsertFailed'))
  }
}

export async function insertChart(ctx: ActionCtx, kind: ChartPresetDef['kind']): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const data = chartSampleData(kind)
  const w = Math.round(slide.widthPx * 0.62)
  const h = Math.round(slide.heightPx * 0.62)
  const r = await window.slidesApi.addChart({
    slideIndex: current,
    kind,
    categories: data.categories,
    series: data.series,
    xPx: Math.round((slide.widthPx - w) / 2),
    yPx: Math.round((slide.heightPx - h) / 2),
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(t('appStatusChartInserted'))
  }
}

export async function insertSmartArt(ctx: ActionCtx, def: SmartArtDef): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const w = Math.round(slide.widthPx * 0.7)
  const h = Math.round(slide.heightPx * 0.5)
  const r = await window.slidesApi.addSmartArt({
    slideIndex: current,
    layout: def.layout,
    items: def.defaultItems,
    xPx: Math.round((slide.widthPx - w) / 2),
    yPx: Math.round((slide.heightPx - h) / 2),
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(t('appStatusSmartArtInserted', { name: def.label }))
  }
}

export async function insertWordArt(ctx: ActionCtx, preset: WordArtPreset): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const w = 520
  const h = 100
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind: 'textbox',
    xPx: Math.round((slide.widthPx - w) / 2),
    yPx: Math.round((slide.heightPx - h) / 2),
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
    paragraphs: [
      {
        align: 'center',
        runs: [
          {
            text: t('appWordArtPlaceholder'),
            fontSize: 40,
            bold: preset.bold,
            italic: preset.italic,
            color: preset.fill,
            ...(preset.outline ? { outline: preset.outline } : {}),
          },
        ],
      },
    ],
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(t('appStatusWordArtInserted'))
  }
}

/** Date-time / slide number: <a:fld> dynamic fields, auto-refreshed when PowerPoint opens the file. */
export async function insertField(ctx: ActionCtx, type: 'datetime' | 'slidenum'): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const isDate = type === 'datetime'
  const w = isDate ? 240 : 100
  const h = 44
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind: 'textbox',
    xPx: Math.round((slide.widthPx - w) / 2),
    yPx: Math.round((slide.heightPx - h) / 2),
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
    paragraphs: [
      {
        align: 'center',
        runs: [
          {
            text: isDate ? new Date().toLocaleDateString() : String(current + 1),
            fontSize: 18,
            field: isDate ? 'datetime1' : 'slidenum',
          },
        ],
      },
    ],
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(isDate ? t('appStatusDateInserted') : t('appStatusSlideNumInserted'))
  }
}

export async function openLinkDialog(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide) return
  // Editing text: the link applies to the selection (run-level); the keep-edit ribbon
  // button / dialog keep the editor alive, saveEditSelection preserved the range
  if (document.activeElement instanceof HTMLElement && document.activeElement.isContentEditable) {
    saveEditSelection()
    ctx.setLinkDialog({ sourceId: null, run: true, initial: selectionLink() })
    return
  }
  if (ctx.selectedIds.length !== 1) return
  const sourceId = ctx.selectedIds[0]!
  const initial = await window.slidesApi.getLink(ctx.current, sourceId)
  ctx.setLinkDialog({ sourceId, initial })
}

export async function applyLink(ctx: ActionCtx, target: LinkTargetOp | null): Promise<void> {
  if (!ctx.linkDialog) return
  if (ctx.linkDialog.run) {
    const ok = applySelectionLink(target)
    ctx.setLinkDialog(null)
    if (ok) ctx.setStatus(target ? t('appStatusLinkSet') : t('appStatusLinkRemoved'))
    return
  }
  if (!ctx.linkDialog.sourceId) return
  const updated = await window.slidesApi.setLink({
    slideIndex: ctx.current,
    sourceId: ctx.linkDialog.sourceId,
    target,
  })
  ctx.setLinkDialog(null)
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setStatus(target ? t('appStatusLinkSet') : t('appStatusLinkRemoved'))
  }
}

/** Zoom link (simplified Zoom): a button shape with an internal jump link. */
export async function insertZoom(ctx: ActionCtx, target: number): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const w = 200
  const h = 60
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind: 'roundRect',
    xPx: slide.widthPx - w - 28,
    yPx: slide.heightPx - h - 28,
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
    fillColor: '#4472C4',
    paragraphs: [
      {
        align: 'center',
        runs: [
          {
            text: t('appZoomButtonText', { page: target + 1 }),
            fontSize: 17,
            bold: true,
            color: '#FFFFFF',
          },
        ],
      },
    ],
  })
  if (!r) return
  const linked = await window.slidesApi.setLink({
    slideIndex: current,
    sourceId: r.sourceId,
    target: { kind: 'slide', slideIndex: target },
  })
  ctx.applySlide(current, linked ?? r.slide)
  ctx.setStatus(t('appStatusZoomInserted', { page: target + 1 }))
}

export async function openHeaderFooter(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide) return
  ctx.setHfDialog(await window.slidesApi.getHeaderFooter(ctx.current))
}

export async function applyHf(
  ctx: ActionCtx,
  opts: { footer: string | null; slideNum: boolean; date: string | null; dateAuto: boolean },
): Promise<void> {
  ctx.setHfDialog(null)
  const updated = await window.slidesApi.applyHeaderFooter({ ...opts, fitWidthPx: FIT_WIDTH })
  if (updated) {
    ctx.setSlides(updated)
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusHfApplied'))
  } else {
    ctx.setStatus(t('appStatusHfUnchanged'))
  }
}

/** Equations: approximated as Cambria Math italic text (Unicode math symbols) for pptx compatibility. */
export async function insertEquation(ctx: ActionCtx, text: string): Promise<void> {
  ctx.setEqDialogOpen(false)
  const { slide, current } = ctx
  if (!slide) return
  const w = 420
  const h = 84
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind: 'textbox',
    xPx: Math.round((slide.widthPx - w) / 2),
    yPx: Math.round((slide.heightPx - h) / 2),
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
    paragraphs: [
      {
        align: 'center',
        runs: [{ text, fontSize: 28, italic: true, fontFamily: 'Cambria Math' }],
      },
    ],
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(t('appStatusEquationInserted'))
  }
}

export async function insertMediaFile(ctx: ActionCtx, kind: 'video' | 'audio'): Promise<void> {
  if (!ctx.slide) return
  const r = await window.slidesApi.insertMedia(ctx.current, kind, FIT_WIDTH)
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(kind === 'video' ? t('appStatusVideoInserted') : t('appStatusAudioInserted'))
  }
}

export async function insertModel3dFile(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide) return
  const r = await window.slidesApi.insertModel3d(ctx.current, FIT_WIDTH)
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setStatus(t('appStatusModel3dInserted'))
  }
}

/** Screen recording: getDisplayMedia + MediaRecorder; when stopped, inserted into the current page as webm video. */
export async function toggleScreenRecord(ctx: ActionCtx): Promise<void> {
  if (ctx.recorderRef.current) {
    ctx.recorderRef.current.rec.stop()
    return
  }
  if (!ctx.slide) return
  try {
    const stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false })
    const mime = MediaRecorder.isTypeSupported('video/webm;codecs=vp9')
      ? 'video/webm;codecs=vp9'
      : 'video/webm'
    const rec = new MediaRecorder(stream, { mimeType: mime })
    const chunks: Blob[] = []
    rec.ondataavailable = (ev) => {
      if (ev.data.size > 0) chunks.push(ev.data)
    }
    const slideIndex = ctx.current
    rec.onstop = async () => {
      stream.getTracks().forEach((t) => t.stop())
      ctx.recorderRef.current = null
      ctx.setRecording(false)
      const blob = new Blob(chunks, { type: 'video/webm' })
      if (blob.size === 0) {
        ctx.setStatus(t('appStatusRecordingEmpty'))
        return
      }
      // Blob → base64 (chunked to avoid call stack overflow)
      const buf = new Uint8Array(await blob.arrayBuffer())
      let bin = ''
      for (let i = 0; i < buf.length; i += 0x8000) {
        bin += String.fromCharCode(...buf.subarray(i, i + 0x8000))
      }
      const r = await window.slidesApi.addMediaBytes({
        slideIndex,
        kind: 'video',
        base64: btoa(bin),
        ext: 'webm',
        fitWidthPx: FIT_WIDTH,
        name: `Screen recording ${new Date().toLocaleTimeString()}`,
      })
      if (r) {
        ctx.applySlide(slideIndex, r.slide)
        ctx.setSelectedIds([r.sourceId])
        ctx.setStatus(t('appStatusRecordingInserted'))
      }
    }
    // Also clean up when the user ends via the system "Stop sharing"
    stream.getVideoTracks()[0]?.addEventListener('ended', () => {
      if (rec.state !== 'inactive') rec.stop()
    })
    rec.start(1000)
    ctx.recorderRef.current = { rec, stream }
    ctx.setRecording(true)
    ctx.setStatus(t('appStatusRecording'))
  } catch {
    ctx.setStatus(t('appStatusRecordingUnavailable'))
  }
}
