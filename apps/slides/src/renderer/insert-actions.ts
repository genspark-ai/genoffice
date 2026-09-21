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
import { fileExt } from '../shared/media-kinds'
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
import {
  EQUATION_BODY_PR,
  EQUATION_FONT_FAMILY,
  EQUATION_FONT_PT,
  equationInsertFrame,
  graphicFrameInsertFrame,
} from './insert-defaults'
import { defaultShapeStyle } from './default-shape'
import { tableInsertSpec } from './table-insert'
import { textBoxInsertSpec } from './textbox-insert'
import { WORDART_FONT_PT, wordArtInsertSpec } from './wordart-insert'
import { insertSlideZooms } from './zoom-actions'

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
    ...(isLine ? { stroke: { color: '#000000', widthPt: 1 } } : defaultShapeStyle()),
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

/** Draw-mode commit for Insert > Text Box: empty one-line box at the gesture, straight into typing. */
export async function insertTextBoxAt(ctx: ActionCtx, rect: DrawRect): Promise<void> {
  const { slide, current } = ctx
  if (!slide) return
  const spec = textBoxInsertSpec(rect)
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind: 'textbox',
    xPx: spec.x,
    yPx: spec.y,
    wPx: spec.w,
    hPx: spec.h,
    fitWidthPx: FIT_WIDTH,
    bodyPr: spec.bodyPr,
  })
  if (!r) return
  ctx.applySlide(current, r.slide)
  ctx.setSelectedIds([r.sourceId])
  ctx.setEditing({ sourceId: r.sourceId, discardIfEmpty: true })
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
  const spec = tableInsertSpec(slide, rows)
  const r = await window.slidesApi.addTable({
    slideIndex: current,
    rows,
    cols,
    xPx: spec.x,
    yPx: spec.y,
    wPx: spec.w,
    hPx: spec.h,
    rowHeightEmu: spec.rowHeightEmu,
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
  const frame = graphicFrameInsertFrame(slide)
  const r = await window.slidesApi.addChart({
    slideIndex: current,
    kind,
    categories: data.categories,
    series: data.series,
    xPx: frame.x,
    yPx: frame.y,
    wPx: frame.w,
    hPx: frame.h,
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
  const frame = graphicFrameInsertFrame(slide)
  const r = await window.slidesApi.addSmartArt({
    slideIndex: current,
    layout: def.layout,
    items: def.defaultItems,
    xPx: frame.x,
    yPx: frame.y,
    wPx: frame.w,
    hPx: frame.h,
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
  const spec = wordArtInsertSpec(slide)
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind: 'textbox',
    xPx: spec.x,
    yPx: spec.y,
    wPx: spec.w,
    hPx: spec.h,
    fitWidthPx: FIT_WIDTH,
    bodyPr: spec.bodyPr,
    paragraphs: [
      {
        align: 'center',
        runs: [
          {
            text: t('appWordArtPlaceholder'),
            fontSize: WORDART_FONT_PT,
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
    ctx.setEditing({ sourceId: r.sourceId, selectAll: true })
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

/** Single-target Slide Zoom, a compatibility wrapper over the cascading insert. */
export async function insertZoom(ctx: ActionCtx, target: number): Promise<void> {
  await insertSlideZooms(ctx, [target])
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

/**
 * Equations: approximated as Cambria Math italic text (Unicode math symbols) for pptx compatibility.
 * PowerPoint drops a 1 in square centered on the slide that auto-fits the math and opens it for editing.
 */
export async function insertEquation(ctx: ActionCtx, text: string): Promise<void> {
  ctx.setEqDialogOpen(false)
  const { slide, current } = ctx
  if (!slide) return
  const frame = equationInsertFrame(slide)
  const r = await window.slidesApi.addElement({
    slideIndex: current,
    kind: 'textbox',
    xPx: frame.x,
    yPx: frame.y,
    wPx: frame.w,
    hPx: frame.h,
    fitWidthPx: FIT_WIDTH,
    bodyPr: EQUATION_BODY_PR,
    paragraphs: [
      {
        runs: [
          { text, fontSize: EQUATION_FONT_PT, italic: true, fontFamily: EQUATION_FONT_FAMILY },
        ],
      },
    ],
  })
  if (r) {
    ctx.applySlide(current, r.slide)
    ctx.setSelectedIds([r.sourceId])
    ctx.setEditing({ sourceId: r.sourceId })
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

/**
 * A video/audio file dropped on the canvas, embedded like the Insert dialog does but
 * centered on the drop point. Electron hands us the file's path, which lets main
 * build the poster frame from the system thumbnail; a path-less File falls back to bytes.
 */
export async function insertDroppedMedia(
  ctx: ActionCtx,
  file: File,
  kind: 'video' | 'audio',
  atPx: { x: number; y: number },
): Promise<void> {
  if (!ctx.slide) return
  const slideIndex = ctx.current
  const path = window.desktop.getPathForFile(file)
  const source = path
    ? { path }
    : { base64: bytesToBase64(new Uint8Array(await file.arrayBuffer())) }
  const r = await window.slidesApi.addMediaBytes({
    slideIndex,
    kind,
    ext: fileExt(file.name),
    fitWidthPx: FIT_WIDTH,
    name: file.name,
    centerPx: atPx,
    ...source,
  })
  if (!r) return
  ctx.applySlide(slideIndex, r.slide)
  ctx.setSelectedIds([r.sourceId])
  ctx.setStatus(
    r.warning ?? (kind === 'video' ? t('appStatusVideoInserted') : t('appStatusAudioInserted')),
  )
}

/** Chunked: spreading a large array into fromCharCode would blow the call stack. */
export function bytesToBase64(bytes: Uint8Array): string {
  let bin = ''
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  }
  return btoa(bin)
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
    const { width, height } = stream.getVideoTracks()[0]?.getSettings() ?? {}
    rec.onstop = async () => {
      stream.getTracks().forEach((t) => t.stop())
      ctx.recorderRef.current = null
      ctx.setRecording(false)
      const blob = new Blob(chunks, { type: 'video/webm' })
      if (blob.size === 0) {
        ctx.setStatus(t('appStatusRecordingEmpty'))
        return
      }
      const r = await window.slidesApi.addMediaBytes({
        slideIndex,
        kind: 'video',
        base64: bytesToBase64(new Uint8Array(await blob.arrayBuffer())),
        ext: 'webm',
        fitWidthPx: FIT_WIDTH,
        name: `Screen recording ${new Date().toLocaleTimeString()}`,
        ...(width && height ? { natural: { width, height } } : {}),
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
