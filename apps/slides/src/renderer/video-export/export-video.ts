/**
 * Export as Video: plays the show offline — every visible slide for its time, transitions,
 * animation steps run back to back, recorded narration on the audio track — and encodes the
 * frames into an MP4 next to the deck. Progress and cancellation flow through the dialog.
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import { buildSteps, computeNodeStates } from '../animation-play'
import type { FrameRenderer } from './frames'
import { AUDIO_SAMPLE_RATE, EncoderUnsupportedError, Mp4Writer } from './encoder'
import {
  animCursorAt,
  buildTimeline,
  segmentAt,
  totalDurationMs,
  type SlideInput,
  type SlideSegment,
} from './timeline'
import { drawTransition } from './transitions'

export const VIDEO_FPS = 30

export interface VideoExportOptions {
  /** Output height; the width follows the slide aspect */
  height: 1080 | 720
  defaultSlideMs: number
  useTimings: boolean
}

export type VideoExportResult =
  | { ok: true; path: string }
  | { ok: false; cancelled: true }
  | { ok: false; cancelled?: false; unsupported?: boolean; error: string }

interface Narration {
  samples: Float32Array<ArrayBuffer>
}

/** A slide's recorded narration clips, decoded and resampled to the track rate, laid end to end. */
async function loadNarration(slideIndex: number, ids: string[]): Promise<Narration | null> {
  const parts: Float32Array<ArrayBuffer>[] = []
  for (const id of ids) {
    const media = await window.slidesApi.getMediaData(slideIndex, id).catch(() => null)
    if (!media || media.kind !== 'audio') continue
    // The CSP keeps fetch() off data: URLs; an external link has no bytes to decode
    const b64 = media.dataUrl.startsWith('data:') ? media.dataUrl.split(',', 2)[1] : undefined
    if (!b64) continue
    try {
      const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))
      const probe = new OfflineAudioContext(1, 1, AUDIO_SAMPLE_RATE)
      const decoded = await probe.decodeAudioData(bytes.buffer)
      parts.push(mixdown(decoded))
    } catch {
      // an undecodable clip exports as silence
    }
  }
  if (!parts.length) return null
  const total = parts.reduce((n, p) => n + p.length, 0)
  const samples = new Float32Array(total)
  let at = 0
  for (const p of parts) {
    samples.set(p, at)
    at += p.length
  }
  return { samples }
}

function mixdown(audio: AudioBuffer): Float32Array<ArrayBuffer> {
  const out = new Float32Array(audio.length)
  for (let c = 0; c < audio.numberOfChannels; c++) {
    const ch = audio.getChannelData(c)
    for (let i = 0; i < out.length; i++) out[i]! += ch[i]! / audio.numberOfChannels
  }
  return out
}

/** Even output size at the requested height, keeping the slide aspect. */
export function outputSize(slide: RenderSlide, height: number): { width: number; height: number } {
  const width = Math.round(((slide.widthPx / slide.heightPx) * height) / 2) * 2
  return { width, height }
}

export async function exportVideo(
  input: {
    slides: RenderSlide[]
    images: Map<string, HTMLImageElement>
    path: string
  },
  opts: VideoExportOptions,
  onProgress: (fraction: number) => void,
  signal: AbortSignal,
): Promise<VideoExportResult> {
  const visible = input.slides.map((s, i) => ({ slide: s, i })).filter((e) => !e.slide.hidden)
  if (!visible.length) return { ok: false, error: 'no slides' }
  const cancelled = (): VideoExportResult => ({ ok: false, cancelled: true })

  let writer: Mp4Writer | undefined
  let frames: FrameRenderer | undefined
  try {
    const prefetch = await Promise.all(
      visible.map(async ({ i }) => {
        const [transition, anims, timing, keys] = await Promise.all([
          window.slidesApi.getTransition(i),
          window.slidesApi.getAnimations(i),
          window.slidesApi.getShowTiming(i),
          window.slidesApi.getShapeKeys(i),
        ])
        const narration =
          opts.useTimings && timing.narrationIds.length
            ? await loadNarration(i, timing.narrationIds)
            : null
        return { transition, anims, timing, keys, narration }
      }),
    )
    if (signal.aborted) return cancelled()

    const shown: RenderSlide[] = visible.map(({ slide }, k) => {
      const hide = new Set(prefetch[k]!.timing.narrationIds)
      return hide.size
        ? { ...slide, nodes: slide.nodes.filter((n) => !hide.has(n.sourceId)) }
        : slide
    })
    const inputs: SlideInput[] = visible.map((_, k) => {
      const p = prefetch[k]!
      return {
        index: k,
        transition: p.transition,
        advanceMs: p.timing.advanceMs,
        narrationMs: p.narration ? (p.narration.samples.length / AUDIO_SAMPLE_RATE) * 1000 : 0,
        steps: buildSteps(p.anims.filter((a) => a.effect !== 'motionPath' || !!a.motionPath)),
      }
    })
    const segments = buildTimeline(inputs, opts)
    const totalMs = totalDurationMs(segments)
    const totalFrames = Math.max(1, Math.round((totalMs / 1000) * VIDEO_FPS))
    const hasAudio = prefetch.some((p) => p.narration)

    const first = shown[0]!
    const size = outputSize(first, opts.height)
    const mp4 = await Mp4Writer.create({
      path: input.path,
      ...size,
      fps: VIDEO_FPS,
      audio: hasAudio,
    })
    writer = mp4
    // Loaded on demand: the Konva stages pull react-konva into every importer of file-actions otherwise
    const { FrameRenderer } = await import('./frames')
    const renderer = new FrameRenderer(size.width, size.height, input.images)
    frames = renderer
    const compose = document.createElement('canvas')
    compose.width = size.width
    compose.height = size.height
    const cctx = compose.getContext('2d')!
    const black = document.createElement('canvas')
    black.width = size.width
    black.height = size.height
    black.getContext('2d')!.fillRect(0, 0, size.width, size.height)

    /** Resting frames of each slide by animation cursor, so a static hold renders once. */
    const restCache = new Map<string, HTMLCanvasElement>()
    const restFrame = (seg: SlideSegment, played: number): HTMLCanvasElement => {
      const key = `${seg.index}:${played}`
      let c = restCache.get(key)
      if (!c) {
        const states = computeNodeStates(seg.steps, played, null, first.heightPx, first.widthPx)
        c = renderer.slideFrame(shown[seg.index]!, states)
        // Only the entering and the fully played frames are revisited (transitions, holds)
        if (restCache.size > 8) restCache.delete(restCache.keys().next().value!)
        restCache.set(key, c)
      }
      return c
    }
    /** Last frame of the previous segment: the page the transition leaves. */
    const leavingFrame = (seg: SlideSegment): HTMLCanvasElement => {
      const prev = segments[seg.index - 1]
      if (!prev) return black
      return restFrame(prev, prev.steps.length)
    }

    const audioFed = new Set<number>()
    const feedAudio = async (seg: SlideSegment) => {
      if (!mp4.hasAudio || audioFed.has(seg.index)) return
      audioFed.add(seg.index)
      // Each segment contributes exactly its length so the tracks never drift apart
      const want = Math.round(((seg.startMs + seg.durationMs) / 1000) * AUDIO_SAMPLE_RATE)
      const have = Math.round(mp4.audioSeconds * AUDIO_SAMPLE_RATE)
      const buf = new Float32Array(Math.max(0, want - have))
      const narration = prefetch[seg.index]!.narration
      if (narration) {
        const offset = Math.round((seg.transitionMs / 1000) * AUDIO_SAMPLE_RATE)
        const n = Math.min(narration.samples.length, Math.max(0, buf.length - offset))
        buf.set(narration.samples.subarray(0, n), offset)
      }
      await mp4.addAudio(buf)
    }

    for (let f = 0; f < totalFrames; f++) {
      if (signal.aborted) {
        await mp4.abort()
        return cancelled()
      }
      const tMs = (f * 1000) / VIDEO_FPS
      const seg = segmentAt(segments, tMs)!
      await feedAudio(seg)
      const tRel = tMs - seg.startMs
      let frame: HTMLCanvasElement
      if (seg.transition && tRel < seg.transitionMs) {
        const u = tRel / seg.transitionMs
        const prevSeg = segments[seg.index - 1]
        if (seg.transition === 'morph' && prevSeg) {
          const from = prefetch[prevSeg.index]!
          const to = prefetch[seg.index]!
          frame = renderer.morphFrame(
            shown[prevSeg.index]!,
            shown[seg.index]!,
            from.keys,
            to.keys,
            u,
          )
        } else {
          const kind = seg.transition === 'morph' ? 'fade' : seg.transition
          drawTransition(
            cctx,
            leavingFrame(seg),
            restFrame(seg, 0),
            kind,
            u,
            size.width,
            size.height,
          )
          frame = compose
        }
      } else {
        const cursor = animCursorAt(seg, tRel)
        frame =
          cursor.activeMs == null
            ? restFrame(seg, cursor.played)
            : renderer.slideFrame(
                shown[seg.index]!,
                computeNodeStates(
                  seg.steps,
                  cursor.played,
                  cursor.activeMs,
                  first.heightPx,
                  first.widthPx,
                ),
              )
      }
      await mp4.addFrame(frame)
      if (f % 5 === 0) {
        onProgress(f / totalFrames)
        // Let the dialog repaint between frames
        await new Promise((r) => setTimeout(r, 0))
      }
    }
    await mp4.finish()
    onProgress(1)
    return { ok: true, path: input.path }
  } catch (e) {
    await writer?.abort()
    if (e instanceof EncoderUnsupportedError)
      return { ok: false, unsupported: true, error: e.message }
    return { ok: false, error: String(e) }
  } finally {
    frames?.dispose()
  }
}
