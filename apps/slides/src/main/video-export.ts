import { execFile } from 'node:child_process'
import { mkdtemp, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/** Export to MP4 (genoffice#2508): visible slides are rendered by the renderer to
 * PNG frames; this module writes the frames + a concat list and muxes them with
 * ffmpeg into a 1080p H.264 MP4. No cloud service is involved — the only
 * external dependency is the user's local ffmpeg binary. */

export const VIDEO_EXPORT_WIDTH = 1920
export const VIDEO_EXPORT_HEIGHT = 1080
export const VIDEO_EXPORT_FPS = 30
export const VIDEO_EXPORT_DEFAULT_DURATION_SEC = 3
export const VIDEO_EXPORT_MIN_DURATION_SEC = 0.5
export const VIDEO_EXPORT_MAX_DURATION_SEC = 30

export const FFMPEG_MISSING_ERROR =
  'ffmpeg not found on PATH. Install ffmpeg (https://ffmpeg.org/download.html) ' +
  'or set the FFMPEG_PATH environment variable to the ffmpeg binary, then try again.'

export function clampSlideDurationSec(value: number | undefined): number {
  if (!Number.isFinite(value)) return VIDEO_EXPORT_DEFAULT_DURATION_SEC
  return Math.min(Math.max(value as number, VIDEO_EXPORT_MIN_DURATION_SEC), VIDEO_EXPORT_MAX_DURATION_SEC)
}

/**
 * Per-slide durations in slide order. Rehearsed timings (durationsSec) win when
 * present; the rest fall back to slideDurationSec, then the 3s default.
 */
export function resolveVideoDurations(
  count: number,
  durationsSec?: number[],
  slideDurationSec?: number,
): number[] {
  const fallback = clampSlideDurationSec(slideDurationSec)
  return Array.from({ length: Math.max(0, count) }, (_, i) => {
    const d = durationsSec?.[i]
    return Number.isFinite(d) ? clampSlideDurationSec(d) : fallback
  })
}

/** Frame file names in slide order: frame-001.png … (pad >= 3 digits). */
export function buildVideoFrameNames(count: number): string[] {
  const pad = Math.max(3, String(Math.max(0, count)).length)
  return Array.from({ length: Math.max(0, count) }, (_, i) =>
    `frame-${String(i + 1).padStart(pad, '0')}.png`,
  )
}

/**
 * ffmpeg concat-demuxer list. Every frame but the last carries its duration;
 * the last frame is repeated so its duration takes effect.
 */
export function buildConcatList(frameNames: string[], durationsSec: number[]): string {
  const durations = resolveVideoDurations(frameNames.length, durationsSec)
  const lines: string[] = []
  frameNames.forEach((name, i) => {
    lines.push(`file '${name}'`)
    lines.push(`duration ${durations[i]}`)
  })
  if (frameNames.length > 0) lines.push(`file '${frameNames[frameNames.length - 1]}'`)
  return lines.join('\n') + '\n'
}

/** Scale/pad to 1080p (even dims for yuv420p) — black bars keep the slide aspect. */
export function buildVideoFilter(width = VIDEO_EXPORT_WIDTH, height = VIDEO_EXPORT_HEIGHT): string {
  return `scale=${width}:${height}:force_original_aspect_ratio=decrease,pad=${width}:${height}:(ow-iw)/2:(oh-ih)/2,format=yuv420p`
}

export function buildFfmpegArgs(listPath: string, outPath: string): string[] {
  return [
    '-y',
    '-f', 'concat',
    '-safe', '0',
    '-i', listPath,
    '-vf', buildVideoFilter(),
    '-r', String(VIDEO_EXPORT_FPS),
    '-c:v', 'libx264',
    '-pix_fmt', 'yuv420p',
    '-movflags', '+faststart',
    outPath,
  ]
}

export interface VideoFfmpegRun {
  ok: boolean
  stderr?: string
  error?: string
}

function defaultRunFfmpeg(bin: string, args: string[]): Promise<VideoFfmpegRun> {
  return new Promise((resolve) => {
    execFile(bin, args, { maxBuffer: 16 * 1024 * 1024 }, (error, _stdout, stderr) => {
      if (!error) return resolve({ ok: true, stderr: String(stderr ?? '') })
      const code = (error as NodeJS.ErrnoException).code
      if (code === 'ENOENT') return resolve({ ok: false, error: FFMPEG_MISSING_ERROR })
      resolve({ ok: false, error: String(error), stderr: String(stderr ?? '') })
    })
  })
}

function defaultCheckFfmpeg(bin: string): Promise<boolean> {
  return new Promise((resolve) => {
    execFile(bin, ['-version'], { maxBuffer: 1024 * 1024 }, (error) => resolve(!error))
  })
}

export interface ExportSlidesVideoOptions {
  /** base64 PNG per visible slide, in slide order */
  pngsBase64: string[]
  /** Target mp4 absolute path (chosen via pickExportVideoPath) */
  filePath: string
  /** Per-slide durations in seconds (e.g. rehearsed timings); missing entries use the default */
  durationsSec?: number[]
  /** Fallback per-slide duration in seconds (default 3) */
  slideDurationSec?: number
  /** Override the ffmpeg binary (default: $FFMPEG_PATH or "ffmpeg" on PATH) */
  ffmpegBin?: string
  /** Injected for tests: run ffmpeg and report success */
  runFfmpeg?: (bin: string, args: string[]) => Promise<VideoFfmpegRun>
  /** Injected for tests: whether the ffmpeg binary is usable */
  checkFfmpeg?: (bin: string) => Promise<boolean>
}

export interface ExportSlidesVideoResult {
  ok: boolean
  path?: string
  error?: string
}

export async function exportSlidesVideo(opts: ExportSlidesVideoOptions): Promise<ExportSlidesVideoResult> {
  const pngs = opts.pngsBase64 ?? []
  if (pngs.length === 0) return { ok: false, error: 'No slides to export.' }
  if (!opts.filePath) return { ok: false, error: 'No output path for video export.' }

  const bin = opts.ffmpegBin ?? process.env.FFMPEG_PATH ?? 'ffmpeg'
  const checkFfmpeg = opts.checkFfmpeg ?? defaultCheckFfmpeg
  if (!(await checkFfmpeg(bin))) return { ok: false, error: FFMPEG_MISSING_ERROR }

  const names = buildVideoFrameNames(pngs.length)
  const durations = resolveVideoDurations(pngs.length, opts.durationsSec, opts.slideDurationSec)
  const runFfmpeg = opts.runFfmpeg ?? defaultRunFfmpeg

  let tempDir: string | null = null
  try {
    tempDir = await mkdtemp(join(tmpdir(), 'genoffice-slides-video-'))
    // frames live next to the list so relative `file 'frame-…'` entries resolve
    for (let i = 0; i < pngs.length; i++) {
      if (!pngs[i]) throw new Error(`Slide ${i + 1} has no rendered frame.`)
      await writeFile(join(tempDir, names[i]), Buffer.from(pngs[i], 'base64'))
    }
    const listPath = join(tempDir, 'concat.txt')
    await writeFile(listPath, buildConcatList(names, durations), 'utf8')

    const run = await runFfmpeg(bin, buildFfmpegArgs(listPath, opts.filePath))
    if (!run.ok) {
      if (run.error === FFMPEG_MISSING_ERROR) return { ok: false, error: FFMPEG_MISSING_ERROR }
      const detail = run.stderr?.trim() ? `: ${run.stderr.trim().slice(-500)}` : ''
      return { ok: false, error: `ffmpeg failed${detail}` }
    }
    // guard against a runner that reports success without writing the file
    try {
      const st = await stat(opts.filePath)
      if (!st.isFile() || st.size === 0)
        return { ok: false, error: 'ffmpeg produced an empty video file.' }
    } catch {
      return { ok: false, error: 'ffmpeg did not produce the video file.' }
    }
    return { ok: true, path: opts.filePath }
  } catch (err) {
    return { ok: false, error: String(err) }
  } finally {
    if (tempDir) await rm(tempDir, { recursive: true, force: true })
  }
}
