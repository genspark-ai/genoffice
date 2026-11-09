import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { PNG } from 'pngjs'
import { afterEach, describe, expect, it } from 'vitest'

import {
  buildConcatList,
  buildFfmpegArgs,
  buildVideoFilter,
  buildVideoFrameNames,
  clampSlideDurationSec,
  exportSlidesVideo,
  resolveVideoDurations,
  FFMPEG_MISSING_ERROR,
} from '../src/main/video-export'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function outputPath(name = 'export.mp4'): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-slides-video-test-'))
  roots.push(root)
  return join(root, name)
}

function solidPngBase64(red: number, green: number, blue: number): string {
  const png = new PNG({ width: 16, height: 9 })
  for (let i = 0; i < png.data.length; i += 4) {
    png.data[i] = red
    png.data[i + 1] = green
    png.data[i + 2] = blue
    png.data[i + 3] = 255
  }
  return PNG.sync.write(png).toString('base64')
}

describe('video export helpers', () => {
  it('names frames in slide order', () => {
    expect(buildVideoFrameNames(3)).toEqual(['frame-001.png', 'frame-002.png', 'frame-003.png'])
    expect(buildVideoFrameNames(0)).toEqual([])
  })

  it('resolves durations: rehearsed timings win, default 3s otherwise', () => {
    expect(resolveVideoDurations(3)).toEqual([3, 3, 3])
    expect(resolveVideoDurations(3, [5])).toEqual([5, 3, 3])
    expect(resolveVideoDurations(2, [5, 7])).toEqual([5, 7])
    expect(resolveVideoDurations(2, undefined, 4)).toEqual([4, 4])
    expect(resolveVideoDurations(0)).toEqual([])
  })

  it('clamps durations to 0.5..30s and rejects non-finite values', () => {
    expect(clampSlideDurationSec(0)).toBe(0.5)
    expect(clampSlideDurationSec(100)).toBe(30)
    expect(clampSlideDurationSec(Number.NaN)).toBe(3)
    expect(clampSlideDurationSec(undefined)).toBe(3)
    expect(clampSlideDurationSec(2.5)).toBe(2.5)
  })

  it('builds a concat list that preserves slide order and repeats the last frame', () => {
    const list = buildConcatList(['frame-001.png', 'frame-002.png'], [5, 2])
    expect(list).toBe(
      "file 'frame-001.png'\nduration 5\nfile 'frame-002.png'\nduration 2\nfile 'frame-002.png'\n",
    )
    expect(buildConcatList([], [])).toBe('\n')
  })

  it('builds 1080p H.264 ffmpeg args', () => {
    const args = buildFfmpegArgs('/tmp/concat.txt', '/tmp/out.mp4')
    expect(args).toContain('-f')
    expect(args).toContain('concat')
    expect(args).toContain('-i')
    expect(args[args.indexOf('-i') + 1]).toBe('/tmp/concat.txt')
    expect(args).toContain('libx264')
    expect(args).toContain('yuv420p')
    expect(args).toContain('+faststart')
    expect(args[args.length - 1]).toBe('/tmp/out.mp4')
    expect(args).toContain('-y')
    const filter = buildVideoFilter()
    expect(filter).toContain('scale=1920:1080')
    expect(filter).toContain('pad=1920:1080')
    expect(filter).toContain('format=yuv420p')
  })
})

describe('exportSlidesVideo', () => {
  it('fails clearly when ffmpeg is missing', async () => {
    const filePath = await outputPath()
    const result = await exportSlidesVideo({
      pngsBase64: [solidPngBase64(255, 0, 0)],
      filePath,
      checkFfmpeg: async () => false,
    })
    expect(result.ok).toBe(false)
    expect(result.error).toBe(FFMPEG_MISSING_ERROR)
    expect(existsSync(filePath)).toBe(false)
  })

  it('fails on an empty deck without running ffmpeg', async () => {
    let ran = false
    const result = await exportSlidesVideo({
      pngsBase64: [],
      filePath: await outputPath(),
      checkFfmpeg: async () => true,
      runFfmpeg: async () => {
        ran = true
        return { ok: true }
      },
    })
    expect(result.ok).toBe(false)
    expect(result.error).toContain('No slides')
    expect(ran).toBe(false)
  })

  it('writes frames in order, runs ffmpeg, and writes the mp4', async () => {
    const filePath = await outputPath()
    const frames = [solidPngBase64(255, 0, 0), solidPngBase64(0, 255, 0)]
    let seenArgs: string[] | null = null
    let seenList: string | null = null
    let tempDirAtRun: string | null = null

    const result = await exportSlidesVideo({
      pngsBase64: frames,
      filePath,
      durationsSec: [5, 2],
      checkFfmpeg: async () => true,
      runFfmpeg: async (_bin, args) => {
        seenArgs = args
        const listPath = args[args.indexOf('-i') + 1]!
        tempDirAtRun = dirname(listPath)
        seenList = await readFile(listPath, 'utf8')
        const frameText = await readFile(join(tempDirAtRun, 'concat.txt'), 'utf8')
        expect(frameText).toContain('frame-001.png')
        expect(frameText).toContain('frame-002.png')
        // slide order: frame 001 bytes decode to red, frame 002 to green
        const first = PNG.sync.read(Buffer.from(await readFile(join(tempDirAtRun, 'frame-001.png'))))
        const second = PNG.sync.read(Buffer.from(await readFile(join(tempDirAtRun, 'frame-002.png'))))
        expect(first.data[0]).toBe(255)
        expect(second.data[1]).toBe(255)
        await writeFile(filePath, Buffer.from('FAKE-MP4'))
        return { ok: true }
      },
    })

    expect(result).toEqual({ ok: true, path: filePath })
    expect(seenArgs![seenArgs!.length - 1]).toBe(filePath)
    expect(seenArgs).toContain('libx264')
    expect(seenList).toContain('duration 5')
    expect(seenList).toContain('duration 2')
    expect(basename(tempDirAtRun!)).toMatch(/^genoffice-slides-video-/)
    // temp frames are cleaned up after success
    expect(existsSync(tempDirAtRun!)).toBe(false)
    await expect(readFile(filePath)).resolves.toEqual(Buffer.from('FAKE-MP4'))
  })

  it('reports ffmpeg failures with stderr detail', async () => {
    const result = await exportSlidesVideo({
      pngsBase64: [solidPngBase64(0, 0, 255)],
      filePath: await outputPath(),
      checkFfmpeg: async () => true,
      runFfmpeg: async () => ({ ok: false, error: 'Error: ffmpeg failed', stderr: 'boom detail' }),
    })
    expect(result.ok).toBe(false)
    expect(result.error).toContain('ffmpeg failed')
    expect(result.error).toContain('boom detail')
  })

  it('reports success without an output file as a failure', async () => {
    const result = await exportSlidesVideo({
      pngsBase64: [solidPngBase64(0, 0, 255)],
      filePath: await outputPath(),
      checkFfmpeg: async () => true,
      runFfmpeg: async () => ({ ok: true }),
    })
    expect(result.ok).toBe(false)
    expect(result.error).toContain('did not produce')
  })
})
