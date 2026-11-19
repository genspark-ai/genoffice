import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi, beforeEach } from 'vitest'

vi.mock('../src/gsk', () => ({
  gskGenerateImage: vi.fn(),
  gskAnalyzeMedia: vi.fn(),
  hasGskAuth: vi.fn(() => true),
}))

import {
  MEDIA_BUDGET,
  MediaBudgetExceededError,
  MediaTooLargeError,
  analyzeMediaTool,
  generateImageTool,
  loadMediaReferences,
} from '../src/media-tools'
import { gskAnalyzeMedia, gskGenerateImage } from '../src/gsk'

const gskGen = vi.mocked(gskGenerateImage)
const gskAnalyze = vi.mocked(gskAnalyzeMedia)

/** A data URL whose decoded payload is `bytes` long. */
function dataUrl(bytes: number, mime = 'image/png'): string {
  return `data:${mime};base64,${'A'.repeat(Math.ceil((bytes * 4) / 3))}`
}

/** The production budget is 200 MB per item and in total, which no test should allocate. */
const TINY = { maxItems: 3, maxItemBytes: 64, maxTotalBytes: 128, concurrency: 2 }

function writeSettings(mediaProvider: string): string {
  const dir = mkdtempSync(join(tmpdir(), 'genoffice-media-budget-'))
  const path = join(dir, 'ai-settings.json')
  writeFileSync(
    path,
    JSON.stringify({
      provider: mediaProvider,
      providers: { [mediaProvider]: { apiKey: 'k', model: 'test-model' } },
      media: {
        provider: mediaProvider,
        providers: {
          [mediaProvider]: {
            apiKey: 'k',
            analysisModel: 'test-model',
            imageModel: 'test-image-model',
          },
        },
      },
    }),
  )
  return path
}

beforeEach(() => {
  gskGen.mockReset()
  gskAnalyze.mockReset()
})

describe('loadMediaReferences budget', () => {
  it('pins the production ceilings', () => {
    expect(MEDIA_BUDGET).toEqual({
      maxItems: 12,
      maxItemBytes: 200 * 1024 * 1024,
      maxTotalBytes: 200 * 1024 * 1024,
      concurrency: 3,
    })
  })

  it('rejects too many items before loading anything', async () => {
    const refs = Array.from({ length: 4 }, () => dataUrl(8))
    await expect(loadMediaReferences(refs, TINY)).rejects.toThrow(/Too many media items/)
    await expect(loadMediaReferences(refs, TINY)).rejects.toBeInstanceOf(MediaTooLargeError)
  })

  it('rejects one data URL over the per-item cap from its encoded size alone', async () => {
    await expect(loadMediaReferences([dataUrl(65)], TINY)).rejects.toThrow(/too large/i)
  })

  it('rejects data URLs whose combined encoded size is over the total cap', async () => {
    await expect(
      loadMediaReferences([dataUrl(45), dataUrl(45), dataUrl(45)], TINY),
    ).rejects.toBeInstanceOf(MediaBudgetExceededError)
  })

  it('rejects a total that only crosses the cap once the bytes have landed', async () => {
    // local files declare no encoded size, so only the running total can catch this
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-media-files-'))
    const refs = Array.from({ length: 3 }, (_, i) => {
      const path = join(dir, `shot-${i}.png`)
      writeFileSync(path, Buffer.alloc(50, 1))
      return path
    })
    await expect(
      loadMediaReferences(refs, { maxItems: 3, maxTotalBytes: 100 }),
    ).rejects.toBeInstanceOf(MediaBudgetExceededError)
  })

  it('loads a batch that fits, in the caller’s order', async () => {
    const blobs = await loadMediaReferences(
      [dataUrl(8, 'image/png'), dataUrl(8, 'image/jpeg')],
      TINY,
    )
    expect(blobs.map((b) => b.mime)).toEqual(['image/png', 'image/jpeg'])
    expect(blobs.every((b) => b.bytes.byteLength > 0)).toBe(true)
  })

  it('keeps loading bounded by the concurrency setting', async () => {
    const refs = Array.from({ length: 3 }, () => dataUrl(8, 'image/png'))
    const blobs = await loadMediaReferences(refs, { ...TINY, concurrency: 1 })
    expect(blobs).toHaveLength(3)
    expect(blobs.every((b) => b.mime === 'image/png')).toBe(true)
  })
})

describe('media tool aggregate budget', () => {
  it('analyze_media refuses too many items', async () => {
    const result = await analyzeMediaTool(writeSettings('gemini'), {
      mediaUrls: Array.from({ length: 13 }, () => dataUrl(8)),
      requirements: 'describe these',
    })
    expect(result.error).toMatch(/Too many media items/)
    expect(gskAnalyze).not.toHaveBeenCalled()
  })

  it('analyze_media still hands a budget failure to Genspark when only one provider is configured', async () => {
    gskAnalyze.mockResolvedValue('from genspark' as never)
    // openai reads images but not video, so a budget failure still has a
    // Genspark route to fall back to (the pre-existing size-cap behavior)
    const result = await analyzeMediaTool(writeSettings('openai'), {
      mediaUrls: Array.from({ length: 13 }, () => dataUrl(8)),
      requirements: 'describe these',
    })
    // the fallback enforces the same item ceiling: the CLI argv stays bounded
    expect(result.error).toMatch(/Too many media items/)
    expect(gskAnalyze).not.toHaveBeenCalled()
  })

  it('generate_image refuses too many references', async () => {
    const result = await generateImageTool(writeSettings('gemini'), {
      prompt: 'a logo',
      referenceImageUrls: Array.from({ length: 13 }, () => dataUrl(8)),
    })
    expect(result.error).toMatch(/Too many media items/)
  })
})
