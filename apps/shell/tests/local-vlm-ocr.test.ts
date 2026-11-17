import { describe, expect, it, vi } from 'vitest'
import { systemTimezoneOffset } from '@genoffice/ai-provider/agy-ocr'
import {
  AgyOcrJob,
  type OcrCandidate,
  type OcrJobHost,
} from '../src/main/document-memory/agy-ocr-job'
import { OcrStateStore } from '../src/main/document-memory/agy-ocr-state'
import { recognizeWithLocalVlm } from '../src/main/fork/local-vlm-ocr'
import {
  AGY_OCR_SETTINGS_KEY,
  DEFAULT_AGY_OCR_SETTINGS,
  agyOcrSettingsFrom,
  mergeAgyOcrSettings,
} from '../src/shared/fork/agy-ocr'

describe('recognizeWithLocalVlm', () => {
  it('splits multi-page batches to max 2 pages per request and normalizes single-page output', async () => {
    const calls: number[] = []
    const res = await recognizeWithLocalVlm(
      {
        model: 'LFM3-VL-2B-Q4_K_M',
        pages: [1, 2, 3],
        images: [Buffer.from('a'), Buffer.from('b'), Buffer.from('c')],
        signal: new AbortController().signal,
      },
      {
        ensureEndpoint: async () => ({
          baseUrl: 'http://127.0.0.1:44444/v1',
          model: 'LFM3-VL-2B-Q4_K_M',
        }),
        fetchFn: (async (_url, init) => {
          const body = JSON.parse(String(init?.body)) as {
            messages: Array<{ content: Array<{ type: string }> }>
          }
          const imageCount = body.messages[0]!.content.filter((c) => c.type === 'image_url').length
          calls.push(imageCount)
          if (imageCount === 2) {
            return new Response(
              JSON.stringify({
                choices: [
                  {
                    message: {
                      content: '=== PAGE 1 ===\n001| Xin chao\n=== PAGE 2 ===\n001| Trang hai',
                    },
                  },
                ],
                usage: { prompt_tokens: 100, completion_tokens: 20 },
              }),
              { status: 200 },
            )
          }
          // Single-page reply omitting === PAGE 3 === marker still wraps cleanly.
          return new Response(
            JSON.stringify({
              choices: [{ message: { content: 'Trang ba khong dau marker' } }],
              usage: { prompt_tokens: 50, completion_tokens: 10 },
            }),
            { status: 200 },
          )
        }) as typeof fetch,
      },
    )

    expect(calls).toEqual([2, 1])
    expect(res.text).toContain('=== PAGE 1 ===')
    expect(res.text).toContain('=== PAGE 2 ===')
    expect(res.text).toContain('=== PAGE 3 ===')
    expect(res.text).toContain('001| Trang ba khong dau marker')
    expect(res.usage).toEqual({ inputTokens: 150, outputTokens: 30 })
  })
})

describe('AgyOcrJob with local-vlm escalation fallback', () => {
  it('merges escalationBackend setting while preserving legacy defaults', () => {
    expect(agyOcrSettingsFrom({}).escalationBackend).toBeUndefined()
    expect(
      mergeAgyOcrSettings(DEFAULT_AGY_OCR_SETTINGS, { escalationBackend: 'local-vlm' })
        .escalationBackend,
    ).toBe('local-vlm')
    expect(
      agyOcrSettingsFrom({ [AGY_OCR_SETTINGS_KEY]: { escalationBackend: 'bogus' } })
        .escalationBackend,
    ).toBeUndefined()
  })

  it('skips Antigravity quota checks and runs local-vlm recognizer when agy is unavailable', async () => {
    let storedJson = ''
    const saved: Array<{ pages: number[]; text: string[] }> = []
    const readUsage = vi.fn(async () => null)
    const recognizeAgy = vi.fn(async () => ({ text: '=== PAGE 1 ===\n001| from agy' }))
    const recognizeLocal = vi.fn(async ({ pages }: { pages: number[] }) => ({
      text: pages.map((p) => `=== PAGE ${p} ===\n001| Local OCR page ${p}`).join('\n'),
      usage: { inputTokens: 40, outputTokens: 12 },
    }))

    const candidate: OcrCandidate = {
      path: 'C:\\docs\\scan.pdf',
      sizeBytes: 1024,
      lastOpenedAt: 1000,
      mtimeMs: 1000,
      pagesDone: 0,
      totalPages: 1,
    }

    const host: OcrJobHost = {
      isEnabled: () => true,
      candidates: () => [candidate],
      documentById: () => ({ path: candidate.path }),
      pagesDone: () => [],
      render: async () => ({
        ok: true,
        hash: 'abc',
        mtimeMs: 1000,
        sizeBytes: 1024,
        totalPages: 1,
        pages: [
          {
            page: 1,
            jpeg: new Uint8Array([1, 2, 3]),
            width: 100,
            height: 100,
            source: 'rendered',
          },
        ],
      }),
      savePages: async (_path, _meta, pages) => {
        saved.push({ pages: pages.map((p) => p.page), text: pages.map((p) => p.text) })
        return { ok: true }
      },
      reindex: () => undefined,
      reindexNow: async () => undefined,
    }

    const state = new OcrStateStore(
      'memory://agy-ocr-state.json',
      () => 1_700_000_000_000,
      systemTimezoneOffset,
      {
        read: () => (storedJson ? storedJson : undefined),
        write: (_p, data) => {
          storedJson = data
        },
      },
    )

    const job = new AgyOcrJob({
      settings: () => ({ ...DEFAULT_AGY_OCR_SETTINGS, onlyOnAC: false, onlyWhenIdle: false }),
      host,
      state,
      recognize: recognizeAgy,
      recognizeLocal,
      resolveBackend: () => 'local-vlm',
      readUsage,
      policy: () => ({ paused: false, onBattery: false }),
      idleSeconds: () => 999,
      now: () => 1_700_000_000_000,
      timezoneOffset: systemTimezoneOffset,
      every: () => () => undefined,
    })

    expect(job.status().activeBackend).toBe('local-vlm')
    expect(job.status().activity).toEqual({ kind: 'working' })

    await job.tick()

    expect(readUsage).not.toHaveBeenCalled()
    expect(recognizeAgy).not.toHaveBeenCalled()
    expect(recognizeLocal).toHaveBeenCalledTimes(1)
    expect(saved).toEqual([{ pages: [1], text: ['Local OCR page 1'] }])
  })
})
