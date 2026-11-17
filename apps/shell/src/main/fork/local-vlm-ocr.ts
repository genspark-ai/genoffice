import { ensureLocalVlmEndpoint } from '@genoffice/ai-provider/local-vlm-fallback'
import { LOCAL_VLM_SAMPLING_DEFAULTS } from '@genoffice/ai-provider/local-vlm-meta'
import { OCR_PAGE_MARKER_PREFIX } from '@genoffice/ai-provider/agy-ocr'
import type { OcrRecognizeInput } from '../document-memory/agy-ocr-job'

const LOCAL_VLM_OCR_MAX_PAGES_PER_CALL = 2

export interface LocalVlmOcrDeps {
  ensureEndpoint?: (signal?: AbortSignal) => Promise<{ baseUrl: string; model: string }>
  fetchFn?: typeof fetch
}

function buildLocalVlmPagePrompt(pages: readonly number[]): string {
  const markerExample = `${OCR_PAGE_MARKER_PREFIX}${pages[0] ?? 1} ===`
  return [
    'Transcribe all visible text on each scanned page image exactly, preserving Vietnamese diacritics and keeping table rows on separate lines.',
    `Output pages in order. Start each page with a line \`${OCR_PAGE_MARKER_PREFIX}n ===\` (for example \`${markerExample}\`), followed by numbered text lines like \`001| text\`.`,
    'No commentary, no markdown code fences, no translation, no summary.',
    'If a page has no readable text, output only its marker line.',
    `Pages in image order: ${pages.join(', ')}.`,
  ].join('\n')
}

function normalizeSinglePageFallback(raw: string, pageNumber: number): string {
  const marker = `${OCR_PAGE_MARKER_PREFIX}${pageNumber} ===`
  if (raw.includes(OCR_PAGE_MARKER_PREFIX)) return raw
  const lines = raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, i) => `${String(i + 1).padStart(3, '0')}| ${line}`)
  return `${marker}\n${lines.join('\n')}`
}

/**
 * Recognizes scanned PDF page JPEGs using the local LFM3-VL-2B llama-server endpoint.
 * Splits batches into at most 2 pages per request to fit comfortably inside the local context window.
 */
export async function recognizeWithLocalVlm(
  input: OcrRecognizeInput,
  deps: LocalVlmOcrDeps = {},
): Promise<{
  text: string
  usage?: { inputTokens?: number; outputTokens?: number; thinkingTokens?: number }
}> {
  const ensure = deps.ensureEndpoint ?? ensureLocalVlmEndpoint
  const fetchFn = deps.fetchFn ?? fetch
  const endpoint = await ensure(input.signal)
  const chunks: string[] = []
  let inputTokens = 0
  let outputTokens = 0

  for (let i = 0; i < input.pages.length; i += LOCAL_VLM_OCR_MAX_PAGES_PER_CALL) {
    const subPages = input.pages.slice(i, i + LOCAL_VLM_OCR_MAX_PAGES_PER_CALL)
    const subImages = input.images.slice(i, i + LOCAL_VLM_OCR_MAX_PAGES_PER_CALL)
    const content: unknown[] = [
      { type: 'text', text: buildLocalVlmPagePrompt(subPages) },
      ...subImages.map((bytes) => ({
        type: 'image_url',
        image_url: { url: `data:image/jpeg;base64,${Buffer.from(bytes).toString('base64')}` },
      })),
    ]

    const response = await fetchFn(`${endpoint.baseUrl}/chat/completions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: endpoint.model,
        messages: [{ role: 'user', content }],
        max_tokens: 4096,
        ...LOCAL_VLM_SAMPLING_DEFAULTS,
      }),
      signal: input.signal,
    })
    if (!response.ok) {
      const detail = await response.text().catch(() => '')
      throw new Error(
        `Local LFM3-VL-2B OCR returned HTTP ${response.status}${detail ? `: ${detail.slice(0, 200)}` : ''}`,
      )
    }
    const json = (await response.json()) as {
      choices?: Array<{ message?: { content?: unknown } }>
      usage?: { prompt_tokens?: number; completion_tokens?: number }
    }
    const rawContent = json.choices?.[0]?.message?.content
    const rawText =
      typeof rawContent === 'string'
        ? rawContent
        : Array.isArray(rawContent)
          ? rawContent
              .map((part) =>
                part &&
                typeof part === 'object' &&
                typeof (part as { text?: unknown }).text === 'string'
                  ? (part as { text: string }).text
                  : '',
              )
              .join('')
          : ''
    if (!rawText.trim()) {
      throw new Error('Local LFM3-VL-2B OCR returned an empty response')
    }
    const normalized =
      subPages.length === 1
        ? normalizeSinglePageFallback(rawText.trim(), subPages[0]!)
        : rawText.trim()
    chunks.push(normalized)
    if (typeof json.usage?.prompt_tokens === 'number') inputTokens += json.usage.prompt_tokens
    if (typeof json.usage?.completion_tokens === 'number')
      outputTokens += json.usage.completion_tokens
  }

  return {
    text: chunks.join('\n'),
    ...(inputTokens > 0 || outputTokens > 0 ? { usage: { inputTokens, outputTokens } } : {}),
  }
}
