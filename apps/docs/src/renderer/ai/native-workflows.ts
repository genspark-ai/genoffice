import type { AgentToolCall, ToolExecution } from '@genoffice/agent-core'
import type { DocsRuntimeClient } from '../runtime/docs-runtime'
import type { DocsErrorCode, DocsOperationResult } from '../../shared/operation-result'
import { DocsOperationError } from '../../shared/operation-result'
import type { AiDocWriter } from './doc-writer'
import { getDocsHost, getNativeAgentServices } from '../host/docs-host'
import { t } from '../i18n/locale'

type Execution = ToolExecution & Partial<Pick<DocsOperationResult, 'errorCode' | 'committed'>>
const refused = (
  summary: string,
  output: string,
  errorCode: DocsErrorCode = 'execution_failure',
): Execution => ({ summary, output, errorCode, isError: true, mutated: false, committed: false })

/** Native Agent orchestration consumes the same document interface as an external Harness. */
export function createNativeDocsWorkflows(
  document: DocsRuntimeClient,
  getWriter?: () => AiDocWriter | undefined,
) {
  const execute = (name: string, input: Record<string, unknown>, signal?: AbortSignal) =>
    document.execute({ id: crypto.randomUUID(), name, input }, signal)

  async function write(call: AgentToolCall, signal?: AbortSignal): Promise<Execution> {
    const summary = t('aiSumWriteDocument')
    const writer = getWriter?.()
    if (!writer) return refused(summary, 'The native writer is not available', 'unavailable')
    const plan = String(call.input.plan ?? '').trim()
    if (!plan) return refused(summary, 'plan must not be empty', 'invalid_input')
    const started = await execute(
      'start_document_draft',
      {
        ...(call.input.afterBlockIndex !== undefined
          ? { afterBlockIndex: call.input.afterBlockIndex }
          : {}),
        ...(call.input.replaceDocument !== undefined
          ? { replaceDocument: call.input.replaceDocument }
          : {}),
      },
      signal,
    )
    if (started.isError) return started
    const { draftId } = JSON.parse(started.output) as { draftId: string }
    const abort = new AbortController()
    const modelSignal = AbortSignal.any([
      document.signal,
      abort.signal,
      ...(signal ? [signal] : []),
    ])
    let sequence = 0
    let closed = false
    let stopped = false
    let latest = ''
    let failure: DocsOperationResult | undefined
    let updates = Promise.resolve()
    const update = (html: string) => {
      if (closed || document.signal.aborted || stopped) return
      latest = html
      const next = ++sequence
      updates = updates.then(async () => {
        if (failure) return
        const result = await execute('update_document_draft', { draftId, sequence: next, html })
        if (result.isError) {
          failure = result
          abort.abort()
        }
      })
    }
    try {
      const result = await writer.write(
        {
          plan,
          title: typeof call.input.title === 'string' ? call.input.title : undefined,
          context: typeof call.input.context === 'string' ? call.input.context : undefined,
          instruction: '',
        },
        {
          onProgress: update,
          onStopped: async () => {
            await updates
            if (failure) throw new DocsOperationError(failure.output, failure)
            document.signal.throwIfAborted()
            const result = await execute('stop_document_draft', { draftId })
            if (result.isError) throw new DocsOperationError(result.output, result)
            stopped = true
          },
        },
        modelSignal,
      )
      await updates
      if (failure) return failure
      if (!result.ok || !result.html)
        return refused(
          t('aiSumWriteDocumentFailed'),
          `The writer produced no accepted content (${result.error ?? 'no output'}); the document is unchanged.`,
        )
      if (stopped && latest !== result.html)
        return refused(
          t('aiSumWriteDocumentFailed'),
          'The writer changed its content after stopping',
          'stale_target',
        )
      if (!stopped) {
        update(result.html)
        await updates
      }
      if (failure) return failure
      const settled = await execute('finish_document_draft', {
        draftId,
        decision: result.truncated ? 'retain_partial' : 'finalize',
      })
      if (settled.isError) return settled
      const state = JSON.parse(settled.output) as {
        result: { blocks: number; validatedCharacters?: number }
      }
      return {
        ...settled,
        output:
          `Content written (${state.result.blocks} block(s), ${state.result.validatedCharacters ?? result.html.length} chars). Block indexes have changed; read fresh context.` +
          (result.truncated ? ' INCOMPLETE: the user chose to keep the partial output.' : ''),
        summary: result.truncated ? t('aiSumWriteDocumentPartial') : summary,
      }
    } finally {
      closed = true
      abort.abort()
      // A settled draft is idempotent; failed or abandoned work releases only its own preview.
      await execute('finish_document_draft', { draftId, decision: 'discard' })
    }
  }

  async function run(call: AgentToolCall, signal?: AbortSignal): Promise<Execution> {
    const modelSignal = AbortSignal.any([document.signal, ...(signal ? [signal] : [])])
    if (modelSignal.aborted) return refused(call.name, 'The operation was canceled', 'canceled')
    switch (call.name) {
      case 'web_search': {
        const query = String(call.input.query ?? '').trim()
        if (!query) return refused(t('aiSumWebSearch'), 'query must not be empty', 'invalid_input')
        const result = await getNativeAgentServices().webSearch(
          query,
          Number(call.input.maxResults) || 6,
        )
        if (result.method === 'error')
          return refused(
            t('aiSumWebSearch'),
            `web search failed: ${result.error ?? 'unknown error'}`,
          )
        const lines = result.answer ? [`Direct answer: ${result.answer}\n`] : []
        result.results.forEach((item, index) =>
          lines.push(`${index + 1}. ${item.title}\n   ${item.url}\n   ${item.snippet}`),
        )
        return {
          output: lines.join('\n') || '(no results)',
          mutated: false,
          committed: false,
          summary: t('aiSumWebSearchDone', { query, count: result.results.length }),
        }
      }
      case 'image_search': {
        const query = String(call.input.query ?? '').trim()
        if (!query)
          return refused(t('aiSumImageSearch'), 'query must not be empty', 'invalid_input')
        const result = await getNativeAgentServices().imageSearch(
          query,
          Number(call.input.maxResults) || 8,
        )
        if (result.method === 'error')
          return refused(
            t('aiSumImageSearch'),
            `image search failed: ${result.error ?? 'unknown error'}`,
          )
        return {
          output:
            result.images
              .map(
                (item, index) =>
                  `${index + 1}. ${item.title || '(untitled)'} [${item.width ?? '?'}x${item.height ?? '?'}]\n   ${item.imageUrl}`,
              )
              .join('\n') || '(no images)',
          mutated: false,
          committed: false,
          summary: t('aiSumImageSearchDone', { query, count: result.images.length }),
        }
      }
      case 'analyze_media': {
        const requirements = String(call.input.requirements ?? '').trim()
        if (!requirements)
          return refused(t('aiSumAnalyzeMedia'), 'requirements must not be empty', 'invalid_input')
        const urls = Array.isArray(call.input.mediaUrls)
          ? call.input.mediaUrls.map(String).filter(Boolean)
          : []
        if (call.input.blockIndex !== undefined) {
          const image = await execute(
            'read_document_media',
            { blockIndex: call.input.blockIndex },
            modelSignal,
          )
          if (image.isError) return image
          if (!image.media)
            return refused(
              t('aiSumAnalyzeMedia'),
              'The document did not return image bytes',
              'host_failure',
            )
          urls.unshift(`data:${image.media.mimeType};base64,${image.media.data}`)
        }
        if (!urls.length)
          return refused(t('aiSumAnalyzeMedia'), 'Give blockIndex or mediaUrls', 'invalid_input')
        const result = await getNativeAgentServices().analyzeMedia({
          mediaUrls: urls,
          requirements,
        })
        if (!result.text)
          return refused(t('aiSumAnalyzeMedia'), result.error ?? 'media analysis failed')
        return {
          output:
            result.text.length > 6000 ? result.text.slice(0, 6000) + '\n…(truncated)' : result.text,
          mutated: false,
          committed: false,
          summary: t('aiSumAnalyzeMediaDone'),
        }
      }
      case 'generate_image': {
        const prompt = String(call.input.prompt ?? '').trim()
        if (!prompt)
          return refused(t('aiSumGenerateImage'), 'prompt must not be empty', 'invalid_input')
        const caps = await execute('get_editor_capabilities', {}, modelSignal)
        if (caps.isError) return caps
        if (!JSON.parse(caps.output).editable)
          return refused(t('aiSumGenerateImage'), 'The document is read-only', 'permission_denied')
        const aspectRatio = String(call.input.aspectRatio ?? '').trim()
        const generated = await getNativeAgentServices().aiGenerateImage({
          prompt,
          ...(aspectRatio ? { aspectRatio } : {}),
        })
        if (modelSignal.aborted)
          return refused(
            t('aiSumGenerateImage'),
            'The operation was canceled; no image was inserted',
            'canceled',
          )
        if (!generated.url)
          return refused(t('aiSumGenerateImage'), generated.error ?? 'image generation failed')
        const image = await getDocsHost().fetchImage(generated.url)
        if (!image)
          return refused(
            t('aiSumGenerateImage'),
            'The generated image could not be read',
            'host_failure',
          )
        const inserted = await execute(
          'insert_image',
          {
            url: `data:${image.mime};base64,${image.base64}`,
            maxWidthPx: Number(call.input.maxWidthPx) || 480,
            label: 'Image (AI)',
          },
          modelSignal,
        )
        return inserted.isError ? inserted : { ...inserted, summary: t('aiSumInsertedGenImage') }
      }
      case 'write_document':
        return write(call, signal)
      default:
        return refused(call.name, 'Unknown native document workflow', 'unavailable')
    }
  }

  return {
    execute(call: AgentToolCall, signal?: AbortSignal): Promise<Execution> {
      const cancellation =
        call.name === 'write_document'
          ? document.signal
          : AbortSignal.any([document.signal, ...(signal ? [signal] : [])])
      return new Promise((resolve) => {
        const canceled = () =>
          resolve(refused(call.name, 'The document operation was canceled', 'canceled'))
        if (cancellation.aborted) {
          canceled()
          return
        }
        cancellation.addEventListener('abort', canceled, { once: true })
        void run(call, signal)
          .then(resolve, (error) =>
            resolve(
              error instanceof DocsOperationError
                ? error.result
                : refused(call.name, error instanceof Error ? error.message : String(error)),
            ),
          )
          .finally(() => cancellation.removeEventListener('abort', canceled))
      })
    },
  }
}
