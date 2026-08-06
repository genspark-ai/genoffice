import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { httpBodyDetail } from './http-error'
import { OPENROUTER_BASE_URL } from './providers'
import type { AiProviderConfig, AiProviderId } from './types'

export interface StreamCallbacks {
  onDelta: (text: string) => void
  onToolCall: (call: AgentToolCall) => void
  onStopReason?: (reason: string) => void
  onActivity?: () => void
  signal: AbortSignal
}

export class AiCreditsError extends Error {
  constructor(notice: string) {
    super(notice)
    this.name = 'AiCreditsError'
  }
}

export async function* sseLines(
  body: NodeJS.ReadableStream | ReadableStream<Uint8Array>,
  onBytes?: () => void,
): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  const reader = (body as ReadableStream<Uint8Array>).getReader()
  let buffer = ''
  while (true) {
    const { done, value } = await reader.read()
    if (done) break
    onBytes?.()
    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop() ?? ''
    for (const line of lines) yield line
  }
  if (buffer) yield buffer
}

function openAiMessages(system: string, messages: AgentMessage[]): unknown[] {
  const out: unknown[] = [{ role: 'system', content: system }]
  for (const m of messages) {
    if (m.role === 'user') {
      if (m.images?.length) {
        out.push({
          role: 'user',
          content: [
            ...(m.text ? [{ type: 'text', text: m.text }] : []),
            ...m.images.map((img) => ({
              type: 'image_url',
              image_url: { url: `data:${img.mime};base64,${img.base64}` },
            })),
          ],
        })
      } else out.push({ role: 'user', content: m.text })
    } else if (m.role === 'assistant') {
      out.push({
        role: 'assistant',
        content: m.text || null,
        ...(m.toolCalls?.length
          ? { tool_calls: m.toolCalls.map((c) => ({ id: c.id, type: 'function', function: { name: c.name, arguments: JSON.stringify(c.input) } })) }
          : {}),
      })
    } else {
      for (const r of m.results) out.push({ role: 'tool', tool_call_id: r.id, content: r.output })
    }
  }
  return out
}

function parseArgs(raw: string): Record<string, unknown> {
  if (!raw.trim()) return {}
  try { return JSON.parse(raw) as Record<string, unknown> } catch { return { _raw: raw } }
}

async function streamOpenAiCompatible(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const response = await fetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    signal: cb.signal,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
    },
    body: JSON.stringify({
      model: config.model,
      max_tokens: maxTokens,
      messages: openAiMessages(system, messages),
      ...(tools.length ? { tools: tools.map((t) => ({ type: 'function', function: { name: t.name, description: t.description, parameters: t.inputSchema } })) } : {}),
      temperature: 0.3,
      stream: true,
    }),
  })
  cb.onActivity?.()
  if (!response.ok || !response.body) throw new Error(`HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)

  const pending = new Map<number, { id: string; name: string; args: string }>()
  let finishReason: string | undefined
  for await (const line of sseLines(response.body, () => cb.onActivity?.())) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload || payload === '[DONE]') continue
    const event = JSON.parse(payload) as {
      error?: { message?: string }
      choices?: Array<{ delta?: { content?: string; tool_calls?: Array<{ index: number; id?: string; function?: { name?: string; arguments?: string } }> }; finish_reason?: string | null }>
    }
    if (event.error) throw new Error(event.error.message ?? 'OpenRouter stream error')
    const choice = event.choices?.[0]
    if (!choice) continue
    if (choice.delta?.content) cb.onDelta(choice.delta.content)
    for (const tc of choice.delta?.tool_calls ?? []) {
      const p = pending.get(tc.index) ?? { id: tc.id ?? crypto.randomUUID(), name: '', args: '' }
      if (tc.id) p.id = tc.id
      if (tc.function?.name) p.name += tc.function.name
      if (tc.function?.arguments) p.args += tc.function.arguments
      pending.set(tc.index, p)
    }
    if (choice.finish_reason) finishReason = choice.finish_reason
  }
  for (const [, p] of [...pending.entries()].sort(([a], [b]) => a - b)) {
    if (p.name) cb.onToolCall({ id: p.id, name: p.name, input: parseArgs(p.args) })
  }
  if (finishReason === 'length') cb.onStopReason?.('max_tokens')
  else if (finishReason) cb.onStopReason?.(finishReason)
}

const OPENAI_COMPATIBLE_BASE_URLS: Partial<Record<AiProviderId, string>> = {
  openrouter: OPENROUTER_BASE_URL,
  deepseek: 'https://api.deepseek.com/v1',
  openai: 'https://api.openai.com/v1',
}

export async function streamForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  if (provider === 'custom') {
    if (!config.baseUrl) throw new Error('A custom provider requires a Base URL')
    return streamOpenAiCompatible(config.baseUrl, config, system, messages, tools, maxTokens, cb)
  }
  const baseUrl = OPENAI_COMPATIBLE_BASE_URLS[provider]
  if (!baseUrl) throw new Error(`Provider ${provider} is not enabled in this OpenRouter-focused build`)
  return streamOpenAiCompatible(baseUrl, config, system, messages, tools, maxTokens, cb)
}
