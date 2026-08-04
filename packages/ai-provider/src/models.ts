import type { AiProviderConfig, AiProviderId } from './types'
import { AI_PROVIDERS } from './providers'

export interface ModelListEntry {
  id: string
  label?: string
  /** zero-cost on this provider (OpenRouter :free / zero price, Gemini free-tier routes) */
  free?: boolean
}

/**
 * Known-free Gemini API routes. Google's models endpoint exposes no pricing, so
 * this curated set (kept in sync with the pricing page) drives the "(free)"
 * badge and the free-only filter; it changes when Google reshuffles tiers.
 */
export const GEMINI_FREE_MODELS = new Set([
  'gemini-2.5-flash',
  'gemini-2.5-flash-lite',
  'gemini-2.5-flash-live-preview',
  'gemini-3-flash-preview',
  'gemini-3.1-flash-lite',
  'gemini-3.5-flash',
  'gemini-3.6-flash',
])

/** OpenRouter free router: picks a currently-free model that fits the request */
export const OPENROUTER_FREE_ROUTER = 'openrouter/free'

async function fetchJson(url: string, init?: RequestInit): Promise<unknown> {
  const response = await fetch(url, init)
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${(await response.text()).slice(0, 500)}`)
  }
  return response.json()
}

function entries(ids: string[], freeIds?: Set<string>): ModelListEntry[] {
  return ids.map((id) => ({ id, free: freeIds?.has(id) ?? false }))
}

async function listOpenRouter(freeOnly: boolean): Promise<ModelListEntry[]> {
  const json = (await fetchJson('https://openrouter.ai/api/v1/models')) as {
    data?: Array<{ id: string; name?: string; pricing?: { prompt?: string; completion?: string } }>
  }
  const out: ModelListEntry[] = []
  for (const m of json.data ?? []) {
    const zero =
      m.pricing?.prompt === '0' && m.pricing?.completion === '0'
    const free = zero || m.id.endsWith(':free')
    if (freeOnly && !free) continue
    out.push({ id: m.id, label: m.name, free })
  }
  out.sort((a, b) => (a.free === b.free ? a.id.localeCompare(b.id) : a.free ? -1 : 1))
  if (freeOnly) {
    out.unshift({ id: OPENROUTER_FREE_ROUTER, label: 'Free models router (auto-picks a free model)' })
  }
  return out
}

async function listGemini(apiKey: string, freeOnly: boolean): Promise<ModelListEntry[]> {
  if (!apiKey) throw new Error('Enter a Gemini API key first')
  const json = (await fetchJson(
    `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(apiKey)}`,
  )) as { models?: Array<{ name?: string; displayName?: string; supportedGenerationMethods?: string[] }> }
  const out: ModelListEntry[] = []
  for (const m of json.models ?? []) {
    const id = (m.name ?? '').replace(/^models\//, '')
    if (!id) continue
    if (!m.supportedGenerationMethods?.includes('generateContent')) continue
    if (/(image|tts|embedding|audio|veo|computer-use)/.test(id)) continue
    const free = GEMINI_FREE_MODELS.has(id)
    if (freeOnly && !free) continue
    out.push({ id, label: m.displayName ?? id, free })
  }
  out.sort((a, b) => (a.free === b.free ? a.id.localeCompare(b.id) : a.free ? -1 : 1))
  return out
}

async function listAnthropic(apiKey: string): Promise<ModelListEntry[]> {
  if (!apiKey) throw new Error('Enter a Claude API key first')
  const json = (await fetchJson('https://api.anthropic.com/v1/models', {
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
  })) as { data?: Array<{ id?: string; display_name?: string }> }
  const out: ModelListEntry[] = []
  for (const m of json.data ?? []) {
    if (!m.id) continue
    if (/(tool|embedding)/.test(m.id)) continue
    out.push({ id: m.id, label: m.display_name, free: false })
  }
  return out
}

async function listOpenAi(apiKey: string): Promise<ModelListEntry[]> {
  if (!apiKey) throw new Error('Enter an OpenAI API key first')
  const json = (await fetchJson('https://api.openai.com/v1/models', {
    headers: { Authorization: `Bearer ${apiKey}` },
  })) as { data?: Array<{ id?: string }> }
  const out: ModelListEntry[] = []
  const nonChat = /(whisper|tts|dall-e|embedding|moderation|audio|realtime|gpt-image|batch|search|rerank|computer-use|transcription|speech)/i
  for (const m of json.data ?? []) {
    if (!m.id) continue
    if (!/^(gpt-|o1|o2|o3|o4|o5|chatgpt-|gpt-oss)/.test(m.id)) continue
    if (nonChat.test(m.id)) continue
    out.push({ id: m.id, free: false })
  }
  out.sort((a, b) => a.id.localeCompare(b.id))
  return out
}

async function listDeepSeek(apiKey: string): Promise<ModelListEntry[]> {
  if (!apiKey) throw new Error('Enter a DeepSeek API key first')
  const json = (await fetchJson('https://api.deepseek.com/models', {
    headers: { Authorization: `Bearer ${apiKey}` },
  })) as { data?: Array<{ id?: string }> }
  return (json.data ?? []).filter((m) => m.id).map((m) => ({ id: m.id as string, free: false }))
}

/**
 * Fetch the live model list a provider exposes. OpenRouter's catalog needs no
 * key; the rest use the (already stored) provider key. `freeOnly` keeps zero-cost
 * models for OpenRouter and free-tier Gemini routes.
 */
export async function listModelsForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  opts: { freeOnly?: boolean } = {},
): Promise<ModelListEntry[]> {
  switch (provider) {
    case 'openrouter':
      return listOpenRouter(opts.freeOnly ?? false)
    case 'gemini':
      return listGemini(config.apiKey, opts.freeOnly ?? false)
    case 'anthropic':
      return listAnthropic(config.apiKey)
    case 'openai':
      return listOpenAi(config.apiKey)
    case 'deepseek':
      return listDeepSeek(config.apiKey)
    case 'genspark': {
      const meta = AI_PROVIDERS.find((p) => p.id === 'genspark')
      return entries(meta?.models ?? [], new Set())
    }
    default:
      return []
  }
}
