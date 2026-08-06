import { httpBodyDetail } from './http-error'
import { GENSPARK_LLM_BASE_URLS, gensparkAttributionHeaders } from './providers'
import type { AiChatResponse, AiProviderConfig, AiProviderId } from './types'
import { AI_CHAT_RESPONSE_TIMEOUT_MS, createStreamWatchdog, type StreamWatchdog } from './watchdog'

async function chatAnthropic(
  wd: StreamWatchdog,
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = 'https://api.anthropic.com',
): Promise<AiChatResponse> {
  const response = await fetch(`${baseUrl.replace(/\/$/, '')}/v1/messages`, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.apiKey,
      'anthropic-version': '2023-06-01',
      // Fetch in the Electron main process goes through Chromium's network stack; this header avoids 403.
      'anthropic-dangerous-direct-browser-access': 'true',
      ...gensparkAttributionHeaders(baseUrl),
    },
    body: JSON.stringify({
      model: config.model,
      max_tokens: 8192,
      system,
      messages: [{ role: 'user', content: user }],
    }),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `Claude HTTP ${response.status}: ${httpBodyDetail(await response.text())}`,
    }
  }
  const json = (await response.json()) as { content?: Array<{ type: string; text?: string }> }
  const content = json.content
    ?.filter((c) => c.type === 'text')
    .map((c) => c.text ?? '')
    .join('')
  if (!content) return { ok: false, error: 'Claude returned an empty response' }
  return { ok: true, content }
}

async function chatGemini(
  wd: StreamWatchdog,
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = 'https://generativelanguage.googleapis.com/v1beta',
): Promise<AiChatResponse> {
  const url = `${baseUrl.replace(/\/$/, '')}/models/${config.model}:generateContent`
  const response = await fetch(url, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': config.apiKey,
      ...gensparkAttributionHeaders(baseUrl),
    },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: [{ role: 'user', parts: [{ text: user }] }],
      generationConfig: { temperature: 0.3 },
    }),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `Gemini HTTP ${response.status}: ${httpBodyDetail(await response.text())}`,
    }
  }
  const json = (await response.json()) as {
    candidates?: Array<{ content?: { parts?: Array<{ text?: string }> } }>
  }
  const content = json.candidates?.[0]?.content?.parts?.map((p) => p.text ?? '').join('')
  if (!content) return { ok: false, error: 'Gemini returned an empty response' }
  return { ok: true, content }
}

async function chatOpenAiCompatible(
  wd: StreamWatchdog,
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  user: string,
): Promise<AiChatResponse> {
  const response = await fetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
      ...gensparkAttributionHeaders(baseUrl),
    },
    body: JSON.stringify({
      model: config.model,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user },
      ],
      temperature: 0.3,
    }),
  })
  wd.touch()
  if (!response.ok) {
    return { ok: false, error: `HTTP ${response.status}: ${httpBodyDetail(await response.text())}` }
  }
  const json = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> }
  const content = json.choices?.[0]?.message?.content
  if (!content) return { ok: false, error: 'AI returned an empty response' }
  return { ok: true, content }
}

const OPENAI_COMPATIBLE_BASE_URLS: Partial<Record<AiProviderId, string>> = {
  deepseek: 'https://api.deepseek.com/v1',
  openai: 'https://api.openai.com/v1',
  openrouter: 'https://openrouter.ai/api/v1',
}

/** route a one-shot (non-streaming, non-tool-calling) chat call by provider id */
export async function chatForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  user: string,
  signal?: AbortSignal,
): Promise<AiChatResponse> {
  // non-streaming: the server generates the full answer before the headers arrive,
  // so the connect phase gets the long budget; the body read then gets the idle budget
  const wd = createStreamWatchdog(signal, AI_CHAT_RESPONSE_TIMEOUT_MS)
  return wd.guard(() => {
    switch (provider) {
      case 'genspark':
        if (config.model.startsWith('claude')) {
          return chatAnthropic(wd, config, system, user, GENSPARK_LLM_BASE_URLS.anthropic)
        }
        if (config.model.startsWith('gemini')) {
          return chatGemini(wd, config, system, user, GENSPARK_LLM_BASE_URLS.gemini)
        }
        return chatOpenAiCompatible(wd, GENSPARK_LLM_BASE_URLS.openai, config, system, user)
      case 'anthropic':
        return chatAnthropic(wd, config, system, user)
      case 'gemini':
        return chatGemini(wd, config, system, user)
      case 'deepseek':
      case 'openai':
      case 'openrouter':
        return chatOpenAiCompatible(
          wd,
          OPENAI_COMPATIBLE_BASE_URLS[provider]!,
          config,
          system,
          user,
        )
      case 'custom':
        if (!config.baseUrl)
          return Promise.resolve({
            ok: false as const,
            error: 'A custom provider requires a Base URL',
          })
        return chatOpenAiCompatible(wd, config.baseUrl, config, system, user)
      default:
        return Promise.resolve({ ok: false as const, error: `Unknown provider: ${provider}` })
    }
  })
}
