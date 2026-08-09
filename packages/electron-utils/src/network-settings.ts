/**
 * The user's network settings — outbound proxy and Tavily key — pushed into
 * every consumer that has to be told about them separately.
 *
 * Docs, sheets and slides all read `ai-settings.json` for the model selection,
 * but the network half of that file used to be applied by the shell alone. A
 * standalone sheets or slides window therefore ignored the configured proxy
 * entirely, and the gsk CLI children it spawns — PPT generation, image
 * generation, media analysis — dialled genspark.ai direct on the one kind of
 * network where nothing but the proxy can reach it.
 */
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

import { session } from 'electron'

import {
  defaultAiSettings,
  normalizeProxyUrl,
  resolveAiSettings,
  type AiSettings,
  type LegacyAiSettings,
} from '@genoffice/ai-provider'
import { setGskProxyUrl, setTavilyApiKey } from '@genoffice/ai-search'

const SETTINGS_FILE = 'ai-settings.json'

/** The network-facing subset of the settings file; everything else is the provider's business. */
export interface NetworkSettings {
  proxyUrl: string
  tavilyApiKey: string
}

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* corrupted state file: fall back to defaults */
  }
  return fallback
}

/** Read the network fields out of `<userData>/ai-settings.json`, migrations applied. */
export function readNetworkSettings(userDataPath: string): NetworkSettings {
  const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(
    join(userDataPath, SETTINGS_FILE),
    {},
  )
  const settings = resolveAiSettings(stored, defaultAiSettings())
  return { proxyUrl: settings.proxyUrl ?? '', tavilyApiKey: settings.tavilyApiKey ?? '' }
}

/** last proxy handed to undici/Chromium, so a no-op save does not churn them */
let appliedProxyUrl: string | null = null

/**
 * Route outbound traffic through the user's proxy.
 *
 * Three consumers need telling separately, which is why this is not one call:
 * main-process `fetch` runs on undici and ignores the system proxy entirely;
 * Chromium sessions carry the renderer, sign-in window and the agent browser;
 * and the gsk CLI is a child process that only sees environment variables.
 *
 * An empty url restores direct connections, so clearing the field in the
 * dialog actually turns the proxy off rather than leaving the old one wired.
 */
export async function applyProxy(rawUrl: string): Promise<void> {
  const proxyUrl = normalizeProxyUrl(rawUrl)
  if (proxyUrl === appliedProxyUrl) return
  appliedProxyUrl = proxyUrl
  setGskProxyUrl(proxyUrl)
  try {
    const { ProxyAgent, getGlobalDispatcher, setGlobalDispatcher, Agent } = await import('undici')
    if (proxyUrl) {
      setGlobalDispatcher(new ProxyAgent(proxyUrl))
    } else if (getGlobalDispatcher() instanceof ProxyAgent) {
      setGlobalDispatcher(new Agent())
    }
  } catch (err) {
    console.warn('[proxy] failed to set undici dispatcher:', err)
  }
  try {
    // proxyRules '' clears it; Chromium understands socks5:// here too
    await session.defaultSession.setProxy(proxyUrl ? { proxyRules: proxyUrl } : { mode: 'system' })
  } catch (err) {
    console.warn('[proxy] failed to set session proxy:', err)
  }
  console.log(
    proxyUrl
      ? // strip user:pass before logging
        `[proxy] outbound via ${proxyUrl.replace(/\/\/[^@/]*@/, '//***@')}`
      : '[proxy] direct (system default)',
  )
}

/** current proxy, for callers that need to pass it on (e.g. the agent browser) */
export function currentProxyUrl(): string {
  return appliedProxyUrl ?? ''
}

/**
 * Push the network-facing settings into the modules that hold them as process
 * state. Called on load and after every save so a settings change takes effect
 * without a restart, the same way the provider switch does.
 */
export function applyNetworkSettings(settings: Partial<NetworkSettings>): void {
  setTavilyApiKey(settings.tavilyApiKey ?? '')
  void applyProxy(settings.proxyUrl ?? '')
}

/**
 * Load the persisted network settings at startup. Returns true when the user
 * configured an explicit proxy, which tells the app bootstraps to skip their
 * env-var / system-proxy detection: an explicit choice must win over both, and
 * must also be honoured when it says "no proxy" on a machine whose system
 * proxy would otherwise be picked up.
 */
export function bootstrapNetworkSettings(userDataPath: string): boolean {
  const settings = readNetworkSettings(userDataPath)
  applyNetworkSettings(settings)
  return !!normalizeProxyUrl(settings.proxyUrl)
}
