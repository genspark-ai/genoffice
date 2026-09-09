/**
 * ai:web-search / ai:image-search for the editors' main processes: reads
 * ai-settings.json live and turns the search provider choice into
 * SearchOptions — Genspark keeps the historic chain (gsk when signed in and
 * cloud tools are on, then env keys, then DuckDuckGo); a user Serper / Tavily
 * key runs first and skips gsk.
 */

import {
  activeSearchProvider,
  cloudToolsEnabled,
  type AiSearchProviderId,
  type AiSettings,
} from '@genoffice/ai-provider'
import { imageSearch, webSearch, type SearchOptions } from './index'
import { readAiSettingsFile } from './media-tools'

export function searchOptionsFromSettings(settings: AiSettings): SearchOptions {
  const provider = activeSearchProvider(settings)
  if (provider === 'genspark') return { useGsk: cloudToolsEnabled(settings) }
  const key = settings.search!.providers[provider].apiKey
  return provider === 'tavily'
    ? { useGsk: false, tavilyKey: key, prefer: 'tavily' }
    : { useGsk: false, serperKey: key }
}

export function webSearchTool(settingsPath: string, query: string, maxResults = 6) {
  return webSearch(query, maxResults, searchOptionsFromSettings(readAiSettingsFile(settingsPath)))
}

export function imageSearchTool(settingsPath: string, query: string, maxResults = 8) {
  return imageSearch(query, maxResults, searchOptionsFromSettings(readAiSettingsFile(settingsPath)))
}

/** settings-UI test: one minimal query against the given key must be answered by that backend */
export async function testSearchProvider(
  provider: AiSearchProviderId,
  apiKey: string,
): Promise<{ ok: boolean; error?: string }> {
  if (provider === 'genspark') return { ok: true }
  if (!apiKey) return { ok: false, error: 'API key is empty' }
  const options: SearchOptions =
    provider === 'tavily'
      ? { useGsk: false, tavilyKey: apiKey, serperKey: '', prefer: 'tavily' }
      : { useGsk: false, serperKey: apiKey, tavilyKey: '' }
  const r = await webSearch('GenOffice', 1, options)
  if (r.method === provider) return { ok: true }
  return {
    ok: false,
    error:
      r.method === 'error'
        ? (r.error ?? 'search failed')
        : `${provider} did not answer (key rejected or quota exhausted); fell back to ${r.method}`,
  }
}
