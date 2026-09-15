import {
  activeMediaProvider,
  activeSearchProvider,
  cloudToolsEnabled,
  imageGenerationAvailable,
  mediaAnalysisAvailable,
} from '@genoffice/ai-provider'
import { hasGskAuth, readAiSettingsFile } from '@genoffice/ai-search'
import { aiSettingsPath, prepareCloud } from '../cloud'
import type { CommandDef } from '../registry'
import { appLaunch } from '../resources'

/**
 * What the cloud commands can do on this machine. We report two layers:
 *
 *  - `configured` is true when a BYOK key or a Genspark login is set up —
 *    the user did explicit work to enable this capability.
 *  - `available` is true when at least one path can serve the request, which
 *    for web/image search INCLUDES the zero-config DuckDuckGo fallback that
 *    the @genoffice/agent-skills parsers hit by default. Image generation
 *    and media analysis have no DDG-style fallback, so their `available`
 *    mirrors `configured`.
 *
 * Agents that "check once before planning" need both layers: `configured` for
 * cost / quality planning, `available` so they don't write off features that
 * actually work.
 */
export const capabilitiesCommand: CommandDef = {
  name: 'capabilities',
  summary:
    'Report which cloud features (search, image search, image generation, media analysis) are configured in GenOffice, whether unkeyed fallbacks (DuckDuckGo) make them available without setup, and whether the app is installed.',
  usage: 'capabilities',
  async run(_args, ctx) {
    await prepareCloud(ctx.env)
    const settings = readAiSettingsFile(aiSettingsPath(ctx.env))
    const gsk = hasGskAuth() && cloudToolsEnabled(settings)
    const searchProvider = activeSearchProvider(settings)
    const keyedSearch = searchProvider !== 'genspark'
    const searchConfigured = gsk || keyedSearch
    // Search and image search both fall back to DuckDuckGo HTML scraping —
    // the @genoffice/agent-skills parsers handle that without a key.
    const searchAvailable = searchConfigured || true
    const imageSearchConfigured = gsk || searchProvider === 'serper'
    const imageSearchAvailable = imageSearchConfigured || true
    const imageGenConfigured = imageGenerationAvailable(settings, hasGskAuth())
    // No DDG fallback for image generation — a vendor or gsk credits are required.
    const imageGenReady = imageGenConfigured
    const mediaAnalysisConfigured = mediaAnalysisAvailable(settings, hasGskAuth())
    const mediaAnalysisReady = mediaAnalysisConfigured
    const via = (byok: string | null | undefined) => (byok ? byok : gsk ? 'genspark' : null)
    const detail = {
      search: {
        configured: searchConfigured,
        available: searchAvailable,
        via: keyedSearch ? searchProvider : gsk ? 'genspark' : 'duckduckgo',
        fallback: searchConfigured ? 'duckduckgo' : undefined,
      },
      image_search: {
        configured: imageSearchConfigured,
        available: imageSearchAvailable,
        via: searchProvider === 'serper' ? 'serper' : gsk ? 'genspark' : 'duckduckgo',
        fallback: imageSearchConfigured ? 'duckduckgo' : undefined,
      },
      image_generation: {
        configured: imageGenConfigured,
        available: imageGenReady,
        via: imageGenReady ? via(activeMediaProvider(settings, 'image')) : null,
      },
      media_analysis: {
        configured: mediaAnalysisConfigured,
        available: mediaAnalysisReady,
        via: mediaAnalysisReady ? via(activeMediaProvider(settings, 'analysis')) : null,
      },
      app: { available: appLaunch(ctx.env) !== null },
      settings_path: aiSettingsPath(ctx.env),
    }
    // "available" drives the headline — a capability is usable if EITHER
    // layer can serve it, even when the user hasn't configured anything.
    const on = Object.entries(detail)
      .filter(([k, v]) => k !== 'settings_path' && (v as { available: boolean }).available)
      .map(([k]) => k)
    return {
      summary: on.length
        ? `available: ${on.join(', ')}`
        : 'no cloud feature available; the app is not installed',
      detail,
    }
  },
}
