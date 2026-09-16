import { imageSearchTool, webSearchTool } from '@genoffice/ai-search'
import { flagBool, flagString } from '../args'
import { aiSettingsPath, prepareCloud } from '../cloud'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

export const searchCommand: CommandDef = {
  name: 'search',
  summary:
    'Web or image search through the provider configured in GenOffice (Genspark, Serper, Tavily).',
  usage: 'search <query> [--images] [--max <n>]',
  options: [
    { name: 'images', description: 'search images instead of web pages' },
    { name: 'max', value: 'n', description: 'result count (default 6 web, 8 images; at most 20)' },
  ],
  async run(args, ctx) {
    const query = args.positionals.join(' ').trim()
    if (!query)
      throw new CliError(EXIT.usage, 'missing <query>', undefined, { reason: 'missing_argument' })
    const max = resultCount(flagString(args, 'max'))
    await prepareCloud(ctx.env)
    const settings = aiSettingsPath(ctx.env)
    if (flagBool(args, 'images')) {
      const r = await imageSearchTool(settings, query, max ?? 8)
      if (r.method === 'error') throw new CliError(EXIT.app, r.error ?? 'image search failed')
      return {
        summary: `${r.images.length} images for "${query}" via ${r.method}`,
        detail: { method: r.method, images: r.images },
      }
    }
    const r = await webSearchTool(settings, query, max ?? 6)
    if (r.method === 'error') throw new CliError(EXIT.app, r.error ?? 'search failed')
    return {
      summary: `${r.results.length} results for "${query}" via ${r.method}`,
      detail: {
        method: r.method,
        ...(r.answer ? { answer: r.answer } : {}),
        results: r.results,
      },
    }
  },
}

/** `--max` clamped to 1..20; undefined when absent so each mode keeps its own default. */
export function resultCount(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined
  const n = Number(raw)
  if (!Number.isFinite(n))
    throw new CliError(EXIT.usage, `--max must be a number, got "${raw}"`, undefined, {
      reason: 'invalid_argument',
    })
  return Math.min(20, Math.max(1, Math.floor(n)))
}
