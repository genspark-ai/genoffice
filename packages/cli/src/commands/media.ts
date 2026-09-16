import { fileURLToPath } from 'node:url'
import { analyzeMediaTool } from '@genoffice/ai-search'
import { flagString } from '../args'
import { aiSettingsPath, prepareCloud } from '../cloud'
import { resolveInput } from '../fs'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

const DEFAULT_ASK = 'Describe this media in detail: subject, text, layout, notable details.'

export const mediaCommand: CommandDef = {
  name: 'media',
  summary:
    'Describe or answer a question about an image, video or audio file (or URL) with the configured provider.',
  usage: 'media <file|url> [--ask <question>]',
  options: [
    { name: 'ask', value: 'question', description: `what to extract (default: "${DEFAULT_ASK}")` },
  ],
  async run(args, ctx) {
    const ref = args.positionals[0]
    if (!ref)
      throw new CliError(EXIT.usage, 'missing <file|url>', undefined, {
        reason: 'missing_argument',
      })
    const target = /^https?:\/\//i.test(ref)
      ? ref
      : resolveInput(ref.startsWith('file:') ? fileURLToPath(ref) : ref, ctx)
    await prepareCloud(ctx.env)
    const r = await analyzeMediaTool(aiSettingsPath(ctx.env), {
      mediaUrls: [target],
      requirements: flagString(args, 'ask') ?? DEFAULT_ASK,
    })
    if (r.text === undefined) throw new CliError(EXIT.app, r.error ?? 'media analysis failed')
    const failure = providerFailure(r.text)
    if (failure) throw new CliError(EXIT.conversion, `media analysis failed: ${failure}`)
    const text = analysisText(r.text)
    return {
      summary: text,
      detail: { source: target, text, ...(text !== r.text ? { raw: r.text } : {}) },
    }
  },
}

/** Genspark reports a fetch/analysis failure as { status: "error", error|message } per file. */
export function providerFailure(text: string): string | null {
  if (!text.trimStart().startsWith('{')) return null
  try {
    const parsed = JSON.parse(text) as Record<string, Record<string, unknown>>
    const entries = Object.values(parsed).filter((v) => v && typeof v === 'object')
    // only an explicit error status counts; other shapes fall through to analysisText
    const failed = entries.find((v) => v.status === 'error' || v.status === 'failed')
    if (!failed) return null
    return String(failed.error ?? failed.message ?? failed.status ?? 'provider error')
  } catch {
    return null
  }
}

/** Genspark answers with a JSON map of upload → { analysis }; BYOK providers with prose. */
export function analysisText(text: string): string {
  if (!text.trimStart().startsWith('{')) return text
  try {
    const parsed = JSON.parse(text) as Record<string, { analysis?: unknown }>
    const parts = Object.values(parsed)
      .map((v) => (typeof v?.analysis === 'string' ? v.analysis.trim() : ''))
      .filter(Boolean)
    return parts.length ? parts.join('\n\n') : text
  } catch {
    return text
  }
}
