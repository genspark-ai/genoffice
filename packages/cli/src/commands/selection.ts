import { controlEndpoint, controlRequest } from '../control'
import { resolveInput } from '../fs'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

export const selectionCommand: CommandDef = {
  name: 'selection',
  summary:
    "What the user has selected in the GenOffice editor showing this file (the user's own pointer for 'this one', 'here').",
  usage: 'selection <file> --json',
  async run(args, ctx) {
    const path = resolveInput(args.positionals[0], ctx)
    const endpoint = controlEndpoint(ctx.env)
    if (!endpoint) {
      throw new CliError(EXIT.app, 'GenOffice is not running', undefined, {
        reason: 'app_unavailable',
        suggestion: `run \`genoffice open ${path}\` and ask the user to select something`,
      })
    }
    const result = await controlRequest(endpoint, { cmd: 'selection', path })
    return { summary: summarize(result), detail: result }
  },
}

function summarize(result: Record<string, unknown>): string {
  if (result.none) return 'nothing is selected'
  if (Array.isArray(result.elements)) {
    return result.elements.length
      ? `slide ${result.slide}: ${result.elements.join(', ')}`
      : `slide ${result.slide}, no element selected`
  }
  if (Array.isArray(result.blocks)) {
    const [a, b] = result.blocks as number[]
    const text = typeof result.text === 'string' && result.text ? `: ${preview(result.text)}` : ''
    return (a === b ? `block ${a}` : `blocks ${a}-${b}`) + text
  }
  if (typeof result.range === 'string') return `${result.sheet}!${result.range}`
  if (typeof result.page === 'number') {
    return (
      `page ${result.page}` + (typeof result.text === 'string' ? `: ${preview(result.text)}` : '')
    )
  }
  return JSON.stringify(result)
}

function preview(text: string): string {
  const flat = text.replace(/\s+/g, ' ').trim()
  return flat.length > 80 ? `${flat.slice(0, 79)}…` : flat
}
