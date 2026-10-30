import { controlEndpoint, controlRequest } from '../control'
import type { ControlRequest } from '../control-protocol'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

/**
 * The `desktop` domain: the running GenOffice app's Home surface over the
 * control socket. Read actions answer a live snapshot; writes land in the
 * same validated stores the desktop UI uses, so the running app stays
 * consistent with what a script sees.
 */

interface ParsedDesktop {
  request: Extract<ControlRequest, { cmd: 'desktop' }>
  usageError?: never
}

const ACTIONS = ['get-state', 'set-wallpaper', 'todo add <text>', 'todo list', 'windows list']

function parseDesktopArgs(positionals: string[]): ParsedDesktop {
  const [action, ...rest] = positionals
  switch ((action ?? '').toLowerCase()) {
    case 'get-state':
      return { request: { cmd: 'desktop', action: 'get-state' } }
    case 'set-wallpaper': {
      const wallpaper = rest.join(' ').trim()
      if (!wallpaper) {
        throw new CliError(
          EXIT.usage,
          'set-wallpaper needs a wallpaper name or image path',
          undefined,
          {
            reason: 'missing_argument',
            suggestion:
              'use "default", a builtin wallpaper name (see `desktop get-state` detail.prefs), or an absolute image path',
          },
        )
      }
      return { request: { cmd: 'desktop', action: 'set-wallpaper', wallpaper } }
    }
    case 'todo': {
      const verb = (rest[0] ?? '').toLowerCase()
      if (verb === 'list') return { request: { cmd: 'desktop', action: 'todo-list' } }
      if (verb === 'add') {
        const text = rest.slice(1).join(' ').trim()
        if (!text) {
          throw new CliError(EXIT.usage, 'todo add needs the item text', undefined, {
            reason: 'missing_argument',
            suggestion: 'genoffice desktop todo add "call the accountant"',
          })
        }
        return { request: { cmd: 'desktop', action: 'todo-add', text } }
      }
      throw unknownAction(['todo', ...rest].join(' '))
    }
    case 'todo-add': {
      const text = rest.join(' ').trim()
      if (!text) {
        throw new CliError(EXIT.usage, 'todo-add needs the item text', undefined, {
          reason: 'missing_argument',
          suggestion: 'genoffice desktop todo add "water the plants"',
        })
      }
      return { request: { cmd: 'desktop', action: 'todo-add', text } }
    }
    case 'todo-list':
      return { request: { cmd: 'desktop', action: 'todo-list' } }
    case 'windows': {
      if ((rest[0] ?? '').toLowerCase() === 'list') {
        return { request: { cmd: 'desktop', action: 'windows-list' } }
      }
      throw unknownAction(['windows', ...rest].join(' '))
    }
    case 'windows-list':
      return { request: { cmd: 'desktop', action: 'windows-list' } }
    default:
      throw unknownAction(action ?? '')
  }
}

function unknownAction(action: string): CliError {
  return new CliError(
    EXIT.usage,
    `unknown desktop action: ${action || '(none)'}`,
    { actions: ACTIONS },
    { reason: 'invalid_argument', suggestion: `use one of: ${ACTIONS.join(', ')}` },
  )
}

export const desktopCommand: CommandDef = {
  name: 'desktop',
  summary:
    "The running GenOffice app's Home desktop: wallpaper, icon layout summary, todo widget and open windows — read them, or change the wallpaper and add todos.",
  usage:
    'desktop <get-state | set-wallpaper <name|path> | todo add <text> | todo list | windows list> [--json]',
  async run(args, ctx) {
    const { request } = parseDesktopArgs(args.positionals)
    const endpoint = controlEndpoint(ctx.env)
    if (!endpoint) {
      throw new CliError(EXIT.app, 'GenOffice is not running', undefined, {
        reason: 'app_unavailable',
        suggestion: 'start GenOffice (or `genoffice open <file>`) and retry',
      })
    }
    const result = await controlRequest(endpoint, request)
    return { summary: summarize(request, result), detail: result }
  },
}

function summarize(
  request: Extract<ControlRequest, { cmd: 'desktop' }>,
  result: Record<string, unknown>,
): string {
  switch (request.action) {
    case 'get-state': {
      const prefs = (result.prefs ?? {}) as { wallpaper?: unknown }
      const layout = (result.layout ?? {}) as { positioned?: unknown }
      const todos = Array.isArray(result.todos) ? result.todos.length : 0
      const windows = Array.isArray(result.windows) ? result.windows.length : 0
      const undone = Array.isArray(result.todos)
        ? (result.todos as Array<{ done?: unknown }>).filter((t) => t.done !== true).length
        : 0
      return (
        `wallpaper: ${String(prefs.wallpaper ?? '?')} · ` +
        `${String(layout.positioned ?? 0)} positioned icons · ` +
        `${undone}/${todos} open todos · ${windows} window tabs`
      )
    }
    case 'set-wallpaper':
      return `wallpaper set to ${String(result.wallpaper ?? '')}`
    case 'todo-add': {
      const count = Array.isArray(result.todos) ? result.todos.length : 0
      return `added todo "${request.text}" (${count} on the list)`
    }
    case 'todo-list': {
      const todos = Array.isArray(result.todos)
        ? (result.todos as Array<{ text?: unknown; done?: unknown }>)
        : []
      if (todos.length === 0) return 'no todos'
      return todos
        .map((t) => `${t.done === true ? '[x]' : '[ ]'} ${String(t.text ?? '')}`)
        .join('\n')
    }
    case 'windows-list': {
      const windows = Array.isArray(result.windows)
        ? (result.windows as Array<{ kind?: unknown; title?: unknown; filePath?: unknown }>)
        : []
      if (windows.length === 0) return 'no open windows'
      return windows
        .map(
          (w) =>
            `${String(w.kind ?? '?')}\t${String(w.title ?? '')}` +
            (typeof w.filePath === 'string' ? `\t${w.filePath}` : ''),
        )
        .join('\n')
    }
  }
}
