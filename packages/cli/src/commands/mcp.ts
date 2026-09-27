import { existsSync } from 'node:fs'
import { homedir } from 'node:os'
import { resolve } from 'node:path'
import {
  mcpAgentIds,
  mcpConfigFile,
  mcpSnippet,
  readMcpEntry,
  removeMcpEntry,
  writeMcpEntry,
  type McpEntryState,
} from '../agent-mcp'
import { agentTarget, detectAgents, type AgentId } from '../agent-skills'
import { flagBool, flagString } from '../args'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type Warning } from '../result'
import { didYouMean } from '../suggest'
import { mcpLaunchFromLauncher, type McpLaunch } from '../mcp-launch'
import { launcherPath } from './install'

const SUBCOMMANDS = ['install', 'uninstall', 'list']

export const mcpCommand: CommandDef = {
  name: 'mcp',
  summary:
    "Serve every command as a Model Context Protocol tool: on stdio for an MCP client on this machine (`claude mcp add --transport stdio genoffice -- genoffice mcp`, Cursor, Claude Desktop), or with --http as a Streamable HTTP server that clients on other machines reach by URL. Over HTTP, files travel too: PUT /files/<name> uploads one and every tool takes http(s) URLs in place of paths; outputs come back as download URLs and, when small, as embedded resources. Ops, specs and Markdown are passed inline either way; a new deck goes through deck_start, deck_page and deck_build. `mcp install <agent|all>` registers the stdio server in a coding agent's MCP config (Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf); `mcp uninstall <agent>` removes it; `mcp list` shows where each agent stands. The ops, cells and data parameters of the apply and create tools are advertised with the typed per-op schema of `guide <domain> --json`; --compact-schemas (or GENOFFICE_MCP_COMPACT_SCHEMAS=1) advertises them as plain arrays for clients with a small context budget.",
  usage:
    'mcp [--http <port> [--host <addr>] [--token <secret>]] [--compact-schemas] | mcp install <agent|all> [--dir <path>] [--force] | mcp uninstall <agent|all> [--dir <path>] [--force] | mcp list',
  quiet: (args) => args.positionals.length === 0,
  options: [
    {
      name: 'http',
      value: 'port',
      description: 'serve Streamable HTTP on this port instead of stdio (0 picks a free port)',
    },
    {
      name: 'host',
      value: 'addr',
      description: 'address to bind with --http (default: 127.0.0.1; 0.0.0.0 for other machines)',
    },
    {
      name: 'token',
      value: 'secret',
      description:
        'with --http, require this bearer token on every request (default: $GENOFFICE_MCP_TOKEN, else none)',
    },
    {
      name: 'compact-schemas',
      description:
        'advertise ops, cells and data as untyped arrays instead of the per-op schema (smaller tools/list; also GENOFFICE_MCP_COMPACT_SCHEMAS=1)',
    },
    {
      name: 'dir',
      value: 'path',
      description:
        "install/uninstall: treat this folder as the agent's config directory (one agent only)",
    },
    {
      name: 'force',
      description:
        'install/uninstall: also write when the agent is not detected or the genoffice entry starts another program',
    },
  ],
  async run(args, ctx) {
    const sub = args.positionals[0]
    if (sub === undefined) {
      for (const flag of ['dir', 'force']) {
        if (flagBool(args, flag)) {
          throw new CliError(
            EXIT.usage,
            `--${flag} needs \`mcp install\` or \`mcp uninstall\``,
            undefined,
            {
              reason: 'invalid_argument',
            },
          )
        }
      }
      return serve(args, ctx)
    }
    for (const flag of ['http', 'host', 'token', 'compact-schemas']) {
      if (flagBool(args, flag)) {
        throw new CliError(
          EXIT.usage,
          `--${flag} serves MCP; it does not combine with \`mcp ${sub}\``,
          undefined,
          {
            reason: 'invalid_argument',
          },
        )
      }
    }
    switch (sub) {
      case 'list':
        return list(ctx)
      case 'install':
      case 'uninstall':
        return edit(sub, args.positionals[1], flagString(args, 'dir'), flagBool(args, 'force'), ctx)
      default:
        throw new CliError(
          EXIT.usage,
          `unknown mcp subcommand: ${sub}`,
          { supported: SUBCOMMANDS },
          {
            reason: 'invalid_argument',
            suggestion: didYouMean(sub, SUBCOMMANDS)
              ? `did you mean \`${didYouMean(sub, SUBCOMMANDS)}\`?`
              : 'run `mcp` alone to serve, or `mcp install <agent|all>`, `mcp uninstall <agent>`, `mcp list`',
          },
        )
    }
  },
}

async function serve(args: Parameters<CommandDef['run']>[0], ctx: CommandContext) {
  const http = flagString(args, 'http')
  const compactSchemas =
    flagBool(args, 'compact-schemas') || envFlag(ctx.env.GENOFFICE_MCP_COMPACT_SCHEMAS)
  if (http === undefined) {
    // lazy: the server imports runCli, which registers this command
    const { serveStdio } = await import('../mcp/server')
    await serveStdio({ cwd: ctx.cwd, env: ctx.env, log: ctx.log, compactSchemas })
    return { summary: 'mcp session ended' }
  }
  const port = Number(http)
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    throw new CliError(EXIT.usage, `--http needs a port number, got ${http}`, undefined, {
      reason: 'invalid_argument',
    })
  }
  const tokenFlag = flagString(args, 'token')
  // `--token ""` — e.g. `--token "$MY_TOKEN"` with the variable unset —
  // would reach authorized() as "no auth configured": the operator believes
  // the lock is on while every request passes. Fail closed instead.
  if (tokenFlag !== undefined && tokenFlag.trim() === '') {
    throw new CliError(
      EXIT.usage,
      '--token needs a non-empty value (omit the flag, or set GENOFFICE_MCP_TOKEN)',
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  const { serveHttp } = await import('../mcp/http')
  await serveHttp({
    cwd: ctx.cwd,
    env: ctx.env,
    log: ctx.log,
    port,
    host: flagString(args, 'host'),
    token: tokenFlag ?? (ctx.env.GENOFFICE_MCP_TOKEN?.trim() || undefined),
    compactSchemas,
  })
  return { summary: 'mcp server stopped' }
}

function envFlag(value: string | undefined): boolean {
  return /^(1|true|yes|on)$/i.test(value?.trim() ?? '')
}

type InstallStatus = 'installed' | 'updated' | 'unchanged' | 'occupied' | 'manual' | 'not_detected'
type UninstallStatus = 'removed' | 'absent' | 'occupied' | 'manual' | 'not_detected'

interface AgentRow {
  agent: AgentId
  label: string
  detected: boolean
  config: string
  registered: boolean
  /** what the existing genoffice entry starts, when there is one */
  command: string | null
}

interface Target {
  id: AgentId
  label: string
  file: string
  detected: boolean
}

function home(env: NodeJS.ProcessEnv): string {
  return env.GENOFFICE_HOME || homedir()
}

/** The entry every config gets: the absolute launcher, or on Windows the app run as Node on the bundle. */
function launcher(): { launcher: string; launch: McpLaunch } {
  const path = launcherPath()
  if (!path) throw new CliError(EXIT.app, 'cannot locate the genoffice launcher')
  return {
    launcher: path,
    launch: mcpLaunchFromLauncher(path, { platform: process.platform, exists: existsSync }),
  }
}

function row(t: Target, state: McpEntryState): AgentRow {
  return {
    agent: t.id,
    label: t.label,
    detected: t.detected,
    config: t.file,
    registered: state.status === 'registered' || state.status === 'stale',
    command: state.command,
  }
}

function list(ctx: CommandContext) {
  const { launcher: bin, launch } = launcher()
  const detected = new Set(detectAgents(ctx.env, home(ctx.env)).map((a) => a.id))
  const agents = mcpAgentIds().map((id) => {
    const t = target(id, undefined, detected, ctx)
    const state = readMcpEntry(id, t.file, launch)
    return { ...row(t, state), status: state.status }
  })
  const found = agents.filter((a) => a.detected)
  return {
    summary: found.length
      ? `${found.length} agent(s) detected, ${found.filter((a) => a.registered).length} with genoffice registered: ${found
          .map((a) => `${a.agent} ${a.status}`)
          .join(', ')}`
      : 'no coding agent detected on this machine',
    detail: { launcher: bin, launch, agents },
  }
}

function target(
  id: AgentId,
  dir: string | undefined,
  detected: Set<AgentId>,
  ctx: CommandContext,
): Target {
  const opts = { env: ctx.env, home: home(ctx.env), ...(dir ? { dir: resolve(ctx.cwd, dir) } : {}) }
  return {
    id,
    label: agentTarget(id, ctx.env, home(ctx.env))!.label,
    file: mcpConfigFile(id, opts)!,
    detected: dir !== undefined || detected.has(id),
  }
}

function edit(
  verb: 'install' | 'uninstall',
  name: string | undefined,
  dir: string | undefined,
  force: boolean,
  ctx: CommandContext,
) {
  const ids = mcpAgentIds()
  if (!name) {
    throw new CliError(EXIT.usage, `mcp ${verb} needs an agent name or \`all\``, undefined, {
      reason: 'missing_argument',
      suggestion: `agents: ${ids.join(', ')}`,
    })
  }
  if (name === 'all' && dir !== undefined) {
    throw new CliError(EXIT.usage, '--dir applies to one agent, not `all`', undefined, {
      reason: 'invalid_argument',
      suggestion: 'name the agent whose config lives in that folder',
    })
  }
  if (name !== 'all' && !ids.includes(name as AgentId)) {
    const guess = didYouMean(name, ids)
    throw new CliError(
      EXIT.usage,
      `unknown agent: ${name}`,
      { supported: [...ids, 'all'] },
      {
        reason: 'unsupported',
        suggestion: guess ? `did you mean \`${guess}\`?` : `agents: ${ids.join(', ')}, or all`,
      },
    )
  }
  const { launcher: bin, launch } = launcher()
  const detected = new Set(detectAgents(ctx.env, home(ctx.env)).map((a) => a.id))
  const targets = (name === 'all' ? ids : [name as AgentId]).map((id) =>
    target(id, dir, detected, ctx),
  )
  const single = targets.length === 1 ? targets[0]! : null
  if (single && !single.detected && !force) {
    throw new CliError(
      EXIT.usage,
      `${single.label} is not installed on this machine (${
        verb === 'install'
          ? `${single.file} would be written`
          : `nothing to remove at ${single.file}`
      })`,
      { agent: single.id, config: single.file },
      {
        reason: 'unsupported',
        suggestion: `repeat with --force to ${verb === 'install' ? 'write' : 'edit'} the file anyway`,
      },
    )
  }
  const forceHint =
    verb === 'install'
      ? 'repeat with --force to replace it'
      : 'repeat with --force to remove the foreign entry'
  const manualHint =
    verb === 'install'
      ? 'paste detail.snippet into the file yourself'
      : 'remove the genoffice entry from the file yourself'

  const warnings: Warning[] = []
  const rows: Array<AgentRow & { status: InstallStatus | UninstallStatus; snippet?: string }> = []
  for (const t of targets) {
    const before = readMcpEntry(t.id, t.file, launch)
    let status: InstallStatus | UninstallStatus
    if (!t.detected && !force) status = 'not_detected'
    else if (before.status === 'manual') status = 'manual'
    else if (before.status === 'occupied' && !force) status = 'occupied'
    else if (verb === 'install') {
      if (before.status === 'registered') status = 'unchanged'
      else {
        writeMcpEntry(t.id, t.file, launch)
        status = before.status === 'absent' ? 'installed' : 'updated'
      }
    } else status = removeMcpEntry(t.id, t.file) ? 'removed' : 'absent'
    const after = status === 'manual' ? before : readMcpEntry(t.id, t.file, launch)
    const r = { ...row(t, after), status }
    rows.push(
      status === 'manual' && verb === 'install' ? { ...r, snippet: mcpSnippet(t.id, launch) } : r,
    )
    if (status === 'occupied' || status === 'manual') {
      warnings.push({
        code: status === 'occupied' ? 'entry_occupied' : 'manual_edit_needed',
        message:
          status === 'occupied'
            ? `${t.id}: the genoffice entry in ${t.file} starts ${before.command ?? 'something else'}`
            : `${t.id}: ${t.file} could not be edited; ${verb === 'install' ? 'add' : 'remove'} the entry by hand`,
        suggestion:
          status === 'occupied'
            ? forceHint
            : verb === 'install'
              ? 'see detail.agents[].snippet'
              : manualHint,
      })
    }
  }
  if (single && (rows[0]!.status === 'occupied' || rows[0]!.status === 'manual')) {
    const r = rows[0]!
    throw new CliError(
      EXIT.file,
      r.status === 'occupied'
        ? `${single.id}: the genoffice entry in ${single.file} starts ${r.command ?? 'something else'}; not ${verb === 'install' ? 'replacing' : 'removing'} it`
        : `${single.id}: ${single.file} could not be edited safely`,
      {
        agent: single.id,
        config: single.file,
        status: r.status,
        command: r.command,
        ...(r.snippet ? { snippet: r.snippet } : {}),
      },
      { reason: 'output_exists', suggestion: r.status === 'occupied' ? forceHint : manualHint },
    )
  }
  const changed = rows.filter(
    (r) => r.status === 'installed' || r.status === 'updated' || r.status === 'removed',
  )
  const skipped = rows.filter((r) => r.status === 'not_detected').length
  const blocked = rows.filter((r) => r.status === 'occupied' || r.status === 'manual').length
  const tail = [
    skipped ? `${skipped} agent(s) not detected` : '',
    blocked ? `${blocked} agent(s) left alone (see warnings)` : '',
  ]
    .filter(Boolean)
    .join(', ')
  const nothing = verb === 'install' ? 'nothing registered' : 'nothing removed'
  const summary = changed.length
    ? `${verb === 'install' ? 'registered genoffice mcp for' : 'removed genoffice mcp from'} ${changed
        .map((r) => `${r.agent} (${r.status})`)
        .join(', ')}${tail ? `; ${tail}` : ''}`
    : rows.some((r) => r.status === 'unchanged')
      ? `genoffice mcp already registered${tail ? `; ${tail}` : ''}`
      : rows.some((r) => r.status === 'absent')
        ? `genoffice mcp was not registered${tail ? `; ${tail}` : ''}`
        : blocked
          ? `${nothing}: ${tail}`
          : `${nothing}: no coding agent detected on this machine`
  return {
    summary,
    ...(warnings.length ? { warnings } : {}),
    detail: { launcher: bin, launch, agents: rows },
  }
}
