import { existsSync, lstatSync, mkdirSync, readFileSync, readlinkSync } from 'node:fs'
import { basename, dirname, isAbsolute, join } from 'node:path'
import { agentConfigDir, type AgentId } from './agent-skills'
import { realizedPath, writeOutput } from './fs'
import type { McpLaunch } from './mcp-launch'

/**
 * Registering `genoffice mcp` (stdio) in each coding agent's MCP config, the
 * MCP twin of agent-skills.ts. One key per file is set or removed; an entry
 * named genoffice that starts something other than a genoffice launcher is
 * never replaced unless forced. Paths and shapes follow each agent's docs.
 */

export const MCP_SERVER_NAME = 'genoffice'

export type McpEntryStatus =
  /** no genoffice entry in the file */
  | 'absent'
  /** genoffice entry starting our launcher, exactly as this CLI would write it */
  | 'registered'
  /** a genoffice launcher, but another path or extra fields */
  | 'stale'
  /** an entry named genoffice that starts something else */
  | 'occupied'
  /** the file cannot be edited safely (unparsable, or a shape we do not know) */
  | 'manual'

export interface McpEntryState {
  status: McpEntryStatus
  /** the program the existing entry starts, when readable */
  command: string | null
}

interface JsonShape {
  kind: 'json'
  /** top-level key holding the server map */
  key: string
  entry: (launch: McpLaunch) => Record<string, unknown>
  /** what an existing entry starts */
  launchOf: (entry: unknown) => { command: string | null; args: string[] }
}

interface TomlShape {
  kind: 'toml'
}

interface McpAgentDef {
  id: AgentId
  /** config file, relative to the agent's dotfolder unless the agent keeps it elsewhere */
  file: (configDir: string, home: string, explicit: boolean) => string
  shape: JsonShape | TomlShape
}

const strings = (v: unknown): string[] =>
  Array.isArray(v) ? v.filter((x): x is string => typeof x === 'string') : []

const plainLaunch: JsonShape['launchOf'] = (entry) =>
  isRecord(entry)
    ? {
        command: typeof entry.command === 'string' ? entry.command : null,
        args: strings(entry.args),
      }
    : { command: null, args: [] }

const withEnv = (launch: McpLaunch) => (launch.env ? { env: launch.env } : {})

const stdioEntry: JsonShape = {
  kind: 'json',
  key: 'mcpServers',
  entry: (l) => ({ type: 'stdio', command: l.command, args: l.args, ...withEnv(l) }),
  launchOf: plainLaunch,
}

const bareEntry: JsonShape = {
  kind: 'json',
  key: 'mcpServers',
  entry: (l) => ({ command: l.command, args: l.args, ...withEnv(l) }),
  launchOf: plainLaunch,
}

const MCP_AGENTS: readonly McpAgentDef[] = [
  {
    id: 'claude-code',
    // `claude mcp add --scope user` writes ~/.claude.json, or <CLAUDE_CONFIG_DIR>/.claude.json
    file: (dir, home, explicit) => join(explicit ? dir : home, '.claude.json'),
    shape: stdioEntry,
  },
  { id: 'codex', file: (dir) => join(dir, 'config.toml'), shape: { kind: 'toml' } },
  { id: 'cursor', file: (dir) => join(dir, 'mcp.json'), shape: stdioEntry },
  { id: 'gemini', file: (dir) => join(dir, 'settings.json'), shape: bareEntry },
  {
    id: 'copilot',
    file: (dir) => join(dir, 'mcp-config.json'),
    shape: {
      kind: 'json',
      key: 'mcpServers',
      entry: (l) => ({
        type: 'local',
        command: l.command,
        args: l.args,
        ...withEnv(l),
        tools: ['*'],
      }),
      launchOf: plainLaunch,
    },
  },
  {
    id: 'opencode',
    file: (dir) => join(dir, 'opencode.json'),
    shape: {
      kind: 'json',
      key: 'mcp',
      entry: (l) => ({
        type: 'local',
        command: [l.command, ...l.args],
        enabled: true,
        ...(l.env ? { environment: l.env } : {}),
      }),
      launchOf: (entry) => {
        const words = isRecord(entry) ? strings(entry.command) : []
        return { command: words[0] ?? null, args: words.slice(1) }
      },
    },
  },
  { id: 'windsurf', file: (dir) => join(dir, 'mcp_config.json'), shape: bareEntry },
]

export function mcpAgentIds(): AgentId[] {
  return MCP_AGENTS.map((a) => a.id)
}

export interface McpConfigOptions {
  env?: NodeJS.ProcessEnv
  home?: string
  /** use this folder as the agent's dotfolder instead of the probed one */
  dir?: string
}

/** Where `id` keeps its MCP servers; null for an unknown agent. */
export function mcpConfigFile(id: AgentId, opts: McpConfigOptions = {}): string | null {
  const def = MCP_AGENTS.find((a) => a.id === id)
  if (!def) return null
  const env = opts.env ?? process.env
  const home = opts.home ?? homeOf(env)
  const dir = opts.dir ?? agentConfigDir(id, env, home)
  if (!dir) return null
  const explicit = opts.dir !== undefined || (id === 'claude-code' && !!env.CLAUDE_CONFIG_DIR)
  return def.file(dir, home, explicit)
}

function homeOf(env: NodeJS.ProcessEnv): string {
  return env.GENOFFICE_HOME || env.HOME || env.USERPROFILE || ''
}

/** What a person pastes when the file cannot be edited for them. */
export function mcpSnippet(id: AgentId, launch: McpLaunch): string {
  const def = MCP_AGENTS.find((a) => a.id === id)!
  if (def.shape.kind === 'toml') return tomlTable(launch)
  return JSON.stringify(
    { [def.shape.key]: { [MCP_SERVER_NAME]: def.shape.entry(launch) } },
    null,
    2,
  )
}

export function readMcpEntry(id: AgentId, file: string, launch: McpLaunch): McpEntryState {
  const def = MCP_AGENTS.find((a) => a.id === id)!
  if (!existsSync(file)) {
    return def.shape.kind === 'json' && existsSync(jsoncTwin(file))
      ? { status: 'manual', command: null }
      : { status: 'absent', command: null }
  }
  const text = readFileSync(file, 'utf-8')
  if (def.shape.kind === 'toml') return tomlState(text, launch)
  if (!text.trim()) return { status: 'absent', command: null }
  const doc = parseJsonObject(text)
  if (!doc) return { status: 'manual', command: null }
  const map = doc[def.shape.key]
  if (map === undefined) return { status: 'absent', command: null }
  if (!isRecord(map)) return { status: 'manual', command: null }
  const entry = map[MCP_SERVER_NAME]
  if (entry === undefined) return { status: 'absent', command: null }
  const { command, args } = def.shape.launchOf(entry)
  if (JSON.stringify(entry) === JSON.stringify(def.shape.entry(launch))) {
    return { status: 'registered', command }
  }
  return { status: isGenofficeLauncher(command, args) ? 'stale' : 'occupied', command }
}

/** Set the genoffice entry (replacing any existing one); the caller has already decided that is allowed. */
export function writeMcpEntry(id: AgentId, file: string, launch: McpLaunch): void {
  const def = MCP_AGENTS.find((a) => a.id === id)!
  const text = existsSync(file) ? readFileSync(file, 'utf-8') : ''
  if (def.shape.kind === 'toml') {
    const stripped = removeTomlTable(text)
    const body = stripped.length && !stripped.endsWith('\n') ? `${stripped}\n` : stripped
    save(file, `${body}${body.trim().length ? '\n' : ''}${tomlTable(launch)}`)
    return
  }
  const doc = text.trim() ? parseJsonObject(text) : {}
  if (!doc) throw new Error(`${file} is not a JSON object`)
  const map = isRecord(doc[def.shape.key]) ? (doc[def.shape.key] as Record<string, unknown>) : {}
  doc[def.shape.key] = { ...map, [MCP_SERVER_NAME]: def.shape.entry(launch) }
  save(file, `${JSON.stringify(doc, null, 2)}\n`)
}

/** Remove the genoffice entry; false when there was none. Other keys stay. */
export function removeMcpEntry(id: AgentId, file: string): boolean {
  const def = MCP_AGENTS.find((a) => a.id === id)!
  if (!existsSync(file)) return false
  const text = readFileSync(file, 'utf-8')
  if (def.shape.kind === 'toml') {
    const next = removeTomlTable(text)
    if (next === text) return false
    save(file, next)
    return true
  }
  if (!text.trim()) return false
  const doc = parseJsonObject(text)
  if (!doc) throw new Error(`${file} is not a JSON object`)
  const map = doc[def.shape.key]
  if (!isRecord(map) || !(MCP_SERVER_NAME in map)) return false
  const { [MCP_SERVER_NAME]: _dropped, ...rest } = map
  doc[def.shape.key] = rest
  save(file, `${JSON.stringify(doc, null, 2)}\n`)
  return true
}

const LAUNCHER_NAMES = new Set(['genoffice', 'genoffice.cmd', 'genoffice.exe'])

/**
 * Any genoffice launcher, whatever install it came from (an older app path is
 * ours to update): the launcher scripts, or the app run as Node on genoffice.cjs.
 */
export function isGenofficeLauncher(command: string | null, args: string[] = []): boolean {
  if (command && LAUNCHER_NAMES.has(fileName(command))) return true
  return args.some((a) => fileName(a) === 'genoffice.cjs')
}

/** basename for either path flavour: Windows entries are read on any host */
function fileName(p: string): string {
  return p.slice(Math.max(p.lastIndexOf('/'), p.lastIndexOf('\\')) + 1).toLowerCase()
}

/** A symlinked config (stow, nix) is written through to its target so the link survives the rename. */
function save(file: string, content: string): void {
  const real = followLinks(file)
  mkdirSync(dirname(real), { recursive: true })
  writeOutput(real, content)
}

/**
 * Like realpath, but a dangling link resolves to where its target would be.
 * Each hop is taken from the real parent directory, so a relative link inside
 * a symlinked folder lands where the filesystem would take it.
 */
function followLinks(file: string): string {
  let cur = join(realizedPath(dirname(file)), basename(file))
  for (let i = 0; i < 32; i++) {
    let target: string
    try {
      if (!lstatSync(cur).isSymbolicLink()) return cur
      target = readlinkSync(cur)
    } catch {
      return cur
    }
    const next = isAbsolute(target) ? target : join(dirname(cur), target)
    cur = join(realizedPath(dirname(next)), basename(next))
  }
  return cur
}

function jsoncTwin(file: string): string {
  return file.replace(/\.json$/, '.jsonc')
}

function parseJsonObject(text: string): Record<string, unknown> | null {
  try {
    const raw: unknown = JSON.parse(text)
    return isRecord(raw) ? raw : null
  } catch {
    return null
  }
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return !!v && typeof v === 'object' && !Array.isArray(v)
}

const TOML_HEADER =
  /^\s*\[\s*mcp_servers\s*\.\s*(?:"genoffice"|'genoffice'|genoffice)\s*(\.[^\]]*)?\]/

/** `env` becomes the `[mcp_servers.genoffice.env]` sub-table Codex reads its `env` map from. */
function tomlTable(launch: McpLaunch): string {
  const head = `[mcp_servers.${MCP_SERVER_NAME}]\ncommand = ${JSON.stringify(launch.command)}\nargs = ${JSON.stringify(launch.args)}\n`
  const env = Object.entries(launch.env ?? {})
  if (!env.length) return head
  const body = env.map(([k, v]) => `${k} = ${JSON.stringify(v)}`).join('\n')
  return `${head}\n[mcp_servers.${MCP_SERVER_NAME}.env]\n${body}\n`
}

/** The `[mcp_servers.genoffice]` table and its sub-tables (`[mcp_servers.genoffice.env]`), as line ranges. */
function tomlTableLines(lines: string[]): Array<[number, number]> {
  const out: Array<[number, number]> = []
  for (let i = 0; i < lines.length; i++) {
    if (!TOML_HEADER.test(lines[i]!)) continue
    let end = i + 1
    while (end < lines.length && !/^\s*\[/.test(lines[end]!)) end++
    out.push([i, end])
    i = end - 1
  }
  return out
}

const meaningful = (lines: string[]) =>
  lines.map((l) => l.trim()).filter((l) => l && !l.startsWith('#'))

function tomlState(text: string, launch: McpLaunch): McpEntryState {
  const lines = text.split('\n')
  const ranges = tomlTableLines(lines)
  if (!ranges.length) return { status: 'absent', command: null }
  const own = ranges.find(([start]) => !TOML_HEADER.exec(lines[start]!)?.[1])
  const block = own ? lines.slice(own[0], own[1]) : []
  const command = tomlString(block, 'command')
  const written = ranges.flatMap(([s, e]) => meaningful(lines.slice(s, e)))
  if (written.join('\n') === meaningful(tomlTable(launch).split('\n')).join('\n')) {
    return { status: 'registered', command }
  }
  const args = tomlStrings(block, 'args')
  return { status: isGenofficeLauncher(command, args) ? 'stale' : 'occupied', command }
}

function tomlValue(block: string[], key: string): string | null {
  const line = block.find((l) => new RegExp(`^\\s*${key}\\s*=`).test(l))
  return line ? line.replace(new RegExp(`^\\s*${key}\\s*=\\s*`), '').trim() : null
}

function tomlString(block: string[], key: string): string | null {
  const m = /^("(?:[^"\\]|\\.)*"|'[^']*')\s*(?:#.*)?$/.exec(tomlValue(block, key) ?? '')
  if (!m) return null
  const lit = m[1]!
  if (lit.startsWith("'")) return lit.slice(1, -1)
  try {
    return JSON.parse(lit) as string
  } catch {
    return null
  }
}

function tomlStrings(block: string[], key: string): string[] {
  const raw = tomlValue(block, key)
  if (!raw) return []
  try {
    return strings(JSON.parse(raw.replace(/,\s*\]$/, ']')))
  } catch {
    return []
  }
}

function removeTomlTable(text: string): string {
  const lines = text.split('\n')
  const ranges = tomlTableLines(lines)
  if (!ranges.length) return text
  const keep: string[] = []
  let i = 0
  for (const [start, end] of ranges) {
    keep.push(...lines.slice(i, start))
    i = end
  }
  keep.push(...lines.slice(i))
  return keep.join('\n').replace(/\n{3,}/g, '\n\n')
}
