import { createHash } from 'node:crypto'
import {
  existsSync,
  lstatSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
  type Stats,
} from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'

/**
 * Installing the bundled `genoffice` skill into the coding agents found on this
 * machine, shared by Settings → Integrations and `genoffice skill`. Every write
 * is a click in that pane or an explicit command; nothing installs on its own.
 * Pure functions over an injected home / env so the state table is
 * unit-testable without touching the real dotfolders.
 */

export const SKILL_NAME = 'genoffice'

/** app-settings.json key under which every install this machine's GenOffice wrote is remembered */
export const LEDGER_KEY = 'agentSkillInstalls'

export type AgentId =
  'claude-code' | 'codex' | 'cursor' | 'gemini' | 'copilot' | 'opencode' | 'windsurf'

export interface AgentTarget {
  id: AgentId
  /** product name, shown as is (not translated) */
  label: string
  /** global skills directory the agent scans */
  skillsDir: string
}

export type SkillInstallStatus =
  /** no genoffice/ folder in the skills directory */
  | 'missing'
  /** written by this app, same version and bytes as the bundled skill */
  | 'installed'
  /** written by this app, older than the bundled skill */
  | 'outdated'
  /** written by this app, edited since (bytes differ from what was written) */
  | 'modified'
  /** our skill by front matter, but not written by this app (other host, npx, by hand) */
  | 'foreign'
  /** a newer skill version than the one bundled, whoever wrote it */
  | 'newer'
  /** genoffice/ exists but does not hold our skill */
  | 'occupied'

export interface SkillInstallState {
  status: SkillInstallStatus
  /** `<skillsDir>/genoffice/SKILL.md` */
  path: string
  installedVersion?: string
  /** installed version is older than the bundled one (foreign rows offer an update then) */
  older?: boolean
}

interface AgentDef {
  id: AgentId
  label: string
  /** the folder whose presence means the tool is installed */
  probe: (env: NodeJS.ProcessEnv, home: string) => string
  /** the skills directory, relative to the probe folder unless absolute */
  skills: (probe: string) => string
}

const AGENTS: readonly AgentDef[] = [
  {
    id: 'claude-code',
    label: 'Claude Code',
    probe: (env, home) => env.CLAUDE_CONFIG_DIR || join(home, '.claude'),
    skills: (p) => join(p, 'skills'),
  },
  {
    id: 'codex',
    label: 'Codex',
    probe: (env, home) => env.CODEX_HOME || join(home, '.codex'),
    skills: (p) => join(p, 'skills'),
  },
  {
    id: 'cursor',
    label: 'Cursor',
    probe: (_env, home) => join(home, '.cursor'),
    skills: (p) => join(p, 'skills'),
  },
  {
    id: 'gemini',
    label: 'Gemini CLI',
    probe: (_env, home) => join(home, '.gemini'),
    skills: (p) => join(p, 'skills'),
  },
  {
    id: 'copilot',
    label: 'GitHub Copilot',
    probe: (_env, home) => join(home, '.copilot'),
    skills: (p) => join(p, 'skills'),
  },
  {
    id: 'opencode',
    label: 'OpenCode',
    probe: (env, home) => join(env.XDG_CONFIG_HOME || join(home, '.config'), 'opencode'),
    skills: (p) => join(p, 'skills'),
  },
  {
    id: 'windsurf',
    label: 'Windsurf',
    probe: (_env, home) => join(home, '.codeium', 'windsurf'),
    skills: (p) => join(p, 'skills'),
  },
]

export function agentIds(): AgentId[] {
  return AGENTS.map((a) => a.id)
}

/** Agents whose dotfolder exists; the folder standing in for "installed" is a heuristic, so a stale one only adds a row. */
export function detectAgents(
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): AgentTarget[] {
  const out: AgentTarget[] = []
  for (const a of AGENTS) {
    const probe = a.probe(env, home)
    if (isDir(probe)) out.push({ id: a.id, label: a.label, skillsDir: a.skills(probe) })
  }
  return out
}

/** The dotfolder probed for `id` (env overrides applied): where the agent keeps its own config files. */
export function agentConfigDir(
  id: AgentId,
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): string | null {
  const a = AGENTS.find((x) => x.id === id)
  return a ? a.probe(env, home) : null
}

export function agentTarget(
  id: AgentId,
  env: NodeJS.ProcessEnv = process.env,
  home = homedir(),
): AgentTarget | null {
  const a = AGENTS.find((x) => x.id === id)
  if (!a) return null
  return { id: a.id, label: a.label, skillsDir: a.skills(a.probe(env, home)) }
}

export interface SkillFrontmatter {
  name: string
  version: string
}

/** `name` and `metadata.version` from the YAML front matter; enough to tell our skill (and its version) from anything else. */
export function parseSkillFrontmatter(text: string): SkillFrontmatter | null {
  const m = /^---\r?\n([\s\S]*?)\r?\n---/.exec(text)
  if (!m) return null
  const name = /^name:\s*(\S+)\s*$/m.exec(m[1]!)?.[1]
  const version = /^\s+version:\s*['"]?(\d+\.\d+\.\d+)['"]?\s*$/m.exec(m[1]!)?.[1] ?? ''
  return name ? { name, version } : null
}

export function sha256(bytes: Uint8Array | string): string {
  return createHash('sha256').update(bytes).digest('hex')
}

/** semver-ish: numeric major.minor.patch; anything unparsable sorts lowest */
export function compareVersions(a: string, b: string): -1 | 0 | 1 {
  const pa = parts(a)
  const pb = parts(b)
  for (let i = 0; i < 3; i++) {
    if (pa[i]! < pb[i]!) return -1
    if (pa[i]! > pb[i]!) return 1
  }
  return 0
}

function parts(v: string): number[] {
  const m = /^(\d+)\.(\d+)\.(\d+)/.exec(v.trim())
  return m ? [Number(m[1]), Number(m[2]), Number(m[3])] : [-1, -1, -1]
}

export interface LedgerEntry {
  version: string
  sha256: string
  at: string
  channel: 'app' | 'zip' | 'cli'
}

/** absolute SKILL.md path → what this app wrote there */
export type SkillLedger = Record<string, LedgerEntry>

export function ledgerFromSettings(settings: Record<string, unknown>): SkillLedger {
  const raw = settings[LEDGER_KEY]
  return raw && typeof raw === 'object' && !Array.isArray(raw) ? { ...(raw as SkillLedger) } : {}
}

export interface BundledSkill {
  bytes: Uint8Array
  text: string
  version: string
  sha256: string
}

export function bundledSkillFrom(bytes: Uint8Array): BundledSkill {
  const text = Buffer.from(bytes).toString('utf-8')
  const fm = parseSkillFrontmatter(text)
  return { bytes, text, version: fm?.version ?? '0.0.0', sha256: sha256(bytes) }
}

export function skillFilePath(skillsDir: string): string {
  return join(skillsDir, SKILL_NAME, 'SKILL.md')
}

export function readInstallState(
  skillsDir: string,
  bundled: BundledSkill,
  ledger: SkillLedger,
): SkillInstallState {
  const path = skillFilePath(skillsDir)
  const dir = join(skillsDir, SKILL_NAME)
  if (!existsSync(dir)) return { status: 'missing', path }
  let text: string
  try {
    text = readFileSync(path, 'utf-8')
  } catch {
    return { status: 'occupied', path }
  }
  const fm = parseSkillFrontmatter(text)
  if (!fm || fm.name !== SKILL_NAME) return { status: 'occupied', path }
  const installedVersion = fm.version
  const cmp = compareVersions(installedVersion, bundled.version)
  if (cmp > 0) return { status: 'newer', path, installedVersion }
  const ours = ledger[path]
  if (!ours) return { status: 'foreign', path, installedVersion, older: cmp < 0 }
  if (sha256(text) !== ours.sha256)
    return { status: 'modified', path, installedVersion, older: cmp < 0 }
  if (cmp < 0) return { status: 'outdated', path, installedVersion, older: true }
  return { status: 'installed', path, installedVersion }
}

/** Copy the bundled SKILL.md byte for byte and remember the write. Returns the file path. */
export function installSkill(
  skillsDir: string,
  bundled: BundledSkill,
  ledger: SkillLedger,
  channel: LedgerEntry['channel'] = 'app',
): string {
  const path = skillFilePath(skillsDir)
  let existing: Stats | undefined
  try {
    existing = lstatSync(path)
  } catch {}
  if (existing?.isSymbolicLink()) {
    throw new Error(`${path} is a symlink; refusing to overwrite whatever it points at`)
  }
  mkdirSync(join(skillsDir, SKILL_NAME), { recursive: true })
  writeFileSync(path, bundled.bytes)
  ledger[path] = {
    version: bundled.version,
    sha256: bundled.sha256,
    at: new Date().toISOString(),
    channel,
  }
  return path
}

/** Remove only what this app wrote; a folder holding other files keeps them. */
export function uninstallSkill(skillsDir: string, ledger: SkillLedger): boolean {
  const path = skillFilePath(skillsDir)
  if (!ledger[path]) return false
  rmSync(path, { force: true })
  const dir = join(skillsDir, SKILL_NAME)
  try {
    if (readdirSync(dir).length === 0) rmSync(dir, { recursive: true })
  } catch {}
  delete ledger[path]
  return true
}

/** `genoffice/SKILL.md` inside a zip: the layout claude.ai and the desktop apps accept for an uploaded skill. */
export async function buildSkillZip(bundled: BundledSkill): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(`${SKILL_NAME}/SKILL.md`, bundled.bytes, { date: new Date(0) })
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function isDir(path: string): boolean {
  try {
    return statSync(path).isDirectory()
  } catch {
    return false
  }
}
