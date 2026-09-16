import { mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import {
  agentIds,
  agentTarget,
  bundledSkillFrom,
  detectAgents,
  installSkill,
  LEDGER_KEY,
  ledgerFromSettings,
  readInstallState,
  type AgentId,
  type AgentTarget,
  type BundledSkill,
  type SkillInstallState,
  type SkillLedger,
} from '../agent-skills'
import { flagBool, flagString } from '../args'
import { genofficeUserDataDir } from '../gui'
import type { CommandContext, CommandDef } from '../registry'
import { bundledSkillPath } from '../resources'
import { CliError, EXIT, type Warning } from '../result'
import { didYouMean } from '../suggest'

/**
 * The headless twin of Settings → Integrations: list the coding agents on this
 * machine and copy the bundled skill into their skills directory. Writes go
 * into the same ledger the app keeps (userData/app-settings.json), so the pane
 * shows a CLI install as "installed", not "foreign".
 */
export const skillCommand: CommandDef = {
  name: 'skill',
  summary:
    'List the coding agents found on this machine and install or update the bundled genoffice skill for them.',
  usage: 'skill list | skill install <agent|all> [--dir <path>] [--force] | skill path',
  options: [
    { name: 'list', description: 'same as `skill list`' },
    {
      name: 'dir',
      value: 'path',
      description: 'install: write into this skills directory instead',
    },
    {
      name: 'force',
      description:
        'install: also write when the agent is not detected, the installed copy is newer, or it was edited by hand',
    },
  ],
  async run(args, ctx) {
    const sub = flagBool(args, 'list') ? 'list' : (args.positionals[0] ?? 'list')
    switch (sub) {
      case 'list':
        return list(ctx)
      case 'path': {
        const path = skillFile()
        return { summary: path, detail: { path, version: bundled().version } }
      }
      case 'install':
        return install(args.positionals[1], flagString(args, 'dir'), flagBool(args, 'force'), ctx)
      default:
        throw new CliError(
          EXIT.usage,
          `unknown skill subcommand: ${sub}`,
          { supported: ['list', 'install', 'path'] },
          {
            reason: 'invalid_argument',
            suggestion: didYouMean(sub, ['list', 'install', 'path'])
              ? `did you mean \`${didYouMean(sub, ['list', 'install', 'path'])}\`?`
              : 'use `skill list`, `skill install <agent|all>` or `skill path`',
          },
        )
    }
  },
}

interface AgentRow {
  /** null for a folder given with --dir */
  agent: AgentId | null
  label: string
  detected: boolean
  skills_dir: string
  status: SkillInstallState['status']
  installed_version: string | null
  up_to_date: boolean
}

function list(ctx: CommandContext) {
  const skill = bundled()
  const ledger = readLedger(ctx.env)
  const detected = new Set(detectAgents(ctx.env, home(ctx.env)).map((a) => a.id))
  const agents: AgentRow[] = agentIds().map((id) => {
    const target = agentTarget(id, ctx.env, home(ctx.env))!
    return row(target, detected.has(id), readInstallState(target.skillsDir, skill, ledger))
  })
  const current = agents.filter((a) => a.detected && a.up_to_date).length
  const found = agents.filter((a) => a.detected)
  return {
    summary: found.length
      ? `${found.length} agent(s) detected, ${current} with skill ${skill.version}: ${found
          .map((a) => `${a.agent} ${a.installed_version ?? '-'}`)
          .join(', ')}`
      : 'no coding agent detected on this machine',
    detail: { bundled_version: skill.version, skill_path: skillFile(), agents },
  }
}

function install(
  name: string | undefined,
  dir: string | undefined,
  force: boolean,
  ctx: CommandContext,
) {
  if (!name && !dir) {
    throw new CliError(
      EXIT.usage,
      'skill install needs an agent name, `all`, or --dir',
      undefined,
      {
        reason: 'missing_argument',
        suggestion: `agents: ${agentIds().join(', ')}`,
      },
    )
  }
  const skill = bundled()
  const ledger = readLedger(ctx.env)
  const detected = new Set(detectAgents(ctx.env, home(ctx.env)).map((a) => a.id))
  const targets: Array<{ target: AgentTarget | null; skillsDir: string; detected: boolean }> = []
  if (dir) {
    targets.push({ target: null, skillsDir: resolve(ctx.cwd, dir), detected: true })
  } else if (name === 'all') {
    for (const id of agentIds()) {
      if (!detected.has(id)) continue
      const t = agentTarget(id, ctx.env, home(ctx.env))!
      targets.push({ target: t, skillsDir: t.skillsDir, detected: true })
    }
    if (!targets.length) {
      throw new CliError(EXIT.usage, 'no coding agent detected on this machine', undefined, {
        reason: 'unsupported',
        suggestion: 'install into a folder with --dir <skills directory>',
      })
    }
  } else {
    const t = agentTarget(name as AgentId, ctx.env, home(ctx.env))
    if (!t) {
      const guess = didYouMean(name!, agentIds())
      throw new CliError(
        EXIT.usage,
        `unknown agent: ${name}`,
        { supported: [...agentIds(), 'all'] },
        {
          reason: 'unsupported',
          suggestion: guess
            ? `did you mean \`${guess}\`?`
            : `agents: ${agentIds().join(', ')}, or all`,
        },
      )
    }
    if (!detected.has(t.id) && !force) {
      throw new CliError(
        EXIT.usage,
        `${t.label} is not installed on this machine (${t.skillsDir} would be created)`,
        { agent: t.id, skills_dir: t.skillsDir },
        { reason: 'unsupported', suggestion: 'repeat with --force to create the folder anyway' },
      )
    }
    targets.push({ target: t, skillsDir: t.skillsDir, detected: detected.has(t.id) })
  }

  // every target is checked before the first write, so a refusal never leaves a half-done batch
  for (const t of targets) {
    const before = readInstallState(t.skillsDir, skill, ledger)
    const blocked = refusal(before, force)
    if (blocked) {
      throw new CliError(
        EXIT.file,
        `${t.target?.id ?? t.skillsDir}: ${blocked} (${before.path})`,
        { agent: t.target?.id ?? null, skills_dir: t.skillsDir, status: before.status },
        { reason: 'output_exists', suggestion: refusalHint(before.status) },
      )
    }
  }
  const warnings: Warning[] = []
  const written: string[] = []
  const rows: Array<AgentRow & { action: 'written' | 'unchanged' }> = []
  let dirty = false
  for (const t of targets) {
    const label = t.target?.id ?? t.skillsDir
    let action: 'written' | 'unchanged' = 'unchanged'
    if (readInstallState(t.skillsDir, skill, ledger).status === 'installed') {
      warnings.push({
        code: 'already_up_to_date',
        message: `${label}: skill ${skill.version} already installed`,
      })
    } else {
      installSkill(t.skillsDir, skill, ledger, 'cli')
      dirty = true
      action = 'written'
      written.push(label)
    }
    rows.push({
      ...row(t.target, t.detected, readInstallState(t.skillsDir, skill, ledger)),
      action,
    })
  }
  if (dirty) writeLedger(ctx.env, ledger)
  return {
    summary: written.length
      ? `installed skill ${skill.version} for ${written.join(', ')}`
      : `skill ${skill.version} already up to date`,
    ...(warnings.length ? { warnings } : {}),
    detail: { bundled_version: skill.version, agents: rows },
  }
}

function refusal(state: SkillInstallState, force: boolean): string | null {
  if (state.status === 'occupied') return 'the genoffice folder there holds something else'
  if (force) return null
  if (state.status === 'newer')
    return `a newer skill ${state.installedVersion} is installed; not downgrading`
  if (state.status === 'modified') return 'the installed skill was edited by hand; not overwriting'
  return null
}

function refusalHint(status: SkillInstallState['status']): string {
  if (status === 'occupied') return 'move that folder away, or install with --dir elsewhere'
  return 'repeat with --force to replace it'
}

function row(target: AgentTarget | null, detected: boolean, state: SkillInstallState): AgentRow {
  return {
    agent: target?.id ?? null,
    label: target?.label ?? state.path,
    detected,
    skills_dir: target?.skillsDir ?? state.path.replace(/[\\/]genoffice[\\/]SKILL\.md$/, ''),
    status: state.status,
    installed_version: state.installedVersion ?? null,
    up_to_date: state.status === 'installed' || state.status === 'newer',
  }
}

/** test seam: the home the agent dotfolders are probed under */
function home(env: NodeJS.ProcessEnv): string {
  return env.GENOFFICE_HOME || homedir()
}

function skillFile(): string {
  const path = bundledSkillPath()
  if (!path)
    throw new CliError(EXIT.app, 'the bundled skill (skills/genoffice/SKILL.md) was not found')
  return path
}

function bundled(): BundledSkill {
  return bundledSkillFrom(readFileSync(skillFile()))
}

function settingsPath(env: NodeJS.ProcessEnv): string {
  return join(genofficeUserDataDir(env), 'app-settings.json')
}

function readSettings(env: NodeJS.ProcessEnv): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath(env), 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw as Record<string, unknown>
  } catch {}
  return {}
}

function readLedger(env: NodeJS.ProcessEnv): SkillLedger {
  return ledgerFromSettings(readSettings(env))
}

/** Same file and key the app uses; replaced atomically so a concurrent app write never sees a half file. */
function writeLedger(env: NodeJS.ProcessEnv, ledger: SkillLedger): void {
  const path = settingsPath(env)
  const next = { ...readSettings(env), [LEDGER_KEY]: ledger }
  const tmp = `${path}.${process.pid}.tmp`
  mkdirSync(dirname(path), { recursive: true })
  try {
    writeFileSync(tmp, JSON.stringify(next, null, 2), { encoding: 'utf8', flag: 'wx' })
    renameSync(tmp, path)
  } catch (err) {
    try {
      unlinkSync(tmp)
    } catch {}
    throw err
  }
}
