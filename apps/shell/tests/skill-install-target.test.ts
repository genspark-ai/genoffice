import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Guards for the skill-install IPC (apps/shell/src/main/integrations-ipc.ts):
 * a renderer-supplied `dir` may only be a directory the main process vouched
 * for, or a compromised renderer can make it write anywhere on disk.
 */
const h = vi.hoisted(() => ({
  handlers: new Map<string, (event: unknown, ...args: unknown[]) => unknown>(),
  installed: [] as string[],
  dialogResult: { canceled: true, filePaths: [] as string[] },
}))

vi.mock('electron', () => ({
  app: { isPackaged: false, getPath: () => tmpdir() },
  clipboard: { writeText: () => {} },
  dialog: {
    showOpenDialog: async () => h.dialogResult,
    showSaveDialog: async () => ({ canceled: true, filePath: undefined }),
  },
  ipcMain: {
    handle: (channel: string, fn: (event: unknown, ...args: unknown[]) => unknown) => {
      h.handlers.set(channel, fn)
    },
  },
}))

vi.mock('@genoffice/cli/agent-skills', () => {
  const skillsDir = '/home/me/.claude/skills'
  return {
    LEDGER_KEY: 'agentSkillInstalls',
    bundledSkillFrom: () => ({ name: 'genoffice', version: '1.0.0', text: 'x' }),
    buildSkillZip: async () => Buffer.alloc(0),
    detectAgents: () => [{ id: 'claude-code', label: 'Claude Code', skillsDir }],
    agentTarget: (id: string) =>
      id === 'claude-code' ? { id, label: 'Claude Code', skillsDir } : null,
    installSkill: (dir: string) => {
      h.installed.push(dir)
    },
    uninstallSkill: () => false,
    ledgerFromSettings: () => ({ version: '1.0.0', installs: {} }),
    readInstallState: (dir: string) => ({ status: 'missing', path: `${dir}/genoffice/SKILL.md` }),
  }
})

vi.mock('@genoffice/cli/install', () => ({ inspectCliLink: () => ({ linked: false }) }))

import { registerIntegrationsIpc } from '../src/main/integrations-ipc'
import { INTEGRATIONS_CHANNELS } from '../src/shared/integrations-api'

const AGENT_SKILLS_DIR = '/home/me/.claude/skills'

function invoke(channel: string, ...args: unknown[]): unknown {
  const handler = h.handlers.get(channel)
  if (!handler) throw new Error(`no handler registered for ${channel}`)
  return handler(null, ...args)
}

let dir = ''

beforeEach(() => {
  h.handlers.clear()
  h.installed.length = 0
  h.dialogResult = { canceled: true, filePaths: [] }
  dir = mkdtempSync(join(tmpdir(), 'genoffice-integrations-'))
  writeFileSync(join(dir, 'SKILL.md'), '---\nname: genoffice\n---\n')
  writeFileSync(join(dir, 'package.json'), JSON.stringify({ version: '0.0.0' }))
  registerIntegrationsIpc({
    settingsPath: () => join(dir, 'app-settings.json'),
    window: () => null,
    cliDir: dir,
    skillPath: join(dir, 'SKILL.md'),
    cliPackageJson: join(dir, 'package.json'),
  })
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('integrations:install-skill', () => {
  it('installs into an agent target resolved in the main process', () => {
    invoke(INTEGRATIONS_CHANNELS.installSkill, { agentId: 'claude-code' })
    expect(h.installed).toEqual([AGENT_SKILLS_DIR])
  })

  it('installs into a directory the user picked in the folder dialog', async () => {
    h.dialogResult = { canceled: false, filePaths: ['/Volumes/work/skills'] }
    const picked = await invoke(INTEGRATIONS_CHANNELS.pickSkillDir, 'Pick')
    expect(picked).toBe('/Volumes/work/skills')
    invoke(INTEGRATIONS_CHANNELS.installSkill, { dir: '/Volumes/work/skills' })
    expect(h.installed).toEqual(['/Volumes/work/skills'])
  })

  it('rejects a renderer-supplied directory the main process never vouched for', async () => {
    await expect(
      Promise.resolve().then(() =>
        invoke(INTEGRATIONS_CHANNELS.installSkill, { dir: '/etc/cron.d' }),
      ),
    ).rejects.toThrow('unknown skill target')
    expect(h.installed).toEqual([])
  })
})
