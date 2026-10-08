import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Guards for the skills IPC (apps/shell/src/main/integrations-ipc.ts): a
 * renderer-supplied `dir` may only be a directory the main process vouched
 * for, or a compromised renderer can make it write anywhere on disk; and a
 * renderer-supplied `path` may only be a SKILL.md out of those same roots, or
 * it becomes an arbitrary file read.
 */
const h = vi.hoisted(() => ({
  handlers: new Map<string, (event: unknown, ...args: unknown[]) => unknown>(),
  installed: [] as string[],
  dialogResult: { canceled: true, filePaths: [] as string[] },
}))

vi.mock('electron', () => ({
  // the skills handlers read the app's userData for their root; everything else
  // (downloads) is not exercised here
  app: {
    isPackaged: false,
    getPath: (name: string) => (name === 'userData' ? userData : tmpdir()),
  },
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
/** stands in for app.getPath('userData') — where the skills root now lives, so
 *  that no imported SKILL.md lands among the user's documents */
let userData = ''
/** a fake home so the agent-directory half of the scan cannot read the real one */
let fakeHome = ''
let realHome: string | undefined

beforeEach(() => {
  h.handlers.clear()
  h.installed.length = 0
  h.dialogResult = { canceled: true, filePaths: [] }
  dir = mkdtempSync(join(tmpdir(), 'genoffice-integrations-'))
  userData = mkdtempSync(join(tmpdir(), 'genoffice-userdata-'))
  fakeHome = mkdtempSync(join(tmpdir(), 'genoffice-fakehome-'))
  realHome = process.env.HOME
  process.env.HOME = fakeHome
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
  rmSync(userData, { recursive: true, force: true })
  rmSync(fakeHome, { recursive: true, force: true })
  if (realHome === undefined) delete process.env.HOME
  else process.env.HOME = realHome
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

/**
 * These read through to a fake home, so every assertion is about the skills the
 * test planted rather than about whatever this machine really has installed.
 */
describe('integrations:list-skills', () => {
  const plant = (root: string, name: string) => {
    const at = join(root, name)
    mkdirSync(at, { recursive: true })
    writeFileSync(
      join(at, 'SKILL.md'),
      `---\nname: ${name}\ndescription: Turn a pdf into html\n---\n\nStep one.\n`,
    )
    return join(at, 'SKILL.md')
  }
  const ours = (name: string) => plant(join(userData, 'skills'), name)
  const agents = (name: string) => plant(join(fakeHome, '.agents', 'skills'), name)

  it('offers a skill under the save directory, and copies nothing', () => {
    const path = ours('pdf-to-html')
    const before = readdirSync(join(userData, 'skills')).sort()
    const found = invoke(INTEGRATIONS_CHANNELS.listSkills) as { path: string }[]
    const skill = found.find((s) => s.path === path)
    expect(skill).toMatchObject({ name: 'pdf-to-html', source: 'genoffice' })
    expect(skill.relevance.relevant).toBe(true)
    // listing is a menu: the folder is exactly as it was afterwards
    expect(readdirSync(join(userData, 'skills')).sort()).toEqual(before)
  })

  it('offers an agent skill as a candidate, without importing it', () => {
    const path = agents('pptx-builder')
    const found = invoke(INTEGRATIONS_CHANNELS.listSkills) as {
      path: string
      source: string
      agent?: string
    }[]
    expect(found.find((s) => s.path === path)).toMatchObject({
      source: 'agent',
      agent: 'agents',
    })
    // it did not move itself into our folder just by being seen
    expect(existsSync(join(userData, 'skills', 'pptx-builder'))).toBe(false)
  })

  it('reads the body of a skill it just offered', () => {
    const path = agents('pdf-to-html')
    expect(invoke(INTEGRATIONS_CHANNELS.skillBody, path)).toBe('Step one.\n')
  })

  it('refuses a path that is not one of those skills', () => {
    ours('pdf-to-html')
    const outside = join(userData, 'secrets.txt')
    writeFileSync(outside, 'private')
    expect(() => invoke(INTEGRATIONS_CHANNELS.skillBody, outside)).toThrow('unknown skill')
    // the same file, reached by climbing out of the skills root
    expect(() =>
      invoke(INTEGRATIONS_CHANNELS.skillBody, join(userData, 'skills', '..', 'secrets.txt')),
    ).toThrow('unknown skill')
    expect(() => invoke(INTEGRATIONS_CHANNELS.skillBody, '/etc/passwd')).toThrow('unknown skill')
    expect(() => invoke(INTEGRATIONS_CHANNELS.skillBody, null)).toThrow('unknown skill')
  })

  it('refuses another file that happens to sit in a skills folder', () => {
    const at = join(userData, 'skills', 'pdf-to-html')
    mkdirSync(at, { recursive: true })
    const note = join(at, 'reference.md')
    writeFileSync(note, 'a helper note')
    expect(() => invoke(INTEGRATIONS_CHANNELS.skillBody, note)).toThrow('unknown skill')
  })
})

describe('integrations:import-skill', () => {
  it('copies a picked candidate into our folder, and reports it as ours', () => {
    const src = join(fakeHome, '.agents', 'skills', 'pptx-builder')
    mkdirSync(src, { recursive: true })
    writeFileSync(
      join(src, 'SKILL.md'),
      '---\nname: pptx-builder\ndescription: Build a pptx\n---\n\nDo it.\n',
    )
    writeFileSync(join(src, 'grid.md'), 'layout rules')

    const imported = invoke(INTEGRATIONS_CHANNELS.importSkill, join(src, 'SKILL.md'))
    expect(imported).toMatchObject({ name: 'pptx-builder', source: 'genoffice' })
    expect(existsSync(join(userData, 'skills', 'pptx-builder', 'grid.md'))).toBe(true)

    // and on the next listing it is ours, not a candidate
    const found = invoke(INTEGRATIONS_CHANNELS.listSkills) as {
      name: string
      source: string
    }[]
    expect(found.find((s) => s.name === 'pptx-builder')?.source).toBe('genoffice')
  })

  it('refuses to import from a path that is not a skill', () => {
    expect(() => invoke(INTEGRATIONS_CHANNELS.importSkill, '/etc/hosts')).toThrow('unknown skill')
    expect(() => invoke(INTEGRATIONS_CHANNELS.importSkill, undefined)).toThrow('unknown skill')
  })

  it('refuses to replace a skill already in our folder', () => {
    const at = join(userData, 'skills', 'shared')
    mkdirSync(at, { recursive: true })
    writeFileSync(join(at, 'SKILL.md'), '---\nname: shared\ndescription: Mine\n---\n\nmine\n')
    const src = join(fakeHome, '.agents', 'skills', 'shared')
    mkdirSync(src, { recursive: true })
    writeFileSync(join(src, 'SKILL.md'), '---\nname: shared\ndescription: Theirs\n---\n\ntheirs\n')

    expect(() => invoke(INTEGRATIONS_CHANNELS.importSkill, join(src, 'SKILL.md'))).toThrow(
      'already there',
    )
    expect(invoke(INTEGRATIONS_CHANNELS.skillBody, join(at, 'SKILL.md'))).toBe('mine\n')
  })
})
