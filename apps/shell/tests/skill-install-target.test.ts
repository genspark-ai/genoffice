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
/** a fake home, so nothing here can touch the real one */
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
    recommendedSkillsDir: join(dir, 'recommended'),
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

  it('lists a skill in our folder, and copies nothing', async () => {
    const path = ours('pdf-to-html')
    const before = readdirSync(join(userData, 'skills')).sort()
    const found = (await invoke(INTEGRATIONS_CHANNELS.listSkills)) as {
      path: string
      name: string
      relevance: { relevant: boolean }
    }[]
    const skill = found.find((s) => s.path === path)
    expect(skill?.name).toBe('pdf-to-html')
    expect(skill?.relevance.relevant).toBe(true)
    expect(readdirSync(join(userData, 'skills')).sort()).toEqual(before)
  })

  it('does not look in a coding agent directory', async () => {
    agents('pptx-builder')
    const found = (await invoke(INTEGRATIONS_CHANNELS.listSkills)) as { name: string }[]
    expect(found.map((s) => s.name)).toEqual([])
  })

  it('reads the body of a skill it just listed', () => {
    const path = ours('pdf-to-html')
    expect(invoke(INTEGRATIONS_CHANNELS.skillBody, path)).toBe('Step one.\n')
  })

  it('refuses a path that is not one of those skills', () => {
    ours('pdf-to-html')
    const outside = join(userData, 'secrets.txt')
    writeFileSync(outside, 'private')
    expect(() => invoke(INTEGRATIONS_CHANNELS.skillBody, outside)).toThrow('unknown skill')
    expect(() =>
      invoke(INTEGRATIONS_CHANNELS.skillBody, join(userData, 'skills', '..', 'secrets.txt')),
    ).toThrow('unknown skill')
    // an agent's SKILL.md is a real skill file, but not one this app listed
    expect(() => invoke(INTEGRATIONS_CHANNELS.skillBody, agents('theirs'))).toThrow('unknown skill')
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

describe('integrations:import-skill-folder / -zip', () => {
  const plantSrc = (name: string, description: string, body: string) => {
    const src = join(fakeHome, 'Downloads', name)
    mkdirSync(src, { recursive: true })
    writeFileSync(
      join(src, 'SKILL.md'),
      `---\nname: ${name}\ndescription: ${description}\n---\n\n${body}\n`,
    )
    return src
  }

  it('copies the folder picked in the dialog into our folder', async () => {
    const src = plantSrc('pptx-builder', 'Build a pptx', 'Do it.')
    writeFileSync(join(src, 'grid.md'), 'layout rules')
    h.dialogResult = { canceled: false, filePaths: [src] }

    const imported = (await invoke(INTEGRATIONS_CHANNELS.importSkillFolder, 'Pick')) as {
      name: string
    }
    expect(imported.name).toBe('pptx-builder')
    expect(existsSync(join(userData, 'skills', 'pptx-builder', 'grid.md'))).toBe(true)
    const found = (await invoke(INTEGRATIONS_CHANNELS.listSkills)) as { name: string }[]
    expect(found.map((s) => s.name)).toEqual(['pptx-builder'])
  })

  it('returns null and copies nothing when the dialog is cancelled', async () => {
    plantSrc('pptx-builder', 'Build a pptx', 'Do it.')
    expect(await invoke(INTEGRATIONS_CHANNELS.importSkillFolder, 'Pick')).toBeNull()
    expect(await invoke(INTEGRATIONS_CHANNELS.importSkillZip, 'Pick')).toBeNull()
    expect(existsSync(join(userData, 'skills'))).toBe(false)
  })

  it('takes the source only from the dialog, never from the renderer', async () => {
    // the renderer's argument is a dialog title; a path there is just a title
    const src = plantSrc('sneaky', 'Build a pptx', 'Do it.')
    h.dialogResult = { canceled: true, filePaths: [] }
    expect(await invoke(INTEGRATIONS_CHANNELS.importSkillFolder, src)).toBeNull()
    expect(existsSync(join(userData, 'skills', 'sneaky'))).toBe(false)
  })

  it('refuses to replace a skill already in our folder', async () => {
    const at = join(userData, 'skills', 'shared')
    mkdirSync(at, { recursive: true })
    writeFileSync(join(at, 'SKILL.md'), '---\nname: shared\ndescription: Mine\n---\n\nmine\n')
    h.dialogResult = { canceled: false, filePaths: [plantSrc('shared', 'Theirs', 'theirs')] }

    await expect(invoke(INTEGRATIONS_CHANNELS.importSkillFolder, 'Pick')).rejects.toThrow(
      'already there',
    )
    expect(invoke(INTEGRATIONS_CHANNELS.skillBody, join(at, 'SKILL.md'))).toBe('mine\n')
  })
})

describe('integrations:list-recommended-skills / install-recommended-skill', () => {
  const ship = (name: string, version = '1.0.0') => {
    const at = join(dir, 'recommended', name)
    mkdirSync(at, { recursive: true })
    writeFileSync(
      join(at, 'SKILL.md'),
      `---\nname: ${name}\ndescription: Write a docx report\nmetadata:\n  version: ${version}\n---\n\nShipped.\n`,
    )
    return join(at, 'SKILL.md')
  }

  it('lists shipped skills without installing them, and lets the body be read', async () => {
    const path = ship('business-report')
    const list = (await invoke(INTEGRATIONS_CHANNELS.listRecommendedSkills)) as {
      name: string
      state: string
      version: string
    }[]
    expect(list).toEqual([
      expect.objectContaining({ name: 'business-report', state: 'missing', version: '1.0.0' }),
    ])
    expect(existsSync(join(userData, 'skills'))).toBe(false)
    expect(invoke(INTEGRATIONS_CHANNELS.skillBody, path)).toBe('Shipped.\n')
  })

  it('installs by name into our folder and reports it installed afterwards', async () => {
    ship('business-report')
    const got = (await invoke(
      INTEGRATIONS_CHANNELS.installRecommendedSkill,
      'business-report',
    )) as { name: string }
    expect(got.name).toBe('business-report')
    expect(existsSync(join(userData, 'skills', 'business-report', 'SKILL.md'))).toBe(true)
    const list = (await invoke(INTEGRATIONS_CHANNELS.listRecommendedSkills)) as { state: string }[]
    expect(list[0]?.state).toBe('installed')
  })

  it('refuses a name it does not ship, whatever the renderer sends', () => {
    ship('business-report')
    for (const bad of ['other', '../recommended/business-report', 42, null]) {
      expect(() => invoke(INTEGRATIONS_CHANNELS.installRecommendedSkill, bad)).toThrow(
        'unknown skill',
      )
    }
  })
})
