import { describe, expect, it, vi, beforeEach } from 'vitest'

/**
 * How the AI panel's skill doors treat plugins: `ai:skills-list` merges the
 * user's own skills with the enabled plugins' components (user first, the
 * user's entry winning a name collision), and `ai:skill-text` resolves a
 * `<plugin>:<name>` entry through that same merged listing — read-only. The
 * settings editor's write doors (`ai:skill-raw-text`, `ai:skill-save`) must
 * refuse a plugin entry: those are the real `skillRawText`/`writeSkill`, which
 * only know user skills.
 */

const handles = new Map<string, (...args: never[]) => unknown>()
vi.mock('electron', () => ({
  app: { getPath: () => '/app/userData', isPackaged: true, getAppPath: () => '/app' },
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      handles.set(channel, handler),
  },
  clipboard: {},
  dialog: {},
}))

const usable = vi.fn()
const readBody = vi.fn()
vi.mock('../src/main/imported-skills', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/main/imported-skills')>()),
  usableSkills: (...a: unknown[]) => usable(...a),
  readSkillBody: (...a: unknown[]) => readBody(...a),
}))

const pluginList = vi.fn()
const pluginBody = vi.fn()
vi.mock('../src/main/imported-plugins', () => ({
  pluginSkills: (...a: unknown[]) => pluginList(...a),
  readPluginSkillBody: (...a: unknown[]) => pluginBody(...a),
}))

const runnerRun = vi.fn()
vi.mock('../src/main/mcp/cli-runner', () => ({
  createCliRunner: () => ({ run: (...a: unknown[]) => runnerRun(...a) }),
}))

import { registerUserSkillsIpc } from '../src/main/user-cli'

const invoke = (channel: string, ...args: unknown[]) =>
  handles.get(channel)!({ sender: { id: 7 } } as never, ...args)

const USER_SKILL = {
  name: 'batch-pdf',
  description: 'user skill',
  path: '/app/userData/skills/batch-pdf/SKILL.md',
  relevance: { relevant: true, matched: ['pdf'] },
}
const PLUGIN_SKILL = {
  name: 'acme:helper',
  description: 'plugin skill',
  path: '/app/userData/plugins/acme/skills/helper/SKILL.md',
  relevance: { relevant: true, matched: ['docx'] },
  plugin: 'acme',
}

beforeEach(() => {
  handles.clear()
  usable.mockReset().mockResolvedValue([USER_SKILL])
  readBody.mockReset().mockReturnValue('the user body')
  pluginList.mockReset().mockResolvedValue([PLUGIN_SKILL])
  pluginBody.mockReset().mockReturnValue('the plugin body')
  runnerRun
    .mockReset()
    .mockResolvedValue({ ok: true, code: 0, json: undefined, stdout: 'o', stderr: '' })
  registerUserSkillsIpc({
    cliRunner: () => ({ run: runnerRun }) as never,
    folderRoots: () => ['/home/u/GenOffice'],
    senderFile: (id) => (id === 7 ? '/docs/open/report.docx' : undefined),
  })
})

describe('ai:skills-list', () => {
  it('lists the user skills first, then the plugin skills', async () => {
    const out = (await invoke('ai:skills-list')) as { name: string }[]
    expect(out.map((s) => s.name)).toEqual(['batch-pdf', 'acme:helper'])
    expect(pluginList).toHaveBeenCalledWith('/app/userData')
  })

  it('drops the plugin skill on a name collision — the user skill wins', async () => {
    pluginList.mockResolvedValue([{ ...PLUGIN_SKILL, name: 'batch-pdf' }])
    const out = (await invoke('ai:skills-list')) as { name: string; description: string }[]
    expect(out).toEqual([{ name: 'batch-pdf', description: 'user skill', relevant: true }])
  })
})

describe('ai:skill-text', () => {
  it('resolves a <plugin>:<name> entry through the merged listing, read-only', async () => {
    await expect(invoke('ai:skill-text', 'acme:helper')).resolves.toBe('the plugin body')
    expect(pluginBody).toHaveBeenCalledWith(PLUGIN_SKILL.path)
    expect(readBody).not.toHaveBeenCalled()
  })

  it('reads a user skill through the skill reader, not the plugin one', async () => {
    await expect(invoke('ai:skill-text', 'batch-pdf')).resolves.toBe('the user body')
    expect(readBody).toHaveBeenCalledWith(USER_SKILL.path)
    expect(pluginBody).not.toHaveBeenCalled()
  })

  it('resolves a collision to the user skill', async () => {
    pluginList.mockResolvedValue([{ ...PLUGIN_SKILL, name: 'batch-pdf' }])
    await expect(invoke('ai:skill-text', 'batch-pdf')).resolves.toBe('the user body')
    expect(readBody).toHaveBeenCalledWith(USER_SKILL.path)
    expect(pluginBody).not.toHaveBeenCalled()
  })

  it('still refuses a name the listing never showed', async () => {
    await expect(invoke('ai:skill-text', 'acme:missing')).rejects.toThrow(
      'no installed skill named "acme:missing"',
    )
    await expect(invoke('ai:skill-text', '/etc/passwd')).rejects.toThrow('no installed skill named')
  })
})

describe('the settings editor doors keep plugin skills out', () => {
  it('ai:skill-raw-text refuses a plugin skill: it is not a user skill', async () => {
    await expect(invoke('ai:skill-raw-text', 'acme:helper')).rejects.toThrow('unknown skill')
  })

  it('ai:skill-save refuses a plugin skill name', () => {
    // a sync handler: the refusal is a throw, exactly what the renderer sees rejected
    expect(() =>
      invoke('ai:skill-save', {
        name: 'acme:helper',
        description: 'd',
        body: 'b',
      }),
    ).toThrow('cannot be used as a skill name')
  })
})
