import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

const REPO = resolve(__dirname, '../../..')
const bundledVersion = /^\s+version:\s*(\S+)/m.exec(
  readFileSync(join(REPO, 'skills/genoffice/SKILL.md'), 'utf-8'),
)![1]!

function fakeMachine(agents: string[]) {
  const home = tempDir()
  for (const dot of agents) mkdirSync(join(home, dot), { recursive: true })
  const userData = join(home, 'userData')
  mkdirSync(userData)
  return { home, env: { GENOFFICE_HOME: home, GENOFFICE_USER_DATA: userData } }
}

describe('genoffice skill', () => {
  it('path prints the bundled SKILL.md', async () => {
    const r = await run(['skill', 'path', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ version: bundledVersion })
    expect(existsSync(r.json().detail.path)).toBe(true)
  })

  it('list shows every agent with detection and install state', async () => {
    const m = fakeMachine(['.claude', '.cursor'])
    const r = await run(['skill', 'list', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    const agents = r.json().detail.agents as Array<Record<string, unknown>>
    expect(agents.map((a) => a.agent)).toContain('codex')
    const claude = agents.find((a) => a.agent === 'claude-code')!
    expect(claude).toMatchObject({
      detected: true,
      status: 'missing',
      installed_version: null,
      up_to_date: false,
      skills_dir: join(m.home, '.claude', 'skills'),
    })
    expect(agents.find((a) => a.agent === 'codex')).toMatchObject({ detected: false })
    expect(r.json().detail.bundled_version).toBe(bundledVersion)
    expect((await run(['skill', '--list'], { env: m.env })).stdout).toContain('2 agent(s) detected')
  })

  it('install writes the skill, records it in the app ledger and is idempotent', async () => {
    const m = fakeMachine(['.claude'])
    const first = await run(['skill', 'install', 'claude-code', '--json'], { env: m.env })
    expect(first.code).toBe(0)
    const path = join(m.home, '.claude', 'skills', 'genoffice', 'SKILL.md')
    expect(readFileSync(path, 'utf-8')).toBe(
      readFileSync(join(REPO, 'skills/genoffice/SKILL.md'), 'utf-8'),
    )
    expect(first.json().detail.agents[0]).toMatchObject({
      agent: 'claude-code',
      action: 'written',
      status: 'installed',
      installed_version: bundledVersion,
    })
    const settings = JSON.parse(
      readFileSync(join(m.env.GENOFFICE_USER_DATA, 'app-settings.json'), 'utf-8'),
    )
    expect(settings.agentSkillInstalls[path]).toMatchObject({
      version: bundledVersion,
      channel: 'cli',
    })

    const again = await run(['skill', 'install', 'claude-code', '--json'], { env: m.env })
    expect(again.code).toBe(0)
    expect(again.json().warnings).toEqual([expect.objectContaining({ code: 'already_up_to_date' })])
    expect(again.json().detail.agents[0].action).toBe('unchanged')

    const listed = await run(['skill', 'list', '--json'], { env: m.env })
    expect(listed.json().detail.agents.find((a: any) => a.agent === 'claude-code')).toMatchObject({
      status: 'installed',
      up_to_date: true,
    })
  })

  it('install all covers only detected agents; unknown or absent agents are refused', async () => {
    const m = fakeMachine(['.claude', '.codex'])
    const all = await run(['skill', 'install', 'all', '--json'], { env: m.env })
    expect(all.code).toBe(0)
    expect(
      all
        .json()
        .detail.agents.map((a: any) => a.agent)
        .sort(),
    ).toEqual(['claude-code', 'codex'])

    const typo = await run(['skill', 'install', 'claud-code', '--json'], { env: m.env })
    expect(typo.code).toBe(1)
    expect(typo.json()).toMatchObject({ error: 'unsupported' })
    expect(typo.json().suggestion).toContain('claude-code')

    const absent = await run(['skill', 'install', 'cursor', '--json'], { env: m.env })
    expect(absent.code).toBe(1)
    expect(absent.json().error).toBe('unsupported')
    expect(existsSync(join(m.home, '.cursor'))).toBe(false)

    const forced = await run(['skill', 'install', 'cursor', '--force', '--json'], { env: m.env })
    expect(forced.code).toBe(0)
    expect(existsSync(join(m.home, '.cursor', 'skills', 'genoffice', 'SKILL.md'))).toBe(true)
  })

  it('refuses to downgrade a newer skill or overwrite a hand-edited one without --force', async () => {
    const m = fakeMachine(['.claude'])
    const dir = join(m.home, '.claude', 'skills', 'genoffice')
    mkdirSync(dir, { recursive: true })
    const path = join(dir, 'SKILL.md')
    writeFileSync(path, '---\nname: genoffice\nmetadata:\n  version: 99.0.0\n---\nfuture\n')
    const newer = await run(['skill', 'install', 'claude-code', '--json'], { env: m.env })
    expect(newer.code).toBe(2)
    expect(newer.json()).toMatchObject({ error: 'output_exists', detail: { status: 'newer' } })

    writeFileSync(path, '---\nname: other-skill\n---\nx\n')
    const occupied = await run(['skill', 'install', 'claude-code', '--force', '--json'], {
      env: m.env,
    })
    expect(occupied.code).toBe(2)
    expect(occupied.json().detail.status).toBe('occupied')

    writeFileSync(path, '---\nname: genoffice\nmetadata:\n  version: 0.0.1\n---\nold\n')
    const foreign = await run(['skill', 'install', 'claude-code', '--json'], { env: m.env })
    expect(foreign.code).toBe(0)
    expect(foreign.json().detail.agents[0].action).toBe('written')

    writeFileSync(path, readFileSync(path, 'utf-8') + '\nedited\n')
    const modified = await run(['skill', 'install', 'claude-code', '--json'], { env: m.env })
    expect(modified.code).toBe(2)
    expect(modified.json().detail.status).toBe('modified')
    const forced = await run(['skill', 'install', 'claude-code', '--force', '--json'], {
      env: m.env,
    })
    expect(forced.code).toBe(0)
  })

  it('installs into an arbitrary folder with --dir, resolved against the working directory', async () => {
    const m = fakeMachine([])
    const dir = join(m.home, 'elsewhere')
    const r = await run(['skill', 'install', '--dir', 'elsewhere', '--json'], {
      env: m.env,
      cwd: m.home,
    })
    expect(r.code).toBe(0)
    expect(existsSync(join(dir, 'genoffice', 'SKILL.md'))).toBe(true)
    expect(r.json().detail.agents[0]).toMatchObject({ agent: null, skills_dir: dir })
    const settings = JSON.parse(
      readFileSync(join(m.env.GENOFFICE_USER_DATA, 'app-settings.json'), 'utf-8'),
    )
    expect(Object.keys(settings.agentSkillInstalls)).toEqual([join(dir, 'genoffice', 'SKILL.md')])
    const none = await run(['skill', 'install', 'all', '--json'], { env: m.env })
    expect(none.code).toBe(1)
  })

  it('creates the app data folder for the ledger on a machine that never ran the app', async () => {
    const m = fakeMachine(['.claude'])
    const env = { ...m.env, GENOFFICE_USER_DATA: join(m.home, 'never', 'launched') }
    const r = await run(['skill', 'install', 'claude-code', '--json'], { env })
    expect(r.code).toBe(0)
    expect(existsSync(join(env.GENOFFICE_USER_DATA, 'app-settings.json'))).toBe(true)
  })

  it('install all refuses before writing anything when one target is blocked', async () => {
    const m = fakeMachine(['.claude', '.codex'])
    const codex = join(m.home, '.codex', 'skills', 'genoffice')
    mkdirSync(codex, { recursive: true })
    writeFileSync(join(codex, 'SKILL.md'), '---\nname: other-skill\n---\nx\n')
    const r = await run(['skill', 'install', 'all', '--json'], { env: m.env })
    expect(r.code).toBe(2)
    expect(r.json().detail).toMatchObject({ agent: 'codex', status: 'occupied' })
    expect(existsSync(join(m.home, '.claude', 'skills', 'genoffice'))).toBe(false)
    expect(existsSync(join(m.env.GENOFFICE_USER_DATA, 'app-settings.json'))).toBe(false)
  })
})
