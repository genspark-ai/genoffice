import { existsSync, lstatSync, mkdirSync, readFileSync, symlinkSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

const LAUNCHER = resolve(__dirname, '..', 'bin', 'genoffice')

function fakeMachine(agents: string[]) {
  const home = tempDir()
  for (const dot of agents) mkdirSync(join(home, dot), { recursive: true })
  return { home, env: { GENOFFICE_HOME: home } as NodeJS.ProcessEnv }
}

const readJson = (path: string) => JSON.parse(readFileSync(path, 'utf-8'))

describe('genoffice mcp list', () => {
  it('shows every agent, detected or not, and whether genoffice is registered', async () => {
    const m = fakeMachine(['.cursor', '.gemini'])
    writeFileSync(
      join(m.home, '.gemini', 'settings.json'),
      JSON.stringify({ mcpServers: { genoffice: { command: LAUNCHER, args: ['mcp'] } } }),
    )
    const r = await run(['mcp', 'list', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    expect(r.json().detail.launcher).toBe(LAUNCHER)
    const agents = r.json().detail.agents as Array<Record<string, unknown>>
    expect(agents.map((a) => a.agent).sort()).toEqual(
      ['claude-code', 'codex', 'copilot', 'cursor', 'gemini', 'opencode', 'windsurf'].sort(),
    )
    expect(agents.find((a) => a.agent === 'cursor')).toMatchObject({
      detected: true,
      registered: false,
      status: 'absent',
      config: join(m.home, '.cursor', 'mcp.json'),
    })
    expect(agents.find((a) => a.agent === 'gemini')).toMatchObject({
      detected: true,
      registered: true,
      status: 'registered',
    })
    expect(agents.find((a) => a.agent === 'codex')).toMatchObject({ detected: false })
    expect(agents.find((a) => a.agent === 'claude-code')!.config).toBe(join(m.home, '.claude.json'))
    expect((await run(['mcp', 'list'], { env: m.env })).stdout).toContain('2 agent(s) detected')
  })
})

describe('genoffice mcp install', () => {
  it('merges into an existing JSON config, keeping other servers, and is idempotent', async () => {
    const m = fakeMachine(['.cursor'])
    const file = join(m.home, '.cursor', 'mcp.json')
    writeFileSync(
      file,
      JSON.stringify({ mcpServers: { other: { url: 'https://x.test/mcp' } }, theme: 'dark' }),
    )
    const first = await run(['mcp', 'install', 'cursor', '--json'], { env: m.env })
    expect(first.code).toBe(0)
    expect(first.json().detail.agents[0]).toMatchObject({
      agent: 'cursor',
      status: 'installed',
      registered: true,
      command: LAUNCHER,
    })
    const text = readFileSync(file, 'utf-8')
    expect(text.endsWith('\n')).toBe(true)
    expect(text).toContain('\n  "mcpServers": {\n')
    expect(JSON.parse(text)).toEqual({
      mcpServers: {
        other: { url: 'https://x.test/mcp' },
        genoffice: { type: 'stdio', command: LAUNCHER, args: ['mcp'] },
      },
      theme: 'dark',
    })

    const again = await run(['mcp', 'install', 'cursor', '--json'], { env: m.env })
    expect(again.code).toBe(0)
    expect(again.json().detail.agents[0].status).toBe('unchanged')
    expect(readFileSync(file, 'utf-8')).toBe(text)
  })

  it('creates the file when the agent has none', async () => {
    const m = fakeMachine(['.codeium/windsurf'])
    const r = await run(['mcp', 'install', 'windsurf', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    expect(readJson(join(m.home, '.codeium', 'windsurf', 'mcp_config.json'))).toEqual({
      mcpServers: { genoffice: { command: LAUNCHER, args: ['mcp'] } },
    })
  })

  it('appends a TOML table for Codex and replaces a stale one with its sub-tables', async () => {
    const m = fakeMachine(['.codex'])
    const file = join(m.home, '.codex', 'config.toml')
    writeFileSync(
      file,
      'model = "gpt-5"\n\n[mcp_servers.other]\ncommand = "npx"\nargs = ["-y", "x"]\n',
    )
    const first = await run(['mcp', 'install', 'codex', '--json'], { env: m.env })
    expect(first.code).toBe(0)
    expect(first.json().detail.agents[0].status).toBe('installed')
    expect(readFileSync(file, 'utf-8')).toBe(
      'model = "gpt-5"\n\n[mcp_servers.other]\ncommand = "npx"\nargs = ["-y", "x"]\n\n' +
        `[mcp_servers.genoffice]\ncommand = ${JSON.stringify(LAUNCHER)}\nargs = ["mcp"]\n`,
    )
    expect(
      (await run(['mcp', 'install', 'codex', '--json'], { env: m.env })).json().detail.agents[0]
        .status,
    ).toBe('unchanged')

    writeFileSync(
      file,
      '[mcp_servers.genoffice]\ncommand = "/old/GenOffice.app/Contents/Resources/cli/genoffice"\nargs = ["mcp"]\n\n' +
        '[mcp_servers.genoffice.env]\nFOO = "1"\n\n[mcp_servers.other]\ncommand = "npx"\n\n[projects."/tmp/x"]\ntrust_level = "trusted"\n',
    )
    const listed = await run(['mcp', 'list', '--json'], { env: m.env })
    expect(listed.json().detail.agents.find((a: any) => a.agent === 'codex')).toMatchObject({
      status: 'stale',
      registered: true,
      command: '/old/GenOffice.app/Contents/Resources/cli/genoffice',
    })
    const replaced = await run(['mcp', 'install', 'codex', '--json'], { env: m.env })
    expect(replaced.code).toBe(0)
    expect(replaced.json().detail.agents[0].status).toBe('updated')
    expect(readFileSync(file, 'utf-8')).toBe(
      '[mcp_servers.other]\ncommand = "npx"\n\n[projects."/tmp/x"]\ntrust_level = "trusted"\n\n' +
        `[mcp_servers.genoffice]\ncommand = ${JSON.stringify(LAUNCHER)}\nargs = ["mcp"]\n`,
    )
  })

  it('refuses an entry that starts another program unless --force', async () => {
    const m = fakeMachine(['.cursor', '.codex'])
    const file = join(m.home, '.cursor', 'mcp.json')
    const before = JSON.stringify({
      mcpServers: { genoffice: { command: 'npx', args: ['other'] } },
    })
    writeFileSync(file, before)
    writeFileSync(
      join(m.home, '.codex', 'config.toml'),
      '[mcp_servers.genoffice]\nurl = "https://x.test"\n',
    )

    const r = await run(['mcp', 'install', 'cursor', '--json'], { env: m.env })
    expect(r.code).toBe(2)
    expect(r.json()).toMatchObject({
      error: 'output_exists',
      detail: { agent: 'cursor', status: 'occupied', command: 'npx' },
    })
    expect(readFileSync(file, 'utf-8')).toBe(before)

    const all = await run(['mcp', 'install', 'all', '--json'], { env: m.env })
    expect(all.code).toBe(0)
    const rows = all.json().detail.agents
    expect(rows.find((a: any) => a.agent === 'cursor').status).toBe('occupied')
    expect(rows.find((a: any) => a.agent === 'codex').status).toBe('occupied')
    expect(rows.find((a: any) => a.agent === 'gemini').status).toBe('not_detected')
    expect(all.json().warnings.map((w: any) => w.code)).toEqual([
      'entry_occupied',
      'entry_occupied',
    ])
    expect(readFileSync(file, 'utf-8')).toBe(before)

    const forced = await run(['mcp', 'install', 'cursor', '--force', '--json'], { env: m.env })
    expect(forced.code).toBe(0)
    expect(forced.json().detail.agents[0].status).toBe('updated')
    expect(readJson(file).mcpServers.genoffice.command).toBe(LAUNCHER)
  })

  it('install all writes only detected agents; an undetected agent needs --force', async () => {
    const m = fakeMachine(['.gemini', '.copilot', '.config/opencode'])
    const r = await run(['mcp', 'install', 'all', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    const rows = r.json().detail.agents
    expect(
      rows
        .filter((a: any) => a.status === 'installed')
        .map((a: any) => a.agent)
        .sort(),
    ).toEqual(['copilot', 'gemini', 'opencode'])
    expect(rows.filter((a: any) => a.status === 'not_detected').length).toBe(4)
    expect(readJson(join(m.home, '.copilot', 'mcp-config.json')).mcpServers.genoffice).toEqual({
      type: 'local',
      command: LAUNCHER,
      args: ['mcp'],
      tools: ['*'],
    })
    expect(readJson(join(m.home, '.config', 'opencode', 'opencode.json')).mcp.genoffice).toEqual({
      type: 'local',
      command: [LAUNCHER, 'mcp'],
      enabled: true,
    })
    expect(existsSync(join(m.home, '.cursor'))).toBe(false)

    const absent = await run(['mcp', 'install', 'cursor', '--json'], { env: m.env })
    expect(absent.code).toBe(1)
    expect(absent.json().error).toBe('unsupported')
    const forced = await run(['mcp', 'install', 'cursor', '--force', '--json'], { env: m.env })
    expect(forced.code).toBe(0)
    expect(existsSync(join(m.home, '.cursor', 'mcp.json'))).toBe(true)
  })

  it('writes Claude Code user scope into .claude.json at the top level, honouring CLAUDE_CONFIG_DIR', async () => {
    const m = fakeMachine(['.claude'])
    const file = join(m.home, '.claude.json')
    writeFileSync(file, JSON.stringify({ numStartups: 3, projects: { '/p': { mcpServers: {} } } }))
    const r = await run(['mcp', 'install', 'claude-code', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    expect(readJson(file)).toEqual({
      numStartups: 3,
      projects: { '/p': { mcpServers: {} } },
      mcpServers: { genoffice: { type: 'stdio', command: LAUNCHER, args: ['mcp'] } },
    })

    const custom = join(m.home, 'cc-config')
    mkdirSync(custom)
    const env = { ...m.env, CLAUDE_CONFIG_DIR: custom }
    const r2 = await run(['mcp', 'install', 'claude-code', '--json'], { env })
    expect(r2.code).toBe(0)
    expect(readJson(join(custom, '.claude.json')).mcpServers.genoffice.command).toBe(LAUNCHER)
  })

  it('--dir points one agent at a custom config folder, resolved against the working directory', async () => {
    const m = fakeMachine([])
    const r = await run(['mcp', 'install', 'gemini', '--dir', 'cfg', '--json'], {
      env: m.env,
      cwd: m.home,
    })
    expect(r.code).toBe(0)
    expect(r.json().detail.agents[0]).toMatchObject({
      agent: 'gemini',
      status: 'installed',
      config: join(m.home, 'cfg', 'settings.json'),
    })
    expect(readJson(join(m.home, 'cfg', 'settings.json')).mcpServers.genoffice.args).toEqual([
      'mcp',
    ])
    const all = await run(['mcp', 'install', 'all', '--dir', 'cfg', '--json'], { env: m.env })
    expect(all.code).toBe(1)
  })

  it('leaves a file it cannot parse alone and hands back the snippet', async () => {
    const m = fakeMachine(['.cursor', '.config/opencode'])
    writeFileSync(join(m.home, '.cursor', 'mcp.json'), '{ "mcpServers": { broken ')
    writeFileSync(join(m.home, '.config', 'opencode', 'opencode.jsonc'), '// comments\n{}\n')
    const r = await run(['mcp', 'install', 'cursor', '--json'], { env: m.env })
    expect(r.code).toBe(2)
    expect(r.json().detail.status).toBe('manual')
    expect(JSON.parse(r.json().detail.snippet).mcpServers.genoffice.command).toBe(LAUNCHER)
    expect(readFileSync(join(m.home, '.cursor', 'mcp.json'), 'utf-8')).toBe(
      '{ "mcpServers": { broken ',
    )

    const all = await run(['mcp', 'install', 'all', '--json'], { env: m.env })
    expect(all.code).toBe(0)
    const opencode = all.json().detail.agents.find((a: any) => a.agent === 'opencode')
    expect(opencode.status).toBe('manual')
    expect(opencode.snippet).toContain('"enabled": true')
    expect(existsSync(join(m.home, '.config', 'opencode', 'opencode.json'))).toBe(false)
  })
})

describe('genoffice mcp install summaries', () => {
  it('says what blocked the batch instead of claiming no agent was detected', async () => {
    const m = fakeMachine(['.cursor', '.gemini'])
    writeFileSync(
      join(m.home, '.cursor', 'mcp.json'),
      JSON.stringify({ mcpServers: { genoffice: { command: 'npx' } } }),
    )
    writeFileSync(join(m.home, '.gemini', 'settings.json'), '{ not json')
    const r = await run(['mcp', 'install', 'all', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    expect(r.json().summary).toBe(
      'nothing registered: 5 agent(s) not detected, 2 agent(s) left alone (see warnings)',
    )
    expect(r.json().warnings).toHaveLength(2)
    const u = await run(['mcp', 'uninstall', 'all', '--json'], { env: m.env })
    expect(u.json().summary).toBe(
      'nothing removed: 5 agent(s) not detected, 2 agent(s) left alone (see warnings)',
    )
  })

  it('treats an empty config file as having no entry', async () => {
    const m = fakeMachine(['.cursor'])
    const file = join(m.home, '.cursor', 'mcp.json')
    writeFileSync(file, '\n')
    const listed = await run(['mcp', 'list', '--json'], { env: m.env })
    expect(listed.json().detail.agents.find((a: any) => a.agent === 'cursor').status).toBe('absent')
    const gone = await run(['mcp', 'uninstall', 'cursor', '--json'], { env: m.env })
    expect(gone.code).toBe(0)
    expect(gone.json().detail.agents[0].status).toBe('absent')
    expect(readFileSync(file, 'utf-8')).toBe('\n')
    const r = await run(['mcp', 'install', 'cursor', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    expect(r.json().detail.agents[0].status).toBe('installed')
    expect(readJson(file).mcpServers.genoffice.command).toBe(LAUNCHER)
  })
})

describe('genoffice mcp with a symlinked config', () => {
  it('writes through the link to its target and keeps the link', async () => {
    const m = fakeMachine(['.cursor', '.claude'])
    const store = join(m.home, 'dotfiles')
    mkdirSync(store)
    const target = join(store, 'mcp.json')
    writeFileSync(target, JSON.stringify({ mcpServers: { other: { command: 'x' } } }))
    const link = join(m.home, '.cursor', 'mcp.json')
    symlinkSync(target, link)
    const claudeTarget = join(store, 'claude.json')
    symlinkSync(claudeTarget, join(m.home, '.claude.json'))

    const r = await run(['mcp', 'install', 'all', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    expect(lstatSync(link).isSymbolicLink()).toBe(true)
    expect(readJson(target).mcpServers).toEqual({
      other: { command: 'x' },
      genoffice: { type: 'stdio', command: LAUNCHER, args: ['mcp'] },
    })
    expect(lstatSync(join(m.home, '.claude.json')).isSymbolicLink()).toBe(true)
    expect(readJson(claudeTarget).mcpServers.genoffice.command).toBe(LAUNCHER)

    const u = await run(['mcp', 'uninstall', 'cursor', '--json'], { env: m.env })
    expect(u.json().detail.agents[0].status).toBe('removed')
    expect(lstatSync(link).isSymbolicLink()).toBe(true)
    expect(readJson(target)).toEqual({ mcpServers: { other: { command: 'x' } } })
  })
})

describe('genoffice mcp with a symlinked config directory', () => {
  it('follows a relative link from the real directory, not the lexical one', async () => {
    const m = fakeMachine([])
    const dotfiles = join(m.home, 'dotfiles')
    mkdirSync(join(dotfiles, 'cursor'), { recursive: true })
    mkdirSync(join(dotfiles, 'store'))
    const target = join(dotfiles, 'store', 'mcp.json')
    writeFileSync(target, JSON.stringify({ mcpServers: { other: { command: 'x' } } }))
    symlinkSync(join('..', 'store', 'mcp.json'), join(dotfiles, 'cursor', 'mcp.json'))
    symlinkSync(join(dotfiles, 'cursor'), join(m.home, '.cursor'))

    const r = await run(['mcp', 'install', 'cursor', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    expect(r.json().detail.agents[0].status).toBe('installed')
    expect(readJson(target).mcpServers.genoffice.command).toBe(LAUNCHER)
    expect(lstatSync(join(m.home, '.cursor')).isSymbolicLink()).toBe(true)
    expect(lstatSync(join(dotfiles, 'cursor', 'mcp.json')).isSymbolicLink()).toBe(true)
    expect(existsSync(join(m.home, 'store'))).toBe(false)
  })
})

describe('genoffice mcp uninstall', () => {
  it('words refusals for removing, not writing', async () => {
    const m = fakeMachine(['.gemini', '.cursor'])
    writeFileSync(
      join(m.home, '.gemini', 'settings.json'),
      JSON.stringify({ mcpServers: { genoffice: { command: 'npx' } } }),
    )
    writeFileSync(join(m.home, '.cursor', 'mcp.json'), '{ broken')
    const absent = await run(['mcp', 'uninstall', 'codex', '--json'], { env: m.env })
    expect(absent.code).toBe(1)
    expect(absent.json().message).toContain('nothing to remove')
    expect(absent.json().suggestion).toBe('repeat with --force to edit the file anyway')
    const occupied = await run(['mcp', 'uninstall', 'gemini', '--json'], { env: m.env })
    expect(occupied.code).toBe(2)
    expect(occupied.json().message).toContain('not removing it')
    expect(occupied.json().suggestion).toBe('repeat with --force to remove the foreign entry')
    const manual = await run(['mcp', 'uninstall', 'cursor', '--json'], { env: m.env })
    expect(manual.code).toBe(2)
    expect(manual.json().suggestion).toBe('remove the genoffice entry from the file yourself')
    expect(manual.json().detail.snippet).toBeUndefined()
    const all = await run(['mcp', 'uninstall', 'all', '--json'], { env: m.env })
    expect(all.json().warnings.map((w: any) => w.suggestion)).toEqual([
      'remove the genoffice entry from the file yourself',
      'repeat with --force to remove the foreign entry',
    ])
    expect(all.json().warnings[0].message).toContain('remove the entry by hand')
  })

  it('removes only the genoffice key and leaves the rest of the file', async () => {
    const m = fakeMachine(['.cursor', '.codex'])
    const json = join(m.home, '.cursor', 'mcp.json')
    writeFileSync(
      json,
      JSON.stringify({
        mcpServers: {
          other: { url: 'https://x.test/mcp' },
          genoffice: { command: LAUNCHER, args: ['mcp'] },
        },
        theme: 'dark',
      }),
    )
    const toml = join(m.home, '.codex', 'config.toml')
    writeFileSync(
      toml,
      `model = "gpt-5"\n\n[mcp_servers.genoffice]\ncommand = ${JSON.stringify(LAUNCHER)}\nargs = ["mcp"]\n\n[mcp_servers.other]\ncommand = "npx"\n`,
    )
    const r = await run(['mcp', 'uninstall', 'all', '--json'], { env: m.env })
    expect(r.code).toBe(0)
    const rows = r.json().detail.agents
    expect(rows.find((a: any) => a.agent === 'cursor')).toMatchObject({
      status: 'removed',
      registered: false,
    })
    expect(rows.find((a: any) => a.agent === 'codex').status).toBe('removed')
    expect(readJson(json)).toEqual({
      mcpServers: { other: { url: 'https://x.test/mcp' } },
      theme: 'dark',
    })
    expect(readFileSync(toml, 'utf-8')).toBe(
      'model = "gpt-5"\n\n[mcp_servers.other]\ncommand = "npx"\n',
    )

    const again = await run(['mcp', 'uninstall', 'cursor', '--json'], { env: m.env })
    expect(again.code).toBe(0)
    expect(again.json().detail.agents[0].status).toBe('absent')
  })

  it('does not remove an entry that starts another program unless --force', async () => {
    const m = fakeMachine(['.gemini'])
    const file = join(m.home, '.gemini', 'settings.json')
    writeFileSync(file, JSON.stringify({ mcpServers: { genoffice: { command: 'npx' } } }))
    const r = await run(['mcp', 'uninstall', 'gemini', '--json'], { env: m.env })
    expect(r.code).toBe(2)
    expect(r.json().detail.status).toBe('occupied')
    expect(readJson(file).mcpServers.genoffice.command).toBe('npx')
    const forced = await run(['mcp', 'uninstall', 'gemini', '--force', '--json'], { env: m.env })
    expect(forced.code).toBe(0)
    expect(readJson(file)).toEqual({ mcpServers: {} })
  })
})

describe('genoffice mcp argument parsing', () => {
  it('keeps serving flags and subcommands apart', async () => {
    const m = fakeMachine([])
    const mixed = await run(['mcp', 'install', 'all', '--http', '3000', '--json'], { env: m.env })
    expect(mixed.code).toBe(1)
    expect(mixed.json().message).toContain('--http')
    const bare = await run(['mcp', '--force', '--json'], { env: m.env })
    expect(bare.code).toBe(1)
    expect(bare.json().message).toContain('--force')
    const typo = await run(['mcp', 'instal', 'all', '--json'], { env: m.env })
    expect(typo.code).toBe(1)
    expect(typo.json().suggestion).toContain('install')
    const noAgent = await run(['mcp', 'install', '--json'], { env: m.env })
    expect(noAgent.code).toBe(1)
    expect(noAgent.json().error).toBe('missing_argument')
    const unknown = await run(['mcp', 'install', 'claud-code', '--json'], { env: m.env })
    expect(unknown.code).toBe(1)
    expect(unknown.json().suggestion).toContain('claude-code')
  })
})
