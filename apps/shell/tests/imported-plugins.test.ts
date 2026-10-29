import { describe, expect, it } from 'vitest'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  symlinkSync,
  utimesSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import {
  installPluginDir,
  installPluginZip,
  listPlugins,
  pluginSkills,
  pluginsRoot,
  readPluginSkillBody,
  removePlugin,
  setPluginEnabled,
} from '../src/main/imported-plugins'

/**
 * The plugin store on real disk: a scratch userData, plugin folders in a
 * scratch source tree, and the state file the enablement flag lives in.
 */

/** a scratch tree: <root>/userData is the app's own, <root>/src holds folders to pick */
function scaffold() {
  const root = mkdtempSync(join(tmpdir(), 'plugins-import-'))
  const userData = join(root, 'userData')
  const src = join(root, 'src')
  mkdirSync(userData)
  mkdirSync(src)
  return { root, userData, src, cleanup: () => rmSync(root, { recursive: true, force: true }) }
}

const SKILL = (name: string, description: string, body = 'Do the thing.') =>
  `---\nname: ${name}\ndescription: ${description}\n---\n\n${body}\n`

interface PluginOpts {
  /** raw manifest JSON; the default carries every field under the given name */
  manifest?: string
  /** write the manifest as .claude-plugin/plugin.json instead of plugin.json */
  claude?: boolean
  skills?: Record<string, string>
  commands?: Record<string, string>
  /** raw .mcp.json text */
  mcp?: string
}

/** a plugin folder: a manifest under either convention, plus components */
function makePlugin(base: string, name: string, opts: PluginOpts = {}): string {
  const dir = join(base, name)
  mkdirSync(dir, { recursive: true })
  const manifest =
    opts.manifest ??
    JSON.stringify({ name, version: '1.2.0', description: `The ${name} plugin`, author: 'Acme' })
  if (opts.claude) {
    mkdirSync(join(dir, '.claude-plugin'), { recursive: true })
    writeFileSync(join(dir, '.claude-plugin', 'plugin.json'), manifest)
  } else {
    writeFileSync(join(dir, 'plugin.json'), manifest)
  }
  for (const [skill, text] of Object.entries(opts.skills ?? {})) {
    mkdirSync(join(dir, 'skills', skill), { recursive: true })
    writeFileSync(join(dir, 'skills', skill, 'SKILL.md'), text)
  }
  for (const [command, text] of Object.entries(opts.commands ?? {})) {
    mkdirSync(join(dir, 'commands'), { recursive: true })
    writeFileSync(join(dir, 'commands', `${command}.md`), text)
  }
  if (opts.mcp !== undefined) writeFileSync(join(dir, '.mcp.json'), opts.mcp)
  return dir
}

/** hand-drop a plugin folder straight into the plugins root, as a person would */
function dropPlugin(userData: string, name: string, opts: PluginOpts = {}): string {
  mkdirSync(pluginsRoot(userData), { recursive: true })
  return makePlugin(pluginsRoot(userData), name, opts)
}

/** a state file written directly, to test what the reader tolerates */
function dropState(userData: string, text: string): void {
  writeFileSync(join(userData, 'plugins-state.json'), text)
}

const stateOf = (userData: string): Record<string, unknown> =>
  JSON.parse(readFileSync(join(userData, 'plugins-state.json'), 'utf8'))

async function zipOf(files: Record<string, string | Buffer>, dest: string): Promise<string> {
  const zip = new JSZip()
  for (const [name, data] of Object.entries(files)) zip.file(name, data)
  writeFileSync(dest, await zip.generateAsync({ type: 'nodebuffer' }))
  return dest
}

const OLD = new Date(Date.now() - 86_400_000)

describe('pluginsRoot', () => {
  it('sits under userData, beside the skills root', () => {
    expect(pluginsRoot('/app/userData')).toBe(join('/app/userData', 'plugins'))
  })
})

describe('listPlugins', () => {
  it('reads the plain convention at the plugin root', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'plain-one')
    expect(await listPlugins(userData)).toEqual([
      {
        name: 'plain-one',
        version: '1.2.0',
        description: 'The plain-one plugin',
        author: 'Acme',
        enabled: true,
        components: { skills: [], commands: [], mcpServers: [] },
        format: 'plain',
      },
    ])
    cleanup()
  })

  it('reads the Claude convention, and it wins when both exist', async () => {
    const { userData, cleanup } = scaffold()
    const dir = dropPlugin(userData, 'both', { claude: true })
    writeFileSync(join(dir, 'plugin.json'), JSON.stringify({ name: 'the-plain-one' }))
    expect((await listPlugins(userData))[0]).toMatchObject({
      name: 'both',
      format: 'claude',
    })
    cleanup()
  })

  it('reads an author given as an object, and tolerates missing fields', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'authored', {
      manifest: JSON.stringify({
        name: 'authored',
        author: { name: 'Ada Lovelace', email: 'ada@example.com' },
      }),
    })
    expect((await listPlugins(userData))[0]).toMatchObject({
      version: '',
      description: '',
      author: 'Ada Lovelace',
    })
    cleanup()
  })

  it('falls back to the folder name when the manifest name is missing or invalid', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'fallback', { manifest: JSON.stringify({ description: 'no name' }) })
    dropPlugin(userData, 'second', { manifest: JSON.stringify({ name: 'a/b' }) })
    expect((await listPlugins(userData)).map((p) => p.name).sort()).toEqual(['fallback', 'second'])
    cleanup()
  })

  it('skips dot folders, folders without a manifest, and corrupt manifests', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'good', { manifest: JSON.stringify({ name: 'good' }) })
    mkdirSync(join(pluginsRoot(userData), '.cache'), { recursive: true })
    writeFileSync(join(pluginsRoot(userData), '.cache', 'plugin.json'), '{}')
    mkdirSync(join(pluginsRoot(userData), 'no-manifest'), { recursive: true })
    dropPlugin(userData, 'corrupt', { manifest: 'not json{' })
    // a corrupt Claude manifest falls back to a parseable plain one beside it
    const backup = dropPlugin(userData, 'backup', { claude: true, manifest: 'also not json' })
    writeFileSync(join(backup, 'plugin.json'), JSON.stringify({ name: 'backup' }))
    expect((await listPlugins(userData)).map((p) => p.name).sort()).toEqual(['backup', 'good'])
    cleanup()
  })

  it('treats a missing or corrupt state file as everything enabled', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'one')
    dropPlugin(userData, 'two')
    dropState(userData, 'not json at all')
    const byName = async () =>
      (await listPlugins(userData))
        .map((p) => [p.name, p.enabled] as const)
        .sort((a, b) => (a[0] < b[0] ? -1 : 1))
    expect(await byName()).toEqual([
      ['one', true],
      ['two', true],
    ])
    dropState(userData, '{"one": {"enabled": false}}')
    expect(await byName()).toEqual([
      ['one', false],
      ['two', true],
    ])
    cleanup()
  })

  it('lists newest first — the folder mtime an install sets — with name as the tiebreaker', async () => {
    const { userData, cleanup } = scaffold()
    const alpha = dropPlugin(userData, 'alpha')
    const beta = dropPlugin(userData, 'beta')
    utimesSync(alpha, OLD, OLD)
    expect((await listPlugins(userData)).map((p) => p.name)).toEqual(['beta', 'alpha'])
    utimesSync(beta, OLD, OLD)
    expect((await listPlugins(userData)).map((p) => p.name)).toEqual(['alpha', 'beta'])
    cleanup()
  })
})

describe('component summary', () => {
  it('namespaces skills and commands, and records .mcp.json servers display only', async () => {
    const { userData, cleanup } = scaffold()
    const dir = dropPlugin(userData, 'acme', {
      skills: { 'pdf-fix': SKILL('pdf-fix', 'Fix a pdf') },
      commands: { run: 'Run the export.\n' },
      mcp: JSON.stringify({ mcpServers: { fetch: {}, fs: {} } }),
    })
    // the noise a glob must not pick up
    mkdirSync(join(dir, 'skills', '.hidden'), { recursive: true })
    writeFileSync(join(dir, 'skills', '.hidden', 'SKILL.md'), SKILL('hidden', 'x'))
    mkdirSync(join(dir, 'skills', 'nested', 'deep'), { recursive: true })
    writeFileSync(join(dir, 'skills', 'nested', 'deep', 'SKILL.md'), SKILL('deep', 'x'))
    mkdirSync(join(dir, 'skills', 'no-skilmd'), { recursive: true })
    mkdirSync(join(dir, 'commands'), { recursive: true })
    writeFileSync(join(dir, 'commands', 'notes.txt'), 'not markdown')
    writeFileSync(join(dir, 'commands', '.draft.md'), 'dot file')
    expect((await listPlugins(userData))[0].components).toEqual({
      skills: ['acme:pdf-fix'],
      commands: ['acme:run'],
      mcpServers: ['fetch', 'fs'],
    })
    cleanup()
  })

  it('survives a missing, corrupt or server-less .mcp.json', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'silent')
    dropPlugin(userData, 'corrupt-mcp', { mcp: 'not json' })
    dropPlugin(userData, 'other-keys', { mcp: JSON.stringify({ mcpServers: [] }) })
    for (const plugin of await listPlugins(userData)) {
      expect(plugin.components.mcpServers).toEqual([])
    }
    cleanup()
  })
})

describe('installPluginDir', () => {
  it('copies the whole folder in under the name the manifest gives', async () => {
    const { userData, src, cleanup } = scaffold()
    const dir = makePlugin(src, 'picked-folder', {
      manifest: JSON.stringify({
        name: 'acme',
        version: '2.0.0',
        description: 'The acme plugin',
        author: { name: 'Ada' },
      }),
      skills: { helper: SKILL('helper', 'A docx helper') },
    })
    mkdirSync(join(dir, 'reference'), { recursive: true })
    writeFileSync(join(dir, 'reference', 'notes.md'), 'grid rules')
    const installed = installPluginDir(dir, userData)
    expect(installed).toMatchObject({
      name: 'acme',
      version: '2.0.0',
      author: 'Ada',
      enabled: true,
      format: 'plain',
    })
    const dest = join(pluginsRoot(userData), 'acme')
    expect(installed.components.skills).toEqual(['acme:helper'])
    expect(existsSync(join(dest, 'skills', 'helper', 'SKILL.md'))).toBe(true)
    expect(existsSync(join(dest, 'reference', 'notes.md'))).toBe(true)
    expect((await listPlugins(userData)).map((p) => p.name)).toEqual(['acme'])
    cleanup()
  })

  it('names the installed folder after the picked folder when the manifest says nothing', () => {
    const { userData, src, cleanup } = scaffold()
    const dir = makePlugin(src, 'named-by-folder', {
      manifest: JSON.stringify({ description: 'no name here' }),
    })
    expect(installPluginDir(dir, userData).name).toBe('named-by-folder')
    expect(existsSync(join(pluginsRoot(userData), 'named-by-folder'))).toBe(true)
    cleanup()
  })

  it('refuses a folder with no parseable manifest, and copies nothing', () => {
    const { userData, src, cleanup } = scaffold()
    const dir = makePlugin(src, 'empty', { manifest: '' })
    rmSync(join(dir, 'plugin.json'))
    expect(() => installPluginDir(dir, userData)).toThrow('no parseable plugin')
    const corrupt = makePlugin(src, 'corrupt', { manifest: 'no json' })
    expect(() => installPluginDir(corrupt, userData)).toThrow('no parseable plugin')
    expect(existsSync(pluginsRoot(userData))).toBe(false)
    cleanup()
  })

  it('refuses a name that is already installed, by destination or by listing', async () => {
    const { userData, src, cleanup } = scaffold()
    // by destination: the folder plugins/acme is already there
    const dir = makePlugin(src, 'acme')
    installPluginDir(dir, userData)
    expect(() => installPluginDir(dir, userData)).toThrow('already installed')
    expect((await listPlugins(userData)).map((p) => p.name)).toEqual(['acme'])
    // by listing: a hand-dropped folder holds the name even though its own
    // folder is called something else, so no destination collides
    const { userData: otherUserData, src: otherSrc, cleanup: otherCleanup } = scaffold()
    dropPlugin(otherUserData, 'folder-name', { manifest: JSON.stringify({ name: 'acme' }) })
    const picked = makePlugin(otherSrc, 'picked', { manifest: JSON.stringify({ name: 'acme' }) })
    expect(() => installPluginDir(picked, otherUserData)).toThrow('already installed')
    expect((await listPlugins(otherUserData)).map((p) => [p.name, p.format])).toEqual([
      ['acme', 'plain'],
    ])
    expect(existsSync(join(pluginsRoot(otherUserData), 'acme'))).toBe(false)
    cleanup()
    otherCleanup()
  })

  it('refuses a symlinked plugin and one that ships a symlink', () => {
    const { userData, src, cleanup } = scaffold()
    const real = makePlugin(src, 'real')
    symlinkSync(real, join(src, 'linked'))
    expect(() => installPluginDir(join(src, 'linked'), userData)).toThrow('symlink')
    const shipping = makePlugin(src, 'shipping')
    writeFileSync(join(shipping, 'outside.md'), 'elsewhere')
    symlinkSync(join(shipping, 'outside.md'), join(shipping, 'reference.md'))
    expect(() => installPluginDir(shipping, userData)).toThrow('symlink')
    expect(existsSync(join(pluginsRoot(userData), 'real'))).toBe(false)
    const leftovers = existsSync(pluginsRoot(userData)) ? readdirSync(pluginsRoot(userData)) : []
    expect(leftovers.filter((n) => n.startsWith('.'))).toEqual([])
    cleanup()
  })

  it('refuses a plugin over the import size cap', () => {
    const { userData, src, cleanup } = scaffold()
    const dir = makePlugin(src, 'heavy')
    mkdirSync(join(dir, 'corpus'), { recursive: true })
    writeFileSync(join(dir, 'corpus', 'big.bin'), Buffer.alloc(11 * 1024 * 1024, 0))
    expect(() => installPluginDir(dir, userData)).toThrow('import cap')
    expect(existsSync(join(pluginsRoot(userData), 'heavy'))).toBe(false)
    cleanup()
  })
})

describe('installPluginZip', () => {
  it('unpacks a zip with the plugin under one top-level folder', async () => {
    const { root, userData, cleanup } = scaffold()
    const zip = await zipOf(
      {
        'acme/plugin.json': JSON.stringify({ name: 'acme', version: '3.0.0', author: 'Acme' }),
        'acme/skills/helper/SKILL.md': SKILL('helper', 'A docx helper'),
        'acme/commands/run.md': 'Run the export.\n',
        'acme/reference/notes.md': 'grid rules',
        '__MACOSX/acme/._plugin.json': 'junk',
      },
      join(root, 'anything.zip'),
    )
    const installed = await installPluginZip(zip, userData)
    expect(installed).toMatchObject({ name: 'acme', version: '3.0.0', format: 'plain' })
    expect(installed.components).toEqual({
      skills: ['acme:helper'],
      commands: ['acme:run'],
      mcpServers: [],
    })
    const dest = join(pluginsRoot(userData), 'acme')
    expect(existsSync(join(dest, 'reference', 'notes.md'))).toBe(true)
    expect(existsSync(join(dest, '__MACOSX'))).toBe(false)
    cleanup()
  })

  it('unpacks a zip with the manifest at the archive root, and reads the Claude convention', async () => {
    const { root, userData, cleanup } = scaffold()
    const zip = await zipOf(
      {
        '.claude-plugin/plugin.json': JSON.stringify({
          name: 'rooted',
          version: '1.0.0',
          description: 'At the root',
        }),
        'skills/helper/SKILL.md': SKILL('helper', 'A pdf helper'),
      },
      join(root, 'rooted.zip'),
    )
    const installed = await installPluginZip(zip, userData)
    expect(installed).toMatchObject({
      name: 'rooted',
      format: 'claude',
      description: 'At the root',
    })
    expect(existsSync(join(pluginsRoot(userData), 'rooted', 'skills', 'helper', 'SKILL.md'))).toBe(
      true,
    )
    cleanup()
  })

  it('falls back to the zip stem when a root-level manifest names nothing', async () => {
    const { root, userData, cleanup } = scaffold()
    const zip = await zipOf(
      { 'plugin.json': JSON.stringify({ description: 'no name' }) },
      join(root, 'stemmed.zip'),
    )
    expect((await installPluginZip(zip, userData)).name).toBe('stemmed')
    cleanup()
  })

  it('refuses archives with no manifest, an unsafe entry, a symlink, or over the cap', async () => {
    const { root, userData, cleanup } = scaffold()
    const noManifest = await zipOf({ 'a/README.md': 'hi' }, join(root, 'none.zip'))
    await expect(installPluginZip(noManifest, userData)).rejects.toThrow('no plugin manifest')
    const twoTops = await zipOf(
      { 'a/plugin.json': '{}', 'b/plugin.json': '{}' },
      join(root, 'two.zip'),
    )
    await expect(installPluginZip(twoTops, userData)).rejects.toThrow('no plugin manifest')
    const escape = await zipOf(
      { 'acme/plugin.json': '{}', '../escape.txt': 'out' },
      join(root, 'bad.zip'),
    )
    await expect(installPluginZip(escape, userData)).rejects.toThrow()
    expect(existsSync(join(root, 'escape.txt'))).toBe(false)
    const z = new JSZip()
    z.file('acme/plugin.json', '{}')
    z.file('acme/link', '/etc/passwd', { unixPermissions: 0o120777 })
    const linked = join(root, 'link.zip')
    writeFileSync(linked, await z.generateAsync({ type: 'nodebuffer', platform: 'UNIX' }))
    await expect(installPluginZip(linked, userData)).rejects.toThrow('symlink')
    const heavy = await zipOf(
      {
        'acme/plugin.json': '{}',
        'acme/big.bin': Buffer.alloc(11 * 1024 * 1024),
      },
      join(root, 'heavy.zip'),
    )
    await expect(installPluginZip(heavy, userData)).rejects.toThrow('import cap')
    const notZip = join(root, 'plain.zip')
    writeFileSync(notZip, 'plain text')
    await expect(installPluginZip(notZip, userData)).rejects.toThrow('not a zip')
    // every refusal left nothing behind: no plugin folder, no scratch dir
    expect(readdirSync(pluginsRoot(userData))).toEqual([])
    cleanup()
  })

  it('refuses to overwrite an installed plugin of the same name', async () => {
    const { root, userData, cleanup } = scaffold()
    makePlugin(root, 'acme', { manifest: JSON.stringify({ name: 'acme' }) })
    installPluginDir(join(root, 'acme'), userData)
    const zip = await zipOf(
      { 'other/plugin.json': JSON.stringify({ name: 'acme' }) },
      join(root, 'a.zip'),
    )
    await expect(installPluginZip(zip, userData)).rejects.toThrow('already installed')
    expect((await listPlugins(userData)).map((p) => p.name)).toEqual(['acme'])
    cleanup()
  })
})

describe('setPluginEnabled / removePlugin', () => {
  it('persists enablement across a re-list and filters the skills a disabled plugin offers', async () => {
    const { userData, src, cleanup } = scaffold()
    const acme = makePlugin(src, 'acme', {
      skills: { helper: SKILL('helper', 'A docx helper') },
      commands: { run: 'Run the export.\n' },
    })
    const beta = makePlugin(src, 'beta', {
      skills: { other: SKILL('other', 'A pdf helper') },
    })
    installPluginDir(acme, userData)
    installPluginDir(beta, userData)
    await setPluginEnabled(userData, 'acme', false)
    const listed = await listPlugins(userData)
    expect(listed.find((p) => p.name === 'acme')?.enabled).toBe(false)
    expect(listed.find((p) => p.name === 'beta')?.enabled).toBe(true)
    const names = (await pluginSkills(userData)).map((s) => s.name).sort()
    expect(names).toEqual(['beta:other'])
    await setPluginEnabled(userData, 'acme', true)
    expect((await pluginSkills(userData)).map((s) => s.name).sort()).toEqual([
      'acme:helper',
      'acme:run',
      'beta:other',
    ])
    cleanup()
  })

  it('refuses a name the listing never showed', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'acme')
    await expect(setPluginEnabled(userData, 'nope', false)).rejects.toThrow('unknown plugin')
    await expect(setPluginEnabled(userData, '../escape', true)).rejects.toThrow('unknown plugin')
    await expect(setPluginEnabled(userData, 'acme', 'yes' as never)).rejects.toThrow(
      'a plugin name',
    )
    await expect(removePlugin(userData, 'nope')).rejects.toThrow('unknown plugin')
    await expect(removePlugin(userData, '/etc')).rejects.toThrow('unknown plugin')
    expect((await listPlugins(userData)).map((p) => p.name)).toEqual(['acme'])
    cleanup()
  })

  it('removes the folder and its state entry, so a reinstall starts fresh', async () => {
    const { userData, src, cleanup } = scaffold()
    const dir = makePlugin(src, 'acme', { skills: { helper: SKILL('helper', 'A docx helper') } })
    installPluginDir(dir, userData)
    await setPluginEnabled(userData, 'acme', false)
    await removePlugin(userData, 'acme')
    expect(existsSync(join(pluginsRoot(userData), 'acme'))).toBe(false)
    expect(await listPlugins(userData)).toEqual([])
    expect(stateOf(userData)).toEqual({})
    expect(installPluginDir(dir, userData).enabled).toBe(true)
    cleanup()
  })

  it('leaves the other plugins and their state alone', async () => {
    const { userData, src, cleanup } = scaffold()
    installPluginDir(makePlugin(src, 'acme'), userData)
    installPluginDir(makePlugin(src, 'beta'), userData)
    await setPluginEnabled(userData, 'acme', false)
    await removePlugin(userData, 'beta')
    expect((await listPlugins(userData)).map((p) => [p.name, p.enabled])).toEqual([['acme', false]])
    expect(stateOf(userData)).toEqual({ acme: { enabled: false } })
    cleanup()
  })
})

describe('pluginSkills', () => {
  it('lists every skill and command of the enabled plugins as <plugin>:<name>', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'acme', {
      skills: { 'pdf-fix': SKILL('pdf-fix', 'Fix a pdf table') },
      commands: { run: 'Run the export.\n' },
    })
    const skills = await pluginSkills(userData)
    expect(skills).toHaveLength(2)
    const [skill, command] = skills
    expect(skill).toMatchObject({
      name: 'acme:pdf-fix',
      description: 'Fix a pdf table',
      plugin: 'acme',
      relevance: { relevant: true, matched: ['pdf'] },
    })
    expect(skill.path).toBe(join(pluginsRoot(userData), 'acme', 'skills', 'pdf-fix', 'SKILL.md'))
    expect(command).toMatchObject({
      name: 'acme:run',
      description: 'Run the export.',
      plugin: 'acme',
    })
    expect(command.path).toBe(join(pluginsRoot(userData), 'acme', 'commands', 'run.md'))
    cleanup()
  })

  it('derives a description when the component has no full frontmatter', async () => {
    const { userData, cleanup } = scaffold()
    dropPlugin(userData, 'acme', {
      commands: {
        hint: '---\ndescription: A docx hint\n---\n\nThe body.\n',
        plain: '# Turn notes into a docx report\n\nStep one.\n',
      },
    })
    const skills = await pluginSkills(userData)
    expect(skills.map((s) => [s.name, s.description])).toEqual([
      ['acme:hint', 'A docx hint'],
      ['acme:plain', 'Turn notes into a docx report'],
    ])
    expect(skills.every((s) => s.relevance.relevant)).toBe(true)
    cleanup()
  })

  it('reads a component body read-only: frontmatter stripped, whole file otherwise', async () => {
    const { userData, cleanup } = scaffold()
    const dir = dropPlugin(userData, 'acme', {
      skills: { helper: SKILL('helper', 'A helper', 'Step one.') },
      commands: {
        plain: 'Just text, no frontmatter.\n',
        hint: '---\ndescription: A hint\n---\n\nThe real body.\n',
      },
    })
    expect(readPluginSkillBody(join(dir, 'skills', 'helper', 'SKILL.md'))).toBe('Step one.\n')
    expect(readPluginSkillBody(join(dir, 'commands', 'plain.md'))).toBe(
      'Just text, no frontmatter.\n',
    )
    expect(readPluginSkillBody(join(dir, 'commands', 'hint.md'))).toBe('The real body.\n')
    cleanup()
  })
})
