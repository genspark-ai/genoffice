import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
  type Dirent,
} from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, dirname, join, sep } from 'node:path'
import JSZip from 'jszip'
import { classifySkill, parseSkillFrontmatter } from '@genoffice/agent-core'
import { atomicWriteFile } from './atomic-write'
import {
  assertSafeZipEntries,
  deriveSkillMeta,
  inflateWithin,
  inspectImportDir,
  isSkillName,
  MAX_IMPORT_BYTES,
} from './imported-skills'
import type { FoundSkill, InstalledPlugin, PluginComponents } from '../shared/found-skill'

/**
 * The plugins a person installed into GenOffice.
 *
 * A plugin is one self-contained folder — a manifest plus prompt-level
 * components (`skills/<name>/SKILL.md`, `commands/<name>.md`) — living under
 * `<userData>/plugins/<name>/`. Two manifest conventions are read: Claude
 * Code's `.claude-plugin/plugin.json` and the plain `plugin.json` at the root,
 * and the whole folder is the unit: it is copied in through the same symlink,
 * size and zip-slip guards a skill import goes through, and removed as one.
 *
 * Components are listed and invoked as `<plugin>:<name>`, so a plugin cannot
 * shadow a user skill (the merge in user-cli keeps the user's). A root
 * `.mcp.json` is recorded for display only — GenOffice's MCP support is a
 * server, not a client, so nothing a plugin declares is ever run.
 */

/** where installed plugins live: one folder per plugin, under userData */
export function pluginsRoot(userDataDir: string): string {
  return join(userDataDir, 'plugins')
}

/** per-plugin enablement, `<userData>/plugins-state.json`; absent means enabled */
function pluginsStatePath(userDataDir: string): string {
  return join(userDataDir, 'plugins-state.json')
}

interface PluginState {
  [name: string]: { enabled: boolean }
}

/**
 * A missing, corrupt or shapeless state file just means every plugin is
 * enabled: enablement is a convenience, never a reason the listing fails.
 */
function readPluginState(userDataDir: string): PluginState {
  let parsed: unknown
  try {
    parsed = JSON.parse(readFileSync(pluginsStatePath(userDataDir), 'utf8'))
  } catch {
    return {}
  }
  if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) return {}
  const state: PluginState = {}
  for (const [name, value] of Object.entries(parsed as Record<string, unknown>)) {
    if (
      value !== null &&
      typeof value === 'object' &&
      typeof (value as { enabled?: unknown }).enabled === 'boolean'
    ) {
      state[name] = { enabled: (value as { enabled: boolean }).enabled }
    }
  }
  return state
}

/** the scratch-then-rename write the imports use, so a crash mid-write costs nothing */
async function writePluginState(userDataDir: string, state: PluginState): Promise<void> {
  await atomicWriteFile(pluginsStatePath(userDataDir), Buffer.from(JSON.stringify(state, null, 2)))
}

/** both manifest conventions, in preference order: Claude's first, plain second */
const MANIFEST_LOCATIONS: ReadonlyArray<{
  rel: readonly string[]
  format: InstalledPlugin['format']
}> = [
  { rel: ['.claude-plugin', 'plugin.json'], format: 'claude' },
  { rel: ['plugin.json'], format: 'plain' },
]

interface Manifest {
  format: InstalledPlugin['format']
  fields: Record<string, unknown>
}

/**
 * The manifest of a plugin folder. `.claude-plugin/plugin.json` wins when both
 * exist; a corrupt one falls through to the plain `plugin.json` beside it.
 * Null when neither parses — the folder is not a plugin, and the listing
 * skips it rather than showing a half-read entry.
 */
function readManifest(dir: string): Manifest | null {
  for (const { rel, format } of MANIFEST_LOCATIONS) {
    let text: string
    try {
      text = readFileSync(join(dir, ...rel), 'utf8')
    } catch {
      continue
    }
    try {
      const parsed: unknown = JSON.parse(text)
      if (parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed)) {
        return { format, fields: parsed as Record<string, unknown> }
      }
    } catch {
      // corrupt JSON of one convention: the other is still tried
    }
  }
  return null
}

/** the manifest an install requires: present, and parseable, before any copy */
function requireManifest(dir: string): Manifest {
  const manifest = readManifest(dir)
  if (!manifest) {
    throw new Error('the folder has no parseable plugin.json or .claude-plugin/plugin.json')
  }
  return manifest
}

/** the manifest's `author`: a plain string, or an object carrying `name` */
function authorOf(fields: Record<string, unknown>): string {
  const author = fields.author
  if (typeof author === 'string') return author
  if (author !== null && typeof author === 'object' && !Array.isArray(author)) {
    const name = (author as Record<string, unknown>).name
    if (typeof name === 'string') return name
  }
  return ''
}

/** a manifest field that is only ever shown, never parsed: string or nothing */
function stringField(value: unknown): string {
  return typeof value === 'string' ? value : ''
}

/**
 * The name a plugin is listed, invoked and keyed by. The manifest's `name`
 * when it is a legal folder name (`isSkillName`: no separators, no Win32
 * characters a folder cannot carry, no leading dot — and no `:`, which the
 * `<plugin>:<name>` namespace needs to stay unambiguous); the folder's own
 * name otherwise. Null when neither yields anything usable.
 */
function effectiveName(manifestName: unknown, folderName: string): string | null {
  if (typeof manifestName === 'string' && isSkillName(manifestName.trim())) {
    return manifestName.trim()
  }
  return isSkillName(folderName) ? folderName : null
}

/** one prompt-level component file: a skill folder's SKILL.md, or a command .md */
interface PluginComponent {
  /** the component's own name: the skill folder's, or the command file's stem */
  name: string
  /** the real file, read when the component is listed or invoked */
  path: string
}

const byName = (a: PluginComponent, b: PluginComponent) =>
  a.name < b.name ? -1 : a.name > b.name ? 1 : 0

/**
 * `skills/<name>/SKILL.md`: one level deep, dot folders never count, and the folder
 * — not the file — is the component's name. A name `isSkillName` refuses is
 * skipped: it could not be invoked as `<plugin>:<name>` without colliding with
 * another plugin's.
 */
function skillComponents(dir: string): PluginComponent[] {
  const root = join(dir, 'skills')
  let entries: Dirent[]
  try {
    entries = readdirSync(root, { withFileTypes: true })
  } catch {
    return []
  }
  const out: PluginComponent[] = []
  for (const entry of entries) {
    if (!entry.isDirectory() || entry.name.startsWith('.') || !isSkillName(entry.name)) continue
    const path = join(root, entry.name, 'SKILL.md')
    if (existsSync(path)) out.push({ name: entry.name, path })
  }
  return out.sort(byName)
}

/** `commands/*.md`: the stem of a file directly inside commands/ */
function commandComponents(dir: string): PluginComponent[] {
  const root = join(dir, 'commands')
  let entries: Dirent[]
  try {
    entries = readdirSync(root, { withFileTypes: true })
  } catch {
    return []
  }
  const out: PluginComponent[] = []
  for (const entry of entries) {
    if (
      !entry.isFile() ||
      entry.name.startsWith('.') ||
      !entry.name.toLowerCase().endsWith('.md')
    ) {
      continue
    }
    const name = entry.name.slice(0, -3)
    if (!isSkillName(name)) continue
    out.push({ name, path: join(root, entry.name) })
  }
  return out.sort(byName)
}

/**
 * The server names in a root `.mcp.json`, display only. GenOffice's MCP
 * support is a server, not a client, so nothing here is ever run: the listing
 * tells the person what the plugin declares and stops there.
 */
function mcpServerNames(dir: string): string[] {
  let parsed: unknown
  try {
    parsed = JSON.parse(readFileSync(join(dir, '.mcp.json'), 'utf8'))
  } catch {
    return []
  }
  const servers =
    parsed !== null && typeof parsed === 'object'
      ? (parsed as Record<string, unknown>).mcpServers
      : undefined
  if (servers === null || typeof servers !== 'object' || Array.isArray(servers)) return []
  return Object.keys(servers as Record<string, unknown>).sort()
}

/** the namespaced component names the listing shows: `<plugin>:<name>` */
function componentSummary(dir: string, pluginName: string): PluginComponents {
  return {
    skills: skillComponents(dir).map((c) => `${pluginName}:${c.name}`),
    commands: commandComponents(dir).map((c) => `${pluginName}:${c.name}`),
    mcpServers: mcpServerNames(dir),
  }
}

/** one plugin on disk: where its folder sits, plus what the listing says about it */
interface ScannedPlugin {
  dir: string
  plugin: InstalledPlugin
}

/** the folder's own mtime, which an install sets and enable/disable never touches */
function mtimeOf(dir: string): number {
  try {
    return statSync(dir).mtimeMs
  } catch {
    return 0
  }
}

function describePlugin(
  dir: string,
  manifest: Manifest,
  state: PluginState,
): InstalledPlugin | null {
  const name = effectiveName(manifest.fields.name, basename(dir))
  if (!name) return null
  return {
    name,
    version: stringField(manifest.fields.version),
    description: stringField(manifest.fields.description),
    author: authorOf(manifest.fields),
    enabled: state[name]?.enabled ?? true,
    components: componentSummary(dir, name),
    format: manifest.format,
  }
}

/**
 * Every plugin under the plugins root: a directory, not a dot folder, with a
 * manifest that parses. Newest first — the folder's mtime is the closest thing
 * to an install time, since copying files in is what sets it — with name as
 * the tiebreaker, so the plugin an install just added shows on top.
 */
function scanPlugins(userDataDir: string): ScannedPlugin[] {
  const root = pluginsRoot(userDataDir)
  let entries: Dirent[]
  try {
    entries = readdirSync(root, { withFileTypes: true })
  } catch {
    return []
  }
  const state = readPluginState(userDataDir)
  const out: ScannedPlugin[] = []
  for (const entry of entries) {
    if (!entry.isDirectory() || entry.name.startsWith('.')) continue
    const dir = join(root, entry.name)
    const manifest = readManifest(dir)
    if (!manifest) continue
    const plugin = describePlugin(dir, manifest, state)
    if (plugin) out.push({ dir, plugin })
  }
  return out.sort((a, b) => {
    const newest = mtimeOf(b.dir) - mtimeOf(a.dir)
    return newest !== 0
      ? newest
      : a.plugin.name < b.plugin.name
        ? -1
        : a.plugin.name > b.plugin.name
          ? 1
          : 0
  })
}

/** the plugins a person has installed, newest first, with what each carries */
export async function listPlugins(userDataDir: string): Promise<InstalledPlugin[]> {
  return scanPlugins(userDataDir).map(({ plugin }) => plugin)
}

/**
 * Refuse a name that is already taken — by a listed plugin of that name, or by
 * any folder already sitting at its destination, listed or not. Refused rather
 * than overwritten, the same rule a skill import follows.
 */
function refuseTaken(name: string, userDataDir: string): void {
  const dest = join(pluginsRoot(userDataDir), name)
  if (existsSync(dest) || scanPlugins(userDataDir).some(({ plugin }) => plugin.name === name)) {
    throw new Error(`a plugin named "${name}" is already installed`)
  }
}

/** re-parse at the destination and hand back what the listing will say */
function readBack(userDataDir: string, dir: string): InstalledPlugin {
  const manifest = readManifest(dir)
  const plugin = manifest ? describePlugin(dir, manifest, readPluginState(userDataDir)) : null
  if (!plugin) throw new Error('the plugin could not be read back')
  return plugin
}

/**
 * Install a folder the user picked. The manifest is read and validated before
 * anything is copied, the folder goes through the same symlink and size walk a
 * skill import does, and it lands under the name it will be listed by — via a
 * scratch dir renamed into place, so a refusal or a failed rename leaves
 * nothing behind. An existing plugin of that name is refused, not overwritten.
 */
export function installPluginDir(dir: string, userDataDir: string): InstalledPlugin {
  const manifest = requireManifest(dir)
  const name = effectiveName(manifest.fields.name, basename(dir))
  if (!name) throw new Error('the plugin has no usable name in its manifest or folder')
  refuseTaken(name, userDataDir)
  inspectImportDir(dir, 'plugin')
  const root = pluginsRoot(userDataDir)
  mkdirSync(root, { recursive: true })
  const scratch = mkdtempSync(join(root, '.import-'))
  const dest = join(root, name)
  try {
    cpSync(dir, scratch, { recursive: true })
    renameSync(scratch, dest)
  } catch (e) {
    rmSync(scratch, { recursive: true, force: true })
    throw e
  }
  return readBack(userDataDir, dest)
}

/**
 * Read a manifest out of the archive, before any byte is written. The same
 * preference as `readManifest`: the Claude entry first, a plain one when that
 * is missing or corrupt.
 */
async function manifestFromEntries(
  files: readonly JSZip.JSZipObject[],
  prefix: string,
): Promise<Manifest> {
  const candidates: Array<{ entry: JSZip.JSZipObject; format: InstalledPlugin['format'] }> = []
  const claude = files.find((f) => f.name === `${prefix}.claude-plugin/plugin.json`)
  if (claude) candidates.push({ entry: claude, format: 'claude' })
  const plain = files.find((f) => f.name === `${prefix}plugin.json`)
  if (plain) candidates.push({ entry: plain, format: 'plain' })
  for (const { entry, format } of candidates) {
    try {
      const parsed: unknown = JSON.parse(await entry.async('string'))
      if (parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed)) {
        return { format, fields: parsed as Record<string, unknown> }
      }
    } catch {
      // corrupt JSON of one convention: the other is still tried
    }
  }
  throw new Error('the archive has no parseable plugin manifest')
}

/**
 * Install a zip the user picked. The plugin may sit at the archive's root
 * (named from its manifest, or the zip's own stem when the manifest says
 * nothing usable) or under exactly one top-level folder (that folder's name is
 * the fallback). The manifest is parsed out of the archive and validated
 * before any byte is written, the entries go through the same zip-slip,
 * symlink and size checks a skill zip does, and everything lands in a scratch
 * dir renamed into place at the end — a refusal half-way leaves nothing behind.
 */
export async function installPluginZip(
  zipPath: string,
  userDataDir: string,
): Promise<InstalledPlugin> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(readFileSync(zipPath))
  } catch {
    throw new Error('the file is not a zip archive')
  }
  const files = Object.values(zip.files).filter((f) => !f.dir && !f.name.startsWith('__MACOSX/'))
  assertSafeZipEntries(files)
  const manifestAt = (prefix: string): boolean =>
    files.some((f) => f.name === `${prefix}.claude-plugin/plugin.json`) ||
    files.some((f) => f.name === `${prefix}plugin.json`)
  const tops = new Set(files.map((f) => f.name.split('/')[0]))
  let prefix = ''
  let folderName: string
  if (manifestAt('')) {
    folderName = basename(zipPath).replace(/\.zip$/i, '')
  } else if (tops.size === 1 && manifestAt(`${[...tops][0]}/`)) {
    prefix = `${[...tops][0]}/`
    folderName = [...tops][0]
  } else {
    throw new Error('the archive has no plugin manifest at its root or under one folder')
  }
  const manifest = await manifestFromEntries(files, prefix)
  const name = effectiveName(manifest.fields.name, folderName)
  if (!name) throw new Error('the plugin has no usable name in its manifest or archive')
  refuseTaken(name, userDataDir)
  const root = pluginsRoot(userDataDir)
  mkdirSync(root, { recursive: true })
  const scratch = mkdtempSync(join(root, '.import-'))
  const dest = join(root, name)
  try {
    let total = 0
    for (const f of files) {
      const bytes = await inflateWithin(f, MAX_IMPORT_BYTES - total, 'plugin')
      total += bytes.length
      const rel = f.name.slice(prefix.length).split('/').join(sep)
      const target = join(scratch, rel)
      mkdirSync(dirname(target), { recursive: true })
      writeFileSync(target, bytes)
    }
    if (!readManifest(scratch)) throw new Error('the plugin could not be read back')
    renameSync(scratch, dest)
  } catch (e) {
    rmSync(scratch, { recursive: true, force: true })
    throw e
  }
  return readBack(userDataDir, dest)
}

/**
 * Enable or disable one plugin. `name` must be a name the listing itself
 * reported — the same only-what-was-vouched-for rule the skills doors follow,
 * so the renderer can only pick from what it was shown, never name a folder.
 */
export async function setPluginEnabled(
  userDataDir: string,
  name: string,
  enabled: boolean,
): Promise<void> {
  if (typeof name !== 'string' || typeof enabled !== 'boolean') {
    throw new Error('a plugin name and an enabled flag are required')
  }
  if (!scanPlugins(userDataDir).some(({ plugin }) => plugin.name === name)) {
    throw new Error('unknown plugin')
  }
  const state = readPluginState(userDataDir)
  state[name] = { enabled }
  await writePluginState(userDataDir, state)
}

/**
 * Delete one plugin — its folder and its state entry, so a future install of
 * the same name starts fresh. `name` must be a listed plugin's, so a removal
 * can only take what the listing vouched for.
 */
export async function removePlugin(userDataDir: string, name: string): Promise<void> {
  if (typeof name !== 'string') throw new Error('unknown plugin')
  const scanned = scanPlugins(userDataDir).find(({ plugin }) => plugin.name === name)
  if (!scanned) throw new Error('unknown plugin')
  rmSync(scanned.dir, { recursive: true, force: true })
  const state = readPluginState(userDataDir)
  if (!(name in state)) return
  delete state[name]
  await writePluginState(userDataDir, state)
}

/**
 * A component as the AI panel sees it: the namespaced name it is invoked by,
 * the description from its frontmatter or the same derivation a
 * frontmatter-less import gets, relevance classified on the component's own
 * name, and the plugin it rode in on.
 */
function describeComponent(
  pluginName: string,
  component: PluginComponent,
  text: string,
): FoundSkill {
  const fm = parseSkillFrontmatter(text)
  const meta = fm
    ? { name: fm.name, description: fm.description }
    : deriveSkillMeta(component.name, text)
  const description = meta?.description || component.name
  return {
    name: `${pluginName}:${component.name}`,
    description,
    path: component.path,
    relevance: classifySkill(meta?.name ?? component.name, description),
    plugin: pluginName,
  }
}

/**
 * Every skill and command of every enabled plugin, as the AI panel lists them:
 * invoked as `<plugin>:<name>`, read from the real files. One plugin's
 * components stay together (skills before commands, each sorted), in listing
 * order. A file that cannot be read is an absence, not an error — a locked
 * component must not hide the plugin's others. Disabled plugins contribute
 * nothing.
 */
export async function pluginSkills(userDataDir: string): Promise<FoundSkill[]> {
  const out: FoundSkill[] = []
  for (const { dir, plugin } of scanPlugins(userDataDir)) {
    if (!plugin.enabled) continue
    for (const component of [...skillComponents(dir), ...commandComponents(dir)]) {
      let text: string
      try {
        text = await readFile(component.path, 'utf8')
      } catch {
        continue
      }
      out.push(describeComponent(plugin.name, component, text))
    }
  }
  return out
}

/**
 * The body of a plugin component for `ai:skill-text`: the prose under the
 * frontmatter, or the whole file when there is none — a command .md often
 * carries only a description, and the derivation sorts that out. Plugin files
 * are read-only through this door: editing happens on a user-owned copy the
 * foreign-import flow makes, never on the plugin's own folder.
 */
export function readPluginSkillBody(path: string): string {
  let text: string
  try {
    text = readFileSync(path, 'utf8')
  } catch {
    throw new Error('the plugin skill could not be read')
  }
  const fm = parseSkillFrontmatter(text)
  if (fm) return fm.body
  return deriveSkillMeta(basename(path).replace(/\.md$/i, ''), text)?.body ?? text
}
