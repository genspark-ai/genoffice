/**
 * pi resource bridge.
 *
 * GenOffice does not re-implement resource discovery. Marketplace skills and
 * plugins are handed to pi's own settings + package manager, so one install is
 * visible to (a) the embedded pi session, (b) the `pi` CLI pointed at the same
 * agent dir, and (c) this server's Settings UI — from a single source of truth
 * on disk:
 *
 *   DATA_DIR/pi-agent/          pi's agent dir (settings.json, auth, npm cache)
 *   DATA_DIR/pi-skills/<id>/    installed marketplace skill (SKILL.md per id)
 *   DATA_DIR/pi-plugins/<id>/   installed marketplace plugin (a pi package)
 *
 * A "plugin" is a pi package: a directory with `package.json` (`pi` manifest or
 * conventional `extensions/` + `skills/` dirs). Plugins that ship an uploaded
 * extension source get a real `extensions/index.ts`; plugins without one are
 * installed as guidance-only packages (a generated SKILL.md) and never fake a
 * tool. Every registration goes through pi's SettingsManager, and every read
 * goes through pi's DefaultPackageManager — no bespoke loader, no duplicated
 * frontmatter rules.
 *
 * Disabled resources are moved out of the discovery path (skills) or registered
 * as `{ source, autoload: false }` packages (plugins) so "disabled" means the
 * agent really stops seeing them instead of only flipping a UI flag.
 */
import {
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { join, resolve, sep } from 'node:path'
import {
  DefaultPackageManager,
  SettingsManager,
  loadSkillsFromDir,
  type PackageSource,
  type Skill,
} from '@earendil-works/pi-coding-agent'
import { DATA_DIR } from '../common/index'

/** Working directory pi resolves relative package paths against. */
export const PI_CWD = join(DATA_DIR, 'pi-cwd')
/** pi's agent dir — holds settings.json, auth.json and the npm package cache. */
export const PI_AGENT_DIR = join(DATA_DIR, 'pi-agent')
/** Skills installed from the marketplace, one `<id>/SKILL.md` per skill. */
export const PI_SKILLS_DIR = join(DATA_DIR, 'pi-skills')
/** Disabled skills are parked here, outside pi's discovery path. */
export const PI_SKILLS_DISABLED_DIR = join(DATA_DIR, 'pi-skills-disabled')
/** Local plugin packages built from marketplace entries. */
export const PI_PLUGIN_DIR = join(DATA_DIR, 'pi-plugins')
/** Scratch space for validating an uploaded SKILL.md before it is published. */
export const PI_STAGING_DIR = join(DATA_DIR, 'pi-staging')

for (const dir of [PI_CWD, PI_AGENT_DIR, PI_SKILLS_DIR, PI_PLUGIN_DIR, PI_STAGING_DIR]) {
  try {
    mkdirSync(dir, { recursive: true })
  } catch {
    /* a read-only data dir must not crash the server at import time */
  }
}

/** Managed roots — resources under these are the ones GenOffice owns. */
const MANAGED_ROOTS = [PI_SKILLS_DIR, PI_PLUGIN_DIR, PI_AGENT_DIR]

function isManagedPath(path: string): boolean {
  const abs = resolve(path)
  return MANAGED_ROOTS.some((root) => abs === root || abs.startsWith(root + sep))
}

/** Open pi's settings for our agent dir. A fresh manager per call keeps
 *  externally edited settings.json visible without a reload dance. */
export function openPiSettings(): SettingsManager {
  return SettingsManager.create(PI_CWD, PI_AGENT_DIR)
}

export function piSettingsPath(): string {
  return join(PI_AGENT_DIR, 'settings.json')
}

function openPackageManager(settings = openPiSettings()): DefaultPackageManager {
  return new DefaultPackageManager({ cwd: PI_CWD, agentDir: PI_AGENT_DIR, settingsManager: settings })
}

/** Plain source string of a settings entry, whether it is a bare string or an
 *  `{ source, autoload }` filter object. */
export function packageSourceOf(entry: PackageSource): string {
  return typeof entry === 'string' ? entry : entry.source
}

// ── skill dir registration ──────────────────────────────────────────

/**
 * Make sure pi's settings point at the marketplace skills directory. Without
 * this the installed SKILL.md files sit on disk and no pi session ever sees
 * them — the exact gap that made "install" a UI-only illusion.
 */
export async function ensureSkillDirRegistered(): Promise<void> {
  const settings = openPiSettings()
  const paths = settings.getSkillPaths()
  if (paths.includes(PI_SKILLS_DIR)) return
  // Keep the user's own pi skill paths untouched; we only add ours.
  settings.setSkillPaths([...paths, PI_SKILLS_DIR])
  await settings.flush()
}

/** Path of an installed skill's directory (active or parked). */
export function skillDir(id: string, enabled = true): string {
  return join(enabled ? PI_SKILLS_DIR : PI_SKILLS_DISABLED_DIR, id)
}

export interface SkillToggleResult {
  ok: boolean
  /** true when a directory actually moved (i.e. the agent's view changed) */
  moved: boolean
  /** where the skill lives after the call */
  path?: string
  error?: string
}

/**
 * Enable/disable one installed skill for real: the SKILL.md directory moves
 * between the discovery root and the parked root. pi's loader only scans the
 * configured directory, so a parked skill is genuinely invisible to the agent.
 */
export async function setSkillEnabled(id: string, enabled: boolean): Promise<SkillToggleResult> {
  await ensureSkillDirRegistered()
  const active = skillDir(id, true)
  const parked = skillDir(id, false)
  try {
    if (!enabled && existsSync(active)) {
      mkdirSync(PI_SKILLS_DISABLED_DIR, { recursive: true })
      rmSync(parked, { recursive: true, force: true })
      renameSync(active, parked)
      return { ok: true, moved: true, path: parked }
    }
    if (enabled && existsSync(parked)) {
      mkdirSync(PI_SKILLS_DIR, { recursive: true })
      rmSync(active, { recursive: true, force: true })
      renameSync(parked, active)
      return { ok: true, moved: true, path: active }
    }
    return { ok: true, moved: false, path: enabled ? active : parked }
  } catch (err) {
    return { ok: false, moved: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/** Remove an installed skill from both the active and parked roots. */
export function removeSkillFromPi(id: string): void {
  rmSync(skillDir(id, true), { recursive: true, force: true })
  rmSync(skillDir(id, false), { recursive: true, force: true })
}

/**
 * Validate a SKILL.md body with pi's own loader: stage it in a scratch
 * directory and check that pi parses it into a skill without diagnostics.
 *
 * Staging matters: pi skips dot-prefixed directories and refuses a skill with
 * no description, so running the publisher's file through the real parser is
 * the only way to promise "this is what the agent will see". The scratch root
 * lives outside the skills directory so a staging copy can never be picked up
 * by a concurrent discovery pass.
 */
export function validateSkillMarkdown(
  id: string,
  markdown: string,
): { ok: true; name: string; description: string } | { ok: false; error: string } {
  const stagedDir = join(PI_STAGING_DIR, `validate-${id}-${Date.now().toString(36)}`)
  try {
    rmSync(stagedDir, { recursive: true, force: true })
    mkdirSync(stagedDir, { recursive: true })
    writeFileSync(join(stagedDir, 'SKILL.md'), markdown, 'utf-8')
    const parsed = loadSkillsFromDir({ dir: stagedDir, source: 'genoffice-validation' })
    const match = parsed.skills.find((s) => resolve(s.filePath) === resolve(join(stagedDir, 'SKILL.md')))
    if (!match) {
      const detail = parsed.diagnostics.map((d) => d.message).join('; ')
      return {
        ok: false,
        error: detail || 'missing a frontmatter description (pi refuses to load a skill without one)',
      }
    }
    if (parsed.diagnostics.length > 0) {
      return { ok: false, error: parsed.diagnostics.map((d) => d.message).join('; ') }
    }
    return { ok: true, name: match.name, description: match.description }
  } finally {
    rmSync(stagedDir, { recursive: true, force: true })
  }
}

// ── plugin packages ─────────────────────────────────────────────────

export interface PluginArtifact {
  filename: string
  content: string
}

export interface PluginPackageInput {
  id: string
  name: string
  description: string
  version: string
  author?: string
  longDescription?: string
  tools?: string[]
  scopes?: string[]
  requirements?: string[]
  category?: string
  homepage?: string
  /** Uploaded pi extension module — the plugin's executable part. */
  artifact?: PluginArtifact
}

export interface PluginPackageInfo {
  ok: boolean
  packageDir?: string
  /** Absolute paths pi will load extensions from (empty for guidance-only). */
  extensions?: string[]
  /** Absolute paths of SKILL.md files the package contributes. */
  skills?: string[]
  /** true when the package ships executable extension code */
  hasCode?: boolean
  error?: string
}

export function pluginPackageDir(id: string): string {
  return join(PI_PLUGIN_DIR, id)
}

/** Only plain extension module names are allowed — no path traversal, no
 *  nested lookups. pi imports `.ts`/`.js` from the package's extensions dir. */
export function sanitizeExtensionFilename(name: string): string | null {
  const base = name.split(/[\\/]/).pop() ?? ''
  if (!/^[A-Za-z0-9._-]{1,64}$/.test(base)) return null
  if (!/\.(ts|mts|js|mjs)$/.test(base)) return null
  if (base.startsWith('.')) return null
  return base
}

function renderPluginSkill(input: PluginPackageInput): string {
  const tools = input.tools?.length ? input.tools.join(', ') : '(none)'
  const scopes = input.scopes?.length ? input.scopes.join(', ') : '(none)'
  const requirements = input.requirements?.length ? input.requirements.join(', ') : '(none)'
  const instructions = input.artifact
    ? [
        `This plugin ships a pi extension (${input.artifact.filename}) that registers its tools. Use them directly; do not invent tool names.`,
        `Tools: ${tools}.`,
        `Required setup: ${requirements}. Ask the user to configure anything missing instead of guessing credentials.`,
      ]
    : [
        `This plugin is installed as guidance only — it ships no executable extension, so its tools (${tools}) are NOT registered.`,
        'Treat the capability as advisory: explain the workflow, and say plainly that the executable part is not installed when the user asks for automation.',
        `Required setup: ${requirements}.`,
      ]
  return [
    '---',
    `name: ${input.id}`,
    `description: ${input.description.replace(/\s+/g, ' ').slice(0, 280)}`,
    `display_name: ${input.name.replace(/\s+/g, ' ')}`,
    '---',
    '',
    `# ${input.name}`,
    '',
    input.longDescription?.trim() || input.description,
    '',
    '## Instructions',
    '',
    ...instructions.map((line) => `- ${line}`),
    '',
    `Permission scopes declared by the plugin: ${scopes}.`,
  ].join('\n')
}

/**
 * Materialize a marketplace plugin as a pi package and register it in pi's
 * settings. Idempotent: re-installing overwrites the package contents and keeps
 * a single settings entry.
 */
export async function installPluginPackage(input: PluginPackageInput): Promise<PluginPackageInfo> {
  const dir = pluginPackageDir(input.id)
  try {
    rmSync(dir, { recursive: true, force: true })
    mkdirSync(dir, { recursive: true })

    const extensionName = input.artifact ? sanitizeExtensionFilename(input.artifact.filename) : null
    if (input.artifact && !extensionName) {
      return { ok: false, error: `Unsupported extension filename "${input.artifact.filename}"` }
    }

    // The `pi` manifest only lists directories that exist: pi reads globs
    // literally, and a manifest entry pointing at a missing dir is a load error.
    const manifest = {
      name: `@genoffice-plugin/${input.id}`,
      version: input.version,
      description: input.description,
      author: input.author || 'GenOffice Marketplace',
      keywords: ['pi-package', 'genoffice-plugin'],
      ...(input.homepage ? { homepage: input.homepage } : {}),
      pi: {
        ...(extensionName ? { extensions: ['./extensions'] } : {}),
        skills: ['./skills'],
      },
    }
    writeFileSync(join(dir, 'package.json'), JSON.stringify(manifest, null, 2), 'utf-8')

    const extensions: string[] = []
    if (extensionName && input.artifact) {
      const extDir = join(dir, 'extensions')
      mkdirSync(extDir, { recursive: true })
      const extPath = join(extDir, extensionName)
      writeFileSync(extPath, input.artifact.content, 'utf-8')
      extensions.push(extPath)
    }

    const skills: string[] = []
    const skillRoot = join(dir, 'skills', input.id)
    mkdirSync(skillRoot, { recursive: true })
    const skillPath = join(skillRoot, 'SKILL.md')
    writeFileSync(skillPath, renderPluginSkill(input), 'utf-8')
    skills.push(skillPath)

    writeFileSync(
      join(dir, 'README.md'),
      `# ${input.name}\n\n${input.longDescription?.trim() || input.description}\n\n` +
        `Installed by GenOffice from the plugin marketplace. Package layout:\n\n` +
        `- \`package.json\` — pi package manifest\n` +
        (extensionName ? `- \`extensions/${extensionName}\` — pi extension module\n` : '') +
        `- \`skills/${input.id}/SKILL.md\` — agent guidance\n`,
      'utf-8',
    )
    writeFileSync(
      join(dir, '.genoffice-plugin.json'),
      JSON.stringify(
        {
          id: input.id,
          version: input.version,
          author: input.author ?? null,
          category: input.category ?? null,
          tools: input.tools ?? [],
          scopes: input.scopes ?? [],
          requirements: input.requirements ?? [],
          hasCode: !!extensionName,
          artifact: extensionName ? input.artifact!.filename : null,
          installedAt: new Date().toISOString(),
        },
        null,
        2,
      ),
      'utf-8',
    )

    await registerPluginPackage(input.id, true)
    return { ok: true, packageDir: dir, extensions, skills, hasCode: !!extensionName }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/** Read the manifest GenOffice wrote next to an installed plugin package. */
export function readPluginManifest(id: string): {
  hasCode: boolean
  artifact: string | null
  installedAt: string | null
} | null {
  const file = join(pluginPackageDir(id), '.genoffice-plugin.json')
  if (!existsSync(file)) return null
  try {
    const raw = JSON.parse(readFileSync(file, 'utf-8')) as {
      hasCode?: boolean
      artifact?: string | null
      installedAt?: string
    }
    return {
      hasCode: raw.hasCode === true,
      artifact: raw.artifact ?? null,
      installedAt: raw.installedAt ?? null,
    }
  } catch {
    return null
  }
}

/**
 * Register (or disable) a plugin package with pi. Disabled packages stay in
 * settings as `{ source, autoload: false }`, which pi resolves to "this package
 * contributes nothing" — the same mechanism `pi config` uses to disable
 * resources from installed packages.
 */
async function registerPluginPackage(id: string, enabled: boolean): Promise<void> {
  const dir = pluginPackageDir(id)
  const settings = openPiSettings()
  const kept = settings.getPackages().filter((entry) => resolve(packageSourceOf(entry)) !== resolve(dir))
  const entry: PackageSource = enabled ? dir : { source: dir, autoload: false }
  settings.setPackages([...kept, entry])
  await settings.flush()
}

export interface PluginToggleResult {
  ok: boolean
  enabled?: boolean
  error?: string
}

export async function setPluginPackageEnabled(id: string, enabled: boolean): Promise<PluginToggleResult> {
  if (!existsSync(pluginPackageDir(id))) {
    return { ok: false, error: `Plugin package "${id}" is not installed` }
  }
  try {
    await registerPluginPackage(id, enabled)
    return { ok: true, enabled }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/**
 * Enable/disable any registered package source (local path or `npm:`/`git:`)
 * with the same `autoload` mechanism pi uses, so a disabled plugin stops
 * contributing extensions/skills on the next resource reload.
 */
export async function setPackageSourceEnabled(
  source: string,
  enabled: boolean,
): Promise<{ ok: boolean; error?: string }> {
  try {
    const settings = openPiSettings()
    const entries = settings.getPackages()
    const present = entries.some((entry) => packageSourceOf(entry) === source)
    if (!present && !enabled) return { ok: true }
    const kept = entries.filter((entry) => packageSourceOf(entry) !== source)
    const next: PackageSource = enabled ? source : { source, autoload: false }
    settings.setPackages([...kept, next])
    await settings.flush()
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

/** Unregister a plugin package and delete its files. */
export async function removePluginPackage(id: string): Promise<{ ok: boolean; error?: string }> {
  const dir = pluginPackageDir(id)
  try {
    const settings = openPiSettings()
    const kept = settings.getPackages().filter((entry) => resolve(packageSourceOf(entry)) !== resolve(dir))
    if (kept.length !== settings.getPackages().length) {
      settings.setPackages(kept)
      await settings.flush()
    }
    rmSync(dir, { recursive: true, force: true })
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

// ── remote pi packages (npm / git / absolute path) ───────────────────

export function isRemotePackageSource(source: string): boolean {
  return /^(npm:|git:|https?:)/.test(source)
}

/**
 * Install a real pi package (npm:/git:) through pi's package manager. The
 * package lands in pi's own cache under the agent dir and is persisted to
 * settings, so its extensions/skills/prompts/themes load exactly like any
 * other pi package.
 */
export async function installPiPackage(
  source: string,
): Promise<{ ok: boolean; source?: string; error?: string }> {
  const trimmed = source.trim()
  if (!/^(npm:|git:|https?:)/.test(trimmed)) {
    return { ok: false, error: 'Package source must start with npm:, git: or https://' }
  }
  try {
    const settings = openPiSettings()
    const manager = openPackageManager(settings)
    await manager.installAndPersist(trimmed, { local: false })
    return { ok: true, source: trimmed }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

export async function removePiPackage(source: string): Promise<{ ok: boolean; error?: string }> {
  try {
    const settings = openPiSettings()
    const manager = openPackageManager(settings)
    await manager.removeAndPersist(source, { local: false })
    return { ok: true }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

// ── resolution: what pi would actually load right now ───────────────

export interface ResolvedResource {
  path: string
  /** file name or skill name, for display */
  name: string
  /** false when the source is registered but disabled/filtered out */
  enabled: boolean
  /** true when the file lives under a GenOffice-managed root */
  managed: boolean
}

export interface PiResourceReport {
  cwd: string
  agentDir: string
  settingsPath: string
  /** pi's own resolution of registered packages + settings paths */
  extensions: ResolvedResource[]
  skills: ResolvedResource[]
  prompts: ResolvedResource[]
  themes: ResolvedResource[]
  /** resources that come from the user's global pi setup, not GenOffice */
  external: { extensions: number; skills: number; prompts: number; themes: number }
  packages: Array<{ source: string; scope: string; filtered: boolean; installedPath?: string }>
  diagnostics: Array<{ type: string; message: string; path?: string }>
}

function resourceName(path: string): string {
  const parts = path.split(sep)
  const base = parts[parts.length - 1] ?? path
  if (base === 'SKILL.md') return parts[parts.length - 2] ?? base
  return base
}

/**
 * Ask pi what it would load for our agent dir. This calls pi's
 * DefaultPackageManager (path-level resolution, no code execution) so the
 * report is the loader's own view rather than a GenOffice re-derivation.
 * Missing remote packages are skipped instead of triggering an install.
 */
export async function resolvePiResources(): Promise<PiResourceReport> {
  await ensureSkillDirRegistered()
  const settings = openPiSettings()
  const manager = openPackageManager(settings)
  const resolved = await manager.resolve(async () => 'skip')

  const map = (list: Array<{ path: string; enabled: boolean }>): ResolvedResource[] =>
    list.map((entry) => ({
      path: entry.path,
      name: resourceName(entry.path),
      enabled: entry.enabled,
      managed: isManagedPath(entry.path),
    }))

  const extensions = map(resolved.extensions)
  const skills = map(resolved.skills)
  const prompts = map(resolved.prompts)
  const themes = map(resolved.themes)

  const external = {
    extensions: extensions.filter((e) => !e.managed).length,
    skills: skills.filter((e) => !e.managed).length,
    prompts: prompts.filter((e) => !e.managed).length,
    themes: themes.filter((e) => !e.managed).length,
  }

  return {
    cwd: PI_CWD,
    agentDir: PI_AGENT_DIR,
    settingsPath: piSettingsPath(),
    extensions,
    skills,
    prompts,
    themes,
    external,
    packages: manager.listConfiguredPackages(),
    diagnostics: settings.drainErrors().map((e) => ({
      type: 'settings',
      message: e.error.message,
      ...(e.path ? { path: e.path } : {}),
    })),
  }
}

/** Installed skill files with their parsed frontmatter, straight from pi's
 *  loader. `enabled` mirrors whether the file sits in the discovery root. */
export function installedSkills(): Array<Skill & { enabled: boolean; managed: boolean }> {
  const parse = (dir: string, enabled: boolean) => {
    if (!existsSync(dir)) return []
    const parsed = loadSkillsFromDir({ dir, source: 'genoffice-marketplace' })
    return parsed.skills.map((skill) => ({ ...skill, enabled, managed: isManagedPath(skill.filePath) }))
  }
  return [...parse(PI_SKILLS_DIR, true), ...parse(PI_SKILLS_DISABLED_DIR, false)]
}

/** Size guard for uploaded artifacts (extension modules and SKILL.md bodies). */
export function artifactBytes(content: string): number {
  return Buffer.byteLength(content, 'utf-8')
}

export function readableBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`
}

/** True when the path exists and is a file (used by upload validation). */
export function isFile(path: string): boolean {
  try {
    return statSync(path).isFile()
  } catch {
    return false
  }
}
