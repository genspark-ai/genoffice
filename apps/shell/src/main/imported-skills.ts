import {
  cpSync,
  existsSync,
  lstatSync,
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
import { lstat, readFile, readdir } from 'node:fs/promises'
import { basename, dirname, isAbsolute, join, resolve, sep } from 'node:path'
import JSZip from 'jszip'
import { classifySkill, parseSkillFrontmatter, type SkillFrontmatter } from '@genoffice/agent-core'
import type {
  FoundSkill,
  ForeignSkillCandidate,
  ForeignSkillSource,
  RecommendedSkill,
} from '../shared/found-skill'

/**
 * The skills a user installed into GenOffice, plus the ones already on the
 * machine.
 *
 * Only `<userData>/skills/<name>/SKILL.md` counts as installed: that is what
 * `usableSkills` lists and what a turn may read. The other agent products'
 * own directories — Claude's commands, Codex prompts, Cursor rules — are
 * scanned too (`scanForeignSkills`), but strictly as import candidates: a
 * person who already wrote a prompt there can pull it in with one click
 * instead of re-authoring it. Nothing foreign runs or lists until it is
 * imported, and importing is the same copy-into-our-folder operation a picked
 * folder or zip goes through.
 */

export type { FoundSkill }

/**
 * Under the app's own userData, not under the save directory: Home lists every
 * `.md` file under the save directory as a document, so a skills folder there
 * put every imported `SKILL.md` into the user's file list.
 */
export function skillsRoot(userDataDir: string): string {
  return join(userDataDir, 'skills')
}

function describe(name: string, path: string, text: string): FoundSkill | null {
  const fm = parseSkillFrontmatter(text)
  if (!fm) return null
  return {
    name,
    description: fm.description,
    path,
    relevance: classifySkill(fm.name, fm.description),
  }
}

/**
 * A permission error on one skill must not hide the others, so each is read
 * independently and a failure is just an absence. Async because this runs on
 * the main process while windows are open.
 */
export async function usableSkills(userDataDir: string): Promise<FoundSkill[]> {
  const root = skillsRoot(userDataDir)
  let entries: string[]
  try {
    entries = await readdir(root)
  } catch {
    return []
  }
  const out: FoundSkill[] = []
  for (const name of entries.sort()) {
    if (name.startsWith('.')) continue
    const path = join(root, name, 'SKILL.md')
    let text: string
    try {
      text = await readFile(path, 'utf8')
    } catch {
      continue
    }
    const skill = describe(name, path, text)
    if (skill) out.push(skill)
  }
  return out
}

/** the body of a skill, for the turn that runs it */
export function readSkillBody(path: string): string {
  const fm = parseSkillFrontmatter(readFileSync(path, 'utf8'))
  return fm?.body ?? ''
}

const MAX_IMPORT_BYTES = 10 * 1024 * 1024

/**
 * A symlink is refused rather than dereferenced: `cpSync` without
 * `dereference` would carry the link across and the imported skill would keep
 * reaching back into whatever it points at, while dereferencing would copy in
 * arbitrarily much from anywhere on disk. The same walk bounds the copy.
 */
function inspectSkillDir(dir: string): void {
  let rootIsLink: boolean
  try {
    rootIsLink = lstatSync(dir).isSymbolicLink()
  } catch {
    throw new Error('the skill directory could not be read')
  }
  if (rootIsLink) throw new Error('the skill directory is a symlink, and imports refuse those')
  let total = 0
  const stack: string[] = [dir]
  while (stack.length > 0) {
    const cur = stack.pop()!
    let entries: Dirent[]
    try {
      entries = readdirSync(cur, { withFileTypes: true })
    } catch {
      throw new Error('the skill directory could not be read')
    }
    for (const entry of entries) {
      if (entry.isSymbolicLink()) {
        throw new Error('the skill contains a symlink, and imports refuse those')
      }
      const full = join(cur, entry.name)
      if (entry.isDirectory()) {
        stack.push(full)
        continue
      }
      let size: number
      try {
        size = statSync(full).size
      } catch {
        throw new Error('the skill directory could not be read')
      }
      total += size
      if (total > MAX_IMPORT_BYTES) {
        throw new Error('the skill is larger than the 10 MB import cap')
      }
    }
  }
}

function readBack(dest: string): FoundSkill {
  const skill = describe(
    basename(dest),
    join(dest, 'SKILL.md'),
    readFileSync(join(dest, 'SKILL.md'), 'utf8'),
  )
  if (!skill) throw new Error('the skill could not be read back')
  return skill
}

function requireSkillMd(dir: string): void {
  let text: string
  try {
    text = readFileSync(join(dir, 'SKILL.md'), 'utf8')
  } catch {
    throw new Error('the folder has no SKILL.md')
  }
  if (!parseSkillFrontmatter(text)) {
    throw new Error('SKILL.md needs name and description frontmatter')
  }
}

/**
 * Copy a folder the user picked into our skills folder.
 *
 * The whole directory goes rather than the SKILL.md alone: a skill routinely
 * ships reference tables, templates and helper scripts next to its
 * instructions. Refuses rather than overwrites when the name is taken.
 */
export function importSkillDir(skillDir: string, userDataDir: string): FoundSkill {
  const dest = join(skillsRoot(userDataDir), basename(skillDir))
  if (existsSync(dest)) throw new Error('a skill of that name is already there')
  requireSkillMd(skillDir)
  inspectSkillDir(skillDir)
  cpSync(skillDir, dest, { recursive: true })
  return readBack(dest)
}

/**
 * The frontmatter block a strict parse refuses for want of a `name` — a
 * Cursor rule's `description` + `globs`, a Claude command's `description` +
 * `allowed-tools`. Rather than own a second parser, a placeholder `name` line
 * is injected at the end of the block and the real parser runs again: it then
 * reads the description, the body and the unknown keys with its own quoting
 * and continuation rules. The placeholder never escapes — only the other
 * fields are used.
 */
function parseUnnamedFrontmatter(text: string): SkillFrontmatter | null {
  const stripped = text.replace(/^\uFEFF/, '')
  const open = /^---[ \t]*\r?\n/.exec(stripped)
  if (!open) return null
  const after = stripped.slice(open[0].length)
  const close = /^---[ \t]*(\r?\n|$)/m.exec(after)
  if (!close) return null
  const block = after.slice(0, close.index)
  return parseSkillFrontmatter(`${open[0]}${block}name: @synthetic\n${after.slice(close.index)}`)
}

/** how long a description squeezed out of the body's first line may run */
const MAX_DERIVED_DESCRIPTION = 160

/**
 * A file stem turned into a folder name: lowercased, whitespace runs to single
 * dashes, the characters a folder name cannot carry replaced the same way,
 * nothing left that `isSkillName` would refuse (a leading dot, a trailing dot
 * or dash that Win32 would strip). The stem of a real basename holds no
 * separators, but a stem such as `Draft: v2?` still needs this. Null when
 * nothing usable remains — `???`, `...`, blank.
 */
function sanitizeSkillStem(stem: string): string | null {
  const cleaned = stem
    .trim()
    .toLowerCase()
    .replace(/[/\\:*?"<>|]+/g, '-')
    .replace(/\s+/g, '-')
    .replace(/-{2,}/g, '-')
    .replace(/^[-.]+/, '')
    .replace(/[.\s-]+$/, '')
  return cleaned !== '' && isSkillName(cleaned) ? cleaned : null
}

/** a heading, list or bold marker is furniture, and the words are what's left */
function stripMarkdownFurniture(line: string): string {
  return line
    .replace(/^#{1,6}[ \t]+/, '')
    .replace(/^(?:[-*+]|\d{1,9}[.)])[ \t]+/, '')
    .replace(/\*\*(.+?)\*\*/g, '$1')
    .replace(/__(.+?)__/g, '$1')
    .trim()
}

/**
 * The first line worth showing as a description: the body's opening line once
 * its markdown markers are stripped, capped, because a description is a label
 * rather than the first paragraph. A code fence or a bare rule is furniture
 * and is skipped rather than quoted.
 */
function firstProseLine(body: string): string {
  for (const raw of body.split(/\r?\n/)) {
    const line = raw.trim()
    if (line === '') continue
    if (/^(```|~~~)/.test(line) || /^(-{3,}|\*{3,}|_{3,})$/.test(line)) continue
    const stripped = stripMarkdownFurniture(line)
    if (stripped === '') continue
    return stripped.length > MAX_DERIVED_DESCRIPTION
      ? `${stripped.slice(0, MAX_DERIVED_DESCRIPTION).trimEnd()}…`
      : stripped
  }
  return ''
}

/** what a scan and its import both need from a file that may not be a skill at all */
export interface DerivedSkillMeta {
  name: string
  description: string
  /** the prose under the frontmatter, or the whole file when there is none */
  body: string
  /** frontmatter keys the parser kept but a skill does not model */
  extra: Record<string, string>
}

/**
 * Derive a skill's name and description from any prompt-shaped file — no
 * frontmatter, partial frontmatter (a description without a name, as a Cursor
 * rule carries), or full skill frontmatter. Pure, so the tricky cases are
 * testable without a disk. The rules:
 *
 * - name: the frontmatter's `name` when the strict parse succeeded and it is a
 *   legal folder name — what a frontmatter-bearing import uses today; else the
 *   sanitized file stem (`sanitizeSkillStem`). Null only when neither yields
 *   anything usable, which is the one case that refuses an import.
 * - description: the frontmatter's, used directly however it was parsed; else
 *   the body's first prose line with its markdown markers stripped and capped
 *   at ~160 characters; else the name itself as a neutral label.
 * - body: the prose under the frontmatter fence, or the whole file when there
 *   is none. `extra`: the unknown frontmatter keys, kept so a synthesized
 *   import loses nothing the source said.
 */
export function deriveSkillMeta(fileStem: string, text: string): DerivedSkillMeta | null {
  const fm = parseSkillFrontmatter(text)
  const parsed = fm ?? parseUnnamedFrontmatter(text)
  const name = fm && isSkillName(fm.name) ? fm.name : sanitizeSkillStem(fileStem)
  if (!name) return null
  const body = parsed?.body ?? text.replace(/^\uFEFF/, '')
  // the parser joins a wrapped value onto one line itself; the collapse is
  // for the odd hand-written description that spans lines
  const fromFrontmatter = parsed?.description.replace(/\s*\r?\n\s*/g, ' ').trim() ?? ''
  const description = fromFrontmatter || firstProseLine(body) || name
  return { name, description, body, extra: parsed?.extra ?? {} }
}

/**
 * The SKILL.md written for a file that had none of its own: the derived name
 * and description up front, the unknown frontmatter keys the source carried
 * kept under them — a Cursor rule's `globs` costs nothing here and losing it
 * would cost the trip back — and the body verbatim under the fence.
 */
function synthesizeSkillDoc(meta: DerivedSkillMeta): string {
  const keys = [
    `name: ${meta.name}`,
    `description: ${meta.description}`,
    ...Object.entries(meta.extra).map(([k, v]) => `${k}: ${v}`),
  ]
  return `---\n${keys.join('\n')}\n---\n\n${meta.body.replace(/\r\n/g, '\n')}`
}

/**
 * The write every single-file import shares: refuse a taken name, land the
 * text in a scratch folder, rename it into place. A refusal or a failed
 * rename leaves nothing behind — the pattern `writeSkill` also uses.
 */
function writeSkillDoc(userDataDir: string, name: string, text: string): FoundSkill {
  if (Buffer.byteLength(text) > MAX_IMPORT_BYTES) {
    throw new Error('the skill is larger than the 10 MB import cap')
  }
  const dest = join(skillsRoot(userDataDir), name)
  if (existsSync(dest)) throw new Error('a skill of that name is already there')
  mkdirSync(skillsRoot(userDataDir), { recursive: true })
  const scratch = mkdtempSync(join(skillsRoot(userDataDir), '.import-'))
  try {
    writeFileSync(join(scratch, 'SKILL.md'), text)
    renameSync(scratch, dest)
  } catch (e) {
    rmSync(scratch, { recursive: true, force: true })
    throw e
  }
  return readBack(dest)
}

/**
 * Import a single markdown file as a skill.
 *
 * The file becomes `<name>/SKILL.md`. A file with skill frontmatter keeps the
 * behavior it always had: the name comes from the frontmatter and falls back
 * to the file stem, and the text is copied verbatim. A file without it — a
 * plain note, a Cursor `.mdc` rule whose frontmatter carries a `description`
 * but no `name` — is imported all the same, with the derivation rules
 * `deriveSkillMeta` documents: the name from the sanitized stem, the
 * description from the frontmatter or the first prose line, and a synthesized
 * `---\nname\ndescription\n---\n\n<body>` doc written through the same
 * scratch-then-rename pattern. Refuses rather than overwrites when the name
 * is taken.
 */
export function importSkillMd(mdPath: string, userDataDir: string): FoundSkill {
  let text: string
  try {
    text = readFileSync(mdPath, 'utf8')
  } catch {
    throw new Error('the file could not be read')
  }
  if (Buffer.byteLength(text) > MAX_IMPORT_BYTES) {
    throw new Error('the skill is larger than the 10 MB import cap')
  }
  const stem = basename(mdPath).replace(/\.(md|mdc)$/i, '')
  const fm = parseSkillFrontmatter(text)
  if (fm) {
    const name = [fm.name.trim(), stem].find((n) => isSkillName(n))
    if (!name) throw new Error('the file name cannot be used as a skill name')
    return writeSkillDoc(userDataDir, name, text)
  }
  const meta = deriveSkillMeta(stem, text)
  if (!meta) throw new Error('the file name cannot be used as a skill name')
  return writeSkillDoc(userDataDir, meta.name, synthesizeSkillDoc(meta))
}

function isSkillName(name: string): boolean {
  return /^[^/\\:*?"<>|]+$/.test(name) && isListedName(name)
}

/** What `usableSkills` lists: one readdir entry, not a dot folder, not a path. */
function isListedName(name: string): boolean {
  return name.length > 0 && !/[/\\]/.test(name) && !name.startsWith('.')
}

/**
 * Inflate one entry, giving up the moment it passes `limit`. A tiny archive can
 * declare and inflate to gigabytes, so the cap has to bite during inflation,
 * not after `async('nodebuffer')` has already handed back the whole payload.
 */
function inflateWithin(entry: JSZip.JSZipObject, limit: number): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    let size = 0
    const stream = entry.nodeStream('nodebuffer')
    stream
      .on('data', (chunk: Buffer) => {
        size += chunk.length
        if (size > limit) {
          stream.pause()
          stream.removeAllListeners('data')
          reject(new Error('the skill is larger than the 10 MB import cap'))
          return
        }
        chunks.push(chunk)
      })
      .on('error', reject)
      .on('end', () => resolve(Buffer.concat(chunks)))
  })
}

const S_IFMT = 0o170000
const S_IFLNK = 0o120000

/**
 * Unpack a zip the user picked into our skills folder.
 *
 * Accepts `SKILL.md` at the archive root (the skill is then named after the
 * file) or under exactly one top-level folder (named after the folder). Entry
 * names are checked before any byte is written: an absolute or `..` name, a
 * symlink entry, or more than the import cap of unpacked bytes refuses the
 * whole archive. Everything lands in a scratch dir first and is renamed into
 * place at the end, so a refusal half-way leaves nothing behind.
 */
export async function importSkillZip(zipPath: string, userDataDir: string): Promise<FoundSkill> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(readFileSync(zipPath))
  } catch {
    throw new Error('the file is not a zip archive')
  }
  const files = Object.values(zip.files).filter((f) => !f.dir && !f.name.startsWith('__MACOSX/'))
  for (const f of files) {
    const parts = f.name.split('/')
    if (isAbsolute(f.name) || f.name.includes('\\') || parts.some((p) => p === '..' || p === '')) {
      throw new Error(`the archive has an unsafe entry name: ${f.name}`)
    }
    if (f.unixPermissions != null && (Number(f.unixPermissions) & S_IFMT) === S_IFLNK) {
      throw new Error('the archive contains a symlink, and imports refuse those')
    }
  }
  const atRoot = files.find((f) => f.name === 'SKILL.md')
  const tops = new Set(files.map((f) => f.name.split('/')[0]))
  let prefix = ''
  let name: string
  if (atRoot) {
    name = basename(zipPath).replace(/\.zip$/i, '')
  } else if (tops.size === 1 && files.some((f) => f.name === `${[...tops][0]}/SKILL.md`)) {
    prefix = `${[...tops][0]}/`
    name = [...tops][0]
  } else {
    throw new Error('the archive has no SKILL.md at its root or under one folder')
  }
  if (!isSkillName(name)) {
    throw new Error('the archive name cannot be used as a skill name')
  }
  const dest = join(skillsRoot(userDataDir), name)
  if (existsSync(dest)) throw new Error('a skill of that name is already there')

  mkdirSync(skillsRoot(userDataDir), { recursive: true })
  const scratch = mkdtempSync(join(skillsRoot(userDataDir), '.import-'))
  try {
    let total = 0
    for (const f of files) {
      const bytes = await inflateWithin(f, MAX_IMPORT_BYTES - total)
      total += bytes.length
      const rel = f.name.slice(prefix.length).split('/').join(sep)
      const target = join(scratch, rel)
      mkdirSync(dirname(target), { recursive: true })
      writeFileSync(target, bytes)
    }
    requireSkillMd(scratch)
    renameSync(scratch, dest)
  } catch (e) {
    rmSync(scratch, { recursive: true, force: true })
    throw e
  }
  return readBack(dest)
}

/**
 * Where the other agent products keep their prompt and skill files, under the
 * home directory. `files` takes every `<ext>` file directly inside; `skills`
 * takes every subfolder that holds a SKILL.md (the folder, not the file, is
 * the skill). A location that is missing or unreadable is simply absent — the
 * scan never throws.
 */
type ForeignLocation =
  | { source: ForeignSkillSource; rel: readonly string[]; kind: 'files'; ext: string }
  | { source: ForeignSkillSource; rel: readonly string[]; kind: 'skills' }

const FOREIGN_LOCATIONS: readonly ForeignLocation[] = [
  { source: 'claude', rel: ['.claude', 'commands'], kind: 'files', ext: '.md' },
  { source: 'claude', rel: ['.claude', 'skills'], kind: 'skills' },
  { source: 'codex', rel: ['.codex', 'prompts'], kind: 'files', ext: '.md' },
  { source: 'cursor', rel: ['.cursor', 'rules'], kind: 'files', ext: '.mdc' },
  { source: 'cursor', rel: ['.cursor', 'commands'], kind: 'files', ext: '.md' },
  { source: 'windsurf', rel: ['.windsurf', 'workflows'], kind: 'files', ext: '.md' },
  { source: 'windsurf', rel: ['.codeium', 'windsurf', 'workflows'], kind: 'files', ext: '.md' },
  { source: 'opencode', rel: ['.config', 'opencode', 'command'], kind: 'files', ext: '.md' },
  { source: 'opencode', rel: ['.config', 'opencode', 'commands'], kind: 'files', ext: '.md' },
  { source: 'pi', rel: ['.pi', 'agent', 'commands'], kind: 'files', ext: '.md' },
  { source: 'pi', rel: ['.pi', 'agent', 'skills'], kind: 'skills' },
]

/** a candidate plus what importing it needs: the folder a `SKILL.md` lives in */
interface ForeignEntry extends ForeignSkillCandidate {
  /** the skill folder, when the source file is its SKILL.md; null for a loose file */
  skillDir: string | null
}

/** a read that treats any failure — missing, unreadable, symlinked, oversized — as an absence */
async function readForeignText(path: string): Promise<string | null> {
  try {
    // lstat, not stat: a symlinked SKILL.md inside a folder would otherwise be
    // followed here and refused only at import time
    const s = await lstat(path)
    if (!s.isFile() || s.size > MAX_IMPORT_BYTES) return null
    return await readFile(path, 'utf8')
  } catch {
    return null
  }
}

async function scanForeignEntries(homeDir: string, userDataDir: string): Promise<ForeignEntry[]> {
  const root = skillsRoot(userDataDir)
  const out: ForeignEntry[] = []
  for (const loc of FOREIGN_LOCATIONS) {
    const dir = join(homeDir, ...loc.rel)
    let entries: Dirent[]
    try {
      entries = await readdir(dir, { withFileTypes: true })
    } catch {
      continue
    }
    const byName = (a: Dirent, b: Dirent) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0)
    for (const entry of entries.sort(byName)) {
      if (entry.name.startsWith('.')) continue
      let sourcePath: string
      let stem: string
      let skillDir: string | null = null
      if (loc.kind === 'skills') {
        if (!entry.isDirectory()) continue
        skillDir = join(dir, entry.name)
        sourcePath = join(skillDir, 'SKILL.md')
        stem = entry.name
      } else {
        if (!entry.isFile() || !entry.name.toLowerCase().endsWith(loc.ext)) continue
        sourcePath = join(dir, entry.name)
        stem = entry.name.slice(0, -loc.ext.length)
      }
      const text = await readForeignText(sourcePath)
      if (text === null) continue
      const meta = deriveSkillMeta(stem, text)
      if (!meta) continue
      out.push({
        name: meta.name,
        description: meta.description,
        source: loc.source,
        sourcePath,
        alreadyImported: existsSync(join(root, meta.name)),
        skillDir,
      })
    }
  }
  const bySourceThenName = (a: ForeignEntry, b: ForeignEntry) =>
    a.source < b.source
      ? -1
      : a.source > b.source
        ? 1
        : a.name < b.name
          ? -1
          : a.name > b.name
            ? 1
            : 0
  return out.sort(bySourceThenName)
}

/**
 * Prompt and skill files the other agent products keep on this machine, as
 * import candidates. Name and description come from the same derivation as a
 * picked markdown file (`deriveSkillMeta`), and `alreadyImported` says whether
 * a folder of that name already sits in ours. Nothing is copied, nothing runs,
 * and a missing or unreadable location is an empty list rather than an error.
 */
export async function scanForeignSkills(
  homeDir: string,
  userDataDir: string,
): Promise<ForeignSkillCandidate[]> {
  const entries = await scanForeignEntries(homeDir, userDataDir)
  return entries.map(({ name, description, source, sourcePath, alreadyImported }) => ({
    name,
    description,
    source,
    sourcePath,
    alreadyImported,
  }))
}

/**
 * The file form of a foreign import: the same synthesized doc a frontmatter-less
 * `.md` gets, from a fresh derivation of the file the scan vouched for.
 */
function importForeignSkillFile(sourcePath: string, userDataDir: string): FoundSkill {
  let text: string
  try {
    text = readFileSync(sourcePath, 'utf8')
  } catch {
    throw new Error('the file could not be read')
  }
  if (Buffer.byteLength(text) > MAX_IMPORT_BYTES) {
    throw new Error('the skill is larger than the 10 MB import cap')
  }
  const stem = basename(sourcePath).replace(/\.(md|mdc)$/i, '')
  const meta = deriveSkillMeta(stem, text)
  if (!meta) throw new Error('the file name cannot be used as a skill name')
  return writeSkillDoc(userDataDir, meta.name, synthesizeSkillDoc(meta))
}

/**
 * The folder form of a foreign import: `importSkillDir`'s approach — the whole
 * folder (references, scripts, everything) copied in after the symlink and
 * size walk, its own SKILL.md kept verbatim, refusal when the name is taken —
 * keyed on the name the scan derived, which is the frontmatter's whenever it
 * is a legal folder name. A skill folder still needs its own parseable
 * SKILL.md, the contract a picked folder has always had; the lenient
 * derivation is for loose files.
 */
function importForeignSkillDir(skillDir: string, name: string, userDataDir: string): FoundSkill {
  const dest = join(skillsRoot(userDataDir), name)
  if (existsSync(dest)) throw new Error('a skill of that name is already there')
  requireSkillMd(skillDir)
  inspectSkillDir(skillDir)
  cpSync(skillDir, dest, { recursive: true })
  return readBack(dest)
}

/**
 * Import one of the scanner's candidates.
 *
 * The path is renderer input and is not trusted: the scan runs again, and only
 * a path that fresh scan vouched for is imported — `isKnownSkillPath`'s
 * philosophy, applied at the moment of writing. A loose file is written as the
 * synthesized doc; a skill folder — a `SKILL.md` under `~/.claude/skills` — is
 * imported whole. A name that is already ours is refused, and the error says so.
 */
export async function importForeignSkill(
  sourcePath: string,
  homeDir: string,
  userDataDir: string,
): Promise<FoundSkill> {
  if (typeof sourcePath !== 'string' || !isAbsolute(sourcePath)) {
    throw new Error('that skill was not offered by the scan')
  }
  const entries = await scanForeignEntries(homeDir, userDataDir)
  const target = resolve(sourcePath)
  const entry = entries.find((e) => resolve(e.sourcePath) === target)
  if (!entry) throw new Error('that skill was not offered by the scan')
  return entry.skillDir
    ? importForeignSkillDir(entry.skillDir, entry.name, userDataDir)
    : importForeignSkillFile(entry.sourcePath, userDataDir)
}

/**
 * Write a skill from the settings editor: one SKILL.md under the skills root,
 * built from the fields the pane collected.
 *
 * The name follows the import rule (`isSkillName`), plus the one shape that
 * rule lets through but a created folder cannot safely carry: a trailing dot
 * or space, which Win32 strips, so the folder on disk would stop matching the
 * name it was written under. Overwriting is opt-in — the editor loads the raw
 * skill first, so the user has seen what is being replaced — and it keeps the
 * folder's other files (a folder import can ship reference tables next to its
 * SKILL.md); only the SKILL.md is rewritten. Like every write here, the text
 * lands in a scratch dir and is renamed into place, so a failure leaves the
 * previous copy untouched and no `.write-` folder behind.
 */
export function writeSkill(
  userDataDir: string,
  name: string,
  doc: { description: string; body: string },
  opts?: { allowOverwrite?: boolean },
): FoundSkill {
  if (
    typeof name !== 'string' ||
    !isSkillName(name) ||
    /[. \t]$/.test(name) ||
    name !== name.trim()
  ) {
    throw new Error('the name cannot be used as a skill name')
  }
  if (
    typeof doc !== 'object' ||
    doc === null ||
    typeof doc.description !== 'string' ||
    typeof doc.body !== 'string'
  ) {
    throw new Error('a description and a body are required')
  }
  // the frontmatter is flat, so a description the pane let wrap becomes one line
  const description = doc.description.replace(/\s*\r?\n\s*/g, ' ').trim()
  if (!description) throw new Error('a description is required')
  const body = doc.body.replace(/\r\n/g, '\n')
  const text = `---\nname: ${name}\ndescription: ${description}\n---\n\n${body}`
  if (Buffer.byteLength(text) > MAX_IMPORT_BYTES) {
    throw new Error('the skill is larger than the 10 MB import cap')
  }
  const dest = join(skillsRoot(userDataDir), name)
  const overwriting = existsSync(dest)
  if (overwriting && !opts?.allowOverwrite) {
    throw new Error('a skill of that name is already there')
  }
  mkdirSync(skillsRoot(userDataDir), { recursive: true })
  const scratch = mkdtempSync(join(skillsRoot(userDataDir), '.write-'))
  // the previous copy steps aside rather than being deleted, so a failed swap puts it back
  const parked = `${scratch}.previous`
  let moved = false
  try {
    if (overwriting) {
      inspectSkillDir(dest)
      cpSync(dest, scratch, { recursive: true })
    }
    writeFileSync(join(scratch, 'SKILL.md'), text)
    if (overwriting) {
      renameSync(dest, parked)
      moved = true
    }
    renameSync(scratch, dest)
  } catch (e) {
    if (moved && !existsSync(dest)) renameSync(parked, dest)
    tryRm(scratch)
    throw e
  }
  // the swap is done; a locked leftover must not turn a finished save into a failure
  if (moved) tryRm(parked)
  return readBack(dest)
}

/**
 * The whole SKILL.md of a listed skill — frontmatter included — for the editor
 * to load. Like `removeSkill`, `name` may only select a folder the listing
 * itself reported, so the renderer can read back what it was shown and nothing
 * else: the name-shaped twin of the path guard `isKnownSkillPath`.
 */
export async function skillRawText(userDataDir: string, name: string): Promise<string> {
  if (typeof name !== 'string' || !isListedName(name)) throw new Error('unknown skill')
  const skill = (await usableSkills(userDataDir)).find((s) => s.name === name)
  if (!skill) throw new Error('unknown skill')
  return readFileSync(skill.path, 'utf8')
}

/**
 * Delete one of our skills. `name` must be a plain folder name that holds a
 * SKILL.md under the skills root, so the renderer can only remove what the
 * listing showed it.
 */
export function removeSkill(name: string, userDataDir: string): void {
  if (typeof name !== 'string' || !isListedName(name)) throw new Error('unknown skill')
  const dir = join(skillsRoot(userDataDir), name)
  if (!existsSync(join(dir, 'SKILL.md'))) throw new Error('unknown skill')
  rmSync(dir, { recursive: true, force: true })
}

/**
 * Whether a path is one of the SKILL.md files the listing can hand out.
 *
 * `readSkillBody` takes a path from the renderer, so on its own it reads any
 * file on disk. The name must be SKILL.md and its directory's directory must
 * be the skills root — the exact shape `usableSkills` builds, so what can be
 * read and what was listed are the same set. `dirname` never collapses `..`,
 * so an escaping path keeps its `..` and cannot match the resolved root.
 */
export function isKnownSkillPath(path: unknown, root: string): path is string {
  if (typeof path !== 'string' || path === '' || !isAbsolute(path)) return false
  const file = resolve(path)
  if (basename(file) !== 'SKILL.md') return false
  return resolve(root) === dirname(dirname(file))
}

/** `metadata:\n  version: x` in the frontmatter, the same shape skills/genoffice uses */
export function skillVersion(path: string): string {
  try {
    const fm = parseSkillFrontmatter(readFileSync(path, 'utf8'))
    return /^\s*version:\s*['"]?([^'"\n]+)/m.exec(fm?.extra.metadata ?? '')?.[1]?.trim() ?? ''
  } catch {
    return ''
  }
}

/**
 * The skills we ship as suggestions, with whether the user already has each.
 *
 * Shipping them is not installing them: a recommended skill only runs after
 * the user presses Install, which copies it into our folder like any import.
 * Installed copies are the user's — they may edit them — so an update is
 * offered, never applied on its own.
 */
export async function recommendedSkills(
  recommendedDir: string,
  userDataDir: string,
): Promise<RecommendedSkill[]> {
  let entries: string[]
  try {
    entries = await readdir(recommendedDir)
  } catch {
    return []
  }
  const out: RecommendedSkill[] = []
  for (const name of entries.sort()) {
    const path = join(recommendedDir, name, 'SKILL.md')
    let text: string
    try {
      text = await readFile(path, 'utf8')
    } catch {
      continue
    }
    const skill = describe(name, path, text)
    if (!skill) continue
    const category = parseSkillFrontmatter(text)?.extra.category?.trim() || 'other'
    const version = skillVersion(path)
    const installedAt = join(skillsRoot(userDataDir), name, 'SKILL.md')
    let state: RecommendedSkill['state'] = 'missing'
    if (existsSync(installedAt)) {
      state = skillVersion(installedAt) === version ? 'installed' : 'outdated'
    }
    out.push({ ...skill, category, version, state })
  }
  return out
}

/**
 * Copy one recommended skill into our folder. `name` is matched against the
 * recommended directory's own listing, so the renderer cannot name a path.
 * An existing copy is replaced only here, because the user asked for exactly
 * that by pressing Update.
 */
export function installRecommendedSkill(
  name: string,
  recommendedDir: string,
  userDataDir: string,
): FoundSkill {
  if (!/^[A-Za-z0-9][A-Za-z0-9._-]*$/.test(name)) throw new Error('unknown skill')
  const src = join(recommendedDir, name)
  if (!existsSync(join(src, 'SKILL.md'))) throw new Error('unknown skill')
  const dest = join(skillsRoot(userDataDir), name)
  requireSkillMd(src)
  inspectSkillDir(src)
  mkdirSync(skillsRoot(userDataDir), { recursive: true })
  const scratch = mkdtempSync(join(skillsRoot(userDataDir), '.install-'))
  // the user's copy steps aside rather than being deleted, so a failed swap puts it back
  const parked = `${scratch}.previous`
  let moved = false
  try {
    cpSync(src, scratch, { recursive: true })
    if (existsSync(dest)) {
      renameSync(dest, parked)
      moved = true
    }
    renameSync(scratch, dest)
  } catch (e) {
    if (moved && !existsSync(dest)) renameSync(parked, dest)
    tryRm(scratch)
    throw e
  }
  // the swap is done; a locked leftover must not turn a finished install into a failure
  if (moved) tryRm(parked)
  return readBack(dest)
}

function tryRm(path: string): void {
  try {
    rmSync(path, { recursive: true, force: true })
  } catch {
    /* EPERM/EBUSY under antivirus or a file lock; dot folders are skipped by usableSkills */
  }
}
