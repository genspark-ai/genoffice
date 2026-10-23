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
import { readFile, readdir } from 'node:fs/promises'
import { basename, dirname, isAbsolute, join, resolve, sep } from 'node:path'
import JSZip from 'jszip'
import { classifySkill, parseSkillFrontmatter } from '@genoffice/agent-core'
import type { FoundSkill, RecommendedSkill } from '../shared/found-skill'

/**
 * The skills a user installed into GenOffice.
 *
 * Only `<userData>/skills/<name>/SKILL.md` counts. Nothing is scanned from a
 * coding agent's directories: a developer's `.claude` holds git, test and
 * deploy skills, and a word processor offering those as candidates was noise
 * for developers and an empty panel for everyone else. A skill arrives here
 * one way — the user picks a folder or a zip and presses Import.
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
 * Import a single markdown file as a skill.
 *
 * The file becomes `<name>/SKILL.md`; the name comes from the frontmatter and
 * falls back to the file stem when `name` cannot be a folder name. Refuses
 * rather than overwrites when the name is taken.
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
  const fm = parseSkillFrontmatter(text)
  if (!fm) throw new Error('SKILL.md needs name and description frontmatter')
  const stem = basename(mdPath).replace(/\.md$/i, '')
  const name = [fm.name.trim(), stem].find((n) => isSkillName(n))
  if (!name) throw new Error('the file name cannot be used as a skill name')
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
