/**
 * Resolving `[[target]]` to a file. Syntax (`wikilink.ts`) decides what a link
 * is; this decides what it points at.
 *
 * Obsidian's rule is shortest-path: the nearest note with that name wins, so
 * a link next to its target is unambiguous and only an ambiguous name has to
 * be disambiguated. Reproduced here as a breadth-first walk of the resolution
 * root, which visits a file before anything nested below it — the same
 * ordering a human reads paths in.
 *
 * Headings and block refs ride along in the target (`Note#Heading`,
 * `Note#^a1b2c3`); they select a position inside the note, not a different
 * file, so they are stripped before matching and the fragment is handed back
 * for the caller to apply after opening.
 */
import { existsSync, readdirSync, statSync } from 'node:fs'
import { basename, dirname, extname, isAbsolute, join } from 'node:path'

/** a filesystem the resolver walks; injectable so tests drive a real temp tree */
export interface WikiFs {
  exists(path: string): boolean
  isFile(path: string): boolean
  list(dir: string): { files: string[]; dirs: string[] }
}

/** directories never entered: version control, dependency trees, app sandboxes */
export const SKIP_DIRS = new Set([
  '.git',
  '.obsidian',
  '.trash',
  'node_modules',
  '.venv',
  'venv',
  '__pycache__',
  '.DS_Store',
])

/** note extensions a bare `[[Name]]` may stand for, in preference order */
const NOTE_EXTS = ['.md', '.markdown', '.txt']

/** how deep below the root a link may reach; a walk past this is not a note */
export const MAX_DEPTH = 6

/** entries under one walk step, so a huge library cannot stall the UI */
const MAX_ENTRIES = 20000

export function nodeFs(): WikiFs {
  return {
    exists: (path) => existsSync(path),
    isFile: (path) => {
      try {
        return statSync(path).isFile()
      } catch {
        return false
      }
    },
    list: (dir) => {
      const files: string[] = []
      const dirs: string[] = []
      for (const entry of readdirSync(dir, { withFileTypes: true })) {
        if (entry.isDirectory()) dirs.push(entry.name)
        else if (entry.isFile()) files.push(entry.name)
      }
      return { files, dirs }
    },
  }
}

export interface SplitTarget {
  /** the file part, '' for a same-file `[[#Heading]]` */
  name: string
  /** the `#Heading` / `#^block` fragment, '' when there is none */
  fragment: string
}

export function splitTarget(raw: string): SplitTarget {
  const hash = raw.indexOf('#')
  if (hash < 0) return { name: raw.trim(), fragment: '' }
  return { name: raw.slice(0, hash).trim(), fragment: raw.slice(hash + 1).trim() }
}

export interface ResolvedLink {
  /** absolute path, or '' when the target is this document */
  path: string
  fragment: string
  /** true when a file was found */
  found: boolean
}

/**
 * Whether a target may be joined onto a directory at all.
 *
 * A link is written by whoever shares the note and opened on the reader's
 * machine, so a target that climbs up with `..` — or names an absolute path —
 * would reach any file they have, at any extension, by relative path. The
 * walk only ever joins a name it read out of a directory listing onto that
 * same tree, which is why the check belongs here, on the one place a caller
 * supplied name is turned into a path.
 */
function escapesVault(name: string): boolean {
  if (name === '') return true
  if (isAbsolute(name)) return true
  return name.split(/[\\/]/).some((segment) => segment === '..')
}

/** the candidates a bare name can stand for, in the order they are tried */
function candidatePaths(dir: string, name: string): string[] {
  if (extname(name)) return [join(dir, name)]
  return NOTE_EXTS.map((ext) => join(dir, `${name}${ext}`))
}

function firstExisting(fs: WikiFs, dir: string, name: string): string | null {
  for (const candidate of candidatePaths(dir, name)) {
    if (fs.exists(candidate) && fs.isFile(candidate)) return candidate
  }
  return null
}

export interface ResolveOptions {
  /** absolute path of the document holding the links; '' for an untitled one */
  fromPath: string
  /** the root a bare name may reach beyond the document's own directory */
  root: string
}

/**
 * Every link in one walk. Per-link resolution would stat the root once per
 * link; batching lets a document with fifty links pay for one traversal, and
 * hands the caller the answer for every link so the unresolved ones can be
 * marked without a second look.
 */
export function resolveWikiLinks(
  targets: readonly string[],
  options: ResolveOptions,
  fs: WikiFs = nodeFs(),
): Map<string, ResolvedLink> {
  const out = new Map<string, ResolvedLink>()
  if (targets.length === 0) return out

  const pending: { raw: string; name: string; fragment: string }[] = []
  for (const raw of targets) {
    if (out.has(raw)) continue
    const { name, fragment } = splitTarget(raw)
    out.set(raw, { path: '', fragment, found: false })
    // `[[#Heading]]` addresses this document; it is never a lookup
    if (name === '') {
      out.set(raw, { path: options.fromPath, fragment, found: options.fromPath !== '' })
      continue
    }
    // refused before it reaches the filesystem, so it reports unresolved
    if (escapesVault(name)) continue
    pending.push({ raw, name, fragment })
  }
  if (pending.length === 0) return out

  // 1. the document's own directory — a link beside its target needs nothing else
  const remaining: typeof pending = []
  if (options.fromPath) {
    const here = dirname(options.fromPath)
    for (const item of pending) {
      const hit = firstExisting(fs, here, item.name)
      if (hit) out.set(item.raw, { path: hit, fragment: item.fragment, found: true })
      else remaining.push(item)
    }
  } else remaining.push(...pending)

  // 2. breadth-first from the root: a shallower file beats a deeper one of the
  //    same name, which is what makes an unqualified name unambiguous
  if (remaining.length > 0 && options.root) {
    const wanted = new Map<string, typeof remaining>()
    for (const item of remaining) {
      const list = wanted.get(item.name)
      if (list) list.push(item)
      else wanted.set(item.name, [item])
    }
    if (wanted.size > 0) {
      let queue: { dir: string; depth: number }[] = [{ dir: options.root, depth: 0 }]
      let budget = MAX_ENTRIES
      while (queue.length > 0 && wanted.size > 0 && budget > 0) {
        const next: typeof queue = []
        for (const { dir, depth } of queue) {
          let listing: { files: string[]; dirs: string[] }
          try {
            listing = fs.list(dir)
          } catch {
            continue // unreadable directory, same as an unplugged root
          }
          budget -= listing.files.length + listing.dirs.length
          for (const file of listing.files) {
            // indexed under both spellings: a bare `[[Name]]` matches the stem,
            // an explicit `[[Figure 1.png]]` has to match the file as written
            const names = new Set([file, basename(file, extname(file))])
            let taken: string | null = null
            for (const key of names) {
              const items = wanted.get(key)
              if (!items) continue
              for (const item of items) {
                if (out.get(item.raw)?.found) continue
                out.set(item.raw, { path: join(dir, file), fragment: item.fragment, found: true })
                taken = taken ?? key
              }
            }
            if (taken === null) continue
            for (const key of names) if (key === taken) wanted.delete(key)
            if (wanted.size === 0) break
          }
          if (wanted.size === 0 || budget <= 0 || depth + 1 > MAX_DEPTH) continue
          for (const sub of listing.dirs) {
            if (SKIP_DIRS.has(sub)) continue
            next.push({ dir: join(dir, sub), depth: depth + 1 })
          }
        }
        queue = next
      }
    }
  }

  return out
}
