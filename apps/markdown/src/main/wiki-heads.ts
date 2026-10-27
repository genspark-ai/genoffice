import { readFile } from 'node:fs/promises'
import { nodeFs, resolveWikiLinks } from './wiki-resolve'
import type { WikiHeadEntry } from '../shared/ipc'

/**
 * What `[[Note#` offers: the target note's headings and the block ids it
 * already carries. Read-only — the suggester lists what is there, it never
 * writes an id into another file.
 */

const HEADING_RE = /^(#{1,6})\s+(.+?)\s*#*\s*$/
const FENCE_RE = /^ {0,3}(`{3,}|~{3,})/
const TRAILING_ID_RE = /(?:^|\s)\^([A-Za-z0-9-]+)\s*$/
const LONE_ID_RE = /^\^([A-Za-z0-9-]+)\s*$/

/** each kind is capped on its own, so a heading-heavy file still shows its blocks */
export const MAX_HEAD_ENTRIES_PER_KIND = 200

/**
 * Headings and block ids of a note's raw text, in document order. Fenced code
 * and the frontmatter block are skipped; a marker alone on a line (the form a
 * table or code block takes) borrows the previous non-blank line as its
 * preview.
 */
export function parseWikiHeads(raw: string): WikiHeadEntry[] {
  const lines = raw.split('\n')
  const out: WikiHeadEntry[] = []
  let headings = 0
  let blocks = 0
  let fence: string | null = null
  // `---` frontmatter opens the file; skip through its closing fence
  let at = 0
  if (lines[0]?.trimEnd() === '---') {
    at = 1
    while (at < lines.length && lines[at].trimEnd() !== '---') at++
    at++
  }
  // the last non-blank line outside fences: a lone marker stands after a
  // blank line (that is Obsidian's layout) and still borrows its block
  let previous = ''
  for (; at < lines.length; at++) {
    const line = lines[at]!
    const open = FENCE_RE.exec(line)
    if (open) {
      if (fence === null) fence = open[1]![0]!
      // the closing fence is the same character, at least as long
      else if (line.trim().startsWith(fence)) fence = null
      previous = ''
      continue
    }
    if (fence !== null) continue
    if (line.trim() === '') continue
    const lone = LONE_ID_RE.exec(line)
    if (lone) {
      if (blocks < MAX_HEAD_ENTRIES_PER_KIND) {
        out.push({ kind: 'block', id: lone[1]!, text: previous })
        blocks++
      }
      previous = line
      continue
    }
    const heading = HEADING_RE.exec(line)
    if (heading) {
      // a heading line may carry its own marker; the heading is addressed
      // by its text, the marker is a block id on the heading (a heading is
      // a textblock, and both the resolver and the writer treat it so)
      const marker = TRAILING_ID_RE.exec(heading[2]!)
      const text = heading[2]!.replace(/\s*\^[A-Za-z0-9-]+\s*$/, '').trim()
      if (text) {
        if (headings < MAX_HEAD_ENTRIES_PER_KIND) {
          out.push({ kind: 'heading', text, level: heading[1]!.length })
          headings++
        }
        if (marker && blocks < MAX_HEAD_ENTRIES_PER_KIND) {
          out.push({
            kind: 'block',
            id: marker[1]!,
            // marker.index sits inside the heading text, not the line
            text: `${heading[1]} ${heading[2]!.slice(0, marker.index).trim()}`.trim(),
          })
          blocks++
        }
      }
      previous = line
      continue
    }
    const trailing = TRAILING_ID_RE.exec(line)
    if (trailing) {
      if (blocks < MAX_HEAD_ENTRIES_PER_KIND) {
        out.push({ kind: 'block', id: trailing[1]!, text: line.slice(0, trailing.index).trim() })
        blocks++
      }
    }
    previous = line
  }
  return out
}

/** resolve the typed target the way the resolver does, then parse that file */
export async function wikiHeads(
  target: string,
  fromPath: string,
  root: string,
): Promise<WikiHeadEntry[]> {
  if (!target) return []
  let hit: { found: boolean; path: string } | undefined
  try {
    hit = (await resolveWikiLinks([target], { fromPath, root }, nodeFs())).get(target)
  } catch {
    return []
  }
  if (!hit?.found) return []
  const raw = await readFile(hit.path, 'utf8').catch(() => null)
  return raw === null ? [] : parseWikiHeads(raw)
}
