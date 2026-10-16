import type { FeatureRecord } from './features-skill'

/**
 * The knowledge kind's contract (K7): a source an assistant may consult.
 *
 * It is deliberately local and read-only: a folder of documents or one file, so
 * the assistant can quote something the user actually has. Fetching addresses, a
 * database, or anything that reaches outside the machine is not this kind — those
 * would need their own trust story, and a knowledge source must not become a way
 * for a feature to read whatever it likes: the host hands over one entry at a time,
 * from inside the declared source, with the same reader and the same bounds as
 * read_document.
 */

export const KNOWLEDGE_SOURCES = ['folder', 'file'] as const
export type KnowledgeSource = (typeof KNOWLEDGE_SOURCES)[number]

/**
 * How a source is indexed. 'full-text' (the default) extracts every contributed
 * file's text into the knowledge base's search index; 'names' lists the files
 * (they stay readable page by page and citable) but their content is not
 * searched — the honest choice for a huge folder. Whichever mode is set,
 * retrieval itself stays the knowledge base's own tools; a feature only decides
 * how its source is indexed.
 */
export const KNOWLEDGE_INDEX_MODES = ['full-text', 'names'] as const
export type KnowledgeIndexMode = (typeof KNOWLEDGE_INDEX_MODES)[number]

export interface FeatureKnowledge {
  source: KnowledgeSource
  /** absolute path to the folder or the file */
  path: string
  /** optional short label the assistant shows when it quotes */
  label?: string
  /**
   * which files of a folder source join the knowledge base, as name globs
   * (`*.md`, `report-*`, `?draft`): an entry must match one of these. Absent
   * means every readable file. (A file source has no list to filter.)
   */
  include?: string[]
  /** name globs left out of the corpus even when an include pattern matched */
  exclude?: string[]
  /** how the source is indexed; absent means full-text */
  index?: KnowledgeIndexMode
}

/** One thing inside a source, as the assistant sees it: a name, never a path. */
export interface KnowledgeEntry {
  name: string
  bytes: number
}

export interface KnowledgeValidationError {
  field: string
  message: string
}

export type KnowledgeValidation =
  { ok: true; knowledge: FeatureKnowledge } | { ok: false; errors: KnowledgeValidationError[] }

/** bounds of the per-source file rules a feature may declare */
const MAX_GLOBS = 12
const GLOB_MAX_CHARS = 120

/** Validate one glob list: a bounded array of short non-empty patterns. */
function validateGlobs(value: unknown, field: string, errors: KnowledgeValidationError[]): void {
  if (value === undefined) return
  if (!Array.isArray(value) || value.length === 0) {
    errors.push({ field, message: 'needs a non-empty array of name patterns' })
    return
  }
  if (value.length > MAX_GLOBS) {
    errors.push({ field, message: 'has more than ' + MAX_GLOBS + ' patterns' })
  }
  for (const pattern of value) {
    if (typeof pattern !== 'string' || pattern.trim().length === 0) {
      errors.push({ field, message: 'needs non-empty string patterns' })
      return
    }
    if (pattern.length > GLOB_MAX_CHARS) {
      errors.push({ field, message: 'has a pattern longer than ' + GLOB_MAX_CHARS + ' characters' })
      return
    }
    // a source is one folder: entries are names, never paths, so a separator
    // in a pattern can only mean the author expected path matching
    if (/[/\\]/.test(pattern)) {
      errors.push({
        field,
        message: 'matches file names (like "*.md"), not paths: drop the slashes',
      })
      return
    }
  }
}

/** Validate a knowledge payload from the assistant or a mod file. */
export function validateKnowledge(value: unknown): KnowledgeValidation {
  const errors: KnowledgeValidationError[] = []
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ok: false, errors: [{ field: 'knowledge', message: 'needs a knowledge object' }] }
  }
  const input = value as Record<string, unknown>
  const unknown = Object.keys(input).filter(
    (key) =>
      key !== 'source' &&
      key !== 'path' &&
      key !== 'label' &&
      key !== 'include' &&
      key !== 'exclude' &&
      key !== 'index',
  )
  for (const key of unknown) {
    errors.push({ field: 'knowledge.' + key, message: 'is not a knowledge field' })
  }
  if (
    typeof input.source !== 'string' ||
    !(KNOWLEDGE_SOURCES as readonly string[]).includes(input.source)
  ) {
    errors.push({
      field: 'knowledge.source',
      message: 'must be one of ' + KNOWLEDGE_SOURCES.join(', '),
    })
  }
  if (typeof input.path !== 'string' || input.path.trim().length === 0) {
    errors.push({ field: 'knowledge.path', message: 'needs the path to read from' })
  } else if (!isAbsolutePath(input.path)) {
    errors.push({
      field: 'knowledge.path',
      message: 'must be an absolute path (no ~, no relative paths, no addresses)',
    })
  }
  if (input.label !== undefined) {
    if (typeof input.label !== 'string' || input.label.trim().length === 0) {
      errors.push({ field: 'knowledge.label', message: 'needs a non-empty string' })
    } else if (input.label.length > 60) {
      errors.push({ field: 'knowledge.label', message: 'is longer than 60 characters' })
    }
  }
  validateGlobs(input.include, 'knowledge.include', errors)
  validateGlobs(input.exclude, 'knowledge.exclude', errors)
  let index: KnowledgeIndexMode | undefined
  if (input.index !== undefined) {
    if (
      typeof input.index !== 'string' ||
      !(KNOWLEDGE_INDEX_MODES as readonly string[]).includes(input.index)
    ) {
      errors.push({
        field: 'knowledge.index',
        message:
          'must be one of ' +
          KNOWLEDGE_INDEX_MODES.join(', ') +
          ' (got ' +
          String(input.index) +
          ')',
      })
    } else {
      index = input.index as KnowledgeIndexMode
    }
  }
  if (errors.length) return { ok: false, errors }
  const globsOf = (value: unknown): string[] | undefined =>
    Array.isArray(value) ? (value as string[]).map((pattern) => pattern.trim()) : undefined
  const include = globsOf(input.include)
  const exclude = globsOf(input.exclude)
  return {
    ok: true,
    knowledge: {
      source: input.source as KnowledgeSource,
      path: (input.path as string).trim(),
      ...(input.label === undefined ? {} : { label: input.label as string }),
      ...(include ? { include } : {}),
      ...(exclude ? { exclude } : {}),
      ...(index ? { index } : {}),
    },
  }
}

/** Absolute on both platforms the suite ships to, and never a scheme. */
function isAbsolutePath(path: string): boolean {
  const trimmed = path.trim()
  if (trimmed.startsWith('~')) return false
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(trimmed)) return false
  return trimmed.startsWith('/') || /^[a-zA-Z]:[\\/]/.test(trimmed)
}

/** The knowledge sources a registry holds. */
export function knowledgeIn(features: readonly FeatureRecord[]): FeatureRecord[] {
  return features.filter((feature) => feature.knowledge)
}

/**
 * One line per source, for the assistant's context (names, never contents).
 * The file rules ride along, so the assistant sees how a source is narrowed.
 */
export function describeKnowledge(features: readonly FeatureRecord[]): string[] {
  return knowledgeIn(features).map((feature) => {
    const knowledge = feature.knowledge!
    const bits = [knowledgeGlobSummary(knowledge)]
    if (knowledge.index === 'names') bits.push('names only')
    const rules = bits.filter(Boolean).join(', ')
    return (
      feature.id +
      ' | ' +
      (knowledge.label ?? feature.name) +
      ' | ' +
      knowledge.source +
      (rules ? ' (' + rules + ')' : '')
    )
  })
}

/** A short human summary of a source's include/exclude rules; empty when none. */
export function knowledgeGlobSummary(knowledge: FeatureKnowledge): string {
  const bits: string[] = []
  if (knowledge.include?.length) bits.push('only ' + knowledge.include.join(', '))
  if (knowledge.exclude?.length) bits.push('not ' + knowledge.exclude.join(', '))
  return bits.join(', ')
}

/**
 * Name-glob match for a source's file rules: `*` spans any run of characters,
 * `?` exactly one, case-insensitive (the hosts' name lookups are too). Names,
 * never paths — a folder source is one folder.
 */
export function knowledgeGlobMatches(pattern: string, name: string): boolean {
  const regex = new RegExp(
    '^' +
      [...pattern]
        .map((ch) => {
          if (ch === '*') return '.*'
          if (ch === '?') return '.'
          return ch.replace(/[.+^${}()|[\]\\]/g, '\\$&')
        })
        .join('') +
      '$',
    'i',
  )
  return regex.test(name)
}

/**
 * Apply a source's file rules to its entries: an entry joins the corpus when it
 * matches an include pattern (or the source declares none) and matches no
 * exclude pattern. Excluded names also stop resolving in pathForEntry, because
 * that guard checks against this listing.
 */
export function filterKnowledgeEntries<T extends { name: string }>(
  knowledge: FeatureKnowledge,
  entries: readonly T[],
): T[] {
  const include = knowledge.include ?? []
  const exclude = knowledge.exclude ?? []
  return entries.filter(
    (entry) =>
      (include.length === 0 ||
        include.some((pattern) => knowledgeGlobMatches(pattern, entry.name))) &&
      !exclude.some((pattern) => knowledgeGlobMatches(pattern, entry.name)),
  )
}
