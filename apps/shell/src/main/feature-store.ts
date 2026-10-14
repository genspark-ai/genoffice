import { randomUUID } from 'node:crypto'
import {
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { join } from 'node:path'
import {
  validateAppearance,
  validateCommand,
  validateHook,
  validateKnowledge,
  validateScript,
  validateSlot,
  type FeatureDraft,
  type FeaturePatch,
  type FeatureRecord,
} from '@genoffice/agent-core'

/**
 * Feature records on disk: one JSON file per feature under a directory the caller
 * owns. Pure functions over an injected directory, so the registry is testable
 * without electron or userData.
 *
 * A feature is a record, not an implementation: no kind executes yet (see
 * docs/mods-pr1-mod-interface.md). Writes are atomic (temp file beside the
 * destination, then rename), and a file that does not parse is skipped rather
 * than failing the whole listing.
 */

/** A registry is a person's own list, not a package index. */
export const FEATURE_MAX_COUNT = 200
const NAME_MAX = 200
const DESCRIPTION_MAX = 2000
const KIND_MAX = 60
/** ids also name files, so they are restricted to a path-safe shape */
const ID_PATTERN = /^[a-z0-9][a-z0-9-]{0,79}$/

export function isFeatureId(id: string): boolean {
  return ID_PATTERN.test(id)
}

function slack(value: unknown, max: number): string {
  return typeof value === 'string' ? value.trim().slice(0, max) : ''
}

function slug(name: string): string {
  const base = name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60)
  return base || 'feature'
}

function isRecord(value: unknown): value is FeatureRecord {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const r = value as Record<string, unknown>
  const base =
    typeof r.id === 'string' &&
    isFeatureId(r.id) &&
    typeof r.name === 'string' &&
    typeof r.description === 'string' &&
    typeof r.kind === 'string' &&
    typeof r.createdAt === 'string' &&
    typeof r.updatedAt === 'string'
  if (!base) return false
  // a record that carries a slot or an appearance must carry a valid one: a
  // malformed tree would otherwise reach the renderer and be drawn half-way
  return (
    (r.slot === undefined || validateSlot(r.slot).ok) &&
    (r.appearance === undefined || validateAppearance(r.appearance).ok) &&
    (r.script === undefined || validateScript(r.script).ok) &&
    (r.hook === undefined || validateHook(r.hook).ok) &&
    (r.knowledge === undefined || validateKnowledge(r.knowledge).ok) &&
    (r.command === undefined || validateCommand(r.command).ok)
  )
}

function featurePath(dir: string, id: string): string {
  if (!isFeatureId(id)) throw new Error('invalid feature id: ' + id)
  return join(dir, id + '.json')
}

/**
 * Write a record under the id it already has. Imports use this: a pack's identity
 * has to survive a trip, so the second import of the same pack replaces the same
 * feature instead of arriving as "-2".
 */
export function putFeature(dir: string, record: FeatureRecord, now = new Date()): FeatureRecord {
  if (!isRecord(record)) {
    // the parameter is typed, but a record read from a pack is still input
    throw new Error('invalid feature record')
  }
  if (!listFeatures(dir).some((feature) => feature.id === record.id)) {
    const existing = listFeatures(dir).length
    if (existing >= FEATURE_MAX_COUNT) {
      throw new Error('the feature list is full (' + FEATURE_MAX_COUNT + ')')
    }
  }
  const stamp = now.toISOString()
  const next: FeatureRecord = {
    ...record,
    name: slack(record.name, NAME_MAX),
    description: slack(record.description, DESCRIPTION_MAX),
    kind: slack(record.kind, KIND_MAX),
    updatedAt: stamp,
  }
  writeRecord(dir, next)
  return next
}

function writeRecord(dir: string, record: FeatureRecord): void {
  mkdirSync(dir, { recursive: true })
  const path = featurePath(dir, record.id)
  // the temp file lives beside the destination so rename never crosses filesystems
  const temp = path + '.' + process.pid + '.' + randomUUID() + '.tmp'
  try {
    writeFileSync(temp, JSON.stringify(record, null, 2), {
      encoding: 'utf8',
      flag: 'wx',
      flush: true,
    })
    renameSync(temp, path)
  } catch (error) {
    try {
      unlinkSync(temp)
    } catch {
      // the write may have failed before the temp file existed
    }
    throw error
  }
}

/** Every readable feature, oldest first; malformed files are skipped. */
export function listFeatures(dir: string): FeatureRecord[] {
  let names: string[]
  try {
    names = readdirSync(dir)
  } catch {
    return []
  }
  const out: FeatureRecord[] = []
  for (const name of names) {
    if (!name.endsWith('.json')) continue
    try {
      const parsed: unknown = JSON.parse(readFileSync(join(dir, name), 'utf8'))
      if (isRecord(parsed)) out.push(parsed)
    } catch {
      // a corrupt or half-written record does not break the rest of the registry
    }
  }
  return out.sort((a, b) =>
    a.createdAt === b.createdAt ? a.id.localeCompare(b.id) : a.createdAt.localeCompare(b.createdAt),
  )
}

export function getFeature(dir: string, id: string): FeatureRecord | null {
  if (!isFeatureId(id)) return null
  try {
    const parsed: unknown = JSON.parse(readFileSync(featurePath(dir, id), 'utf8'))
    return isRecord(parsed) ? parsed : null
  } catch {
    return null
  }
}

/** Record a new feature; the id is derived from the name and kept unique. */
export function createFeature(dir: string, draft: FeatureDraft, now = new Date()): FeatureRecord {
  const name = slack(draft.name, NAME_MAX)
  const description = slack(draft.description, DESCRIPTION_MAX)
  if (!name) throw new Error('a feature needs a name')
  if (!description) throw new Error('a feature needs a description')
  const existing = new Set(listFeatures(dir).map((f) => f.id))
  if (existing.size >= FEATURE_MAX_COUNT) {
    throw new Error('the feature list is full (' + FEATURE_MAX_COUNT + ')')
  }
  const base = slug(name)
  let id = base
  for (let n = 2; existing.has(id); n++) id = base + '-' + n
  const stamp = now.toISOString()
  if (draft.slot !== undefined && !validateSlot(draft.slot).ok) {
    throw new Error('invalid slot')
  }
  if (draft.appearance !== undefined && !validateAppearance(draft.appearance).ok) {
    throw new Error('invalid appearance')
  }
  if (draft.script !== undefined && !validateScript(draft.script).ok) {
    throw new Error('invalid script')
  }
  if (draft.hook !== undefined && !validateHook(draft.hook).ok) {
    throw new Error('invalid hook')
  }
  if (draft.knowledge !== undefined && !validateKnowledge(draft.knowledge).ok) {
    throw new Error('invalid knowledge')
  }
  if (draft.command !== undefined && !validateCommand(draft.command).ok) {
    throw new Error('invalid command')
  }
  const record: FeatureRecord = {
    id,
    name,
    description,
    kind:
      slack(draft.kind, KIND_MAX) ||
      (draft.slot
        ? 'slot'
        : draft.appearance
          ? 'appearance'
          : draft.script
            ? 'script'
            : draft.hook
              ? 'hook'
              : draft.knowledge
                ? 'knowledge'
                : draft.command
                  ? 'command'
                  : ''),
    ...(draft.slot === undefined ? {} : { slot: draft.slot }),
    ...(draft.appearance === undefined ? {} : { appearance: draft.appearance }),
    ...(draft.script === undefined ? {} : { script: draft.script }),
    ...(draft.hook === undefined ? {} : { hook: draft.hook }),
    ...(draft.knowledge === undefined ? {} : { knowledge: draft.knowledge }),
    ...(draft.command === undefined ? {} : { command: draft.command }),
    createdAt: stamp,
    updatedAt: stamp,
  }
  writeRecord(dir, record)
  return record
}

/** Change the fields given; null when no such feature exists. */
export function updateFeature(
  dir: string,
  id: string,
  patch: FeaturePatch,
  now = new Date(),
): FeatureRecord | null {
  const current = getFeature(dir, id)
  if (!current) return null
  if (patch.slot !== undefined && patch.slot !== null && !validateSlot(patch.slot).ok) {
    throw new Error('invalid slot')
  }
  if (
    patch.appearance !== undefined &&
    patch.appearance !== null &&
    !validateAppearance(patch.appearance).ok
  ) {
    throw new Error('invalid appearance')
  }
  if (patch.script !== undefined && patch.script !== null && !validateScript(patch.script).ok) {
    throw new Error('invalid script')
  }
  if (patch.hook !== undefined && patch.hook !== null && !validateHook(patch.hook).ok) {
    throw new Error('invalid hook')
  }
  if (
    patch.knowledge !== undefined &&
    patch.knowledge !== null &&
    !validateKnowledge(patch.knowledge).ok
  ) {
    throw new Error('invalid knowledge')
  }
  if (patch.command !== undefined && patch.command !== null && !validateCommand(patch.command).ok) {
    throw new Error('invalid command')
  }
  const next: FeatureRecord = {
    ...current,
    ...(patch.name === undefined ? {} : { name: slack(patch.name, NAME_MAX) || current.name }),
    ...(patch.description === undefined
      ? {}
      : { description: slack(patch.description, DESCRIPTION_MAX) || current.description }),
    ...(patch.kind === undefined ? {} : { kind: slack(patch.kind, KIND_MAX) }),
    updatedAt: now.toISOString(),
  }
  if (patch.slot === null) delete next.slot
  else if (patch.slot !== undefined) next.slot = patch.slot
  if (patch.appearance === null) delete next.appearance
  else if (patch.appearance !== undefined) next.appearance = patch.appearance
  if (patch.script === null) delete next.script
  else if (patch.script !== undefined) next.script = patch.script
  if (patch.hook === null) delete next.hook
  else if (patch.hook !== undefined) next.hook = patch.hook
  if (patch.knowledge === null) delete next.knowledge
  else if (patch.knowledge !== undefined) next.knowledge = patch.knowledge
  if (patch.command === null) delete next.command
  else if (patch.command !== undefined) next.command = patch.command
  writeRecord(dir, next)
  return next
}

/** Delete the record; false when there was nothing to delete. */
export function removeFeature(dir: string, id: string): boolean {
  if (!isFeatureId(id)) return false
  try {
    rmSync(featurePath(dir, id))
    return true
  } catch {
    return false
  }
}
