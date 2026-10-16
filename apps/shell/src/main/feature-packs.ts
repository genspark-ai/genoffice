import {
  validateAppearance,
  validateCommand,
  validateHook,
  validateKnowledge,
  validateScript,
  validateSlot,
  type FeatureRecord,
} from '@genoffice/agent-core'
import type { PackDescription, PackLine } from '../shared/feature-channels'
import { getFeature, putFeature, updateFeature } from './feature-store'

/**
 * Feature packs: what a mod is when it leaves this machine.
 *
 * A pack is JSON, and it is treated as untrusted input on the way back in — every
 * record is re-validated with the same validators the registry uses, so a pack
 * cannot smuggle a malformed slot, an invented appearance or a module-shaped script
 * past the boundary. Nothing is installed before the reader has been told what the
 * pack declares, and the description names the things worth a second look (a hook
 * that can stop a save, a knowledge source, code that runs).
 */

export const PACK_KIND = 'genoffice-features'
export const PACK_VERSION = 1
export const PACK_MAX_FEATURES = 50

export interface FeaturePack {
  kind: typeof PACK_KIND
  version: number
  exportedAt: string
  features: FeatureRecord[]
}

export type { PackDescription, PackLine }

/** Write a pack out of the records the user picked. */
export function buildPack(features: readonly FeatureRecord[], now = new Date()): FeaturePack {
  return {
    kind: PACK_KIND,
    version: PACK_VERSION,
    exportedAt: now.toISOString(),
    features: features.slice(0, PACK_MAX_FEATURES).map((feature) => ({ ...feature })),
  }
}

/** What one record declares, in the reader's words rather than ours. */
export function declaresOf(feature: FeatureRecord): string[] {
  const lines: string[] = []
  if (feature.slot) lines.push('draws at ' + feature.slot.at)
  if (feature.appearance) {
    const bits = [
      feature.appearance.background ? 'background ' + feature.appearance.background : '',
      feature.appearance.accent ? 'accent ' + feature.appearance.accent : '',
      feature.appearance.markdown
        ? 'markdown theme ' +
          (typeof feature.appearance.markdown === 'string'
            ? feature.appearance.markdown
            : 'custom (' + feature.appearance.markdown.paper + ')')
        : '',
    ].filter(Boolean)
    if (bits.length) lines.push('chrome: ' + bits.join(', '))
  }
  if (feature.script) {
    lines.push(
      'script' +
        (feature.script.format ? ' (' + feature.script.format + ')' : '') +
        ', ' +
        feature.script.source.length +
        ' characters',
    )
  }
  if (feature.hook) lines.push('observes ' + feature.hook.on)
  if (feature.knowledge) {
    lines.push('knowledge source: ' + feature.knowledge.source + ' at ' + feature.knowledge.path)
  }
  if (feature.command) lines.push('command: ' + feature.command.label)
  return lines
}

/** Describe a pack before anything is installed. */
export function describePack(
  pack: FeaturePack,
  existingIds: readonly string[] = [],
): PackDescription {
  const lines: PackLine[] = pack.features.map((feature) => ({
    id: feature.id,
    name: feature.name,
    kind: feature.kind || 'recorded',
    declares: declaresOf(feature),
  }))
  const warnings: string[] = []
  for (const feature of pack.features) {
    if (feature.hook?.on === 'document.will-save') {
      warnings.push(feature.name + ' can stop a save')
    }
    if (feature.script) warnings.push(feature.name + ' runs code (read-only, sandboxed)')
    if (feature.knowledge) {
      warnings.push(feature.name + ' reads from ' + feature.knowledge.path)
    }
    if (existingIds.includes(feature.id)) {
      warnings.push(feature.name + ' replaces a feature that is already installed')
    }
  }
  return {
    version: pack.version,
    exportedAt: pack.exportedAt,
    count: pack.features.length,
    lines,
    warnings,
  }
}

/** Read a pack: shape, version, and every record through the real validators. */
export function readPack(
  text: string,
): { ok: true; pack: FeaturePack } | { ok: false; errors: string[] } {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    return { ok: false, errors: ['this file is not a feature pack (it is not JSON)'] }
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return { ok: false, errors: ['this file is not a feature pack'] }
  }
  const pack = parsed as Record<string, unknown>
  if (pack.kind !== PACK_KIND) {
    return { ok: false, errors: ['this file is not a feature pack (wrong kind)'] }
  }
  const version = typeof pack.version === 'number' ? pack.version : 0
  if (version > PACK_VERSION) {
    return {
      ok: false,
      errors: ['this pack was written by a newer GenOffice (version ' + version + ')'],
    }
  }
  if (!Array.isArray(pack.features) || pack.features.length === 0) {
    return { ok: false, errors: ['this pack has no features'] }
  }
  if (pack.features.length > PACK_MAX_FEATURES) {
    return { ok: false, errors: ['this pack has more than ' + PACK_MAX_FEATURES + ' features'] }
  }
  const errors: string[] = []
  const features: FeatureRecord[] = []
  pack.features.forEach((raw, index) => {
    const where = 'feature ' + (index + 1) + ': '
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
      errors.push(where + 'is not an object')
      return
    }
    const record = raw as Record<string, unknown>
    for (const key of ['id', 'name', 'description', 'kind']) {
      if (typeof record[key] !== 'string') errors.push(where + key + ' is missing')
    }
    if (record.slot !== undefined && !validateSlot(record.slot).ok) {
      errors.push(where + 'the slot is not valid')
    }
    if (record.appearance !== undefined && !validateAppearance(record.appearance).ok) {
      errors.push(where + 'the appearance is not valid')
    }
    if (record.script !== undefined && !validateScript(record.script).ok) {
      errors.push(where + 'the script is not valid')
    }
    if (record.hook !== undefined && !validateHook(record.hook).ok) {
      errors.push(where + 'the hook is not valid')
    }
    if (record.knowledge !== undefined && !validateKnowledge(record.knowledge).ok) {
      errors.push(where + 'the knowledge source is not valid')
    }
    if (record.command !== undefined && !validateCommand(record.command).ok) {
      errors.push(where + 'the command is not valid')
    }
    features.push({
      id: String(record.id ?? ''),
      name: String(record.name ?? ''),
      description: String(record.description ?? ''),
      kind: String(record.kind ?? ''),
      ...(record.slot === undefined ? {} : { slot: record.slot as FeatureRecord['slot'] }),
      ...(record.appearance === undefined
        ? {}
        : { appearance: record.appearance as FeatureRecord['appearance'] }),
      ...(record.script === undefined ? {} : { script: record.script as FeatureRecord['script'] }),
      ...(record.hook === undefined ? {} : { hook: record.hook as FeatureRecord['hook'] }),
      ...(record.knowledge === undefined
        ? {}
        : { knowledge: record.knowledge as FeatureRecord['knowledge'] }),
      ...(record.command === undefined
        ? {}
        : { command: record.command as FeatureRecord['command'] }),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    })
  })
  if (errors.length) return { ok: false, errors }
  return {
    ok: true,
    pack: {
      kind: PACK_KIND,
      version,
      exportedAt: typeof pack.exportedAt === 'string' ? pack.exportedAt : '',
      features,
    },
  }
}

/**
 * Install a pack. A feature whose id is already taken is **not** touched unless the
 * caller asked to replace it: an import must not quietly overwrite what the user has.
 */
export function installPack(
  dir: string,
  pack: FeaturePack,
  options: { replace?: boolean } = {},
): { installed: string[]; replaced: string[]; skipped: string[] } {
  const installed: string[] = []
  const replaced: string[] = []
  const skipped: string[] = []
  for (const feature of pack.features) {
    const existing = getFeature(dir, feature.id)
    try {
      if (existing && !options.replace) {
        skipped.push(feature.id)
        continue
      }
      if (existing) {
        updateFeature(dir, feature.id, {
          name: feature.name,
          description: feature.description,
          kind: feature.kind,
          slot: feature.slot ?? null,
          appearance: feature.appearance ?? null,
          script: feature.script ?? null,
          hook: feature.hook ?? null,
          knowledge: feature.knowledge ?? null,
          command: feature.command ?? null,
        })
        replaced.push(feature.id)
        continue
      }
      // the pack's own id is kept: identity has to survive the trip, so a second
      // import of the same pack addresses the same features
      putFeature(dir, { ...feature, createdAt: new Date().toISOString() })
      installed.push(feature.id)
    } catch {
      skipped.push(feature.id)
    }
  }
  return { installed, replaced, skipped }
}
