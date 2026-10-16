import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  PACK_KIND,
  PACK_MAX_FEATURES,
  PACK_VERSION,
  buildPack,
  describePack,
  installPack,
  readPack,
} from '../src/main/feature-packs'
import { createFeature, listFeatures } from '../src/main/feature-store'
import type { FeatureRecord } from '@genoffice/agent-core'

/**
 * Feature packs, which are untrusted input on the way in: the reader re-validates
 * every record, and nothing installs before the caller has seen what it declares.
 */

function dir(): string {
  return mkdtempSync(join(tmpdir(), 'genoffice-pack-'))
}

const timer: FeatureRecord = {
  id: 'tidy-headings',
  name: 'Tidy headings',
  description: 'unfold wrapped headings',
  kind: 'script',
  script: { source: '$.log("hi")' },
  createdAt: 'x',
  updatedAt: 'x',
}

const guard: FeatureRecord = {
  id: 'no-untitled',
  name: 'No untitled reports',
  description: 'stop me saving untitled reports',
  kind: 'hook',
  hook: { on: 'document.will-save', script: { source: '$.event.veto("no title")' } },
  createdAt: 'x',
  updatedAt: 'x',
}

describe('feature packs', () => {
  it('round-trips the records it was built from', () => {
    const pack = buildPack([timer, guard], new Date('2026-01-01T00:00:00.000Z'))
    expect(pack.kind).toBe(PACK_KIND)
    expect(pack.version).toBe(PACK_VERSION)
    const read = readPack(JSON.stringify(pack))
    expect(read.ok).toBe(true)
    if (read.ok) {
      expect(read.pack.features.map((feature) => feature.id)).toEqual([
        'tidy-headings',
        'no-untitled',
      ])
      expect(read.pack.features[1]!.hook?.on).toBe('document.will-save')
    }
  })

  it('refuses what is not a pack, and a pack from a newer build', () => {
    expect(readPack('not json').ok).toBe(false)
    expect(readPack('{"kind":"other"}').ok).toBe(false)
    const future = { kind: PACK_KIND, version: PACK_VERSION + 1, features: [timer] }
    const read = readPack(JSON.stringify(future))
    expect(read.ok).toBe(false)
    if (!read.ok) expect(read.errors[0]).toContain('newer GenOffice')
    expect(readPack(JSON.stringify({ kind: PACK_KIND, version: 1, features: [] })).ok).toBe(false)
  })

  it('re-validates every record: a malformed one is refused, not installed', () => {
    const bad = {
      kind: PACK_KIND,
      version: 1,
      exportedAt: '',
      features: [
        timer,
        {
          id: 'broken',
          name: 'Broken',
          description: 'x',
          kind: 'script',
          script: { source: 'export default 1' },
        },
      ],
    }
    const read = readPack(JSON.stringify(bad))
    expect(read.ok).toBe(false)
    if (!read.ok) expect(read.errors[0]).toContain('feature 2')

    const badSlot = {
      kind: PACK_KIND,
      version: 1,
      exportedAt: '',
      features: [
        {
          id: 's',
          name: 'S',
          description: 'x',
          kind: 'slot',
          slot: { at: 'nowhere', element: {} },
        },
      ],
    }
    expect(readPack(JSON.stringify(badSlot)).ok).toBe(false)
  })

  it('caps the pack size', () => {
    const many = {
      kind: PACK_KIND,
      version: 1,
      exportedAt: '',
      features: Array.from({ length: PACK_MAX_FEATURES + 1 }, (_, index) => ({
        ...timer,
        id: 'f' + index,
      })),
    }
    expect(readPack(JSON.stringify(many)).ok).toBe(false)
  })

  it('describes what a pack declares, and warns about the parts that matter', () => {
    const pack = buildPack([timer, guard], new Date('2026-01-01T00:00:00.000Z'))
    const described = describePack(pack, ['tidy-headings'])
    expect(described.count).toBe(2)
    expect(described.lines[0]!.declares).toEqual(['script, 11 characters'])
    expect(described.lines[1]!.declares).toEqual(['observes document.will-save'])
    expect(described.warnings).toContain('No untitled reports can stop a save')
    expect(described.warnings).toContain('Tidy headings runs code (read-only, sandboxed)')
    expect(described.warnings).toContain(
      'Tidy headings replaces a feature that is already installed',
    )
  })

  it('installs into an empty registry, and does not overwrite by default', () => {
    const target = dir()
    const pack = buildPack([timer, guard])
    const first = installPack(target, pack)
    expect(first.installed).toEqual(['tidy-headings', 'no-untitled'])
    expect(
      listFeatures(target)
        .map((feature) => feature.id)
        .sort(),
    ).toEqual(['no-untitled', 'tidy-headings'])

    // a second import leaves what is there alone
    const second = installPack(target, pack)
    expect(second.skipped).toEqual(['tidy-headings', 'no-untitled'])
    expect(second.installed).toEqual([])

    // and with the explicit option it replaces
    const third = installPack(target, pack, { replace: true })
    expect(third.replaced).toEqual(['tidy-headings', 'no-untitled'])
    expect(listFeatures(target)).toHaveLength(2)
  })

  it('keeps a feature installed by the user when a pack only replaces its own', () => {
    const target = dir()
    createFeature(target, { name: 'Mine', description: 'x', script: { source: '$.log(1)' } })
    const before = listFeatures(target).map((feature) => feature.id)
    installPack(target, buildPack([timer]))
    const after = listFeatures(target).map((feature) => feature.id)
    expect(after).toEqual([...before, 'tidy-headings'])
  })
})
