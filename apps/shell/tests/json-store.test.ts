import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { createJsonStore, type JsonStore } from '../src/main/json-store'

/**
 * The shared persistence core every one-file userData store goes through
 * (desktop-ipc.ts, todos-ipc.ts). What is pinned here is the contract the
 * four stores rely on for behavior-zero-regression: a version-1 store's file
 * stays byte-shaped like the pre-factory file (no wrapper, no marker, the
 * legacy payload round-trips), a corrupt or missing file reads as the
 * default, normalize runs on both halves, and the version machinery works
 * the way the migration convention in json-store.ts documents it.
 */

let dir = ''

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-jsonstore-'))
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

interface Item {
  name: string
}

/** a v1 store in the todos shape: bare-array body, junk reads as the default */
const plainStore = (file: string): JsonStore<Item[]> =>
  createJsonStore<Item[]>({
    path: () => join(dir, file),
    version: 1,
    normalize: (raw) =>
      Array.isArray(raw)
        ? raw
            .filter(
              (entry): entry is Item => !!entry && typeof entry === 'object' && 'name' in entry,
            )
            .slice(0, 2)
        : [],
  })

/** the file's parsed content, as the next launch would read it */
const readBody = (file: string): unknown => JSON.parse(readFileSync(join(dir, file), 'utf8'))

describe('a version-1 store keeps the legacy file shape', () => {
  it('writes the payload pretty-printed with no marker or wrapper', async () => {
    await plainStore('plain.json').write([{ name: 'a' }, { name: 'b' }])
    expect(readBody('plain.json')).toEqual([{ name: 'a' }, { name: 'b' }])
    expect(readFileSync(join(dir, 'plain.json'), 'utf8')).toContain('\n  ')
  })

  it('round-trips a file written before the factory existed', () => {
    writeFileSync(join(dir, 'plain.json'), JSON.stringify([{ name: 'legacy' }]), 'utf8')
    expect(plainStore('plain.json').read()).toEqual([{ name: 'legacy' }])
  })

  it('publishes through a scratch file and leaves none behind', async () => {
    await plainStore('plain.json').write([{ name: 'a' }])
    expect(readdirSync(dir)).toEqual(['plain.json'])
  })
})

describe('the read is fault-tolerant', () => {
  it('answers the default for a missing file', () => {
    expect(plainStore('absent.json').read()).toEqual([])
  })

  it('answers the default for a corrupt or foreign file', () => {
    for (const garbage of ['{oops', 'null', '42', '"nope"', '[]]']) {
      writeFileSync(join(dir, 'plain.json'), garbage, 'utf8')
      expect(plainStore('plain.json').read(), `${garbage} should read empty`).toEqual([])
    }
  })
})

describe('normalize runs on both halves', () => {
  it('drops junk entries and clamps on read', () => {
    writeFileSync(
      join(dir, 'plain.json'),
      JSON.stringify([{ name: 'a' }, 'junk', { name: 'b' }, { name: 'c' }, null]),
      'utf8',
    )
    expect(plainStore('plain.json').read()).toEqual([{ name: 'a' }, { name: 'b' }])
  })

  it('normalizes the value before it hits the disk', async () => {
    await plainStore('plain.json').write([
      { name: 'a' },
      'junk' as unknown as Item,
      { name: 'b' },
      { name: 'c' },
    ])
    expect(readBody('plain.json')).toEqual([{ name: 'a' }, { name: 'b' }])
  })
})

describe('the envelope hooks', () => {
  /** a v1 store in the pins shape: `{ items }` body, bare-list domain value */
  const envelopeStore = (file: string): JsonStore<Item[]> =>
    createJsonStore<Item[]>({
      path: () => join(dir, file),
      version: 1,
      normalize: (raw) => (Array.isArray(raw) ? (raw as Item[]).slice(0, 2) : []),
      unwrap: (raw) => (raw as { items?: unknown } | null)?.items,
      wrap: (items) => ({ items }),
    })

  it('round-trips through the wrap/unwrap pair, marker-less', async () => {
    await envelopeStore('env.json').write([{ name: 'a' }])
    expect(readBody('env.json')).toEqual({ items: [{ name: 'a' }] })
    expect(envelopeStore('env.json').read()).toEqual([{ name: 'a' }])
  })

  it('a shapeless or unwrappable body is the default', () => {
    for (const garbage of ['{oops', 'null', '42', '{"items":"nope"}', '{}']) {
      writeFileSync(join(dir, 'env.json'), garbage, 'utf8')
      expect(envelopeStore('env.json').read(), `${garbage} should read empty`).toEqual([])
    }
  })
})

describe('the version machinery', () => {
  /** the migrated shape: v2 moves the list under `items` and caps it at one */
  const v2Store = (file: string): JsonStore<Item[]> =>
    createJsonStore<Item[]>({
      path: () => join(dir, file),
      version: 2,
      normalize: (raw) => (Array.isArray(raw) ? (raw as Item[]).slice(0, 1) : []),
      unwrap: (raw) => (raw as { items?: unknown } | null)?.items,
      wrap: (items) => ({ items }),
      migrate: (raw, from) => {
        if (from !== 1) throw new Error(`unexpected migration from version ${from}`)
        return { items: raw }
      },
    })

  it('stamps the marker into a version-2 write and reads it back without migrating', async () => {
    const store = v2Store('v2.json')
    await store.write([{ name: 'a' }])
    expect(readBody('v2.json')).toEqual({ items: [{ name: 'a' }], storeVersion: 2 })
    expect(store.read()).toEqual([{ name: 'a' }])
  })

  it('walks a marker-less legacy file through migrate, then normalizes', () => {
    // the version-1 body of this shape is the bare array (see the convention)
    writeFileSync(join(dir, 'v2.json'), JSON.stringify([{ name: 'a' }, { name: 'b' }]), 'utf8')
    const migrated = v2Store('v2.json').read()
    // the migrated list is capped by v2's own normalize, proving both ran
    expect(migrated).toEqual([{ name: 'a' }])
  })

  it('a junk marker is ignored, and the body reads as version 1', () => {
    // storeVersion: 'two' is not an integer marker, so the whole object is
    // the version-1 body; unwrap still finds its `items` field
    writeFileSync(join(dir, 'v2.json'), JSON.stringify({ items: [], storeVersion: 'two' }), 'utf8')
    expect(v2Store('v2.json').read()).toEqual([])
  })

  it('runs the chain one step at a time up to the current version', () => {
    const steps: number[] = []
    const store = createJsonStore<Item[]>({
      path: () => join(dir, 'v3.json'),
      version: 3,
      normalize: (raw) => (Array.isArray(raw) ? (raw as Item[]).slice(0, 1) : []),
      unwrap: (raw) => (raw as { items?: unknown } | null)?.items,
      wrap: (items) => ({ items }),
      migrate: (raw, from) => {
        steps.push(from)
        return from === 1 ? { items: raw } : { ...raw, migratedAt: from }
      },
    })
    writeFileSync(join(dir, 'v3.json'), JSON.stringify([{ name: 'a' }]), 'utf8')
    expect(store.read()).toEqual([{ name: 'a' }])
    expect(steps).toEqual([1, 2])
  })

  it('a body stamped with a newer version than the store reads through normalize', () => {
    writeFileSync(
      join(dir, 'v2.json'),
      JSON.stringify({ items: [{ name: 'from-the-future' }], storeVersion: 9 }),
      'utf8',
    )
    expect(v2Store('v2.json').read()).toEqual([{ name: 'from-the-future' }])
  })
})
