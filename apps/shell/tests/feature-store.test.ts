import { mkdtempSync, readdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  FEATURE_MAX_COUNT,
  createFeature,
  getFeature,
  isFeatureId,
  listFeatures,
  removeFeature,
  updateFeature,
} from '../src/main/feature-store'

function dir(): string {
  return mkdtempSync(join(tmpdir(), 'genoffice-features-'))
}

const at = (iso: string): Date => new Date(iso)

describe('feature store', () => {
  it('starts empty and records a feature with a slug id', () => {
    const d = dir()
    expect(listFeatures(d)).toEqual([])
    const created = createFeature(
      d,
      { name: 'Weekly Report Formatting', description: 'house style' },
      at('2026-01-01T00:00:00.000Z'),
    )
    expect(created.id).toBe('weekly-report-formatting')
    expect(created.kind).toBe('')
    expect(listFeatures(d)).toHaveLength(1)
    expect(getFeature(d, created.id)).toMatchObject({ name: 'Weekly Report Formatting' })
  })

  it('keeps ids unique and lists oldest first', () => {
    const d = dir()
    const first = createFeature(
      d,
      { name: 'Same name', description: 'one' },
      at('2026-01-01T00:00:00.000Z'),
    )
    const second = createFeature(
      d,
      { name: 'Same name', description: 'two' },
      at('2026-01-02T00:00:00.000Z'),
    )
    expect(first.id).toBe('same-name')
    expect(second.id).toBe('same-name-2')
    expect(listFeatures(d).map((f) => f.id)).toEqual(['same-name', 'same-name-2'])
  })

  it('updates only what it is given and stamps updatedAt', () => {
    const d = dir()
    const created = createFeature(
      d,
      { name: 'Report', description: 'old', kind: 'script' },
      at('2026-01-01T00:00:00.000Z'),
    )
    const changed = updateFeature(
      d,
      created.id,
      { description: 'new' },
      at('2026-01-03T00:00:00.000Z'),
    )
    expect(changed).toMatchObject({ name: 'Report', description: 'new', kind: 'script' })
    expect(changed?.updatedAt).toBe('2026-01-03T00:00:00.000Z')
    expect(changed?.createdAt).toBe('2026-01-01T00:00:00.000Z')
    expect(updateFeature(d, 'nope', { name: 'x' })).toBeNull()
  })

  it('deletes a record and reports a missing one', () => {
    const d = dir()
    const created = createFeature(d, { name: 'Report', description: 'x' })
    expect(removeFeature(d, created.id)).toBe(true)
    expect(listFeatures(d)).toEqual([])
    expect(removeFeature(d, created.id)).toBe(false)
  })

  it('refuses ids that could escape the directory', () => {
    const d = dir()
    expect(isFeatureId('../../etc/passwd')).toBe(false)
    expect(isFeatureId('Has Space')).toBe(false)
    expect(isFeatureId('ok-id-2')).toBe(true)
    expect(getFeature(d, '../../../etc/passwd')).toBeNull()
    expect(removeFeature(d, '../x')).toBe(false)
    expect(updateFeature(d, '../x', { name: 'y' })).toBeNull()
  })

  it('skips a malformed record instead of failing the listing', () => {
    const d = dir()
    createFeature(d, { name: 'Good', description: 'x' })
    writeFileSync(join(d, 'broken.json'), '{ not json')
    writeFileSync(join(d, 'wrong-shape.json'), JSON.stringify({ id: 'wrong-shape', name: 1 }))
    expect(listFeatures(d).map((f) => f.name)).toEqual(['Good'])
  })

  it('requires a name and a description, and caps the list', () => {
    const d = dir()
    expect(() => createFeature(d, { name: '   ', description: 'x' })).toThrow(/name/)
    expect(() => createFeature(d, { name: 'x', description: '  ' })).toThrow(/description/)
    for (let i = 0; i < FEATURE_MAX_COUNT; i++) {
      createFeature(d, { name: 'Feature ' + i, description: 'x' })
    }
    expect(() => createFeature(d, { name: 'one more', description: 'x' })).toThrow(/full/)
  })

  it('round-trips a slot and clears it with an explicit null', () => {
    const d = dir()
    const created = createFeature(d, {
      name: 'File switcher',
      description: 'jump to a recent file',
      slot: {
        at: 'home.sidebar',
        element: { type: 'button', label: 'Files', action: { kind: 'open', path: '/tmp/a.docx' } },
      },
    })
    expect(created.kind).toBe('slot')
    expect(getFeature(d, created.id)?.slot?.at).toBe('home.sidebar')
    expect(updateFeature(d, created.id, { slot: null })?.slot).toBeUndefined()
  })

  it('refuses an invalid slot on write, and skips a malformed one on read', () => {
    const d = dir()
    expect(() =>
      createFeature(d, {
        name: 'Bad',
        description: 'x',
        slot: { at: 'nowhere', element: { type: 'text', value: 'x' } },
      }),
    ).toThrow(/invalid slot/)
    writeFileSync(
      join(d, 'half.json'),
      JSON.stringify({
        id: 'half',
        name: 'Half',
        description: 'x',
        kind: 'slot',
        slot: { at: 'home.card', element: { type: 'nope' } },
        createdAt: 'x',
        updatedAt: 'x',
      }),
    )
    expect(listFeatures(d).map((feature) => feature.id)).toEqual([])
  })

  it('round-trips an appearance and gives it back with an explicit null', () => {
    const d = dir()
    const created = createFeature(d, {
      name: 'Night shift',
      description: 'softer chrome',
      appearance: { background: 'surface-subtle', accent: 'violet' },
    })
    expect(created.kind).toBe('appearance')
    expect(getFeature(d, created.id)?.appearance).toEqual({
      background: 'surface-subtle',
      accent: 'violet',
    })
    expect(updateFeature(d, created.id, { appearance: null })?.appearance).toBeUndefined()
  })

  it('refuses a raw colour on write, and skips one on read', () => {
    const d = dir()
    expect(() =>
      createFeature(d, {
        name: 'Loud',
        description: 'x',
        appearance: { background: '#ff00ff' as never },
      }),
    ).toThrow(/invalid appearance/)
    writeFileSync(
      join(d, 'loud.json'),
      JSON.stringify({
        id: 'loud',
        name: 'Loud',
        description: 'x',
        kind: 'appearance',
        appearance: { accent: 'chartreuse' },
        createdAt: 'x',
        updatedAt: 'x',
      }),
    )
    expect(listFeatures(d).map((feature) => feature.id)).toEqual([])
  })

  it('round-trips a script, refuses a module, and clears it with null', () => {
    const d = dir()
    const created = createFeature(d, {
      name: 'Tidy headings',
      description: 'unfold wrapped headings',
      script: { source: '$.log("hello")' },
    })
    expect(created.kind).toBe('script')
    expect(getFeature(d, created.id)?.script?.source).toBe('$.log("hello")')
    expect(() =>
      createFeature(d, {
        name: 'Module',
        description: 'x',
        script: { source: 'export default 1' },
      }),
    ).toThrow(/invalid script/)
    expect(updateFeature(d, created.id, { script: null })?.script).toBeUndefined()
  })

  it('leaves no temporary file behind', () => {
    const d = dir()
    createFeature(d, { name: 'Report', description: 'x' })
    expect(readdirSync(d)).toEqual(['report.json'])
  })
})
