import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import {
  prefetchFaces,
  resolveDeclaredFace,
  setFontMetricsProvider,
  setResolverFontTable,
} from '../src/renderer/font-resolve'
import {
  lineHeightFactor,
  setDocFontTable,
  tableLineHeightFactor,
} from '../src/renderer/line-metrics'

beforeEach(() => {
  setDocFontTable(null)
  globalThis.__fontResolveDiff = []
})

afterEach(() => setFontMetricsProvider(undefined))

const flush = () => new Promise((r) => setTimeout(r, 0))

describe('Latin cutover', () => {
  it('derives the factors whose table entries were deleted from word-faces.json', () => {
    for (const [name, factor] of [
      ['Poppins', 1.5],
      ['Montserrat', 1.219],
      ['Montserrat Medium', 1.219],
      ['Segoe UI', 1.3301],
      ['Segoe UI Semibold', 1.3301],
      ['Merriweather', 1.257],
      ['Sakkal Majalla', 1.3965],
      ['Lucida Handwriting', 1.3794],
      ['Lucida Console', 1.0],
      ['DM Sans', 1.302],
      ['Lato', 1.2],
      ['Helvetica Neue', 1.193],
      ['Nirmala UI', 1.3301],
      ['Symbol', 1.2251],
      ['Wingdings', 1.1099],
      ['Wingdings 2', 1.0542],
      ['Wingdings 3', 1.1387],
      ['Webdings', 1.0],
    ] as const) {
      expect(lineHeightFactor(name), name).toBe(factor)
    }
  })

  it('lays out a face Word ships that the table never probed for real', () => {
    setDocFontTable([{ name: 'Mangal', panose: '00000400000000000000' }])
    expect(tableLineHeightFactor('Mangal')).toBe(1.22)
    expect(lineHeightFactor('Mangal')).toBe(1.6797)
    expect(lineHeightFactor('Garamond')).toBe(1.125)
  })

  it('keeps a probed table value the resolver disagrees with and records the diff', () => {
    expect(lineHeightFactor('Georgia')).toBe(1.1375)
    expect(lineHeightFactor('Helvetica')).toBe(1.2)
    expect(globalThis.__fontResolveDiff).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ name: 'Georgia', legacy: 1.1375, resolver: 1.1362 }),
        expect.objectContaining({ name: 'Helvetica', legacy: 1.2, resolver: 1 }),
      ]),
    )
  })

  it('East Asian faces take hhea (ascent + descent) x 1.3 (probe 2026-10-22)', () => {
    expect(lineHeightFactor('Malgun Gothic')).toBe(1.7291)
    expect(lineHeightFactor('SimSun')).toBe(1.3)
    expect(lineHeightFactor('Arial Unicode MS')).toBe(1.7417)
    setDocFontTable([{ name: 'Mont', panose: '00000700' }])
    expect(lineHeightFactor('Mont')).toBe(1.22)
    expect(resolveDeclaredFace('Mont')?.substituted).toBe(true)
  })

  it('reads the fontTable through setDocFontTable', () => {
    setResolverFontTable([{ name: 'Montserrat-Bold', altName: 'Times New Roman' }])
    expect(resolveDeclaredFace('Montserrat-Bold')?.family).toMatch(/^times new roman$/i)
    setDocFontTable([{ name: 'Montserrat-Bold', altName: 'Georgia' }])
    expect(resolveDeclaredFace('Montserrat-Bold')?.family).toMatch(/^georgia$/i)
  })
})

describe('installed index failures', () => {
  // Helvetica Neue is only in the macOS snapshot: the chain must not reach it once the IPC failed
  const rejecting = (family: string) =>
    family === 'helvetica neue' ? Promise.reject(new Error('ipc down')) : Promise.resolve(null)

  it('a rejected lookup is cached as unknown and the table answers', async () => {
    const ipc = vi.fn(rejecting)
    setFontMetricsProvider(ipc)
    expect(lineHeightFactor('Helvetica Neue')).toBe(1.193) // snapshot stands in while in flight
    await flush()
    const legacy = tableLineHeightFactor('Helvetica Neue')
    expect(legacy).not.toBe(1.193)
    for (let i = 0; i < 3; i++) expect(lineHeightFactor('Helvetica Neue')).toBe(legacy)
    expect(resolveDeclaredFace('Helvetica Neue')?.substituted).toBe(true)
    await prefetchFaces(['Helvetica Neue'])
    expect(ipc.mock.calls.filter(([f]) => f === 'helvetica neue')).toHaveLength(1)
  })

  it('an answer only drops the cached faces whose chain touched the family', async () => {
    setFontMetricsProvider(rejecting)
    resolveDeclaredFace('Verdana')
    await flush()
    const verdana = resolveDeclaredFace('Verdana')
    expect(verdana?.path).toEqual(['verdana'])
    resolveDeclaredFace('Helvetica Neue')
    await flush()
    expect(resolveDeclaredFace('Verdana')).toBe(verdana)
    expect(resolveDeclaredFace('Helvetica Neue')?.source).not.toBe('system')
  })
})
