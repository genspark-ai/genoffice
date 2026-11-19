import { existsSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { configureMetricsCache, familyVerticalMetrics, findSystemFont } from '../src/index'

const darwin = process.platform === 'darwin'
const hasHelvetica = darwin && existsSync('/System/Library/Fonts/Helvetica.ttc')

const cacheDir = join(tmpdir(), `font-metrics-test-${process.pid}`)
afterEach(() => rmSync(cacheDir, { recursive: true, force: true }))

describe('familyVerticalMetrics', () => {
  it('returns plausible metrics for an installed family', () => {
    if (!hasHelvetica) return
    const m = familyVerticalMetrics('Helvetica')
    expect(m).not.toBeNull()
    expect(m!.unitsPerEm).toBeGreaterThan(0)
    expect(m!.hheaAscender).toBeGreaterThan(0)
    expect(m!.hheaDescender).toBeLessThan(0)
    // Helvetica.ttc carries an OS/2 table: both metric groups present
    expect(m!.winAscent).toBeGreaterThan(0)
    expect(m!.winDescent).toBeGreaterThan(0)
    expect(m!.typoAscender).not.toBeNull()
    expect(typeof m!.useTypoMetrics).toBe('boolean')
  })

  it('matches family names exactly after normalization, never fuzzily', () => {
    expect(familyVerticalMetrics('No Such Font Family ZZZ')).toBeNull()
    if (!hasHelvetica) return
    // NFKC + case + separator folding is the only tolerated variance
    const a = familyVerticalMetrics('Helvetica')
    const b = familyVerticalMetrics('helvetica')
    expect(b).toEqual(a)
    expect(familyVerticalMetrics('Helveticaish')).toBeNull()
  })

  it('resolves ttc faces (metrics come from the Regular face)', () => {
    if (!hasHelvetica) return
    const m = familyVerticalMetrics('Helvetica')
    const bytes = findSystemFont('Helvetica', '')
    expect(m).not.toBeNull()
    expect(bytes).not.toBeNull()
    // same face: the standalone sfnt's hhea ascender equals the metrics value
    const num = bytes!.readUInt16BE(4)
    for (let t = 0; t < num; t++) {
      if (bytes!.toString('latin1', 12 + 16 * t, 12 + 16 * t + 4) !== 'hhea') continue
      const off = bytes!.readUInt32BE(12 + 16 * t + 8)
      expect(bytes!.readInt16BE(off + 4)).toBe(m!.hheaAscender)
    }
  })

  it('persists results keyed by path+mtime+size and reloads them', () => {
    if (!hasHelvetica) return
    configureMetricsCache(cacheDir)
    const first = familyVerticalMetrics('Helvetica')
    expect(first).not.toBeNull()
    const file = join(cacheDir, 'font-vertical-metrics.json')
    expect(existsSync(file)).toBe(true)
    const stored = JSON.parse(readFileSync(file, 'utf8')) as Record<
      string,
      { path: string; mtimeMs: number; size: number; m: unknown }
    >
    expect(stored['helvetica']!.m).toEqual(first)
    // fresh in-memory state must serve the query from the persisted entry
    configureMetricsCache(cacheDir)
    expect(familyVerticalMetrics('Helvetica')).toEqual(first)
  })
})
