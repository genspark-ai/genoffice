import { describe, expect, it } from 'vitest'

import standInBoxes from '../src/renderer/fonts/stand-in-boxes.json'
import { fontBoxCss } from '../src/renderer/line-metrics'

const data = standInBoxes as {
  _generated: string
  bundled: Record<string, { ascent: number; descent: number }>
  installed: Record<string, { ascent: number; descent: number }>
}

// Myanmar / Nastaleeq faces legitimately exceed a 1 em descent
const isBox = (b: { ascent: number; descent: number }) =>
  b.ascent > 0.5 && b.ascent < 3 && b.descent > 0 && b.descent < 2

describe('stand-in-boxes.json (tools/build-stand-in-boxes.py)', () => {
  it('is generated and carries an hhea box per bundled and snapshot family', () => {
    expect(data._generated).toContain('build-stand-in-boxes.py')
    expect(Object.keys(data.bundled).length).toBeGreaterThan(50)
    expect(Object.keys(data.installed).length).toBeGreaterThan(100)
    for (const store of [data.bundled, data.installed])
      for (const [key, box] of Object.entries(store)) {
        expect(key, key).toBe(key.toLowerCase())
        expect(isBox(box), `${key} ${JSON.stringify(box)}`).toBe(true)
      }
  })

  it('overrides and size-adjust fold into the bundled box (Aptos GO = Carlito x 1.063)', () => {
    expect(data.bundled['carlito go']).toEqual({ ascent: 0.75, descent: 0.25 })
    expect(data.bundled['aptos go']).toEqual({ ascent: 0.7988, descent: 0.2663 })
    expect(data.bundled.caladea).toEqual({ ascent: 0.95, descent: 0.222 })
    expect(data.installed['times new roman']).toEqual({ ascent: 0.8911, descent: 0.2163 })
  })

  it('reproduces the former hand table through the chain (Word probe 2026-09-30)', () => {
    const box = (chain: string) => fontBoxCss(chain)[0]
    expect(box("'Aptos','Aptos GO',sans-serif")).toBe('--doc-font-box:1.0651')
    expect(box("'Cambria','Caladea',serif")).toBe('--doc-font-box:1.1720')
    expect(box("'Verdana',sans-serif")).toBe('--doc-font-box:1.2154')
    expect(box("'Tahoma',sans-serif")).toBe('--doc-font-box:1.2070')
    expect(box("'Courier New',monospace")).toBe('--doc-font-box:1.1328')
    expect(box("'Trebuchet MS',sans-serif")).toBe('--doc-font-box:1.1612')
    // heads the old regexes lumped with a sibling now read their own drawn face
    expect(box("'Courier','Menlo','Courier New',monospace")).toBe('--doc-font-box:1.0000')
    expect(box("'Aptos Display','Aptos Display GO',sans-serif")).toBe('--doc-font-box:0.9987')
  })
})
