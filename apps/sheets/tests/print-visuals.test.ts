// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'

import { collectPrintCss, snapshotPrintVisuals } from '../src/renderer/print-visuals'
import type { InstalledVisualFrame } from '../src/renderer/WorkbookVisuals'
import type { WorkbookVisualObject } from '../src/shared/desktop-api'

function mount(css: string, body: string): void {
  document.head.innerHTML = `<style>${css}</style>`
  document.body.innerHTML = body
}

const frame = (id: string): InstalledVisualFrame => ({
  visual: { id, sheetId: 's1', kind: 'chart' } as unknown as WorkbookVisualObject,
  fromRow: 1,
  fromColumn: 3,
  toRow: 12,
  toColumn: 8,
  marginX: 6,
  marginY: 2,
  width: 480,
  height: 288,
})

describe('collectPrintCss', () => {
  it('keeps the rules that reach the visuals, resolves :root variables, drops hover and strangers', () => {
    mount(
      `:root { --label: #555; --dim: #888; }
       .xlsx-chart { border: 1px solid var(--edge, #ccc); }
       .chart-svg text { fill: var(--label); }
       .chart-svg .axis-label { fill: var(--dim); }
       .xlsx-chart:hover { outline: 1px solid red; }
       .ribbon-button { color: blue; }
       @media (min-width: 10px) { .xlsx-chart figcaption { font-size: 9px; } .elsewhere { color: red; } }`,
      `<div data-print-visual="v1"><figure class="xlsx-chart"><svg class="chart-svg"><text>1</text></svg><figcaption>t</figcaption></figure></div>`,
    )
    const root = document.querySelector('[data-print-visual="v1"]')!
    const css = collectPrintCss(document, [root])
    expect(css).toContain('.xlsx-chart {')
    expect(css).toContain('border: 1px solid #ccc')
    expect(css).toContain('fill: #555')
    expect(css).not.toContain('.axis-label')
    expect(css).not.toContain('hover')
    expect(css).not.toContain('ribbon-button')
    expect(css).toMatch(/@media \(min-width: 10px\) \{\s*\.xlsx-chart figcaption/)
    expect(css).not.toContain('.elsewhere')
    expect(css).not.toContain('var(')
  })
})

describe('snapshotPrintVisuals', () => {
  it('clones each installed frame without its controls and carries the anchor geometry', () => {
    mount(
      `.xlsx-chart { color: black; }`,
      `<div data-print-visual="v1" class="xlsx-print-visual"><figure class="xlsx-chart" tabindex="0"><svg></svg><button>x</button><div contenteditable="true">label</div></figure></div>
       <div data-print-visual="gone"></div>`,
    )
    const snap = snapshotPrintVisuals(document, [frame('v1'), frame('missing'), frame('gone')])
    expect(snap.visuals).toHaveLength(1)
    const [v] = snap.visuals
    expect(v).toMatchObject({
      id: 'v1',
      fromRow: 1,
      fromColumn: 3,
      toRow: 12,
      toColumn: 8,
      offsetXPx: 6,
      offsetYPx: 2,
      widthPx: 480,
      heightPx: 288,
    })
    expect(v!.html).toContain('data-print-visual="v1"')
    expect(v!.html).not.toContain('<button')
    expect(v!.html).not.toContain('contenteditable')
    expect(v!.html).not.toContain('tabindex')
    expect(snap.css).toContain('.xlsx-chart')
  })

  it('returns nothing (and no css) when no visual is in the document', () => {
    mount('.xlsx-chart { color: black; }', '<p>no floats</p>')
    expect(snapshotPrintVisuals(document, [frame('v1')])).toEqual({ visuals: [], css: '' })
  })
})
