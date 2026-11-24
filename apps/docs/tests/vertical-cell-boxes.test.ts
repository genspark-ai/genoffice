// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  syncVerticalCellBoxes,
  turnedBoxVars,
  turnedSides,
} from '../src/renderer/editor/vertical-cell-boxes'

const css = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)
const rule = (selector: string): string => {
  const at = css.indexOf(selector)
  expect(at).toBeGreaterThan(-1)
  return css.slice(css.indexOf('{', at) + 1, css.indexOf('}', at))
}

const horizontal = { top: '6px', right: '2px', bottom: '10px', left: '4px' }

describe('paragraph boxes in a vertical cell', () => {
  it('turn a quarter right with tbRl text: space after lands on the left, the indent at the top', () => {
    expect(turnedSides('vertical-rl', horizontal)).toEqual({
      top: '4px',
      right: '6px',
      bottom: '2px',
      left: '10px',
    })
  })

  it('turn a quarter left with btLr text: space after lands on the right, the indent at the bottom', () => {
    expect(turnedSides('sideways-lr', horizontal)).toEqual({
      top: '2px',
      right: '10px',
      bottom: '4px',
      left: '6px',
    })
  })

  it('carry margin and padding of every side as variables', () => {
    const zero = { top: '0px', right: '0px', bottom: '0px', left: '0px' }
    expect(turnedBoxVars('vertical-rl', horizontal, zero)).toBe(
      '--vb-mt:4px;--vb-pt:0px;--vb-mr:6px;--vb-pr:0px;--vb-mb:2px;--vb-pb:0px;--vb-ml:10px;--vb-pl:0px',
    )
  })

  const cell = (wrapper: string, paras: string) => {
    document.body.innerHTML = `<table><tr><td>${wrapper.replace('><', `>${paras}<`)}</td><td><p id="h" style="margin-bottom:10px">h</p></td></tr></table>`
    return document.body
  }

  it('are read from the cell and written once: a second pass changes nothing', () => {
    const root = cell(
      '<div class="cell-vert" style="writing-mode:vertical-rl"></div>',
      '<p id="a" style="margin-bottom:10px;margin-left:4px">a</p><p id="b" style="margin-top:6px">b</p>',
    )
    expect(syncVerticalCellBoxes(root)).toBe(true)
    const a = root.querySelector<HTMLElement>('#a')!
    expect(a.style.getPropertyValue('--vb-ml')).toBe('10px')
    expect(a.style.getPropertyValue('--vb-mt')).toBe('4px')
    expect(a.style.getPropertyValue('--vb-mb')).toBe('0px')
    expect(root.querySelector<HTMLElement>('#b')!.style.getPropertyValue('--vb-mr')).toBe('6px')
    expect(a.dataset.vbox).toBeTruthy()
    expect(root.querySelector('.cell-flow-probe')).toBeNull()
    expect(syncVerticalCellBoxes(root)).toBe(false)
    // the horizontal neighbour is never touched
    expect(root.querySelector<HTMLElement>('#h')!.dataset.vbox).toBeUndefined()
  })

  it('cover an exact-height vertical cell and skip an exact-height horizontal one', () => {
    const vertical = cell(
      '<div class="cell-clip" style="height:20px;writing-mode:sideways-lr"></div>',
      '<p id="a" style="margin-bottom:10px">a</p>',
    )
    expect(syncVerticalCellBoxes(vertical)).toBe(true)
    expect(vertical.querySelector<HTMLElement>('#a')!.style.getPropertyValue('--vb-mr')).toBe(
      '10px',
    )
    const plain = cell('<div class="cell-clip" style="height:20px"></div>', '<p id="a">a</p>')
    expect(syncVerticalCellBoxes(plain)).toBe(false)
    expect(plain.querySelector<HTMLElement>('#a')!.dataset.vbox).toBeUndefined()
  })

  it('a document without vertical cells is left alone', () => {
    document.body.innerHTML =
      '<p style="margin-bottom:10px">a</p><table><tr><td><p>b</p></td></tr></table>'
    const before = document.body.innerHTML
    expect(syncVerticalCellBoxes(document.body)).toBe(false)
    expect(document.body.innerHTML).toBe(before)
  })

  it('the stylesheet applies the turned box only outside the horizontal reading pass', () => {
    const applied = rule(
      ":is(.cell-vert, .cell-clip[style*='writing-mode']):not(.cell-flow-probe) > [data-vbox]",
    )
    expect(applied).toMatch(/margin:\s*var\(--vb-mt\) var\(--vb-mr\) var\(--vb-mb\) var\(--vb-ml\)/)
    expect(applied).toMatch(
      /padding:\s*var\(--vb-pt\) var\(--vb-pr\) var\(--vb-pb\) var\(--vb-pl\)/,
    )
    expect(rule('.cell-flow-probe {')).toMatch(/writing-mode:\s*horizontal-tb\s*!important/)
    expect(rule(".cell-clip[style*='writing-mode'] {")).toMatch(/max-width:\s*100%/)
  })
})

describe('Horizontal in Vertical in the stylesheet', () => {
  it('turns the run a quarter left of a horizontal line, at its own size', () => {
    const run = rule('[data-doc-style][data-ea-vert] {')
    expect(run).toMatch(/display:\s*inline-block/)
    expect(run).toMatch(/writing-mode:\s*var\(--ea-vert-wm, sideways-lr\)/)
    expect(run).toMatch(/rotate:\s*var\(--ea-vert-rot, none\)/)
    expect(run).toMatch(/white-space:\s*nowrap/)
    expect(css).not.toContain('text-combine-upright')
  })

  it('stands upright in text that runs down and upside down in text that runs up', () => {
    const down = rule("[style*='writing-mode:vertical-rl'],")
    expect(down).toMatch(/--ea-vert-wm:\s*horizontal-tb/)
    expect(down).toMatch(/--ea-vert-rot:\s*none/)
    const up = rule("[style*='writing-mode:sideways-lr'],")
    expect(up).toMatch(/--ea-vert-wm:\s*horizontal-tb/)
    expect(up).toMatch(/--ea-vert-rot:\s*180deg/)
  })
})
