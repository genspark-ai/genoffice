import { describe, expect, it } from 'vitest'
import { renderWavedrom } from '../src/renderer/editor/wavedrom'

describe('renderWavedrom', () => {
  it('renders WaveJSON (JSON5 flavoured) into a self-contained svg', async () => {
    const result = await renderWavedrom("{ signal: [{ name: 'clk', wave: 'p...' }] }")
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.svg).toMatch(/^<svg[^>]*class="WaveDrom"/)
    expect(result.svg).toMatch(/viewBox="0 0 \d+ \d+"/)
  })

  it('namespaces the skin classes so the inline stylesheet cannot leak into the page', async () => {
    const result = await renderWavedrom("{ signal: [{ name: 'clk', wave: 'p...' }] }")
    if (!result.ok) throw new Error(result.error)
    const css = /<style[^>]*>([\s\S]*?)<\/style>/.exec(result.svg)![1]!
    expect(css).toContain('.wd-s1{')
    expect(css).toContain('svg.WaveDrom text{')
    expect(css).not.toMatch(/(^|\})\.s1\{/)
    expect(css).not.toMatch(/(^|\})text\{/)
    expect(result.svg).toContain('class="wd-s1"')
    expect(result.svg).not.toContain('class="s1"')
  })

  it('renders register fields too', async () => {
    const result = await renderWavedrom('{ reg: [{ bits: 8, name: "op" }] }')
    expect(result.ok).toBe(true)
  })

  it('reports parse errors and non-WaveJSON objects instead of drawing an empty box', async () => {
    const broken = await renderWavedrom("{ signal: [{ name: 'clk' ]")
    expect(broken.ok).toBe(false)
    if (!broken.ok) expect(broken.error).toMatch(/JSON5/)
    const notWave = await renderWavedrom('{ foo: 1 }')
    expect(notWave.ok).toBe(false)
    const list = await renderWavedrom('[1, 2]')
    expect(list.ok).toBe(false)
  })
})
