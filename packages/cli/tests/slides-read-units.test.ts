import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { defaultRegistry } from '../src/cli'
import { resolveTool, toolShape, TOOLS } from '../src/mcp/tools'
import { run, tempDir } from './helpers'

const INCH = 914400

async function smallDeck(): Promise<string> {
  const dir = tempDir()
  const ops = join(dir, 'create.json')
  writeFileSync(
    ops,
    JSON.stringify([
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: INCH, y: INCH / 2, cx: 4 * INCH, cy: 1.2 * INCH },
        paragraphs: [{ runs: [{ text: 'Units' }] }],
      },
    ]),
  )
  const out = join(dir, 'deck.pptx')
  expect((await run(['create', '--type', 'pptx', '--ops', ops, '--out', out, '--json'])).code).toBe(
    0,
  )
  return out
}

async function read(file: string, ...flags: string[]) {
  const r = await run(['slides', 'read', file, ...flags, '--json'])
  expect(r.code).toBe(0)
  return r.json().detail
}

describe('slides read --units', () => {
  it('keeps the default EMU output unchanged', async () => {
    const out = await smallDeck()
    const plain = await run(['slides', 'read', out, '--json'])
    const explicit = await run(['slides', 'read', out, '--units', 'emu', '--json'])
    expect(explicit.stdout).toBe(plain.stdout)
    const deck = plain.json().detail
    expect(deck.size).toEqual({
      cx: 12192000,
      cy: 6858000,
      inches: { width: 13.33, height: 7.5 },
    })
    expect(deck.pages[0].elements[0].box).toEqual({
      x: INCH,
      y: INCH / 2,
      cx: 4 * INCH,
      cy: 1.2 * INCH,
    })
    expect(deck.units).toMatch(/^EMU \(914400 per inch\); slide ids/)
  })

  it.each([
    ['in', { x: 1, y: 0.5, cx: 4, cy: 1.2 }, { cx: 13.33, cy: 7.5 }],
    ['cm', { x: 2.54, y: 1.27, cx: 10.16, cy: 3.05 }, { cx: 33.87, cy: 19.05 }],
    ['pt', { x: 72, y: 36, cx: 288, cy: 86.4 }, { cx: 960, cy: 540 }],
    ['px', { x: 96, y: 48, cx: 384, cy: 115.2 }, { cx: 1280, cy: 720 }],
  ])('converts boxes and the slide size to %s', async (unit, box, size) => {
    const out = await smallDeck()
    const deck = await read(out, '--units', unit)
    expect(deck.pages[0].elements[0].box).toEqual({ ...box, unit })
    expect(deck.size).toMatchObject({ ...size, unit })
    expect(deck.size.inches).toEqual({ width: 13.33, height: 7.5 })
    expect(deck.units).toContain(`unit: "${unit}"`)
    expect(deck.units).toContain(`"1.5${unit}"`)
  })

  it('rejects an unknown unit before opening the file', async () => {
    const r = await run(['slides', 'read', '/nonexistent.pptx', '--units', 'mm', '--json'])
    expect(r.json()).toMatchObject({ code: 1, error: 'invalid_argument' })
    expect(r.json().message).toContain('emu, pt, px, cm, in')
    expect(r.json().detail.valid_values).toEqual(['emu', 'pt', 'px', 'cm', 'in'])
  })

  it('exposes units as an enum on the slides_read MCP tool', () => {
    const tool = resolveTool(
      TOOLS.find((t) => t.name === 'slides_read')!,
      defaultRegistry(),
    )
    const shape = toolShape(tool)
    expect(shape.units!.safeParse('cm').success).toBe(true)
    expect(shape.units!.safeParse('mm').success).toBe(false)
    expect(shape.units!.safeParse(undefined).success).toBe(true)
  })
})
