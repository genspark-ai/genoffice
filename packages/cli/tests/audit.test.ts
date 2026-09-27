import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { solidPng } from '@genoffice/pptx-engine'
import { run, tempDir } from './helpers'

const INCH = 914400

describe('slides audit', () => {
  it('returns typed findings whose setTransform suggestion fixes the geometry', async () => {
    const dir = tempDir()
    const create = join(dir, 'create.json')
    writeFileSync(
      create,
      JSON.stringify([
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: 9 * INCH, y: INCH, cx: 8 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: 'Far right', fontSize: 24 }] }],
        },
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: INCH, y: 3 * INCH, cx: 3 * INCH, cy: Math.round(0.3 * INCH) },
          paragraphs: [
            {
              runs: [
                {
                  text: 'A paragraph long enough to wrap onto several lines inside a narrow box, ',
                },
              ],
            },
            {
              runs: [
                {
                  text: 'and a second paragraph that surely does not fit into a third of an inch.',
                },
              ],
            },
          ],
        },
      ]),
    )
    const out = join(dir, 'deck.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', create, '--out', out])).code).toBe(0)
    const ids = (await run(['slides', 'read', out, '--json']))
      .json()
      .detail.pages[0].elements.map((e: { id: string }) => e.id) as string[]
    const rotate = join(dir, 'rotate.json')
    writeFileSync(
      rotate,
      JSON.stringify([
        {
          op: 'setTransform',
          target: { slide: 0, el: ids[0] },
          box: { x: 9 * INCH, y: INCH, cx: 8 * INCH, cy: INCH },
          rotDeg: 15,
        },
      ]),
    )
    expect((await run(['slides', 'apply', out, '--ops', rotate, '--json'])).code).toBe(0)

    const first = await run(['slides', 'audit', out, '--json'])
    expect(first.code).toBe(0)
    const detail = first.json().detail
    const issues = detail.issues as {
      id: string
      code: string
      level: string
      path: string
      el: string
      suggest?: { op: string; target: { slide: string; el: string }; box: Record<string, number> }
    }[]
    const oob = issues.find((i) => i.code === 'out_of_bounds')!
    expect(oob).toMatchObject({ level: 'error', id: 'E1' })
    expect(oob.path).toMatch(/^s_\d+\/e_/)
    expect(oob.suggest).toMatchObject({ op: 'setTransform', target: { el: oob.el } })
    expect(oob.suggest!.target.slide).toMatch(/^s_/)
    expect((oob.suggest as { rotDeg?: number }).rotDeg).toBe(15)
    for (const el of new Set(issues.map((i) => i.el))) {
      const boxes = issues
        .filter((i) => i.el === el && i.suggest)
        .map((i) => JSON.stringify(i.suggest))
      expect(new Set(boxes).size).toBeLessThanOrEqual(1)
    }
    const overflow = issues.find((i) => i.code === 'text_overflow')!
    expect(overflow.suggest!.box.cy).toBeGreaterThan(0.3 * INCH)
    expect(detail.slides[0].issues).toHaveLength(issues.length)
    expect(detail.counts.error).toBeGreaterThanOrEqual(2)

    const fix = join(dir, 'fix.json')
    writeFileSync(fix, JSON.stringify(issues.filter((i) => i.suggest).map((i) => i.suggest)))
    expect((await run(['slides', 'apply', out, '--ops', fix, '--json'])).code).toBe(0)
    const second = await run(['slides', 'audit', out, '--json'])
    const codes = (second.json().detail.issues as { code: string }[]).map((i) => i.code)
    expect(codes).not.toContain('out_of_bounds')
    expect(codes).not.toContain('text_overflow')
  })

  it('flags elements fully off the slide and stretched pictures, sparing a well-cropped one', async () => {
    const dir = tempDir()
    const png = join(dir, 'wide.png')
    writeFileSync(png, solidPng(400, 200, [20, 40, 60]))
    const create = join(dir, 'create.json')
    writeFileSync(
      create,
      JSON.stringify([
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: 14 * INCH, y: INCH, cx: 2 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: 'Gone', fontSize: 24 }] }],
        },
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: 12 * INCH, y: 3 * INCH, cx: 2 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: 'Edge', fontSize: 24 }] }],
        },
        {
          op: 'addPicture',
          target: { slide: 0 },
          bytes: png,
          offset: { x: INCH, y: 5 * INCH, cx: 2 * INCH, cy: 2 * INCH },
        },
        {
          op: 'addPicture',
          target: { slide: 0 },
          bytes: png,
          offset: { x: 4 * INCH, y: 5 * INCH, cx: 2 * INCH, cy: 2 * INCH },
        },
      ]),
    )
    const out = join(dir, 'deck.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', create, '--out', out])).code).toBe(0)
    const ids = (await run(['slides', 'read', out, '--json']))
      .json()
      .detail.pages[0].elements.map((e: { id: string }) => e.id) as string[]
    const crop = join(dir, 'crop.json')
    writeFileSync(
      crop,
      JSON.stringify([
        {
          op: 'setPictureSrcRect',
          target: { slide: 0, el: ids[3] },
          srcRect: { l: 0.25, t: 0, r: 0.25, b: 0 },
        },
      ]),
    )
    expect((await run(['slides', 'apply', out, '--ops', crop, '--json'])).code).toBe(0)

    const r = await run(['slides', 'audit', out, '--json'])
    expect(r.code).toBe(0)
    const issues = r.json().detail.issues as {
      code: string
      level: string
      el: string
      expected_ratio?: number
      actual_ratio?: number
      distortion_pct?: number
      suggest?: { box: Record<string, number> }
    }[]
    expect(issues.map((i) => [i.el, i.code])).toEqual([
      [ids[0], 'off_slide'],
      [ids[1], 'out_of_bounds'],
      [ids[2], 'picture_distorted'],
    ])
    const off = issues[0]!
    expect(off.level).toBe('error')
    expect(off.suggest!.box.x + off.suggest!.box.cx).toBeLessThanOrEqual(13.334 * INCH)
    const pic = issues[2]!
    expect(pic).toMatchObject({
      level: 'warning',
      expected_ratio: 2,
      actual_ratio: 1,
      distortion_pct: 50,
    })
    expect(pic.suggest!.box).toMatchObject({ cx: 2 * INCH, cy: INCH })

    const fix = join(dir, 'fix.json')
    writeFileSync(fix, JSON.stringify(issues.map((i) => i.suggest)))
    expect((await run(['slides', 'apply', out, '--ops', fix, '--json'])).code).toBe(0)
    expect((await run(['slides', 'audit', out, '--json'])).json().detail.issues).toEqual([])
  })
})
