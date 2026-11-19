import { chmodSync, existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { run, tempDir, writeMinimalPdf } from './helpers'

/** A stand-in for the GenOffice binary: copies a prepared PDF to --out and prints the envelope. */
function fakeApp(dir: string, pdf: string): string {
  const fake = join(dir, 'fake-genoffice.sh')
  writeFileSync(
    fake,
    `#!/bin/sh\nwhile [ $# -gt 0 ]; do if [ "$1" = "--out" ]; then out="$2"; fi; if [ "$1" = "--to" ]; then to="$2"; fi; shift; done\n[ "$to" = "pdf" ] || exit 9\ncp "${pdf}" "$out"\necho '{"status":"ok","summary":"exported"}'\n`,
  )
  chmodSync(fake, 0o755)
  return fake
}

describe('genoffice render', () => {
  it('rasterizes a PDF directly, one PNG per page, 1-based names and --page', async () => {
    const dir = tempDir()
    const pdf = writeMinimalPdf(join(dir, 'report.pdf'))
    const shots = join(dir, 'shots')
    const r = await run(['render', pdf, '--out', shots, '--scale', '2', '--json'])
    expect(r.code).toBe(0)
    const j = r.json()
    expect(j.detail.via).toBe('pdfium')
    expect(j.detail.files).toEqual([
      { page: 1, path: join(shots, 'report-01.png'), width: 1224, height: 1584 },
    ])
    expect(Array.from(readFileSync(j.detail.files[0].path).slice(0, 4))).toEqual([
      0x89, 0x50, 0x4e, 0x47,
    ])

    const one = await run(['render', pdf, '--out', shots, '--page', '1', '--json'])
    expect(one.code).toBe(0)
    expect(one.json().detail.files).toHaveLength(1)
    const beyond = await run(['render', pdf, '--out', shots, '--page', '2', '--json'])
    expect(beyond.code).toBe(1)
    expect(beyond.json().message).toContain('--page out of range (1-1)')
    const zero = await run(['render', pdf, '--out', shots, '--page', '0', '--json'])
    expect(zero.code).toBe(1)
    expect(zero.json().message).toContain('1 or more')
  })

  it('prints a workbook through the app export before rasterizing', async () => {
    if (process.platform === 'win32') return
    const dir = tempDir()
    const table = join(dir, 'table.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['item', 'qty'],
        ['Apple', 2],
      ]),
    )
    const xlsx = join(dir, 'compatibility-basic.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx])).code).toBe(0)
    const pdf = writeMinimalPdf(join(dir, 'export.pdf'))
    const shots = join(dir, 'shots')
    const r = await run(['render', xlsx, '--out', shots, '--json'], {
      env: { ...process.env, GENOFFICE_APP_BIN: fakeApp(dir, pdf) },
    })
    expect(r.code).toBe(0)
    expect(r.json().detail.via).toContain('headless-export')
    const files = r.json().detail.files as { path: string; width: number }[]
    expect(files).toHaveLength(1)
    expect(files[0]!.path).toBe(join(shots, 'compatibility-basic-01.png'))
    expect(files[0]!.width).toBe(612)
    expect(existsSync(files[0]!.path)).toBe(true)
  })

  it('crops a slide to an element and builds a contact sheet', async () => {
    if (process.platform === 'win32') return
    const dir = tempDir()
    const INCH = 914400
    const create = join(dir, 'create.json')
    writeFileSync(
      create,
      JSON.stringify([
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: INCH, y: INCH, cx: 2 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: 'Crop me' }] }],
        },
      ]),
    )
    const deck = join(dir, 'deck.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', create, '--out', deck])).code).toBe(0)
    const read = (await run(['slides', 'read', deck, '--json'])).json().detail
    const id = read.pages[0].elements[0].id as string
    const pdf = writeMinimalPdf(join(dir, 'export.pdf'))
    const shots = join(dir, 'shots')
    const env = { ...process.env, GENOFFICE_APP_BIN: fakeApp(dir, pdf) }
    const r = await run(
      ['render', deck, '--out', shots, '--el', id, '--pad', '10', '--grid', '--json'],
      { env },
    )
    expect(r.code).toBe(0)
    const files = r.json().detail.files as {
      path: string
      width: number
      height: number
      el?: string
    }[]
    const crop = files.find((f) => f.el === id)!
    expect(crop.path).toBe(join(shots, `deck-01-${id}.png`))
    const pxPerEmu = 612 / read.size.cx
    expect(crop.width).toBe(Math.ceil(3 * INCH * pxPerEmu + 10) - Math.floor(INCH * pxPerEmu - 10))
    expect(existsSync(crop.path)).toBe(true)
    const grid = r.json().detail.grid
    expect(grid.path).toBe(join(shots, 'deck-grid.png'))
    expect(grid.tiles).toEqual([{ page: 1, x: 8, y: 8, w: 320, h: Math.round((792 * 320) / 612) }])
    expect(r.json().summary).toContain('1 crop(s) and a contact sheet')

    const tight = await run(['render', deck, '--out', shots, '--el', id, '--pad', '0', '--json'], {
      env,
    })
    expect(tight.code).toBe(0)
    const tightCrop = (tight.json().detail.files as { el?: string; width: number }[]).find(
      (f) => f.el === id,
    )!
    expect(tightCrop.width).toBe(Math.ceil(3 * INCH * pxPerEmu) - Math.floor(INCH * pxPerEmu))
    const missing = await run(['render', deck, '--out', shots, '--el', 'e_nope', '--json'], { env })
    expect(missing.code).toBe(1)
    expect(missing.json().error).toBe('target_not_found')
    const wrongPage = await run(
      ['render', deck, '--out', shots, '--el', id, '--page', '2', '--json'],
      { env },
    )
    expect(wrongPage.code).toBe(1)
    expect(wrongPage.json().suggestion).toContain('page 1')
    const notPptx = await run(['render', pdf, '--out', shots, '--el', id, '--json'])
    expect(notPptx.json().error).toBe('unsupported')
  })

  it('refuses unknown formats and needs --out', async () => {
    const dir = tempDir()
    const txt = join(dir, 'notes.txt')
    writeFileSync(txt, 'x')
    const r = await run(['render', txt, '--out', join(dir, 'shots'), '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('cannot render .txt')
    expect(r.json().detail.supported).toContain('docx')
    const noOut = await run(['render', writeMinimalPdf(join(dir, 'a.pdf')), '--json'])
    expect(noOut.code).toBe(1)
    expect(noOut.json().message).toContain('--out')
  })
})
