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
