import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir, writeMinimalPdf } from './helpers'

const REPO = resolve(__dirname, '../../..')
const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')
const PPTX = join(REPO, 'packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const XLSX = join(REPO, 'apps/sheets/fixtures/generated/compatibility-basic.xlsx')

describe('genoffice cli', () => {
  it('prints help and usage errors with the documented exit codes', async () => {
    expect((await run([])).code).toBe(1)
    expect((await run(['--help'])).code).toBe(0)
    expect((await run(['help', 'convert'])).stdout).toContain('Usage: genoffice convert')
    const unknown = await run(['frobnicate', '--json'])
    expect(unknown.code).toBe(1)
    expect(unknown.json()).toMatchObject({ status: 'error', code: 1, command: 'frobnicate' })
    expect((await run(['--version'])).stdout).toMatch(/^genoffice \d/)
  })

  it('reports a missing file as exit code 2', async () => {
    const r = await run(['info', '/nonexistent/x.docx', '--json'])
    expect(r.code).toBe(2)
    expect(r.json().message).toContain('file not found')
  })

  it('info describes docx, pptx, csv and markdown without any app process', async () => {
    const docx = await run(['info', DOCX, '--json'])
    expect(docx.code).toBe(0)
    expect(docx.json().detail).toMatchObject({ format: 'docx' })
    expect(docx.json().detail.blocks).toBeGreaterThan(0)

    const pptx = await run(['info', PPTX, '--json'])
    expect(pptx.code).toBe(0)
    expect(pptx.json().detail.slides).toBeGreaterThan(0)
    expect(pptx.json().detail.slide_size_in.width).toBeGreaterThan(0)

    const dir = tempDir()
    const csv = join(dir, 'data.csv')
    writeFileSync(csv, 'name,qty\nApple,3\nPear,5\n')
    const csvInfo = await run(['info', csv, '--json'])
    expect(csvInfo.json().detail).toMatchObject({ rows: 3, columns: 2, delimiter: ',' })

    const md = join(dir, 'notes.md')
    writeFileSync(md, '# Title\n\ntext\n\n## Sub\n')
    const mdInfo = await run(['info', md])
    expect(mdInfo.code).toBe(0)
    expect(mdInfo.stdout).toContain('headings: 2')
  })

  it('info counts pdf pages and flags encrypted files', async () => {
    const dir = tempDir()
    const pdf = writeMinimalPdf(join(dir, 'one.pdf'))
    const r = await run(['info', pdf, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ pages: 1, encrypted: false })
    // a password nobody asked for must not flip the flag
    const pw = await run(['--json', 'info', pdf, '--password', 'x'])
    expect(pw.json().detail).toMatchObject({ pages: 1, encrypted: false })
  })

  it('converts csv to xlsx and refuses to overwrite without --force', async () => {
    const dir = tempDir()
    const csv = join(dir, 'data.csv')
    writeFileSync(csv, 'name,qty\nApple,3\n')
    const first = await run(['convert', csv, '--to', 'xlsx', '--json'])
    expect(first.code).toBe(0)
    const out = first.json().output_path as string
    expect(out).toBe(join(dir, 'data.xlsx'))
    const zip = await JSZip.loadAsync(readFileSync(out))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('Apple')

    const again = await run(['convert', csv, '--to', 'xlsx', '--json'])
    expect(again.code).toBe(2)
    expect((await run(['convert', csv, '--to', 'xlsx', '--force'])).code).toBe(0)
  })

  it('rejects unsupported routes as a usage error', async () => {
    const r = await run(['convert', DOCX, '--to', 'pptx', '--json'])
    expect(r.code).toBe(1)
    expect(r.json().detail.supported).toContain('pdf→docx/pptx/xlsx')
  })

  it('converts pdf to docx through the local engine', async () => {
    const dir = tempDir()
    const pdf = writeMinimalPdf(join(dir, 'one.pdf'), 'Converted by genoffice')
    const r = await run([
      'convert',
      pdf,
      '--to',
      'docx',
      '--out',
      join(dir, 'sub/one.docx'),
      '--json',
    ])
    expect(r.code).toBe(0)
    expect(r.json().detail.pages).toBe(1)
    const out = r.json().output_path as string
    expect(existsSync(out)).toBe(true)
    const zip = await JSZip.loadAsync(readFileSync(out))
    const doc = await zip.file('word/document.xml')!.async('string')
    expect(doc).toContain('Converted by genoffice')
  })

  it.skipIf(!xlsxSidecarPath())('info reads workbook sheets through the xlsx sidecar', async () => {
    const r = await run(['info', XLSX, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.sheets.length).toBeGreaterThan(0)
    expect(r.json().detail.sheets[0]).toMatchObject({ name: expect.any(String) })
  })

  it('open fails with exit code 4 when no app binary is available', async () => {
    // A GenOffice already running on this machine publishes its control endpoint in the
    // user-data dir; `open` would then hand the file to that live instance and exit 0,
    // never reaching the missing-binary path. Isolate the lookup like control.test.ts does.
    const r = await run(['open', DOCX, '--json'], {
      env: {
        GENOFFICE_APP_BIN: '/nonexistent/GenOffice',
        GENOFFICE_USER_DATA: tempDir(),
        GENOFFICE_AUDIT_LOG: 'off',
      },
    })
    expect(r.code).toBe(4)
  })
})

describe('option validation', () => {
  it('rejects an option the command does not define instead of ignoring it', async () => {
    const r = await run(['info', DOCX, '--dry-run', '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('--dry-run')
    expect(r.json().detail.options).toContain('--password')
  })
})
