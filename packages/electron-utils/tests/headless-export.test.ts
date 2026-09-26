import { describe, expect, it, vi } from 'vitest'

import {
  HEADLESS_EXIT,
  formatHeadlessEnvelope,
  headlessExitCode,
  headlessModuleFor,
  parseHeadlessExportArgv,
  pollUntilReady,
  runHeadlessRendererExport,
} from '../src/headless-export'

/** argv as Electron delivers it: binary first, then the app's own switches. */
const argv = (...rest: string[]): string[] => [
  '/Applications/GenOffice.app/Contents/MacOS/GenOffice',
  ...rest,
]

describe('parseHeadlessExportArgv', () => {
  it('reports no headless request when the flag is absent', () => {
    expect(parseHeadlessExportArgv(argv('/tmp/a.docx'))).toEqual({ kind: 'none' })
  })

  it('parses the space-separated form', () => {
    const parsed = parseHeadlessExportArgv(
      argv('--headless-export', '/tmp/a.docx', '--to', 'pdf', '--out', '/tmp/a.pdf'),
    )
    expect(parsed).toEqual({
      kind: 'ok',
      request: { input: '/tmp/a.docx', targetFormat: 'pdf', outPath: '/tmp/a.pdf', json: false },
    })
  })

  it('parses the --flag=value form and the --json switch', () => {
    const parsed = parseHeadlessExportArgv(
      argv('--headless-export=/tmp/a b.pptx', '--to=PDF', '--out=/tmp/out.pdf', '--json'),
    )
    expect(parsed).toEqual({
      kind: 'ok',
      request: { input: '/tmp/a b.pptx', targetFormat: 'pdf', outPath: '/tmp/out.pdf', json: true },
    })
  })

  it('rejects a missing input, target or output', () => {
    expect(
      parseHeadlessExportArgv(argv('--headless-export', '--to', 'pdf', '--out', '/o.pdf')),
    ).toMatchObject({
      kind: 'error',
    })
    expect(
      parseHeadlessExportArgv(argv('--headless-export', '/tmp/a.docx', '--out', '/o.pdf')),
    ).toMatchObject({
      kind: 'error',
    })
    expect(
      parseHeadlessExportArgv(argv('--headless-export', '/tmp/a.docx', '--to', 'pdf')),
    ).toMatchObject({
      kind: 'error',
    })
  })

  it('rejects a target format no module renders', () => {
    const parsed = parseHeadlessExportArgv(
      argv('--headless-export', '/tmp/a.docx', '--to', 'txt', '--out', '/o.txt'),
    )
    expect(parsed).toMatchObject({ kind: 'error', message: expect.stringContaining('txt') })
  })

  it("accepts the non-pdf targets; input/target pairing is the host's job", () => {
    const parsed = parseHeadlessExportArgv(
      argv('--headless-export', '/tmp/a.html', '--to', 'DOCX', '--out', '/o.docx'),
    )
    expect(parsed).toMatchObject({ kind: 'ok', request: { targetFormat: 'docx' } })
  })

  it('keeps the --json switch on the error branch so failures stay machine-readable', () => {
    const parsed = parseHeadlessExportArgv(argv('--headless-export', '--json'))
    expect(parsed).toEqual({ kind: 'error', json: true, message: expect.any(String) })
  })
})

describe('headlessModuleFor', () => {
  it.each([
    ['/x/a.docx', 'docs'],
    ['/x/a.xlsx', 'sheets'],
    ['/x/a.XLSM', 'sheets'],
    ['/x/a.xls', 'sheets'],
    ['/x/a.csv', 'sheets'],
    ['/x/a.tsv', 'sheets'],
    ['/x/a.pptx', 'slides'],
    ['/x/a.md', 'markdown'],
    ['/x/a.markdown', 'markdown'],
    ['/x/a.html', 'html'],
    ['/x/a.htm', 'html'],
  ])('routes %s to %s', (path, module) => {
    expect(headlessModuleFor(path)).toBe(module)
  })

  it('returns null for an unsupported extension', () => {
    expect(headlessModuleFor('/x/a.pdf')).toBeNull()
    expect(headlessModuleFor('/x/a.doc')).toBeNull()
    expect(headlessModuleFor('/x/noext')).toBeNull()
  })
})

describe('formatHeadlessEnvelope', () => {
  it('prints one human sentence without --json', () => {
    const line = formatHeadlessEnvelope({ ok: true, input: '/a.docx', outPath: '/a.pdf' }, false)
    expect(line).toBe('Exported /a.docx to /a.pdf')
    expect(line).not.toContain('\n')
  })

  it('prints exactly one JSON line with --json', () => {
    const line = formatHeadlessEnvelope({ ok: true, input: '/a.docx', outPath: '/a.pdf' }, true)
    expect(line.split('\n')).toHaveLength(1)
    expect(JSON.parse(line)).toEqual({
      status: 'ok',
      summary: 'Exported /a.docx to /a.pdf',
      output_path: '/a.pdf',
    })
  })

  it('folds newlines out of a multi-line failure so the envelope stays one line', () => {
    const line = formatHeadlessEnvelope(
      { ok: false, code: HEADLESS_EXIT.conversionFailure, message: 'boom\n  at frame' },
      true,
    )
    expect(line.split('\n')).toHaveLength(1)
    expect(JSON.parse(line)).toEqual({
      status: 'error',
      summary: 'Export failed: boom at frame',
      error: 'boom at frame',
    })
  })
})

describe('headlessExitCode', () => {
  it('maps outcomes to the genoffice exit-code convention', () => {
    expect(headlessExitCode({ ok: true, input: '/a', outPath: '/b' })).toBe(0)
    expect(headlessExitCode({ ok: false, code: HEADLESS_EXIT.badArgs, message: 'x' })).toBe(1)
    expect(headlessExitCode({ ok: false, code: HEADLESS_EXIT.inputError, message: 'x' })).toBe(2)
    expect(
      headlessExitCode({ ok: false, code: HEADLESS_EXIT.conversionFailure, message: 'x' }),
    ).toBe(3)
  })
})

describe('pollUntilReady', () => {
  it('returns as soon as the predicate is satisfied', async () => {
    let ticks = 0
    await pollUntilReady(() => ++ticks >= 3, 'never ready', { pollMs: 1, timeoutMs: 1000 })
    expect(ticks).toBe(3)
  })

  it('gives up with the stall message rather than hanging forever', async () => {
    await expect(
      pollUntilReady(() => false, 'no document opened', { pollMs: 1, timeoutMs: 20 }),
    ).rejects.toThrow('no document opened')
  })

  it('evaluates a lazy stall message at give-up time', async () => {
    let state = 'loading'
    const failed = pollUntilReady(
      () => false,
      () => `stuck while ${state}`,
      {
        pollMs: 5,
        timeoutMs: 30,
      },
    )
    state = 'paginating'
    await expect(failed).rejects.toThrow('stuck while paginating')
  })
})

describe('runHeadlessRendererExport', () => {
  it('reports ok when the export wrote a file', async () => {
    const exportPdf = vi.fn(() => Promise.resolve(true))
    await expect(
      runHeadlessRendererExport('/out.pdf', () => Promise.resolve(), exportPdf),
    ).resolves.toEqual({ ok: true })
    expect(exportPdf).toHaveBeenCalledWith('/out.pdf')
  })

  it('reports an error when the export produced nothing', async () => {
    await expect(
      runHeadlessRendererExport(
        '/out.pdf',
        () => Promise.resolve(),
        () => Promise.resolve(false),
      ),
    ).resolves.toEqual({ ok: false, error: 'export did not produce a file' })
  })

  it('never throws: a failed wait becomes a report', async () => {
    const exportPdf = vi.fn(() => Promise.resolve(true))
    await expect(
      runHeadlessRendererExport(
        '/out.pdf',
        () => Promise.reject(new Error('no document opened')),
        exportPdf,
      ),
    ).resolves.toEqual({ ok: false, error: 'no document opened' })
    expect(exportPdf).not.toHaveBeenCalled()
  })

  it('never throws: a failed export becomes a report', async () => {
    await expect(
      runHeadlessRendererExport(
        '/out.pdf',
        () => Promise.resolve(),
        () => Promise.reject(new Error('printToPDF failed')),
      ),
    ).resolves.toEqual({ ok: false, error: 'printToPDF failed' })
  })
})
