import { describe, expect, it, vi } from 'vitest'
import { resolve } from 'node:path'

import {
  runHeadlessExport,
  validateHeadlessPaths,
  type HeadlessExporters,
} from '../src/main/headless-export'
import type { HeadlessExportRequest } from '@genoffice/electron-utils'

/**
 * The `--headless-export` host (src/main/headless-export.ts): input checks,
 * extension routing and the exit-code mapping. The per-module exporters are
 * injected, so no Electron window is ever created here.
 */

const request = (
  input: string,
  outPath = '/out/a.pdf',
  targetFormat: HeadlessExportRequest['targetFormat'] = 'pdf',
): HeadlessExportRequest => ({
  input,
  targetFormat,
  outPath,
  json: false,
})

/** The host resolves caller input, so stubs and expectations must use the platform's path shape. */
const p = (path: string): string => resolve(path)

/** Every path in `present` exists and is a file; everything else does not. */
const fsWith = (present: readonly string[], isFile = true) => ({
  exists: (path: string) => present.includes(path),
  isFile: () => isFile,
})

function stubExporters(): { exporters: HeadlessExporters; calls: string[] } {
  const calls: string[] = []
  const make = (name: string) => async (input: string, outPath: string, format: string) => {
    calls.push(`${name}:${input}->${outPath}:${format}`)
  }
  return {
    calls,
    exporters: {
      docs: make('docs'),
      sheets: make('sheets'),
      slides: make('slides'),
      markdown: make('markdown'),
      html: make('html'),
    },
  }
}

describe('validateHeadlessPaths', () => {
  it('routes a readable input to its module', () => {
    const result = validateHeadlessPaths(
      request('/docs/a.docx'),
      fsWith([p('/docs/a.docx'), p('/out')]),
    )
    expect(result).toEqual({
      ok: true,
      input: p('/docs/a.docx'),
      outPath: p('/out/a.pdf'),
      module: 'docs',
    })
  })

  it('rejects a target the input module does not render as bad args (exit 1)', () => {
    const fs = fsWith([p('/docs/a.docx'), p('/decks/a.pptx'), p('/out')])
    expect(
      validateHeadlessPaths(request('/decks/a.pptx', '/out/a.docx', 'docx'), fs),
    ).toMatchObject({ ok: false, code: 1, message: expect.stringContaining('pdf') })
    expect(validateHeadlessPaths(request('/docs/a.docx', '/out/a.html', 'html'), fs)).toMatchObject(
      { ok: true, module: 'docs' },
    )
  })

  it('reports a missing input as an input-file error (exit 2)', () => {
    expect(validateHeadlessPaths(request('/nope.docx'), fsWith([]))).toEqual({
      ok: false,
      code: 2,
      message: expect.stringContaining(p('/nope.docx')),
    })
  })

  it('rejects a directory handed in as the input', () => {
    const result = validateHeadlessPaths(
      request('/docs/a.docx'),
      fsWith([p('/docs/a.docx'), p('/out')], false),
    )
    expect(result).toMatchObject({ ok: false, code: 2, message: expect.stringContaining('file') })
  })

  it('rejects an extension no module can render', () => {
    const result = validateHeadlessPaths(
      request('/docs/a.pdf'),
      fsWith([p('/docs/a.pdf'), p('/out')]),
    )
    expect(result).toMatchObject({ ok: false, code: 2 })
  })

  it('rejects an output path that is the input file, also through a symlink', () => {
    const fs = {
      ...fsWith([p('/docs/a.docx'), p('/docs'), p('/link/a.docx'), p('/link')]),
      realpath: (path: string) => path.replace(p('/link'), p('/docs')),
    }
    expect(
      validateHeadlessPaths(request('/docs/a.docx', '/docs/a.docx', 'html'), fs),
    ).toMatchObject({ ok: false, code: 1, message: expect.stringContaining('differ') })
    expect(
      validateHeadlessPaths(request('/docs/a.docx', '/link/a.docx', 'html'), fs),
    ).toMatchObject({ ok: false, code: 1, message: expect.stringContaining('differ') })
    expect(
      validateHeadlessPaths(request('/docs/a.docx', '/link/b.docx', 'html'), fs),
    ).toMatchObject({ ok: true, outPath: p('/link/b.docx') })
  })

  it('rejects an output path that is an existing directory (exit 1)', () => {
    const fs = {
      exists: (path: string) => [p('/docs/a.docx'), p('/out'), p('/out/a.pdf')].includes(path),
      isFile: (path: string) => path === p('/docs/a.docx'),
    }
    expect(validateHeadlessPaths(request('/docs/a.docx', '/out/a.pdf'), fs)).toMatchObject({
      ok: false,
      code: 1,
      message: expect.stringContaining('directory'),
    })
  })

  it('treats a missing output directory as a bad argument (exit 1)', () => {
    const result = validateHeadlessPaths(
      request('/docs/a.docx', '/gone/a.pdf'),
      fsWith([p('/docs/a.docx')]),
    )
    expect(result).toMatchObject({
      ok: false,
      code: 1,
      message: expect.stringContaining(p('/gone')),
    })
  })
})

describe('runHeadlessExport', () => {
  it('calls the module that owns the extension and reports success', async () => {
    const { exporters, calls } = stubExporters()
    const outcome = await runHeadlessExport(
      request('/decks/a.pptx'),
      exporters,
      fsWith([p('/decks/a.pptx'), p('/out'), p('/out/a.pdf')]),
    )
    expect(calls).toEqual([`slides:${p('/decks/a.pptx')}->${p('/out/a.pdf')}:pdf`])
    expect(outcome).toEqual({
      ok: true,
      input: p('/decks/a.pptx'),
      outPath: p('/out/a.pdf'),
    })
  })

  it.each([
    ['/a.docx', 'docs'],
    ['/a.xlsx', 'sheets'],
    ['/a.pptx', 'slides'],
    ['/a.md', 'markdown'],
    ['/a.html', 'html'],
  ])('routes %s to the %s exporter', async (input, module) => {
    const { exporters, calls } = stubExporters()
    await runHeadlessExport(
      request(input),
      exporters,
      fsWith([p(input), p('/out'), p('/out/a.pdf')]),
    )
    expect(calls[0]?.startsWith(`${module}:`)).toBe(true)
  })

  it('turns a thrown exporter error into a conversion failure (exit 3)', async () => {
    const { exporters } = stubExporters()
    exporters.docs = vi.fn(() => Promise.reject(new Error('renderer stopped')))
    const outcome = await runHeadlessExport(
      request('/a.docx'),
      exporters,
      fsWith([p('/a.docx'), p('/out')]),
    )
    expect(outcome).toEqual({ ok: false, code: 3, message: 'renderer stopped' })
  })

  it('fails when the exporter resolves but wrote nothing', async () => {
    const { exporters } = stubExporters()
    const outcome = await runHeadlessExport(
      request('/a.docx'),
      exporters,
      fsWith([p('/a.docx'), p('/out')]),
    )
    expect(outcome).toMatchObject({
      ok: false,
      code: 3,
      message: expect.stringContaining(p('/out/a.pdf')),
    })
  })

  it('fails when the exporter resolves but left a directory at the output path', async () => {
    const { exporters } = stubExporters()
    const present = [p('/a.docx'), p('/out')]
    exporters.docs = async () => {
      present.push(p('/out/a.pdf'))
    }
    const outcome = await runHeadlessExport(request('/a.docx'), exporters, {
      exists: (path) => present.includes(path),
      isFile: (path) => path === p('/a.docx'),
    })
    expect(outcome).toMatchObject({ ok: false, code: 3 })
  })

  it('fails when the exporter resolves but the output predates the run', async () => {
    const { exporters } = stubExporters()
    const fs = fsWith([p('/a.docx'), p('/out'), p('/out/a.pdf')])
    const fresh = await runHeadlessExport(request('/a.docx'), exporters, {
      ...fs,
      mtimeMs: () => Date.now(),
    })
    expect(fresh).toMatchObject({ ok: true })
    const stale = await runHeadlessExport(request('/a.docx'), exporters, {
      ...fs,
      mtimeMs: () => Date.now() - 60_000,
    })
    expect(stale).toMatchObject({ ok: false, code: 3, message: expect.stringContaining('stale') })
  })

  it('never reaches an exporter when the input is unusable', async () => {
    const { exporters, calls } = stubExporters()
    const outcome = await runHeadlessExport(request('/gone.docx'), exporters, fsWith([]))
    expect(calls).toEqual([])
    expect(outcome).toMatchObject({ ok: false, code: 2 })
  })
})
