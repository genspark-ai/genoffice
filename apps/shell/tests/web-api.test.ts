import { describe, expect, it, vi } from 'vitest'
import { mkdtemp, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { CliRunner, CliRunOutcome } from '../src/main/mcp/cli-runner'

vi.mock('electron', () => ({
  app: { getAppPath: () => '/tmp/nowhere/apps/shell', isPackaged: false },
}))
vi.mock('../../docs/src/main/lazy-media', () => ({ readLazyMedia: async () => null }))
vi.mock('../../docs/src/main/byte-handoff', () => ({ takeHandoff: () => null }))
vi.mock('../../markdown/src/main/markdown-main', () => ({
  serveMarkdownAsset: async () => new Response(null, { status: 403 }),
}))
vi.mock('../../html/src/main/html-main', () => ({
  serveHtmlAsset: async () => new Response(null),
  previewEntryFor: () => null,
}))
vi.mock('../../html/src/main/preview-document', () => ({ buildPreviewDocument: (t: string) => t }))

import { handleApiRequest, type WebApiDeps } from '../src/web/api'

function fakeRunner(fn: (args: string[]) => CliRunOutcome, calls: string[][] = []): CliRunner {
  return {
    run: async (args: string[]) => {
      calls.push(args)
      return fn(args)
    },
  }
}

function deps(
  fn: (args: string[]) => CliRunOutcome,
  calls: string[][] = [],
  dirs: string[] = [],
): WebApiDeps {
  return {
    runner: fakeRunner(fn, calls),
    workDir: async () => {
      const dir = await mkdtemp(join(tmpdir(), 'webapi-test-'))
      dirs.push(dir)
      return dir
    },
  }
}

describe('web conversion API', () => {
  it('converts an upload through the CLI and returns the file', async () => {
    const calls: string[][] = []
    const dirs: string[] = []
    const pdf = Buffer.from('%PDF-1.7 fake pdf bytes')
    const d = deps(
      (args) => {
        // the fake "CLI" writes the output the runner promised
        const out = args[args.indexOf('--out') + 1]
        void writeFile(out, pdf)
        return {
          ok: true,
          code: 0,
          json: { status: 'ok', command: 'convert', summary: 'ok' },
          stdout: '',
        }
      },
      calls,
      dirs,
    )
    const result = await handleApiRequest(
      d,
      'POST',
      '/api/convert?to=pdf&name=report.docx',
      Buffer.from('docx-bytes'),
    )
    expect(result.status).toBe(200)
    expect(result.body).toEqual(pdf)
    expect(result.disposition).toContain('report.pdf')
    expect(calls[0][0]).toBe('convert')
    expect(calls[0].join(' ')).toContain('--force --json')
    // the temp work dir is gone after the request
    await new Promise((r) => setTimeout(r, 30))
    await expect(readdir(dirs[0])).rejects.toThrow()
    void rm(dirs[0], { recursive: true, force: true }).catch(() => {})
  })

  it('maps the CLI exit-code convention onto HTTP statuses', async () => {
    const d = deps(() => ({ ok: false, code: 3, stdout: '', stderr: 'boom' }))
    const result = await handleApiRequest(
      d,
      'POST',
      '/api/convert?to=xlsx&name=a.csv',
      Buffer.from('x'),
    )
    expect(result.status).toBe(422)
    const payload = JSON.parse(String(result.body)) as { status: string }
    expect(payload.status).toBe('error')
  })

  it('passes an info upload through and returns the JSON envelope', async () => {
    const envelope = JSON.stringify({ status: 'ok', command: 'info', summary: '1 sheet' })
    const d = deps(() => ({
      ok: true,
      code: 0,
      json: { status: 'ok', command: 'info', summary: '1 sheet' },
      stdout: envelope,
    }))
    const result = await handleApiRequest(d, 'POST', '/api/info?name=book.xlsx', Buffer.from('xx'))
    expect(result.status).toBe(200)
    expect(String(result.body)).toContain('"1 sheet"')
  })

  it('rejects non-POST, missing name, extensionless names and empty bodies', async () => {
    const d = deps(() => ({ ok: true, code: 0, stdout: '' }))
    expect(
      (await handleApiRequest(d, 'GET', '/api/convert?to=pdf&name=a.docx', Buffer.from('x')))
        .status,
    ).toBe(405)
    expect(
      (await handleApiRequest(d, 'POST', '/api/convert?to=pdf', Buffer.from('x'))).status,
    ).toBe(400)
    expect(
      (await handleApiRequest(d, 'POST', '/api/convert?to=pdf&name=noext', Buffer.from('x')))
        .status,
    ).toBe(400)
    expect(
      (await handleApiRequest(d, 'POST', '/api/convert?to=pdf&name=a.docx', Buffer.alloc(0)))
        .status,
    ).toBe(400)
  })

  it('strips path components from the client-supplied name', async () => {
    const calls: string[][] = []
    const d = deps(
      () => ({
        ok: true,
        code: 0,
        json: { status: 'ok', command: 'info', summary: 'ok' },
        stdout: '',
        stderr: '',
      }),
      calls,
    )
    await handleApiRequest(
      d,
      'POST',
      '/api/info?name=' + encodeURIComponent('../../etc/shadow.txt'),
      Buffer.from('x'),
    )
    expect(calls[0][1]).not.toContain('..')
    expect(calls[0][1].startsWith(join('tmp', ''))).toBe(false)
    expect(calls[0][1].endsWith('shadow.txt')).toBe(true)
  })

  it('enforces the upload cap', async () => {
    const d = deps(() => ({ ok: true, code: 0, stdout: '' }))
    const big = Buffer.alloc(201 * 1024 * 1024, 0)
    const result = await handleApiRequest(d, 'POST', '/api/info?name=a.xlsx', big)
    expect(result.status).toBe(413)
  })

  it('answers unknown api paths and leaves non-api paths alone', async () => {
    const d = deps(() => ({ ok: true, code: 0, stdout: '' }))
    expect(
      (await handleApiRequest(d, 'POST', '/api/other?name=a.md', Buffer.from('x'))).status,
    ).toBe(404)
  })
})
