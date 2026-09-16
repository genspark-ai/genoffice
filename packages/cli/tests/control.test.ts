import { mkdtempSync, writeFileSync } from 'node:fs'
import { createServer, type Server } from 'node:net'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { parseArgs } from '../src/args'
import { parseTarget } from '../src/commands/open'
import { controlEndpoint, controlRequest } from '../src/control'
import type { ControlReply } from '../src/control-protocol'
import { CliError } from '../src/result'
import { run } from './helpers'

const REPO = resolve(__dirname, '../../..')
const PPTX = join(REPO, 'packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')

const servers: Server[] = []
afterEach(() => {
  while (servers.length) servers.pop()!.close()
})

/** A stand-in shell: publishes control.json into `dir` and answers with `reply`. */
async function fakeShell(
  dir: string,
  reply: (request: unknown) => ControlReply,
): Promise<{ env: NodeJS.ProcessEnv; requests: unknown[] }> {
  const requests: unknown[] = []
  const endpoint =
    process.platform === 'win32'
      ? `\\\\.\\pipe\\genoffice-test-${process.pid}-${servers.length}`
      : join(dir, 'control.sock')
  const server = createServer((socket) => {
    let buffer = ''
    socket.setEncoding('utf8')
    socket.on('data', (chunk: string) => {
      buffer += chunk
      if (!buffer.includes('\n')) return
      const envelope = JSON.parse(buffer.slice(0, buffer.indexOf('\n')))
      if (envelope.token !== 'secret') return socket.destroy()
      requests.push(envelope.request)
      socket.end(JSON.stringify(reply(envelope.request)) + '\n')
    })
  })
  servers.push(server)
  await new Promise<void>((r) => server.listen(endpoint, r))
  writeFileSync(
    join(dir, 'control.json'),
    JSON.stringify({ protocol: 1, pid: process.pid, endpoint, token: 'secret' }),
  )
  return { env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_USER_DATA: dir }, requests }
}

describe('open targets', () => {
  const target = (argv: string[], file: string) => parseTarget(parseArgs(argv), file)

  it('maps flags to one target per file type', () => {
    expect(target([], '/d.pptx')).toBeUndefined()
    expect(target(['--slide', '2', '--el', 'e_7'], '/d.pptx')).toEqual({
      kind: 'slide',
      slide: 2,
      el: 'e_7',
    })
    expect(target(['--block', '4'], '/d.docx')).toEqual({ kind: 'block', block: 4 })
    expect(target(['--range', 'Data!B2:D5'], '/d.xlsx')).toEqual({
      kind: 'range',
      range: 'Data!B2:D5',
    })
    expect(target(['--range', 'B2', '--sheet', 'Q1'], '/d.xlsx')).toEqual({
      kind: 'range',
      range: 'B2',
      sheet: 'Q1',
    })
    expect(target(['--page', '3'], '/d.pdf')).toEqual({ kind: 'page', page: 3 })
  })

  it('refuses a flag that belongs to another file type and bad numbers', () => {
    expect(() => target(['--block', '1'], '/d.pptx')).toThrow(CliError)
    expect(() => target(['--slide', '-1'], '/d.pptx')).toThrow(/integer >= 0/)
    expect(() => target(['--page', '0'], '/d.pdf')).toThrow(/integer >= 1/)
    expect(() => target(['--el', 'e_1'], '/d.pptx')).toThrow(/--el needs --slide/)
    expect(() => target(['--page', '1'], '/notes.md')).toThrow(/pptx, docx, xlsx and pdf/)
  })
})

describe('control endpoint', () => {
  it('ignores a missing, malformed or dead-pid control.json', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ctl-'))
    expect(controlEndpoint({ GENOFFICE_USER_DATA: dir })).toBeNull()
    writeFileSync(join(dir, 'control.json'), '{')
    expect(controlEndpoint({ GENOFFICE_USER_DATA: dir })).toBeNull()
    writeFileSync(
      join(dir, 'control.json'),
      JSON.stringify({ protocol: 1, pid: 2 ** 22 - 1, endpoint: '/x', token: 't' }),
    )
    expect(controlEndpoint({ GENOFFICE_USER_DATA: dir })).toBeNull()
    writeFileSync(
      join(dir, 'control.json'),
      JSON.stringify({ protocol: 1, pid: process.pid, endpoint: '/x', token: 't' }),
    )
    expect(controlEndpoint({ GENOFFICE_USER_DATA: dir })).toMatchObject({
      pid: process.pid,
      token: 't',
    })
  })

  it('fails with app_unavailable when nobody listens', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ctl-'))
    await expect(
      controlRequest(
        { protocol: 1, pid: process.pid, endpoint: join(dir, 'missing.sock'), token: 't' },
        { cmd: 'selection', path: '/a.docx' },
        2_000,
      ),
    ).rejects.toMatchObject({ reason: 'app_unavailable', code: 4 })
  })
})

describe('open / selection through the control channel', () => {
  it('open --slide asks the running shell instead of spawning the app', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ctl-'))
    const shell = await fakeShell(dir, () => ({
      ok: true,
      result: { slide: 1, element: 'e_3', type: 'text' },
    }))
    const r = await run(['open', PPTX, '--slide', '1', '--el', 'e_3', '--json'], { env: shell.env })
    expect(r.code).toBe(0)
    expect(r.json()).toMatchObject({
      status: 'ok',
      summary: expect.stringContaining('element e_3 on slide 1'),
      detail: { slide: 1, element: 'e_3', gui_pid: process.pid },
    })
    expect(shell.requests).toEqual([
      { cmd: 'open', path: PPTX, target: { kind: 'slide', slide: 1, el: 'e_3' } },
    ])
  })

  it('relays renderer errors as structured CLI errors', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ctl-'))
    const shell = await fakeShell(dir, () => ({
      ok: false,
      error: {
        reason: 'target_not_found',
        message: 'no element e_9 on slide 0',
        detail: { available: ['e_1', 'e_2'] },
      },
    }))
    const r = await run(['open', PPTX, '--slide', '0', '--el', 'e_9', '--json'], { env: shell.env })
    expect(r.code).toBe(1)
    expect(r.json()).toMatchObject({
      status: 'error',
      error: 'target_not_found',
      detail: { available: ['e_1', 'e_2'] },
      suggestion: expect.stringContaining('slides read'),
    })
  })

  it('selection returns the editor selection and explains when the file is not open', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ctl-'))
    const shell = await fakeShell(dir, (request) =>
      (request as { path: string }).path === DOCX
        ? { ok: true, result: { blocks: [2, 2], text: 'Second paragraph', collapsed: false } }
        : {
            ok: false,
            error: {
              reason: 'file_not_open_in_gui',
              message: 'not open',
              detail: { suggestion: 'genoffice open x' },
            },
          },
    )
    const ok = await run(['selection', DOCX, '--json'], { env: shell.env })
    expect(ok.code).toBe(0)
    expect(ok.json()).toMatchObject({
      summary: 'block 2: Second paragraph',
      detail: { blocks: [2, 2], text: 'Second paragraph' },
    })
    const notOpen = await run(['selection', PPTX, '--json'], { env: shell.env })
    expect(notOpen.code).toBe(2)
    expect(notOpen.json()).toMatchObject({
      error: 'file_not_open_in_gui',
      suggestion: 'genoffice open x',
    })
  })

  it('selection without a running shell is app_unavailable with an open hint', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-ctl-'))
    const r = await run(['selection', DOCX, '--json'], {
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_USER_DATA: dir },
    })
    expect(r.code).toBe(4)
    expect(r.json()).toMatchObject({
      error: 'app_unavailable',
      suggestion: expect.stringContaining(`genoffice open ${DOCX}`),
    })
  })
})
