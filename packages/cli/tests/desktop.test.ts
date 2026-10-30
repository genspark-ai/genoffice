import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { createServer, type Server } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import type { ControlReply } from '../src/control-protocol'
import { run } from './helpers'

/**
 * The `desktop` command against a fake shell: argument parsing (verb and
 * compact forms), the wire request shape, and the error surfaces (app not
 * running, usage errors before any socket is touched).
 */

const servers: Server[] = []
const dirs: string[] = []
let pipeSeq = 0

afterEach(() => {
  while (servers.length) servers.pop()!.close()
  while (dirs.length) rmSync(dirs.pop()!, { recursive: true, force: true })
})

/** A stand-in shell: publishes control.json into a temp dir and answers with `reply`. */
async function fakeShell(
  reply: (request: unknown) => ControlReply,
): Promise<{ env: NodeJS.ProcessEnv; requests: unknown[] }> {
  const dir = mkdtempSync(join(tmpdir(), 'genoffice-cli-desktop-'))
  dirs.push(dir)
  const requests: unknown[] = []
  const endpoint =
    process.platform === 'win32'
      ? `\\\\.\\pipe\\genoffice-desktop-test-${process.pid}-${pipeSeq++}`
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
  await new Promise<void>((resolve) => server.listen(endpoint, resolve))
  writeFileSync(
    join(dir, 'control.json'),
    JSON.stringify({ protocol: 1, pid: process.pid, endpoint, token: 'secret' }),
  )
  return { env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_USER_DATA: dir }, requests }
}

describe('genoffice desktop', () => {
  it('get-state asks the shell and relays the snapshot', async () => {
    const { env, requests } = await fakeShell(() => ({
      ok: true,
      result: {
        prefs: { wallpaper: 'builtin:teal-deep' },
        layout: { positioned: 3, positions: {} },
        todos: [{ id: 't1', text: 'milk', done: false }],
        windows: [{ id: 'home', kind: 'home', title: 'GenOffice', active: true }],
      },
    }))
    const r = await run(['desktop', 'get-state', '--json'], { env })
    expect(r.code).toBe(0)
    expect(requests).toEqual([{ cmd: 'desktop', action: 'get-state' }])
    expect(r.json().status).toBe('ok')
    expect(r.json().detail.prefs.wallpaper).toBe('builtin:teal-deep')
    expect(r.json().summary).toContain('builtin:teal-deep')
    expect(r.json().summary).toContain('3 positioned icons')
    expect(r.json().summary).toContain('1/1 open todos')
  })

  it('todo add joins the words and todo list reads the widget', async () => {
    const { env, requests } = await fakeShell((req) => {
      const request = req as { action: string }
      if (request.action === 'todo-add') {
        return {
          ok: true,
          result: { added: 'buy milk', todos: [{ id: 'x', text: 'buy milk', done: false }] },
        }
      }
      return { ok: true, result: { todos: [{ id: 'x', text: 'buy milk', done: false }] } }
    })
    const added = await run(['desktop', 'todo', 'add', 'buy', 'milk', '--json'], { env })
    expect(added.code).toBe(0)
    expect(requests[0]).toEqual({ cmd: 'desktop', action: 'todo-add', text: 'buy milk' })
    expect(added.json().summary).toContain('buy milk')

    const listed = await run(['desktop', 'todo', 'list', '--json'], { env })
    expect(listed.code).toBe(0)
    expect(requests[1]).toEqual({ cmd: 'desktop', action: 'todo-list' })
    expect(listed.json().summary).toContain('[ ] buy milk')
  })

  it('set-wallpaper sends the value through (builtin name or path)', async () => {
    const { env, requests } = await fakeShell(() => ({
      ok: true,
      result: { wallpaper: 'builtin:warm-dusk' },
    }))
    const r = await run(['desktop', 'set-wallpaper', 'warm-dusk', '--json'], { env })
    expect(r.code).toBe(0)
    expect(requests).toEqual([{ cmd: 'desktop', action: 'set-wallpaper', wallpaper: 'warm-dusk' }])
    expect(r.json().summary).toContain('builtin:warm-dusk')
  })

  it('windows list prints the open tabs', async () => {
    const { env, requests } = await fakeShell(() => ({
      ok: true,
      result: {
        windows: [
          { id: 'home', kind: 'home', title: 'GenOffice', active: true },
          { id: 't2', kind: 'docs', title: 'report.docx', filePath: '/tmp/report.docx' },
        ],
      },
    }))
    const r = await run(['desktop', 'windows', 'list', '--json'], { env })
    expect(r.code).toBe(0)
    expect(requests).toEqual([{ cmd: 'desktop', action: 'windows-list' }])
    expect(r.json().detail.windows).toHaveLength(2)
    expect(r.json().summary).toContain('report.docx')
  })

  it('usage errors never touch the socket', async () => {
    const { env, requests } = await fakeShell(() => ({ ok: true, result: {} }))
    for (const argv of [
      ['desktop', '--json'],
      ['desktop', 'nuke', '--json'],
      ['desktop', 'todo', 'add', '--json'],
      ['desktop', 'set-wallpaper', '--json'],
    ]) {
      const r = await run(argv, { env })
      expect(r.code, argv.join(' ')).toBe(1)
      expect(r.json().error, argv.join(' ')).toMatch(/invalid_argument|missing_argument/)
    }
    expect(requests).toEqual([])
  })

  it('exit 4 when no app is running', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-cli-desktop-'))
    dirs.push(dir)
    const env = { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_USER_DATA: dir }
    const r = await run(['desktop', 'get-state', '--json'], { env })
    expect(r.code).toBe(4)
    expect(r.json().error).toBe('app_unavailable')
  })

  it('relays a shell-side error with its stable reason', async () => {
    const { env } = await fakeShell(() => ({
      ok: false,
      error: {
        reason: 'invalid_argument',
        message: 'unknown wallpaper: nope',
        detail: { available: ['teal-deep'] },
      },
    }))
    const r = await run(['desktop', 'set-wallpaper', 'nope', '--json'], { env })
    expect(r.code).toBe(1)
    expect(r.json().error).toBe('invalid_argument')
    expect(r.json().detail.available).toEqual(['teal-deep'])
  })
})
