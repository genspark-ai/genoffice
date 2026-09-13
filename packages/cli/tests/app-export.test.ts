import { EventEmitter } from 'node:events'
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { PassThrough } from 'node:stream'
import { describe, expect, it } from 'vitest'
import { exportViaApp, parseEnvelope } from '../src/formats/app-export'
import { appLaunch } from '../src/resources'
import { CliError } from '../src/result'
import { run, tempDir } from './helpers'

interface Script {
  stdout?: string
  stderr?: string
  code?: number | null
  hang?: boolean
  ignoreTerm?: boolean
  writeOutput?: boolean
  /** the main process exits but helpers keep the stdio pipes open: no `close` */
  exitOnly?: boolean
}

const signals: string[] = []

function fakeSpawn(script: Script, calls: { command: string; args: string[] }[]) {
  return ((command: string, args: string[]) => {
    calls.push({ command, args })
    const child = new EventEmitter() as EventEmitter & {
      stdout: PassThrough
      stderr: PassThrough
      kill: (signal?: string) => boolean
      killed: boolean
    }
    child.stdout = new PassThrough()
    child.stderr = new PassThrough()
    child.killed = false
    child.kill = (signal?: string) => {
      signals.push(signal ?? 'SIGTERM')
      if (signal === 'SIGTERM' && script.ignoreTerm) return true
      child.killed = true
      setTimeout(() => child.emit('close', null), 1)
      return true
    }
    setTimeout(() => {
      if (script.writeOutput) {
        const out = args[args.indexOf('--out') + 1]!
        writeFileSync(out, 'pdf')
      }
      if (script.stdout) child.stdout.write(script.stdout)
      if (script.stderr) child.stderr.write(script.stderr)
      if (script.hang) return
      child.emit('exit', script.code ?? 0)
      if (script.exitOnly) return
      child.stdout.end()
      child.stderr.end()
      child.emit('close', script.code ?? 0)
    }, 5)
    return child
  }) as unknown as typeof import('node:child_process').spawn
}

const env = { GENOFFICE_APP_BIN: '/Applications/GenOffice.app/Contents/MacOS/GenOffice' }

describe('exportViaApp', () => {
  it('spawns the app in headless-export mode and returns the envelope', async () => {
    const dir = tempDir()
    const out = join(dir, 'a.pdf')
    const calls: { command: string; args: string[] }[] = []
    const r = await exportViaApp('/tmp/a.docx', 'pdf', out, {
      env,
      spawn: fakeSpawn(
        {
          stdout: `log line\n{"status":"ok","summary":"Exported /tmp/a.docx to ${out}","output_path":"${out}"}\n`,
          writeOutput: true,
        },
        calls,
      ),
    })
    expect(r.outputPath).toBe(out)
    expect(calls[0]!.command).toBe(env.GENOFFICE_APP_BIN)
    expect(calls[0]!.args).toEqual([
      '--headless-export',
      '/tmp/a.docx',
      '--to',
      'pdf',
      '--out',
      out,
      '--json',
    ])
  })

  it('maps the app exit codes and surfaces its error message', async () => {
    const dir = tempDir()
    const attempt = (script: Script) =>
      exportViaApp('/tmp/a.docx', 'pdf', join(dir, 'b.pdf'), { env, spawn: fakeSpawn(script, []) })
    await expect(
      attempt({
        code: 2,
        stdout:
          '{"status":"error","summary":"Export failed: no such file","error":"no such file"}\n',
      }),
    ).rejects.toMatchObject({ code: 2, message: 'no such file' })
    await expect(attempt({ code: 3, stderr: 'renderer crashed\n' })).rejects.toMatchObject({
      code: 3,
      message: 'renderer crashed',
    })
    // exit 0 without the file on disk is still a failure
    await expect(
      attempt({ code: 0, stdout: '{"status":"ok","summary":"x"}\n' }),
    ).rejects.toBeInstanceOf(CliError)
  })

  it('terminates a hung export and escalates to SIGKILL when it ignores SIGTERM', async () => {
    const dir = tempDir()
    signals.length = 0
    await expect(
      exportViaApp('/tmp/a.docx', 'pdf', join(dir, 'c.pdf'), {
        env,
        timeoutMs: 30,
        spawn: fakeSpawn({ hang: true }, []),
      }),
    ).rejects.toMatchObject({ code: 3 })
    expect(signals).toEqual(['SIGTERM'])

    signals.length = 0
    await expect(
      exportViaApp('/tmp/a.docx', 'pdf', join(dir, 'd.pdf'), {
        env,
        timeoutMs: 30,
        killGraceMs: 20,
        spawn: fakeSpawn({ hang: true, ignoreTerm: true }, []),
      }),
    ).rejects.toMatchObject({ code: 3 })
    expect(signals).toEqual(['SIGTERM', 'SIGKILL'])
  })

  it('settles once the main process exits even if helpers keep the pipes open', async () => {
    const dir = tempDir()
    const out = join(dir, 'e.pdf')
    const started = Date.now()
    const r = await exportViaApp('/tmp/a.docx', 'pdf', out, {
      env,
      timeoutMs: 5000,
      spawn: fakeSpawn(
        { stdout: '{"status":"ok","summary":"done"}\n', writeOutput: true, exitOnly: true },
        [],
      ),
    })
    expect(r.summary).toBe('done')
    expect(Date.now() - started).toBeLessThan(2000)
  })

  it('keeps a finished export whose app then hung on quit', async () => {
    const dir = tempDir()
    const out = join(dir, 'f.pdf')
    signals.length = 0
    const r = await exportViaApp('/tmp/a.docx', 'pdf', out, {
      env,
      timeoutMs: 30,
      spawn: fakeSpawn(
        { stdout: '{"status":"ok","summary":"done"}\n', writeOutput: true, hang: true },
        [],
      ),
    })
    expect(r.outputPath).toBe(out)
    expect(signals).toEqual(['SIGTERM'])
  })

  it('launches the checkout through the real Electron binary, not the npm shim', () => {
    const launch = appLaunch({})
    expect(launch).not.toBeNull()
    expect(launch!.command).not.toContain('.bin')
    expect(launch!.command).toMatch(
      process.platform === 'darwin' ? /MacOS\/Electron$/ : /electron/i,
    )
    expect(launch!.args[0]).toMatch(/apps\/shell$/)
  })

  it('parses the last JSON line of stdout', () => {
    expect(parseEnvelope('noise\n{"status":"ok","summary":"s"}\ntrailer')).toEqual({
      status: 'ok',
      summary: 's',
    })
    expect(parseEnvelope('nothing here')).toBeNull()
    expect(parseEnvelope('{"status":"weird"}')).toBeNull()
  })

  // Opt-in end-to-end run against the checkout's Electron (needs `npm run build:all`):
  //   GENOFFICE_E2E_APP=1 npx vitest run tests/app-export.test.ts
  it.skipIf(!process.env.GENOFFICE_E2E_APP)(
    'converts a Word document to PDF through the real app',
    async () => {
      const dir = tempDir()
      const docx = join(dir, 'a.docx')
      const { copyFileSync, existsSync } = await import('node:fs')
      copyFileSync(
        join(__dirname, '../../../apps/docs/tests/pagination-corpus/docx/01-simple-english.docx'),
        docx,
      )
      const r = await run(['convert', docx, '--to', 'pdf', '--json'])
      expect(r.code).toBe(0)
      expect(existsSync(join(dir, 'a.pdf'))).toBe(true)
    },
    240_000,
  )

  it.skipIf(!process.env.GENOFFICE_E2E_APP)(
    'round-trips Word → HTML → Word through the real app',
    async () => {
      const dir = tempDir()
      const docx = join(dir, 'a.docx')
      const { copyFileSync, existsSync, readFileSync } = await import('node:fs')
      copyFileSync(
        join(__dirname, '../../../apps/docs/tests/pagination-corpus/docx/01-simple-english.docx'),
        docx,
      )
      const toHtml = await run(['convert', docx, '--to', 'html', '--json'])
      expect(toHtml.code).toBe(0)
      expect(readFileSync(join(dir, 'a.html'), 'utf8')).toContain('<html')
      const back = await run([
        'convert',
        join(dir, 'a.html'),
        '--to',
        'docx',
        '--out',
        join(dir, 'b.docx'),
        '--json',
      ])
      expect(back.code).toBe(0)
      expect(existsSync(join(dir, 'b.docx'))).toBe(true)
    },
    480_000,
  )
})
