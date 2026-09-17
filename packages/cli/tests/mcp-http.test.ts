import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { createServer, type Server } from 'node:http'
import { join, resolve } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { defaultRegistry } from '../src/cli'
import { fetchToFile, FileStore, safeName } from '../src/mcp/files'
import { startHttp, type HttpHandle } from '../src/mcp/http'
import { defaultOut } from '../src/mcp/remote'
import { createContext, disposeContext } from '../src/mcp/run'
import { resolveTools } from '../src/mcp/tools'
import { tempDir } from './helpers'

const REPO = resolve(__dirname, '../../..')
const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')
const TOKEN = 'secret-token'

const registry = defaultRegistry()

type Content = {
  type: string
  text?: string
  uri?: string
  resource?: { uri: string; blob?: string }
}

describe('genoffice mcp --http', () => {
  let handle: HttpHandle
  let client: Client
  let dir: string
  const headers = { authorization: `Bearer ${TOKEN}` }

  beforeAll(async () => {
    dir = tempDir()
    handle = await startHttp({
      port: 0,
      host: '127.0.0.1',
      token: TOKEN,
      cwd: dir,
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_ALLOWED_ROOTS: '' },
      log: () => {},
      registry,
    })
    client = new Client({ name: 'test', version: '0' })
    await client.connect(
      new StreamableHTTPClientTransport(new URL(`${handle.url}/mcp`), {
        requestInit: { headers },
      }),
    )
  })

  afterAll(async () => {
    await client.close().catch(() => undefined)
    await handle.close()
  })

  const call = async (name: string, args: Record<string, unknown>) => {
    const r = await client.callTool({ name, arguments: args })
    const content = r.content as Content[]
    const text = content.find((c) => c.type === 'text')?.text ?? ''
    return { isError: r.isError === true, content, text, json: () => JSON.parse(text) }
  }

  const upload = async (name: string, bytes: Buffer): Promise<string> => {
    const res = await fetch(`${handle.url}/files/${name}`, {
      method: 'PUT',
      headers,
      body: new Uint8Array(bytes),
    })
    expect(res.status).toBe(201)
    const body = (await res.json()) as { url: string; name: string; size: number }
    expect(body).toMatchObject({ name, size: bytes.byteLength })
    return body.url
  }

  it('refuses every route but /health without the bearer token', async () => {
    expect((await fetch(`${handle.url}/health`)).status).toBe(200)
    expect((await fetch(`${handle.url}/files/x.docx`, { method: 'PUT', body: 'x' })).status).toBe(
      401,
    )
    const mcp = await fetch(`${handle.url}/mcp`, {
      method: 'POST',
      headers: { authorization: 'Bearer wrong', 'content-type': 'application/json' },
      body: '{}',
    })
    expect(mcp.status).toBe(401)
  })

  it('tells the model it is remote, drops open, and makes out optional', async () => {
    expect(client.getInstructions()).toContain(`curl -T report.docx ${handle.url}/files/`)
    const { tools } = await client.listTools()
    const names = tools.map((t) => t.name)
    expect(names).toContain('docs_apply')
    expect(names).not.toContain('open')
    const create = tools.find((t) => t.name === 'create_docx')!
    expect(create.inputSchema.required ?? []).not.toContain('out')
    const props = create.inputSchema.properties as Record<string, { description?: string }>
    expect(props.out?.description).toContain('download URL')
    const info = tools.find((t) => t.name === 'info')!
    const infoProps = info.inputSchema.properties as Record<string, { description?: string }>
    expect(infoProps.file?.description).toContain('http(s) URL')
  })

  it('reads an uploaded file through its URL', async () => {
    const url = await upload('report.docx', readFileSync(DOCX))
    expect(url).toMatch(new RegExp(`^${handle.url}/files/[a-f0-9]{16}/report\\.docx$`))
    const r = await call('info', { file: url })
    expect(r.isError).toBe(false)
    expect(r.json()).toMatchObject({ status: 'ok', command: 'info', detail: { format: 'docx' } })
  })

  it('edits an uploaded file and hands the result back as a download and a resource', async () => {
    const url = await upload('edit.docx', readFileSync(DOCX))
    const before = await call('docs_read', { file: url, range: '0', full: true })
    const first = (before.json().detail.items[0].text as string).split(' ')[0]!
    const r = await call('docs_apply', {
      file: url,
      ops: [{ op: 'findReplace', find: first, replace: 'GENOFFICE' }],
    })
    expect(r.isError).toBe(false)
    const ok = r.json()
    expect(ok.output_url).toMatch(/\/files\/[a-f0-9]{16}\/edit\.docx$/)
    const resource = r.content.find((c) => c.type === 'resource')!
    expect(resource.resource?.uri).toBe(ok.output_url)
    expect(resource.resource?.blob).toBeTruthy()

    const download = await fetch(ok.output_url, { headers })
    expect(download.status).toBe(200)
    expect(download.headers.get('content-type')).toContain('wordprocessingml')
    const bytes = Buffer.from(await download.arrayBuffer())
    expect(bytes.equals(Buffer.from(resource.resource!.blob!, 'base64'))).toBe(true)
    const local = join(dir, 'edited.docx')
    writeFileSync(local, bytes)
    const check = await call('docs_read', { file: local, range: '0', full: true })
    expect(check.isError).toBe(true)
    expect(check.json().error).toBe('outside_allowed_roots')
  })

  it('creates a document without an out path and returns the file', async () => {
    const r = await call('create_docx', { markdown: '# Remote\n\nHello from far away.' })
    expect(r.isError).toBe(false)
    const ok = r.json()
    expect(ok.output_path).toMatch(/document\.docx$/)
    expect(ok.output_url).toMatch(/\/document\.docx$/)
    expect(r.content.some((c) => c.type === 'resource')).toBe(true)
    const again = await call('docs_read', { file: ok.output_url, range: '0', full: true })
    expect(again.json().detail.items[0].text).toContain('Remote')
  })

  it('reports an unreachable URL as a file error', async () => {
    const r = await call('info', { file: `${handle.url}/files/0123456789abcdef/gone.docx` })
    expect(r.isError).toBe(true)
    expect(r.json()).toMatchObject({ error: 'file_not_found' })
  })

  it('keeps sessions apart: a relative deck folder is private to its session', async () => {
    const other = new Client({ name: 'other', version: '0' })
    await other.connect(
      new StreamableHTTPClientTransport(new URL(`${handle.url}/mcp`), {
        requestInit: { headers },
      }),
    )
    try {
      const start = await client.callTool({
        name: 'deck_start',
        arguments: {
          dir: 'deck',
          style: '# Style\n\nPalette: #112233 on white.',
          outline: {
            core_hook: 'One page, kept private',
            pages: [
              {
                title: 'Cover',
                type: 'cover',
                layout: 'cover_typography_hero',
                brief:
                  'Cover: the title and the one 2024 figure from the brief, source line under it.',
                image_queries: [],
              },
            ],
          },
        },
      })
      const startText = (start.content as Content[])[0]!.text!
      expect(JSON.parse(startText).status).not.toBe('error')
      const peek = await other.callTool({ name: 'deck_build', arguments: { dir: 'deck' } })
      const peekText = (peek.content as Content[])[0]!.text!
      expect(JSON.parse(peekText)).toMatchObject({ status: 'error', error: 'missing_argument' })
    } finally {
      await other.close()
    }
  })

  it('answers /health and 404s elsewhere', async () => {
    const health = (await (await fetch(`${handle.url}/health`)).json()) as { sessions: number }
    expect(health.sessions).toBeGreaterThanOrEqual(1)
    expect((await fetch(`${handle.url}/nope`, { headers })).status).toBe(404)
  })
})

describe('http mode with GENOFFICE_ALLOWED_ROOTS set', () => {
  it('still lets the tools read uploads', async () => {
    const roots = tempDir()
    const handle = await startHttp({
      port: 0,
      cwd: roots,
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_ALLOWED_ROOTS: roots },
      log: () => {},
      registry,
    })
    const client = new Client({ name: 'roots', version: '0' })
    try {
      await client.connect(new StreamableHTTPClientTransport(new URL(`${handle.url}/mcp`)))
      const res = await fetch(`${handle.url}/files/r.docx`, {
        method: 'PUT',
        body: new Uint8Array(readFileSync(DOCX)),
      })
      const { url } = (await res.json()) as { url: string }
      const r = await client.callTool({ name: 'info', arguments: { file: url } })
      expect(r.isError ?? false).toBe(false)
      const local = join(roots, 'local.docx')
      writeFileSync(local, readFileSync(DOCX))
      const l = await client.callTool({ name: 'info', arguments: { file: local } })
      expect(l.isError ?? false).toBe(false)
    } finally {
      await client.close().catch(() => undefined)
      await handle.close()
    }
  })
})

describe('remote defaults', () => {
  it('fills out for every tool that writes PNGs or creates a file', () => {
    const ctx = createContext({ cwd: tempDir(), env: process.env, log: () => {}, mode: 'http' })
    try {
      const tools = new Map(resolveTools(registry).map((t) => [t.name, t]))
      for (const name of ['render', 'slides_render']) {
        const args: Record<string, unknown> = { file: 'x.pptx' }
        defaultOut(tools.get(name)!, args, ctx)
        expect(String(args.out)).toMatch(new RegExp(`^${ctx.scratchDir}/render-`))
      }
      const created: Record<string, unknown> = { from: 'notes.md' }
      defaultOut(tools.get('create_pdf')!, created, ctx)
      expect(String(created.out)).toMatch(/\/notes\.pdf$/)
      const kept: Record<string, unknown> = { file: 'a.docx', out: 'b.docx' }
      defaultOut(tools.get('docs_apply')!, kept, ctx)
      expect(kept.out).toBe('b.docx')
      const inPlace: Record<string, unknown> = { file: 'a.docx' }
      defaultOut(tools.get('docs_apply')!, inPlace, ctx)
      expect(inPlace.out).toBeUndefined()
    } finally {
      disposeContext(ctx)
    }
  })
})

describe('file store and fetch guard', () => {
  it('sanitises names and expires idle entries', () => {
    expect(safeName('../../etc/passwd')).toBe('passwd')
    expect(safeName('C:\\Users\\me\\Q3 report (final).docx')).toBe('Q3 report (final).docx')
    expect(safeName('   ', 'upload.bin')).toBe('upload.bin')
    const store = new FileStore(join(tempDir(), 'store'), 1000)
    const target = store.uploadTarget('a.txt')
    writeFileSync(target, 'hi')
    const stored = store.expose(target)
    expect(store.expose(target).id).toBe(stored.id)
    expect(store.resolveOwnUrl(`http://h:1/files/${stored.id}/a.txt`)).toBe(target)
    expect(store.resolveOwnUrl(`http://h:1/files/${stored.id}`)).toBeUndefined()
    store.sweep(Date.now() + 5000)
    expect(store.get(stored.id)).toBeUndefined()
    expect(existsSync(target)).toBe(false)
    store.dispose()
  })

  it('downloads by IP and by host name, keeping the remote name, and refuses link-local hosts', async () => {
    const server: Server = createServer((req, res) => {
      if (req.url === '/redirect') {
        res.writeHead(302, { location: '/dir/data.csv' })
        res.end()
        return
      }
      res.writeHead(200, { 'content-type': 'text/csv' })
      res.end('a,b\n1,2\n')
    })
    await new Promise<void>((r) => server.listen(0, r))
    const port = (server.address() as { port: number }).port
    try {
      const dir = tempDir()
      const path = await fetchToFile(`http://127.0.0.1:${port}/redirect`, dir)
      expect(path.endsWith('/data.csv')).toBe(true)
      expect(readFileSync(path, 'utf8')).toBe('a,b\n1,2\n')
      const byName = await fetchToFile(`http://localhost:${port}/dir/data.csv`, dir)
      expect(readFileSync(byName, 'utf8')).toBe('a,b\n1,2\n')
      for (const host of ['169.254.169.254', '[::ffff:a9fe:a9fe]', '[fd00:ec2:0:0:0:0:0:254]']) {
        await expect(fetchToFile(`http://${host}/latest/meta-data`, dir)).rejects.toThrow(
          'refusing to fetch',
        )
      }
      await expect(fetchToFile(`ftp://127.0.0.1/x`, dir)).rejects.toThrow('only http(s)')
      await expect(
        fetchToFile(`http://127.0.0.1:${port}/big.csv`, dir, { maxBytes: 4 }),
      ).rejects.toThrow('exceeds')
    } finally {
      server.close()
    }
  })
})
