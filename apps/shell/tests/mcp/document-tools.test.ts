import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { McpServerService } from '../../src/main/mcp/mcp-server'
import {
  createDocumentTools,
  documentDriver,
  resolveTargetPath,
  sanitizeFileBase,
} from '../../src/main/mcp/tools/document-tools'
import { createSessionHost, createSessionTools } from '../../src/main/mcp/tools/session-tools'
import { fakeCli } from './fake-cli'

/**
 * The docx tool surface over a real MCP session. Files are written to a temp dir
 * that stands in for the app default save folder. The headless tools delegate to
 * the bundled genoffice CLI (a fake here), so these assert the argv/stdin the
 * tools build — document fidelity is the CLI's own test surface.
 */

let service: McpServerService | undefined
let dir: string
let client: Client | undefined
let opened: string[]
let cli: ReturnType<typeof fakeCli>

async function freePort(): Promise<number> {
  const { createServer } = await import('node:http')
  return new Promise((resolve, reject) => {
    const server = createServer()
    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : 0
      server.close(() => resolve(port))
    })
  })
}

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-'))
  opened = []
  cli = fakeCli()
  const port = await freePort()
  service = new McpServerService({
    port,
    tools: createDocumentTools(
      {
        version: '0.9.0-test',
        defaultSaveDir: () => dir,
        openInTab: (filePath) => {
          opened.push(filePath)
        },
        cli: cli.runner,
      },
      createSessionHost(),
    ),
  })
  await service.start()
  client = new Client({ name: 'm3-test', version: '0.0.0' })
  await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))
})

afterEach(async () => {
  await client?.close()
  client = undefined
  await service?.stop()
  service = undefined
  await rm(dir, { recursive: true, force: true })
})

function text(content: unknown): string {
  const arr = content as Array<{ type: string; text?: string }>
  return arr.map((c) => c.text ?? '').join('')
}

describe('M3 docx tools', () => {
  it('exposes exactly the phase-1 tools', async () => {
    const { tools } = await client!.listTools()
    expect(tools.map((t) => t.name).sort()).toEqual([
      'create_docx',
      'get_app_info',
      'open_in_genoffice',
      'read_docx',
    ])
  })

  it('create_docx delegates to the CLI with the markdown staged on disk', async () => {
    const result = await client!.callTool({
      name: 'create_docx',
      arguments: { title: 'Meeting Notes', content: '# Agenda\n\n- item one\n- item two' },
    })
    expect(result.isError).toBeFalsy()

    const payload = JSON.parse(text(result.content)) as { path: string }
    expect(existsSync(payload.path)).toBe(true)
    expect(payload.path.startsWith(dir)).toBe(true)
    expect(payload.path.endsWith('.docx')).toBe(true)

    // `create --type docx --from <staged .md> --out <path>`
    const args = cli.last()
    expect(args.slice(0, 4)).toEqual(['create', '--type', 'docx', '--from'])
    expect(args[4]!.endsWith('.md')).toBe(true)
    expect(cli.lastFrom()).toContain('# Agenda')
  })

  it('create_docx accepts restricted HTML (format:"html")', async () => {
    const result = await client!.callTool({
      name: 'create_docx',
      arguments: { title: 'Web', content: '<h1>Title</h1>', format: 'html' },
    })
    expect(result.isError).toBeFalsy()
    const args = cli.last()
    expect(args[4]!.endsWith('.html')).toBe(true)
    expect(cli.lastFrom()).toBe('<h1>Title</h1>')
  })

  it('create_docx avoids clobbering an existing file', async () => {
    const first = await client!.callTool({
      name: 'create_docx',
      arguments: { title: 'Report', content: 'one' },
    })
    const second = await client!.callTool({
      name: 'create_docx',
      arguments: { title: 'Report', content: 'two' },
    })
    const p1 = JSON.parse(text(first.content)).path as string
    const p2 = JSON.parse(text(second.content)).path as string
    expect(p1).not.toBe(p2)
    expect(p2).toContain('Report-2')
  })

  it('create_docx rejects overwriting an explicit path unless asked', async () => {
    const target = join(dir, 'explicit.docx')
    await writeFile(target, 'placeholder')

    const refused = await client!.callTool({
      name: 'create_docx',
      arguments: { title: 'x', content: 'hi', path: target },
    })
    expect(refused.isError).toBe(true)
    expect(text(refused.content)).toContain('already exists')

    const allowed = await client!.callTool({
      name: 'create_docx',
      arguments: { title: 'x', content: 'hi', path: target, overwrite: true },
    })
    expect(allowed.isError).toBeFalsy()
    // overwrite maps to the CLI's --force
    expect(cli.last()).toContain('--force')
  })

  it('surfaces a CLI failure as a tool error', async () => {
    cli.failNext('markdown rejected')
    const result = await client!.callTool({
      name: 'create_docx',
      arguments: { title: 'Bad', content: '# x' },
    })
    expect(result.isError).toBe(true)
    expect(text(result.content)).toContain('markdown rejected')
  })

  it('read_docx returns the CLI-read text', async () => {
    await writeFile(join(dir, 'doc.docx'), 'placeholder')
    cli.setReadItems([
      { index: 0, type: 'heading', text: 'Title' },
      { index: 1, text: 'Body text here.' },
    ])
    const result = await client!.callTool({
      name: 'read_docx',
      arguments: { path: join(dir, 'doc.docx') },
    })
    expect(result.isError).toBeFalsy()
    const payload = JSON.parse(text(result.content)) as { text: string }
    expect(payload.text).toBe('Title\nBody text here.')
    // --full matters: without it the CLI clips every block to a preview
    expect(cli.last()).toEqual(['docs', 'read', join(dir, 'doc.docx'), '--full'])
  })

  it('read_docx rejects a missing file and a non-docx extension', async () => {
    const missing = await client!.callTool({
      name: 'read_docx',
      arguments: { path: join(dir, 'nope.docx') },
    })
    expect(missing.isError).toBe(true)

    const wrongExt = await client!.callTool({
      name: 'read_docx',
      arguments: { path: join(dir, 'x.txt') },
    })
    expect(wrongExt.isError).toBe(true)
  })

  it('open_in_genoffice calls the injected opener', async () => {
    await writeFile(join(dir, 'open.docx'), 'placeholder')
    const result = await client!.callTool({
      name: 'open_in_genoffice',
      arguments: { path: join(dir, 'open.docx') },
    })
    expect(result.isError).toBeFalsy()
    expect(opened).toEqual([join(dir, 'open.docx')])
  })

  it('get_app_info reports version, save dir and formats', async () => {
    const result = await client!.callTool({ name: 'get_app_info', arguments: {} })
    const payload = JSON.parse(text(result.content)) as {
      version: string
      defaultSaveDir: string
      formats: string[]
    }
    expect(payload.version).toBe('0.9.0-test')
    expect(payload.defaultSaveDir).toBe(dir)
    expect(payload.formats).toEqual(['docx'])
  })
})

describe('background generation gating', () => {
  it('hides create_docx when background is off, keeps the read/open/info tools', () => {
    const names = createDocumentTools(
      {
        version: 'x',
        defaultSaveDir: () => dir,
        background: false,
      },
      createSessionHost(),
    ).map((t) => t.name)
    expect(names).not.toContain('create_docx')
    expect(names).toContain('read_docx')
    expect(names).toContain('open_in_genoffice')
    expect(names).toContain('get_app_info')
  })

  it('exposes create_docx by default and when background is on', () => {
    const base = { version: 'x', defaultSaveDir: () => dir }
    const withHost = (extra: Record<string, unknown> = {}) =>
      createDocumentTools({ ...base, ...extra }, createSessionHost()).map((t) => t.name)
    expect(withHost()).toContain('create_docx')
    expect(withHost({ background: true })).toContain('create_docx')
  })
})

describe('path policy helpers', () => {
  it('sanitizes illegal filename characters', () => {
    expect(sanitizeFileBase('a/b:c*d?')).toBe('a_b_c_d_')
    expect(sanitizeFileBase('   ')).toBe('Untitled')
    expect(sanitizeFileBase('...')).toBe('Untitled')
  })

  it('rejects a relative explicit path', () => {
    expect(() =>
      resolveTargetPath({ version: 'x', defaultSaveDir: () => dir }, 't', 'relative.docx'),
    ).toThrow(/absolute/)
  })
})

/** a DocsControl that records what the tools asked it to do */
function fakeDocsControl(wcId = 42) {
  const calls: Array<{ wcId: number; command: string; payload: unknown }> = []
  let opens = 0
  const control = {
    openBlankTab: async () => {
      opens++
      return wcId
    },
    runCommand: async (id: number, command: string, payload: unknown) => {
      calls.push({ wcId: id, command, payload })
      return { summary: `${command} ok` }
    },
  }
  return { control, calls, opens: () => opens }
}

describe('M6 visible-editing tools (docs control wired)', () => {
  let dir2: string
  let service2: McpServerService | undefined
  let client2: Client | undefined
  let fake: ReturnType<typeof fakeDocsControl>

  beforeEach(async () => {
    dir2 = await mkdtemp(join(tmpdir(), 'genoffice-mcp-visible-'))
    fake = fakeDocsControl()
    const port = await (async () => {
      const { createServer } = await import('node:http')
      return new Promise<number>((resolve, reject) => {
        const server = createServer()
        server.on('error', reject)
        server.listen(0, '127.0.0.1', () => {
          const address = server.address()
          const p = typeof address === 'object' && address ? address.port : 0
          server.close(() => resolve(p))
        })
      })
    })()
    service2 = new McpServerService({
      port,
      tools: (() => {
        const host = createSessionHost()
        return [
          ...createSessionTools([documentDriver(fake.control)], host),
          ...createDocumentTools(
            {
              version: '0.9.0-test',
              defaultSaveDir: () => dir2,
              docs: fake.control,
            },
            host,
          ),
        ]
      })(),
    })
    await service2.start()
    client2 = new Client({ name: 'm6-test', version: '0.0.0' })
    await client2.connect(
      new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)),
    )
  })

  afterEach(async () => {
    await client2?.close()
    client2 = undefined
    await service2?.stop()
    service2 = undefined
    await rm(dir2, { recursive: true, force: true })
  })

  function textOf(content: unknown): string {
    const arr = content as Array<{ type: string; text?: string }>
    return arr.map((c) => c.text ?? '').join('')
  }

  it('registers the visible session tools alongside the file tools', async () => {
    const { tools } = await client2!.listTools()
    expect(tools.map((t) => t.name).sort()).toEqual([
      'apply_ops',
      'create_docx',
      'create_session',
      'get_app_info',
      'insert_content',
      'open_in_genoffice',
      'read_document',
      'read_docx',
      'replace_blocks',
      'save_session',
    ])
  })

  it('create_session opens a blank docx tab and returns a session id', async () => {
    const result = await client2!.callTool({
      name: 'create_session',
      arguments: { family: 'docx' },
    })
    expect(result.isError).toBeFalsy()
    const payload = JSON.parse(textOf(result.content)) as {
      sessionId: number
      ok: boolean
      family: string
    }
    expect(payload.ok).toBe(true)
    expect(payload.family).toBe('docx')
    expect(payload.sessionId).toBe(42)
    expect(fake.opens()).toBe(1)
  })

  it('routes insert_content into the created session tab', async () => {
    await client2!.callTool({ name: 'create_session', arguments: { family: 'docx' } })
    const result = await client2!.callTool({
      name: 'insert_content',
      arguments: { html: '<h1>Title</h1>', afterBlockIndex: 0 },
    })
    expect(result.isError).toBeFalsy()
    expect(fake.calls.at(-1)).toEqual({
      wcId: 42,
      command: 'insert_content',
      payload: { html: '<h1>Title</h1>', afterBlockIndex: 0 },
    })
  })

  it('routes replace_blocks and apply_ops with their payloads', async () => {
    await client2!.callTool({ name: 'create_session', arguments: { family: 'docx' } })
    await client2!.callTool({
      name: 'replace_blocks',
      arguments: { startBlockIndex: 0, endBlockIndex: 1, html: '<p>x</p>' },
    })
    expect(fake.calls.at(-1)?.command).toBe('replace_blocks')
    await client2!.callTool({
      name: 'apply_ops',
      arguments: { ops: [{ op: 'setFont' }], dryRun: true },
    })
    expect(fake.calls.at(-1)).toEqual({
      wcId: 42,
      command: 'apply_ops',
      payload: { ops: [{ op: 'setFont' }], dryRun: true },
    })
  })

  it('save_session forwards path + overwrite and ends the session', async () => {
    await client2!.callTool({ name: 'create_session', arguments: { family: 'docx' } })
    const target = join(dir2, 'out.docx')
    const saved = await client2!.callTool({
      name: 'save_session',
      arguments: { path: target, overwrite: true },
    })
    expect(saved.isError).toBeFalsy()
    expect(fake.calls.at(-1)).toEqual({
      wcId: 42,
      command: 'save_document',
      payload: { path: target, overwrite: true },
    })
    // session ended: a follow-up edit reports an error instead of hitting a stale tab
    const after = await client2!.callTool({
      name: 'insert_content',
      arguments: { html: '<p>late</p>' },
    })
    expect(after.isError).toBe(true)
    expect(textOf(after.content)).toContain('create_session')
  })

  it('editing before create_session is refused', async () => {
    const result = await client2!.callTool({
      name: 'insert_content',
      arguments: { html: '<p>nope</p>' },
    })
    expect(result.isError).toBe(true)
    expect(textOf(result.content)).toContain('create_session')
    expect(fake.calls).toEqual([])
  })

  it('rejects an unknown family value', async () => {
    // only the docx driver is registered here, so the schema enum excludes xlsx
    const bad = await client2!.callTool({ name: 'create_session', arguments: { family: 'xlsx' } })
    expect(bad.isError).toBe(true)
  })
})
