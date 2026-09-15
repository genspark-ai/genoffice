import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { mkdtemp, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { McpServerService } from '../../src/main/mcp/mcp-server'
import {
  createSheetsTools,
  sheetsDriver,
  type SheetsControl,
} from '../../src/main/mcp/tools/sheets-tools'
import { createSessionHost, createSessionTools } from '../../src/main/mcp/tools/session-tools'
import { fakeCli } from './fake-cli'

/**
 * Sheets tool surface over a real MCP session: headless create_xlsx (row
 * matrix, path policy, background gating, CLI delegation) and the visible grid
 * session tools driven through a fake SheetsControl.
 */

let service: McpServerService | undefined
let dir: string
let client: Client | undefined

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
  dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-xlsx-'))
})

afterEach(async () => {
  await client?.close()
  client = undefined
  await service?.stop()
  service = undefined
  await rm(dir, { recursive: true, force: true })
})

async function startService(tools: ReturnType<typeof createSheetsTools>): Promise<void> {
  const port = await freePort()
  service = new McpServerService({ port, tools })
  await service.start()
  client = new Client({ name: 'sheets-test', version: '0.0.0' })
  await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))
}

/** content tools + the shared session tools wired against the same host, as the shell does */
function sessionSurface(control: SheetsControl): ReturnType<typeof createSheetsTools> {
  const host = createSessionHost()
  return [
    ...createSessionTools([sheetsDriver(control)], host),
    ...createSheetsTools({ defaultSaveDir: () => dir, sheets: control }, host),
  ]
}

function baseDeps(background?: boolean): {
  defaultSaveDir: () => string
  background?: boolean
} {
  return { defaultSaveDir: () => dir, ...(background === undefined ? {} : { background }) }
}

function text(content: unknown): string {
  const arr = content as Array<{ type: string; text?: string }>
  return arr.map((c) => c.text ?? '').join('')
}

describe('headless create_xlsx', () => {
  it('delegates to the CLI with the staged row matrix', async () => {
    const cli = fakeCli()
    await startService(createSheetsTools({ ...baseDeps(), cli: cli.runner }, createSessionHost()))
    const result = await client!.callTool({
      name: 'create_xlsx',
      arguments: {
        title: 'Quarterly',
        data: [
          ['Item', 'Qty'],
          ['Widgets', 12],
          ['Gadgets', 3.5],
          ['Mixed, comma', 'plain'],
        ],
        sheetName: 'Data',
      },
    })
    expect(result.isError).toBeFalsy()
    const payload = JSON.parse(text(result.content)) as { path: string }
    expect(existsSync(payload.path)).toBe(true)
    expect(payload.path.endsWith('.xlsx')).toBe(true)

    // `create --type xlsx --from <staged table.json> --out <path>`; the staged
    // JSON carries the sheet name and rows with numbers kept numeric
    const args = cli.last()
    expect(args.slice(0, 4)).toEqual(['create', '--type', 'xlsx', '--from'])
    expect(args[4]!.endsWith('table.json')).toBe(true)
    const table = JSON.parse(cli.lastFrom() ?? '{}') as {
      sheets: Array<{ name: string; rows: unknown[][] }>
    }
    expect(table.sheets[0]!.name).toBe('Data')
    expect(table.sheets[0]!.rows[1]).toEqual(['Widgets', 12])
    expect(table.sheets[0]!.rows[3]![0]).toBe('Mixed, comma')
  })

  it('enforces the shared path policy: absolute paths, extension append, clobber guard', async () => {
    const cli = fakeCli()
    await startService(createSheetsTools({ ...baseDeps(), cli: cli.runner }, createSessionHost()))
    const relative = await client!.callTool({
      name: 'create_xlsx',
      arguments: { title: 'X', data: [[1]], path: 'relative.xlsx' },
    })
    expect(relative.isError).toBe(true)
    expect(text(relative.content)).toMatch(/path must be absolute/)

    const target = join(dir, 'taken')
    const first = await client!.callTool({
      name: 'create_xlsx',
      arguments: { title: 'Taken', data: [[1]], path: target },
    })
    expect(first.isError).toBeFalsy()
    const payload = JSON.parse(text(first.content)) as { path: string }
    expect(payload.path).toBe(`${target}.xlsx`)

    const second = await client!.callTool({
      name: 'create_xlsx',
      arguments: { title: 'Taken', data: [[2]], path: target },
    })
    expect(second.isError).toBe(true)
    expect(text(second.content)).toMatch(/file already exists/)
  })

  it('rejects malformed data', async () => {
    const cli = fakeCli()
    await startService(createSheetsTools({ ...baseDeps(), cli: cli.runner }, createSessionHost()))
    // flat rows fail the input schema before the handler runs
    const flat = await client!.callTool({
      name: 'create_xlsx',
      arguments: { title: 'Flat', data: [1, 2, 3] },
    })
    expect(flat.isError).toBe(true)
    const noTitle = await client!.callTool({
      name: 'create_xlsx',
      arguments: { title: ' ', data: [[1]] },
    })
    expect(noTitle.isError).toBe(true)
  })

  it('is hidden when background is off, present when on (session tools unaffected)', async () => {
    const fake = fakeSheetsControl()
    const cli = fakeCli()
    await startService(
      createSheetsTools(
        { ...baseDeps(false), sheets: fake.control, cli: cli.runner },
        createSessionHost(),
      ),
    )
    let names = (await client!.listTools()).tools.map((t) => t.name)
    expect(names).not.toContain('create_xlsx')
    expect(names).toContain('read_sheet')

    await client!.close()
    await service!.stop()
    service = undefined
    await startService(
      createSheetsTools(
        { ...baseDeps(true), sheets: fake.control, cli: cli.runner },
        createSessionHost(),
      ),
    )
    names = (await client!.listTools()).tools.map((t) => t.name)
    expect(names).toContain('create_xlsx')
  })
})

/** In-memory SheetsControl standing in for the shell's sheets-bridge */
function fakeSheetsControl(): { control: SheetsControl; saved: Array<{ path: string }> } {
  let nextId = 1
  const saved: Array<{ path: string }> = []
  return {
    saved,
    control: {
      openBlankTab: async () => nextId++,
      runCommand: async (wcId, command, payload) => {
        if (wcId <= 0) throw new Error('no workbook')
        const p = payload as { path?: string; addresses?: string[]; ops?: unknown[] }
        if (command === 'save_sheet') {
          // match on the file name, not a literal path: the sentinel has to be a
          // real absolute path on every platform to get past the tool's guard
          if (basename(p.path ?? '') === 'blocked.xlsx')
            throw new Error(`file already exists: ${p.path}`)
          saved.push({ path: p.path ?? '' })
          return { ok: true, path: p.path }
        }
        if (command === 'read_sheet') {
          return p.addresses?.length
            ? { cells: { A1: { value: 'hello', rawValue: 'hello' } } }
            : { context: { sheetId: 's1', wc: wcId, sheets: [] } }
        }
        if (p.ops?.some((o) => (o as { op: string }).op === 'explode')) {
          return { ok: false, reason: 'boom' }
        }
        return { ok: true }
      },
    },
  }
}

describe('visible grid session tools', () => {
  it('expose session lifecycle: create, read, apply, save-closes', async () => {
    const fake = fakeSheetsControl()
    await startService(sessionSurface(fake.control))

    const before = await client!.callTool({ name: 'read_sheet', arguments: {} })
    expect(before.isError).toBe(true)
    expect(text(before.content)).toMatch(/no session is open — call create_session first/)

    const created = await client!.callTool({
      name: 'create_session',
      arguments: { family: 'xlsx' },
    })
    expect(created.isError).toBeFalsy()

    const overview = await client!.callTool({ name: 'read_sheet', arguments: {} })
    expect(overview.isError).toBeFalsy()
    expect(text(overview.content)).toContain('"wc": 1')

    const cells = await client!.callTool({
      name: 'read_sheet',
      arguments: { addresses: ['A1', 'B2'] },
    })
    expect(cells.isError).toBeFalsy()
    expect(text(cells.content)).toContain('hello')

    const applied = await client!.callTool({
      name: 'apply_sheet_ops',
      arguments: { ops: [{ op: 'set_cell', sheetId: 's1', address: 'A1', value: 'x' }] },
    })
    expect(applied.isError).toBeFalsy()

    const saved = await client!.callTool({
      name: 'save_session',
      arguments: { path: join(dir, 'out.xlsx') },
    })
    expect(saved.isError).toBeFalsy()
    expect(fake.saved).toEqual([{ path: join(dir, 'out.xlsx') }])

    // the session ended with the save: further edits must ask for create_session
    const after = await client!.callTool({
      name: 'apply_sheet_ops',
      arguments: { ops: [{ op: 'set_cell' }] },
    })
    expect(after.isError).toBe(true)
    expect(text(after.content)).toMatch(/no session is open/)
  })

  it('surfaces failed applies as errors and forwards dry runs untouched', async () => {
    const fake = fakeSheetsControl()
    await startService(sessionSurface(fake.control))
    await client!.callTool({ name: 'create_session', arguments: { family: 'xlsx' } })

    // a docx target is refused before the bridge is reached (format guard)
    const wrongExt = await client!.callTool({
      name: 'save_session',
      arguments: { path: join(dir, 'out.docx') },
    })
    expect(wrongExt.isError).toBe(true)
    expect(text(wrongExt.content)).toMatch(/spreadsheet session must be saved as \.xlsx/)

    const failed = await client!.callTool({
      name: 'apply_sheet_ops',
      arguments: { ops: [{ op: 'explode' }] },
    })
    expect(failed.isError).toBe(true)
    expect(text(failed.content)).toContain('boom')

    const saved = await client!.callTool({
      name: 'save_session',
      arguments: { path: join(dir, 'blocked.xlsx') },
    })
    expect(saved.isError).toBe(true)
    expect(text(saved.content)).toMatch(/already exists/)
  })

  it('disappear entirely without a control (headless runs keep the file-only surface)', async () => {
    await startService(createSheetsTools(baseDeps(), createSessionHost()))
    const names = (await client!.listTools()).tools.map((t) => t.name)
    for (const tool of ['create_session', 'read_sheet', 'apply_sheet_ops', 'save_session']) {
      expect(names).not.toContain(tool)
    }
  })
})
