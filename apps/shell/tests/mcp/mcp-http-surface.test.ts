import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { createServer, get as httpGet, request as httpRequest } from 'node:http'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { parseDocx } from '@genoffice/docx-engine'
import { openPptx } from '@genoffice/pptx-engine'
import JSZip from 'jszip'
import {
  applyMcpSettings,
  configureMcpRuntime,
  mcpStatus,
  stopMcp,
} from '../../src/main/mcp/app-mcp'
import type { DocsControl } from '../../src/main/mcp/tools/document-tools'
import type { SlidesControl } from '../../src/main/mcp/tools/slides-tools'
import type { SheetsControl } from '../../src/main/mcp/tools/sheets-tools'
import { realCliRunner } from './real-cli'

/**
 * Full-surface MCP acceptance test over the REAL Streamable HTTP transport.
 *
 * Boots the server through the app's own composition path (configureMcpRuntime +
 * applyMcpSettings, i.e. exactly what the Settings toggle drives), then connects
 * an MCP client to `http://127.0.0.1:<port>/mcp` — the same URL an
 * `mcpServers` entry would use. Only the three editor bridges are faked (no
 * Electron windows in a headless run); the transport, handshake, tool
 * registration, schema validation, session host and the real docx/pptx/xlsx
 * headless engines are all exercised for real.
 */

const DEFAULT_NAMES = [
  'apply_ops',
  'apply_sheet_ops',
  'apply_slide_ops',
  'create_session',
  'get_app_info',
  'insert_content',
  'open_in_genoffice',
  'read_deck',
  'read_document',
  'read_docx',
  'read_sheet',
  'replace_blocks',
  'save_session',
].sort()

const BACKGROUND_NAMES = [...DEFAULT_NAMES, 'create_docx', 'create_pptx', 'create_xlsx'].sort()

function freePort(): Promise<number> {
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

/** fake docs bridge that records every command the tools forward */
function docsControl(): { control: DocsControl; calls: Array<Record<string, unknown>> } {
  const calls: Array<Record<string, unknown>> = []
  let wc = 100
  const control: DocsControl = {
    openBlankTab: async () => {
      calls.push({ command: 'open_blank_tab', wcId: ++wc })
      return wc
    },
    runCommand: async (wcId, command, payload) => {
      calls.push({ wcId, command, payload })
      switch (command) {
        case 'read_document':
          return { blocks: [{ index: 0, text: 'Quarterly Report' }], text: 'Quarterly Report' }
        case 'insert_content':
          return { summary: 'inserted', mutated: true }
        case 'replace_blocks':
          return { summary: 'replaced', mutated: true }
        case 'apply_ops':
          return payload && (payload as { dryRun?: boolean }).dryRun
            ? { plan: [{ op: 'setFont' }], mutated: false }
            : { summary: 'applied', mutated: true }
        case 'save_document':
          return { ok: true, path: (payload as { path: string }).path }
        default:
          return {}
      }
    },
  }
  return { control, calls }
}

function slidesControl(): { control: SlidesControl; saved: string[] } {
  const saved: string[] = []
  let wc = 200
  return {
    saved,
    control: {
      openBlankTab: async () => ++wc,
      runTxn: async (_wcId, req) => {
        const ops = req.ops as Array<{ op: string }>
        if (ops.some((o) => o.op === 'explode')) {
          return { applied: false, failures: [{ index: 0, error: 'boom' }] }
        }
        return { applied: true, records: ops.map((o) => ({ op: o.op })) }
      },
      readDeck: async (wcId) => ({ slides: [{ index: 0, wc: wcId, elements: [] }] }),
      saveDeck: async (_wcId, path) => {
        saved.push(path)
        return { path }
      },
    },
  }
}

function sheetsControl(): { control: SheetsControl; saved: string[] } {
  const saved: string[] = []
  let wc = 300
  return {
    saved,
    control: {
      openBlankTab: async () => ++wc,
      runCommand: async (wcId, command, payload) => {
        const p = payload as { path?: string; addresses?: string[]; ops?: Array<{ op: string }> }
        if (command === 'save_sheet') {
          saved.push(p.path ?? '')
          return { ok: true, path: p.path }
        }
        if (command === 'read_sheet') {
          return p.addresses?.length
            ? { cells: { A1: { value: 'hello', rawValue: 'hello' } } }
            : { context: { sheetId: 's1', wc: wcId, sheets: [] } }
        }
        if (p.ops?.some((o) => o.op === 'explode')) return { ok: false, reason: 'boom' }
        return { ok: true, applied: p.ops?.length ?? 0 }
      },
    },
  }
}

let port: number
let workDir: string
let logPath: string
const docs = docsControl()
const slides = slidesControl()
const sheets = sheetsControl()
const openedPaths: string[] = []

/** poll /health so we never race a server (re)start — what a real client does */
async function waitForHealth(p: number, timeoutMs = 10_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    const ok = await new Promise<boolean>((resolve) => {
      const req = httpGet({ hostname: '127.0.0.1', port: p, path: '/health' }, (res) => {
        res.resume()
        resolve(res.statusCode === 200)
      })
      req.on('error', () => resolve(false))
    })
    if (ok) return
    if (Date.now() > deadline) throw new Error('MCP server never became healthy')
    await new Promise((r) => setTimeout(r, 100))
  }
}

async function connect(): Promise<Client> {
  await waitForHealth(port)
  const client = new Client({ name: 'genoffice-acceptance', version: '1.0.0' })
  // exactly the shape in the mcpServers entry under test
  await client.connect(new StreamableHTTPClientTransport(new URL(`http://127.0.0.1:${port}/mcp`)))
  return client
}

async function call(
  client: Client,
  name: string,
  args: Record<string, unknown>,
): Promise<{ isError: boolean; text: string; json: unknown }> {
  const result = (await client.callTool({ name, arguments: args })) as {
    isError?: boolean
    content?: Array<{ text?: string }>
  }
  const text = (result.content ?? []).map((c) => c.text ?? '').join('')
  let json: unknown
  try {
    json = JSON.parse(text)
  } catch {
    json = text
  }
  return { isError: result.isError === true, text, json }
}

beforeAll(async () => {
  port = await freePort()
  workDir = await mkdtemp(join(tmpdir(), 'genoffice-acceptance-'))
  logPath = join(workDir, 'mcp-log.txt')
  configureMcpRuntime({
    version: '0.9.0-acceptance',
    defaultSaveDir: () => workDir,
    openPath: (p) => {
      // pretend only .docx files in the temp dir route to a tab
      const ok = p.startsWith(workDir) && /\.docx$/i.test(p)
      if (ok) openedPaths.push(p)
      return ok
    },
    docsControl: docs.control,
    slidesControl: slides.control,
    sheetsControl: sheets.control,
    cliRunner: realCliRunner(workDir),
    logFilePath: logPath,
  })
  await applyMcpSettings({ enabled: true, port, background: false, logging: true })
})

afterAll(async () => {
  await stopMcp()
  await rm(workDir, { recursive: true, force: true })
})

describe('MCP surface over Streamable HTTP (/mcp)', () => {
  it('completes the full workflow an mcpServers client would drive', async () => {
    const client = await connect()
    try {
      // ── 1. handshake + tools/list ──────────────────────────────────────────
      const listed = (await client.listTools()).tools.map((t) => t.name).sort()
      expect(listed).toEqual(DEFAULT_NAMES)
      expect(mcpStatus().capabilities).toEqual(['docs', 'slides', 'sheets'])

      // ── 2. get_app_info: editor matrix + exposed formats ──────────────────
      const info = await call(client, 'get_app_info', {})
      const infoJson = info.json as {
        name: string
        version: string
        defaultSaveDir: string
        formats: string[]
        families: Array<{ family: string; editor: { open: string[] }; mcp?: unknown }>
      }
      expect(infoJson.name).toBe('GenOffice')
      expect(infoJson.version).toBe('0.9.0-acceptance')
      expect(infoJson.defaultSaveDir).toBe(workDir)
      // background generation is off here, so create_docx/create_pptx/create_xlsx
      // are not registered and get_app_info must not advertise a generation
      // format the client cannot use
      expect(infoJson.formats).toEqual([])
      expect(infoJson.families.find((f) => f.family === 'docx')!.mcp).toEqual({
        save: ['docx'],
        read: 'docx',
      })
      // the slides and sheets families follow the same rule: only `generate` is
      // hidden
      expect(infoJson.families.find((f) => f.family === 'pptx')!.mcp).toEqual({
        save: ['pptx'],
      })
      expect(infoJson.families.find((f) => f.family === 'xlsx')!.mcp).toEqual({
        save: ['xlsx'],
      })
      expect(infoJson.families.map((f) => f.family)).toEqual([
        'docx',
        'xlsx',
        'pptx',
        'md',
        'html',
        'pdf',
      ])
      // the format registry's editor truth survives the wire
      expect(infoJson.families.find((f) => f.family === 'xlsx')!.editor.open).toEqual([
        'xlsx',
        'xlsm',
        'xls',
        'csv',
      ])

      // ── 3. read_docx error paths ──────────────────────────────────────────
      const relative = await call(client, 'read_docx', { path: 'relative.docx' })
      expect(relative.isError).toBe(true)
      expect(relative.text).toMatch(/path must be absolute/)
      const missing = await call(client, 'read_docx', { path: join(workDir, 'nope.docx') })
      expect(missing.isError).toBe(true)
      expect(missing.text).toMatch(/file not found/)
      const wrongExt = await call(client, 'read_docx', { path: join(workDir, 'x.txt') })
      expect(wrongExt.isError).toBe(true)

      // ── 4. visible docx session: create → edit → read ─────────────────────
      const created = await call(client, 'create_session', { family: 'docx' })
      expect(created.isError).toBe(false)
      expect(created.json).toMatchObject({ ok: true, family: 'docx' })

      const inserted = await call(client, 'insert_content', { html: '<h1>Hi</h1>' })
      expect(inserted.isError).toBe(false)
      expect((inserted.json as { mutated: boolean }).mutated).toBe(true)

      const applied = await call(client, 'apply_ops', {
        ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true }],
      })
      expect(applied.isError).toBe(false)

      const dry = await call(client, 'apply_ops', {
        ops: [{ op: 'setFont' }],
        dryRun: true,
      })
      expect(dry.isError).toBe(false)
      expect((dry.json as { mutated: boolean }).mutated).toBe(false)

      const read = await call(client, 'read_document', {})
      expect(read.isError).toBe(false)
      expect((read.json as { text: string }).text).toContain('Quarterly Report')

      // ── 5. cross-family refusal (schema-valid, host-level) ────────────────
      // a second session switches the family; the docx content tool now refuses
      const pptxSession = await call(client, 'create_session', { family: 'pptx' })
      expect(pptxSession.isError).toBe(false)
      const crossEdit = await call(client, 'insert_content', { html: '<p>x</p>' })
      expect(crossEdit.isError).toBe(true)
      expect(crossEdit.text).toMatch(/active session is a presentation/)

      // read_deck now works (the active family matches)
      const deck = await call(client, 'read_deck', {})
      expect(deck.isError).toBe(false)
      const txn = await call(client, 'apply_slide_ops', { ops: [{ op: 'addBlankSlide' }] })
      expect(txn.isError).toBe(false)
      const failedTxn = await call(client, 'apply_slide_ops', { ops: [{ op: 'explode' }] })
      expect(failedTxn.isError).toBe(true)
      expect(failedTxn.text).toContain('boom')

      // ── 6. format guard on save_session ──────────────────────────────────
      const wrongSave = await call(client, 'save_session', { path: join(workDir, 'deck.docx') })
      expect(wrongSave.isError).toBe(true)
      expect(wrongSave.text).toMatch(/presentation session must be saved as \.pptx/)

      const relativeSave = await call(client, 'save_session', { path: 'relative.pptx' })
      expect(relativeSave.isError).toBe(true)
      expect(relativeSave.text).toMatch(/path must be absolute/)

      const savedDeck = await call(client, 'save_session', {
        path: join(workDir, 'deck.pptx'),
        overwrite: true,
      })
      expect(savedDeck.isError).toBe(false)
      expect(slides.saved).toEqual([join(workDir, 'deck.pptx')])

      // the save ended the session
      const afterSave = await call(client, 'apply_slide_ops', { ops: [{ op: 'addBlankSlide' }] })
      expect(afterSave.isError).toBe(true)
      expect(afterSave.text).toMatch(/no session is open/)

      const saveNoSession = await call(client, 'save_session', { path: join(workDir, 'x.pptx') })
      expect(saveNoSession.isError).toBe(true)
      expect(saveNoSession.text).toMatch(/no session is open/)

      // ── 7. sheets session round trip ─────────────────────────────────────
      const sheetSession = await call(client, 'create_session', { family: 'xlsx' })
      expect(sheetSession.isError).toBe(false)
      const overview = await call(client, 'read_sheet', {})
      expect(overview.isError).toBe(false)
      const cells = await call(client, 'read_sheet', { addresses: ['A1'] })
      expect(cells.isError).toBe(false)
      expect(cells.text).toContain('hello')
      const sheetOps = await call(client, 'apply_sheet_ops', {
        ops: [{ op: 'set_cell', sheetId: 's1', address: 'A1', value: 'x' }],
      })
      expect(sheetOps.isError).toBe(false)
      const savedSheet = await call(client, 'save_session', { path: join(workDir, 'book.xlsx') })
      expect(savedSheet.isError).toBe(false)
      expect(sheets.saved).toEqual([join(workDir, 'book.xlsx')])

      // ── 8. open_in_genoffice routes .docx, refuses others ────────────────
      const openableDocx = join(workDir, 'open-me.docx')
      await writeFile(openableDocx, 'placeholder')
      const opened = await call(client, 'open_in_genoffice', { path: openableDocx })
      expect(opened.isError).toBe(false)
      expect(openedPaths).toContain(openableDocx)

      const notOpenable = join(workDir, 'note.txt')
      await writeFile(notOpenable, 'x')
      const refusedOpen = await call(client, 'open_in_genoffice', { path: notOpenable })
      expect(refusedOpen.isError).toBe(true)
      expect(refusedOpen.text).toMatch(/could not open/)
    } finally {
      await client.close()
    }
  })

  it('exposes the headless generators after a background flip, writing real files', async () => {
    // flip the setting: the server rebuilds its tool list on the SAME port
    await applyMcpSettings({ enabled: true, port, background: true, logging: true })
    const client = await connect()
    try {
      const listed = (await client.listTools()).tools.map((t) => t.name).sort()
      expect(listed).toEqual(BACKGROUND_NAMES)

      // with generation on, get_app_info advertises the formats those tools write
      const info = await call(client, 'get_app_info', {})
      expect((info.json as { formats: string[] }).formats.sort()).toEqual(['docx', 'pptx', 'xlsx'])

      // create_docx → a real, reparsable .docx
      const docxPath = join(workDir, 'generated.docx')
      const madeDoc = await call(client, 'create_docx', {
        title: 'Generated',
        content: '# Heading\n\nBody paragraph.',
        path: docxPath,
      })
      expect(madeDoc.isError).toBe(false)
      expect(existsSync(docxPath)).toBe(true)
      const parsed = await parseDocx(new Uint8Array(await readFile(docxPath)))
      expect(parsed.blocks.filter((b) => !b.hidden)[0].type).toBe('heading')

      // read_docx reads it back through the real engine
      const readBack = await call(client, 'read_docx', { path: docxPath })
      expect(readBack.isError).toBe(false)
      expect((readBack.json as { text: string }).text).toContain('Heading')
      expect((readBack.json as { text: string }).text).toContain('Body paragraph')

      // create_pptx → a real 2-slide deck
      const pptxPath = join(workDir, 'generated.pptx')
      const madeDeck = await call(client, 'create_pptx', {
        title: 'Deck',
        outline: '# One\n- a\n\n# Two\n- b',
        path: pptxPath,
      })
      expect(madeDeck.isError).toBe(false)
      expect((madeDeck.json as { path: string }).path).toBe(pptxPath)
      const opened = await openPptx(new Uint8Array(await readFile(pptxPath)))
      expect(opened.deck.slides).toHaveLength(2)

      // create_xlsx → a real workbook with a numeric cell
      const xlsxPath = join(workDir, 'generated.xlsx')
      const madeBook = await call(client, 'create_xlsx', {
        title: 'Book',
        data: [
          ['Item', 'Qty'],
          ['Widget', 12],
        ],
        path: xlsxPath,
      })
      expect(madeBook.isError).toBe(false)
      const zip = await JSZip.loadAsync(await readFile(xlsxPath))
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheetXml).toContain('<v>12</v>')

      // the clobber guard still holds for explicit paths
      const clobber = await call(client, 'create_xlsx', {
        title: 'Book',
        data: [[1]],
        path: xlsxPath,
      })
      expect(clobber.isError).toBe(true)
      expect(clobber.text).toMatch(/file already exists/)
    } finally {
      await client.close()
    }
  })

  it('speaks raw JSON-RPC over /mcp with no SDK in the loop', async () => {
    // the same URL from the mcpServers entry, driven by plain node:http
    const health = await raw('GET', '/health')
    expect(health.status).toBe(200)
    const healthBody = JSON.parse(health.body) as {
      status: string
      transport: string
      port: number
    }
    expect(healthBody.status).toBe('ok')
    expect(healthBody.port).toBe(port)
    console.log('[health]', health.body)

    // initialize → the server hands back a session id header
    const init = await raw('POST', '/mcp', {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: { name: 'raw-probe', version: '1' },
      },
    })
    expect(init.status).toBe(200)
    const sessionId = String(init.headers['mcp-session-id'] ?? '')
    expect(sessionId).not.toBe('')
    console.log('[initialize] mcp-session-id =', sessionId)

    await raw('POST', '/mcp', { jsonrpc: '2.0', method: 'notifications/initialized' }, sessionId)

    const list = await raw(
      'POST',
      '/mcp',
      { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} },
      sessionId,
    )
    const listBody = parseRpc(list.body) as { result?: { tools?: Array<{ name: string }> } }
    const names = (listBody.result?.tools ?? []).map((t) => t.name).sort()
    expect(names).toEqual(BACKGROUND_NAMES) // background is still on from the previous test
    console.log('[tools/list]', names.join(', '))

    // a real tool call over the wire
    const infoCall = await raw(
      'POST',
      '/mcp',
      {
        jsonrpc: '2.0',
        id: 3,
        method: 'tools/call',
        params: { name: 'get_app_info', arguments: {} },
      },
      sessionId,
    )
    const callBody = parseRpc(infoCall.body) as {
      result?: { content?: Array<{ text?: string }> }
    }
    const infoText = (callBody.result?.content ?? []).map((c) => c.text ?? '').join('')
    expect(infoText).toContain('GenOffice')
    console.log('[tools/call get_app_info] ok,', infoText.length, 'chars')

    // a bad session id is refused, not silently accepted
    const badSession = await raw(
      'POST',
      '/mcp',
      { jsonrpc: '2.0', id: 4, method: 'tools/list', params: {} },
      'nope',
    )
    expect(badSession.status).toBe(404)
    console.log('[bad session]', badSession.status, badSession.body)
  })

  it('isolates the active session between two concurrent clients', async () => {
    // background off so the surface is the visible-session one
    await applyMcpSettings({ enabled: true, port, background: false, logging: true })

    const clientA = await connect()
    const clientB = await connect()
    try {
      // A opens a docx session, B opens a pptx session — both succeed
      const a = await call(clientA, 'create_session', { family: 'docx' })
      expect(a.isError).toBe(false)
      const b = await call(clientB, 'create_session', { family: 'pptx' })
      expect(b.isError).toBe(false)

      // A's docx content tool still works (B's session did not clobber A's)
      const aEdit = await call(clientA, 'insert_content', { html: '<p>a</p>' })
      expect(aEdit.isError).toBe(false)
      // B's deck tool still works
      const bRead = await call(clientB, 'read_deck', {})
      expect(bRead.isError).toBe(false)

      // and each client only sees its own family
      const aDeck = await call(clientA, 'read_deck', {})
      expect(aDeck.isError).toBe(true)
      expect(aDeck.text).toMatch(/active session is a Word document/)
      const bEdit = await call(clientB, 'insert_content', { html: '<p>b</p>' })
      expect(bEdit.isError).toBe(true)
      expect(bEdit.text).toMatch(/active session is a presentation/)

      // closing A's session leaves B's intact
      await call(clientA, 'save_session', { path: join(workDir, 'a.docx'), overwrite: true })
      const bStill = await call(clientB, 'read_deck', {})
      expect(bStill.isError).toBe(false)
    } finally {
      await clientA.close()
      await clientB.close()
    }
  })
})

/** raw HTTP exchange, optionally carrying a session id — no MCP SDK involved */
function raw(
  method: 'GET' | 'POST',
  path: string,
  payload?: unknown,
  sessionId?: string,
): Promise<{
  status: number
  headers: Record<string, string | string[] | undefined>
  body: string
}> {
  return new Promise((resolve, reject) => {
    const data = payload === undefined ? undefined : JSON.stringify(payload)
    const req = httpRequest(
      {
        hostname: '127.0.0.1',
        port,
        path,
        method,
        headers: {
          ...(data
            ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) }
            : {}),
          Accept: 'application/json, text/event-stream',
          ...(sessionId ? { 'Mcp-Session-Id': sessionId } : {}),
        },
      },
      (res) => {
        let body = ''
        res.on('data', (c) => (body += c))
        res.on('end', () => resolve({ status: res.statusCode ?? 0, headers: res.headers, body }))
      },
    )
    req.on('error', reject)
    if (data) req.write(data)
    req.end()
  })
}

/** the transport may answer as JSON or as an SSE event stream */
function parseRpc(body: string): unknown {
  const trimmed = body.trim()
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) return JSON.parse(trimmed)
  for (const line of trimmed.split('\n')) {
    if (line.startsWith('data: ')) {
      try {
        return JSON.parse(line.slice(6))
      } catch {
        /* heartbeat / endpoint line */
      }
    }
  }
  return null
}
