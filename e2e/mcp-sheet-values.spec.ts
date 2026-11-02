import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, freeMcpPort } from './helpers'
import { existsSync } from 'node:fs'
import { mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import http from 'node:http'
import JSZip from 'jszip'
import { SHELL_DIR } from './helpers'

/**
 * MCP sheet edits against the real built app: value typing, worksheet
 * addressing, and focus.
 *
 * A formula whose result does not divide evenly used to come back — and land on
 * disk — as the string its column-width-fitted General display produced
 * ("0.333333" with t="str"), because the reader returned the view model's
 * rendered text where the machines needed the model value. Worksheet ids are
 * also minted per session, so an id an agent cached stopped addressing anything
 * after a reopen; ops accept the CLI's worksheet NAME instead.
 */

const APP_ROOT = resolve(SHELL_DIR, '..', '..')

interface RpcResponse {
  status: number
  headers: http.IncomingHttpHeaders
  body: unknown
}

function parseBody(text: string): unknown {
  const trimmed = text.trim()
  if (!trimmed) return null
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      return JSON.parse(trimmed)
    } catch {
      return trimmed
    }
  }
  const out: unknown[] = []
  for (const line of trimmed.split('\n')) {
    if (!line.startsWith('data: ')) continue
    try {
      out.push(JSON.parse(line.slice(6)))
    } catch {
      /* endpoint/heartbeat line */
    }
  }
  return out.length === 1 ? out[0] : out
}

function postMcp(port: number, sessionId: string | null, payload: unknown): Promise<RpcResponse> {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload)
    const headers: Record<string, string | number> = {
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(data),
      Accept: 'application/json, text/event-stream',
    }
    if (sessionId) headers['Mcp-Session-Id'] = sessionId
    const req = http.request(
      { hostname: '127.0.0.1', port, path: '/mcp', method: 'POST', headers },
      (res) => {
        let text = ''
        res.on('data', (c) => (text += c))
        res.on('end', () =>
          resolve({ status: res.statusCode ?? 0, headers: res.headers, body: parseBody(text) }),
        )
      },
    )
    req.on('error', reject)
    req.write(data)
    req.end()
  })
}

async function waitForHealth(port: number, timeoutMs = 20_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    try {
      const ok = await new Promise<boolean>((resolve) => {
        const req = http.get({ hostname: '127.0.0.1', port, path: '/health' }, (res) => {
          res.resume()
          resolve(res.statusCode === 200)
        })
        req.on('error', () => resolve(false))
      })
      if (ok) return
    } catch {
      /* retry */
    }
    if (Date.now() > deadline) throw new Error('MCP server never became healthy')
    await new Promise((r) => setTimeout(r, 250))
  }
}

/** one MCP client over the app's Streamable-HTTP endpoint */
async function connect(port: number) {
  const init = await postMcp(port, null, {
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'e2e', version: '1.0' },
    },
  })
  expect(init.status).toBe(200)
  const sessionId = String(init.headers['mcp-session-id'] ?? '')
  expect(sessionId).not.toBe('')
  await postMcp(port, sessionId, { jsonrpc: '2.0', method: 'notifications/initialized' })
  return async (name: string, args: Record<string, unknown>) => {
    const res = await postMcp(port, sessionId, {
      jsonrpc: '2.0',
      id: Math.floor(Math.random() * 1e6),
      method: 'tools/call',
      params: { name, arguments: args },
    })
    const body = res.body as { result?: { isError?: boolean; content?: Array<{ text?: string }> } }
    const text = body?.result?.content?.map((c) => c.text ?? '').join('') ?? ''
    return { isError: body?.result?.isError === true, text }
  }
}

test.describe('MCP sheet values and addressing', () => {
  test('a repeating formula stays a number in the result, the read and the file', async () => {
    const port = await freeMcpPort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-sheet-values-'))
    const outFile = join(outDir, 'values.xlsx')
    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-sheet-values-userdata-'))
    // seed settings that skip onboarding and turn MCP on (launchShell only
    // writes its own file when onboardingSeen/settings are passed)
    const { writeFile } = await import('node:fs/promises')
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({
        onboardingSeen: true,
        mcpEnabled: true,
        mcpPort: port,
        defaultSaveDir: outDir,
      }),
    )
    const launched = await launchShell({
      videoDir: 'mcp-sheet-values',
      userDataDir,
    })
    const { app } = launched
    try {
      await waitForHealth(port)
      const call = await connect(port)

      const created = await call('create_session', { family: 'xlsx' })
      expect(created.isError, created.text).toBeFalsy()
      await waitForPageWithUrl(app, 'genoffice-app://sheets')

      // =1/3 in a default-width column is the case that rendered as "0.333333"
      const applied = await call('apply_sheet_ops', {
        ops: [
          { op: 'set_formula', sheet: 'Sheet1', address: 'B1', formula: '=1/3' },
          { op: 'set_formula', sheet: 'Sheet1', address: 'B2', formula: '=1/2' },
          { op: 'set_formula', sheet: 'Sheet1', address: 'B3', formula: '=1/7' },
        ],
      })
      expect(applied.isError, applied.text).toBeFalsy()
      const applyPayload = JSON.parse(applied.text) as {
        formulaValues: Array<{ address: string; value: unknown }>
      }
      const byAddress = new Map(applyPayload.formulaValues.map((c) => [c.address, c.value]))
      // the full-precision engine value, not the display text and not a string
      expect(byAddress.get('B1')).toBe(1 / 3)
      expect(byAddress.get('B2')).toBe(0.5)
      expect(byAddress.get('B3')).toBe(1 / 7)

      // read_sheet agrees, and still shows what the user sees on screen
      const read = await call('read_sheet', { addresses: ['B1', 'B2', 'B3'] })
      expect(read.isError, read.text).toBeFalsy()
      const cells = (
        JSON.parse(read.text) as { cells: Record<string, { value: unknown; display?: string }> }
      ).cells
      expect(cells.B1!.value).toBe(1 / 3)
      expect(typeof cells.B1!.value).toBe('number')
      expect(cells.B1!.display).toBe('0.333333')

      const saved = await call('save_session', { path: outFile })
      expect(saved.isError, saved.text).toBeFalsy()
      expect(existsSync(outFile)).toBe(true)

      // on disk: a real number with Excel's own precision, no t="str"
      const zip = await JSZip.loadAsync(await readFile(outFile))
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheetXml).not.toContain('t="str"')
      const b1 = /<c r="B1"[^>]*>(.*?)<\/c>/.exec(sheetXml)?.[1] ?? ''
      expect(b1).toContain('<f>1/3</f>')
      expect(Number(/<v>([^<]*)<\/v>/.exec(b1)?.[1])).toBeCloseTo(1 / 3, 15)
    } finally {
      await closeAndSaveVideo(launched, 'mcp-sheet-values')
    }
  })

  test('addresses worksheets by name after a reopen (ids do not survive)', async () => {
    const port = await freeMcpPort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-sheet-names-'))
    const workbookPath = join(outDir, 'two-sheets.xlsx')
    // build a two-sheet workbook with the repo's CLI, the way a user's file exists
    const { execFileSync } = await import('node:child_process')
    const payload = join(outDir, 'table.json')
    const { writeFile } = await import('node:fs/promises')
    await writeFile(
      payload,
      JSON.stringify({
        sheets: [
          { name: 'Summary', rows: [['total'], [0]] },
          { name: 'Data 2024', rows: [['widget'], [12]] },
        ],
      }),
    )
    execFileSync(
      process.execPath,
      [
        join(APP_ROOT, 'packages', 'cli', 'dist', 'genoffice.cjs'),
        'create',
        '--type',
        'xlsx',
        '--from',
        payload,
        '--out',
        workbookPath,
      ],
      { stdio: 'pipe' },
    )
    expect(existsSync(workbookPath)).toBe(true)

    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-sheet-names-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({
      videoDir: 'mcp-sheet-names',
      userDataDir,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    const { app, page } = launched
    try {
      await waitForHealth(port)
      const call = await connect(port)

      // open the user's file through the app's own routing — NOT a session,
      // so its sheet ids are whatever that open produced
      const opened = await call('open_in_genoffice', { path: workbookPath })
      expect(opened.isError, opened.text).toBeFalsy()
      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      const editorPage = await waitForPageWithUrl(app, 'genoffice-app://sheets')

      const activeSheetName = async (): Promise<string | null> =>
        editorPage.evaluate(() => {
          const api = (
            window as unknown as {
              __genofficeDebug?: {
                univerAPI?: {
                  getActiveWorkbook(): {
                    getActiveSheet(): { getSheetName(): string } | null
                  } | null
                }
              }
            }
          ).__genofficeDebug?.univerAPI
          return api?.getActiveWorkbook()?.getActiveSheet()?.getSheetName() ?? null
        })

      // the workbook opens on its first sheet
      await expect.poll(activeSheetName).toBe('Summary')

      // a name-addressed edit works without ever reading an id. The document
      // argument is how a caller edits the file the *user* has open rather
      // than the session's own blank tab.
      const applied = await call('apply_sheet_ops', {
        document: workbookPath,
        ops: [{ op: 'set_formula', sheet: 'Data 2024', address: 'B2', formula: '=7/3' }],
      })
      expect(applied.isError, applied.text).toBeFalsy()

      // focus followed the edit: the user is looking at the sheet the agent
      // changed, not whichever sheet happened to be active
      await expect.poll(activeSheetName).toBe('Data 2024')

      const read = await call('read_sheet', {
        document: workbookPath,
        sheet: 'Data 2024',
        addresses: ['B2'],
      })
      expect(read.isError, read.text).toBeFalsy()
      const cells = (JSON.parse(read.text) as { cells: Record<string, { value: unknown }> }).cells
      expect(cells.B2!.value).toBe(7 / 3)

      // the edit landed on the named sheet only
      const summary = await call('read_sheet', {
        document: workbookPath,
        sheet: 'Summary',
        addresses: ['B2'],
      })
      const summaryCells = (
        JSON.parse(summary.text) as { cells: Record<string, { value: unknown }> }
      ).cells
      expect(summaryCells.B2?.value ?? null).toBeNull()

      // an unknown sheet name is refused before anything applies, naming the sheets
      const refused = await call('apply_sheet_ops', {
        document: workbookPath,
        ops: [{ op: 'set_cell', sheet: 'Sheet1', address: 'A1', value: 'x' }],
      })
      expect(refused.isError).toBe(true)
      expect(refused.text).toContain('Sheet1')
      expect(refused.text).toContain('Summary')
      expect(refused.text).toContain('Data 2024')

      // copy_range's sourceSheet uses the same vocabulary
      const copied = await call('apply_sheet_ops', {
        document: workbookPath,
        ops: [
          {
            op: 'copy_range',
            sheet: 'Summary',
            sourceSheet: 'Data 2024',
            source: 'A1:A2',
            target: 'C1',
          },
        ],
      })
      expect(copied.isError, copied.text).toBeFalsy()
      const copiedRead = await call('read_sheet', {
        document: workbookPath,
        sheet: 'Summary',
        addresses: ['C1', 'C2'],
      })
      const copiedCells = (
        JSON.parse(copiedRead.text) as { cells: Record<string, { value: unknown }> }
      ).cells
      expect(copiedCells.C1!.value).toBe('widget')
      expect(copiedCells.C2!.value).toBe(12)
    } finally {
      await closeAndSaveVideo(launched, 'mcp-sheet-names')
    }
  })
})
