/**
 * A cell border the user removes is written as "nil". Word lets a cell's
 * "none" fall back to the table's own borders and draws them, so the page
 * (no line) and the file (a line in Word) disagreed. Both writers are covered:
 * Table Design > Borders > No Border and the AI table tool.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import http from 'node:http'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, freeMcpPort } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-table-border-nil'
const SIDES = ['top', 'left', 'bottom', 'right'] as const

const line = (side: string) => `<w:${side} w:val="single" w:sz="12" w:space="0" w:color="000000"/>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/>' +
  `<w:tblBorders>${[...SIDES, 'insideH', 'insideV'].map(line).join('')}</w:tblBorders>` +
  '<w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  [0, 1, 2]
    .map(
      (r) =>
        `<w:tr>${[0, 1, 2]
          .map(
            (c) =>
              `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>r${r}c${c}</w:t></w:r></w:p></w:tc>`,
          )
          .join('')}</w:tr>`,
    )
    .join('') +
  '</w:tbl>'

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string')

/** per cell of the saved table: the value each side's border is written with */
const savedCellBorders = async (docPath: string) =>
  Array.from((await documentXml(docPath)).matchAll(/<w:tc>.*?<\/w:tc>/gs)).map((cell) =>
    Object.fromEntries(
      SIDES.map((side) => [
        side,
        new RegExp(`<w:tcBorders>.*<w:${side} w:val="(\\w+)"`, 's').exec(cell[0])?.[1] ?? null,
      ]),
    ),
  )

/** per row, per cell: the sides that draw a line, as the page shows them */
const drawnSides = (page: Page) =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('.doc-page table tr')).map((tr) =>
      Array.from(tr.querySelectorAll('td, th')).map((td) => {
        const s = getComputedStyle(td)
        return (['Top', 'Left', 'Bottom', 'Right'] as const)
          .filter((k) => s[`border${k}Style`] !== 'none' && parseFloat(s[`border${k}Width`]) > 0)
          .join('')
      }),
    ),
  )

const NIL = { top: 'nil', left: 'nil', bottom: 'nil', right: 'nil' }
const BARE = [0, 1, 2].map(() => ['', '', ''])

async function open(docPath: string, settings?: Record<string, unknown>) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    settings,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page table')).toHaveCount(1, { timeout: 30_000 })
  return { launched, app: launched.app, page }
}

const close = async (session: Awaited<ReturnType<typeof open>>) => {
  await closeAndSaveVideo(session.launched, VIDEO_DIR)
  rmSync(session.launched.userDataDir, { recursive: true, force: true })
}

test('No Border on a bordered table writes nil on every cell, and the file reads back without lines', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-border-nil-')))
  const docPath = join(dir, 'no-border.docx')
  writeFileSync(docPath, await docxWith(para('Before') + TABLE + para('After')))
  let session = await open(docPath)
  try {
    const { app, page } = session
    expect((await drawnSides(page))[1][1]).toBe('TopLeftBottomRight')

    // a bare caret formats the whole table
    await page.locator('.doc-page td', { hasText: 'r1c1' }).click()
    await page.locator('.ribbon-tab', { hasText: 'Table Design' }).click()
    await page.locator('.table-split-caret').click()
    await page.locator('.layout-menu button', { hasText: 'No Border' }).click()
    await expect.poll(() => drawnSides(page)).toEqual(BARE)

    await saveFromMenu(app)
    await expect
      .poll(() => savedCellBorders(docPath), { timeout: 30_000 })
      .toEqual(Array.from({ length: 9 }, () => NIL))
    // the table keeps its own outer lines: under "none" Word would draw them
    expect(await documentXml(docPath)).toMatch(/<w:tblBorders>.*<w:top w:val="single"/s)

    await close(session)
    session = await open(docPath)
    expect(await drawnSides(session.page)).toEqual(BARE)
  } finally {
    await close(session)
    rmSync(dir, { recursive: true, force: true })
  }
})

interface Rpc {
  headers: http.IncomingHttpHeaders
  body: { result?: { isError?: boolean; content?: Array<{ text?: string }> } } | null
}

function postMcp(port: number, sessionId: string | null, payload: unknown): Promise<Rpc> {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload)
    const req = http.request(
      {
        hostname: '127.0.0.1',
        port,
        path: '/mcp',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(data),
          Accept: 'application/json, text/event-stream',
          ...(sessionId ? { 'Mcp-Session-Id': sessionId } : {}),
        },
      },
      (res) => {
        let text = ''
        res.on('data', (c) => (text += c))
        res.on('end', () => {
          const json = text.trim().startsWith('{')
            ? text
            : (text
                .split('\n')
                .find((l) => l.startsWith('data: '))
                ?.slice(6) ?? 'null')
          resolve({ headers: res.headers, body: JSON.parse(json) as Rpc['body'] })
        })
      },
    )
    req.on('error', reject)
    req.end(data)
  })
}

const healthy = (port: number) =>
  new Promise<boolean>((resolve) => {
    const req = http.get({ hostname: '127.0.0.1', port, path: '/health' }, (res) => {
      res.resume()
      resolve(res.statusCode === 200)
    })
    req.on('error', () => resolve(false))
  })

test('the AI table tool writes nil for borders it removes, whether asked with null or with "none"', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-border-nil-')))
  const docPath = join(dir, 'ai-no-border.docx')
  writeFileSync(docPath, await docxWith(para('Before') + TABLE + para('After')))
  const port = await freeMcpPort()
  let session = await open(docPath, { mcpEnabled: true, mcpPort: port })
  try {
    const { app, page } = session
    await expect.poll(() => healthy(port), { timeout: 20_000 }).toBe(true)
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
    const sessionId = String(init.headers['mcp-session-id'] ?? '')
    await postMcp(port, sessionId, { jsonrpc: '2.0', method: 'notifications/initialized' })
    const target = { blockIndexes: [1] }
    const applied = await postMcp(port, sessionId, {
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/call',
      params: {
        name: 'apply_ops',
        arguments: {
          document: docPath,
          ops: [
            { op: 'setTableCellFormat', target, borders: null },
            {
              op: 'setTableCellFormat',
              target,
              range: { row: 2, col: 0, colEnd: 2 },
              borders: { bottom: { style: 'single', width: 1.5 }, top: { style: 'none' } },
            },
          ],
        },
      },
    })
    expect(
      applied.body?.result?.isError,
      applied.body?.result?.content?.map((c) => c.text).join(''),
    ).toBeFalsy()
    const lastRow = ['Bottom', 'Bottom', 'Bottom']
    await expect.poll(() => drawnSides(page)).toEqual([BARE[0], BARE[0], lastRow])

    await saveFromMenu(app)
    await expect
      .poll(() => savedCellBorders(docPath), { timeout: 30_000 })
      .toEqual([
        ...Array.from({ length: 6 }, () => NIL),
        ...Array.from({ length: 3 }, () => ({ ...NIL, bottom: 'single' })),
      ])

    await close(session)
    session = await open(docPath)
    expect(await drawnSides(session.page)).toEqual([BARE[0], BARE[0], lastRow])
  } finally {
    await close(session)
    rmSync(dir, { recursive: true, force: true })
  }
})
