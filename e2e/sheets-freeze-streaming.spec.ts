import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var; the
// spec drives the freeze and reads cell data through the Facade
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Regression for gendesk #680: "a frozen top row unfreezes after scrolling".
 * On a workbook large enough to stream, the viewport eviction's "never evict
 * frozen rows" guard reads the freeze from the parsed file model — which an
 * in-session Freeze Top Row never used to reach, so the pinned row's cells
 * were null-wiped on the next window advance and the header went blank.
 * The freeze mutation now updates the file model, and a frozen-row strip
 * loader refetches the header band when the viewport lives below it.
 */

const ROWS = 120_000
const COLUMNS = 10

/** Minimal valid xlsx, big enough (1.2M value cells) to open in lazy mode. */
async function writeStreamingFixture(path: string): Promise<void> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  const letters = Array.from({ length: COLUMNS }, (_, c) => String.fromCharCode(65 + c))
  const chunks: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      `<dimension ref="A1:${letters[COLUMNS - 1]}${ROWS}"/><sheetData>`,
    `<row r="1">${letters.map((l) => `<c r="${l}1" t="inlineStr"><is><t>HEADER_${l}</t></is></c>`).join('')}</row>`,
  ]
  for (let row = 2; row <= ROWS; row += 1) {
    chunks.push(
      `<row r="${row}">${letters.map((l) => `<c r="${l}${row}"><v>${row}</v></c>`).join('')}</row>`,
    )
  }
  chunks.push('</sheetData></worksheet>')
  zip.file('xl/worksheets/sheet1.xml', chunks.join(''))
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

interface HeaderState {
  a1: string | null
  ySplit: number
}

test.describe('sheets: frozen top row survives streamed scrolling', () => {
  test('in-session Freeze Top Row keeps the header painted past the loaded window', async () => {
    test.setTimeout(240_000)
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-freeze-stream-'))
    const fixture = join(scratch, 'freeze-stream.xlsx')
    await writeStreamingFixture(fixture)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-freeze-streaming',
      openFile: fixture,
    })
    try {
      const { app } = launched
      const sheets = await waitForPageWithUrl(app, '://sheets/')
      await sheets.waitForFunction(
        () => Boolean((window as unknown as Record<string, unknown>).__genofficeDebug),
        null,
        { timeout: 60_000 },
      )
      const headerState = (): Promise<HeaderState> =>
        sheets.evaluate(() => {
          const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
            univerAPI: {
              getActiveWorkbook(): {
                getActiveSheet(): {
                  getRange(
                    r: number,
                    c: number,
                    h: number,
                    w: number,
                  ): {
                    getCellData(): { v?: unknown } | null
                  }
                  getFreeze(): { ySplit: number }
                }
              }
            }
          }
          const ws = debug.univerAPI.getActiveWorkbook().getActiveSheet()
          const data = ws.getRange(0, 0, 1, 1).getCellData()
          return { a1: data?.v == null ? null : String(data.v), ySplit: ws.getFreeze().ySplit }
        })
      await expect(async () => {
        expect((await headerState()).a1).toBe('HEADER_A')
      }).toPass({ timeout: 60_000 })

      // Freeze the top row through the real ribbon path (View tab -> Freeze
      // Panes menu), so the set_freeze op and the mutation listener both run.
      await sheets.getByRole('button', { name: 'View', exact: true }).click()
      await sheets.locator('button[aria-label="Freeze Panes"]').click()
      await sheets.getByRole('listbox').getByText('Freeze Top Row').click()
      await expect(async () => {
        expect((await headerState()).ySplit).toBe(1)
      }).toPass({ timeout: 10_000 })

      // Jump far past the initial window; the advancing window evicts the old
      // one. Before the fix this null-wiped row 1 and A1 went blank.
      await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: {
            getActiveWorkbook(): {
              setActiveRange(range: unknown): void
              getActiveSheet(): {
                getRange(r: number, c: number): unknown
                scrollToCell(row: number, column: number): void
              }
            }
          }
        }
        const workbook = debug.univerAPI.getActiveWorkbook()
        const ws = workbook.getActiveSheet()
        workbook.setActiveRange(ws.getRange(59_999, 0))
        ws.scrollToCell(59_999, 0)
      })
      await sheets.waitForTimeout(4_000)
      await sheets.mouse.move(700, 400)
      await sheets.mouse.wheel(0, 2_000)

      // the frozen header keeps its content through the eviction
      await sheets.waitForTimeout(4_000)
      const after = await headerState()
      expect(after.ySplit).toBe(1)
      expect(after.a1).toBe('HEADER_A')
    } finally {
      await closeAndSaveVideo(launched, 'sheets-freeze-streaming')
    }
  })
})
