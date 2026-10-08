import { resolve } from 'node:path'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { workbookFindCellsResultSchema, workbookRangeResultSchema } from '../src/shared/desktop-api'

const fixture = process.argv[2]
const needle = (process.argv[3] ?? 'needle').toLowerCase()
if (!fixture) throw new Error('usage: tsx bench-find-cells.ts <xlsx> [needle]')
const fixturePath: string = fixture
const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
const client = new XlsxSidecarClient(resolve('native/xlsx-engine/target/release', executable))
const BATCH_CELLS = 18_000
const SCAN_CAP = 400_000

async function main() {
  const opened = (await client.open(resolve(fixturePath))) as {
    sessionId: string
    sheets: { id: string; rowCount: number; columnCount: number }[]
  }
  const sheet = opened.sheets[0]!
  const cells = sheet.rowCount * sheet.columnCount
  console.log(`sheet ${sheet.id}: ${sheet.rowCount} x ${sheet.columnCount} = ${cells} cells`)

  const rangeScan = async (cap: number) => {
    const started = performance.now()
    const batchRows = Math.max(1, Math.floor(BATCH_CELLS / sheet.columnCount))
    let hits = 0
    let scanned = 0
    let bytes = 0
    for (let startRow = 0; startRow < sheet.rowCount; startRow += batchRows) {
      if (scanned >= cap) break
      const endRow = Math.min(startRow + batchRows - 1, sheet.rowCount - 1)
      for (;;) {
        const raw = await client.readRange({
          sessionId: opened.sessionId,
          sheetId: sheet.id,
          range: { startRow, endRow, startColumn: 0, endColumn: sheet.columnCount - 1 },
        })
        bytes += JSON.stringify(raw).length
        const result = workbookRangeResultSchema.parse(raw)
        if (!result.indexingComplete && (result.indexedThroughRow ?? -1) < endRow) continue
        for (const cell of result.cells) {
          const text = cell.value === null ? '' : String(cell.value).toLowerCase()
          if (text.includes(needle) || cell.formula?.toLowerCase().includes(needle)) hits += 1
        }
        break
      }
      scanned += (endRow - startRow + 1) * sheet.columnCount
    }
    return {
      ms: Math.round(performance.now() - started),
      hits,
      scanned,
      mb: (bytes / 1e6).toFixed(1),
    }
  }

  const sidecarFind = async () => {
    const started = performance.now()
    let hits = 0
    let pages = 0
    let resumeAt: unknown
    for (;;) {
      const result = workbookFindCellsResultSchema.parse(
        await client.findCells({
          sessionId: opened.sessionId,
          sheetId: sheet.id,
          query: needle,
          matchCase: false,
          matchEntireCell: false,
          lookIn: 'both',
          wildcards: false,
          ...(resumeAt === undefined ? {} : { resumeAt }),
          limit: 100_000,
        }),
      )
      pages += 1
      hits += result.matches.length
      if (result.complete) break
      resumeAt = result.nextCursor
    }
    return { ms: Math.round(performance.now() - started), hits, pages }
  }

  console.log('before (cold index), range paging capped at 400k cells:', await rangeScan(SCAN_CAP))
  console.log('before (warm), range paging capped at 400k cells:', await rangeScan(SCAN_CAP))
  console.log('before (warm), range paging over the whole sheet:', await rangeScan(Infinity))
  console.log('after (warm), find_cells:', await sidecarFind())
  await client.close(opened.sessionId)
  const reopened = (await client.open(resolve(fixturePath))) as { sessionId: string }
  opened.sessionId = reopened.sessionId
  console.log('after (cold index), find_cells:', await sidecarFind())
  await client.close(opened.sessionId)
  client.stop()
}

main().catch((error) => {
  console.error(error)
  client.stop()
  process.exit(1)
})
