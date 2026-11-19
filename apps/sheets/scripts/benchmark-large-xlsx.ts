import { execFile } from 'node:child_process'
import { mkdir, writeFile } from 'node:fs/promises'
import { platform } from 'node:os'
import { resolve } from 'node:path'
import { promisify } from 'node:util'

import { z } from 'zod'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { workbookRangeResultSchema } from '../src/shared/desktop-api'

const execFileAsync = promisify(execFile)
const fixturePath = process.argv[2] ?? process.env.LARGE_XLSX_FIXTURE

const openResultSchema = z.object({
  sessionId: z.string().uuid(),
  name: z.string(),
  entryCount: z.number().int().nonnegative(),
  sheets: z
    .array(
      z.object({
        id: z.string(),
        name: z.string(),
        rowCount: z.number().int().positive(),
        columnCount: z.number().int().positive(),
      }),
    )
    .min(1),
})

async function main(): Promise<void> {
  if (!fixturePath) {
    throw new Error('Pass a large XLSX path or set LARGE_XLSX_FIXTURE.')
  }
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  const client = new XlsxSidecarClient(resolve('native/xlsx-engine/target/release', executable))
  let sessionId: string | null = null
  let peakSidecarRssKilobytes: number | null = null
  let sampling = false
  const memoryTimer = setInterval(() => {
    const pid = client.getProcessId()
    if (!pid || sampling || process.platform === 'win32') return
    sampling = true
    void readResidentMemory(pid)
      .then((rss) => {
        if (rss !== null) peakSidecarRssKilobytes = Math.max(peakSidecarRssKilobytes ?? 0, rss)
      })
      .finally(() => {
        sampling = false
      })
  }, 100)

  try {
    const openStarted = performance.now()
    const opened = openResultSchema.parse(await client.open(resolve(fixturePath)))
    const openMilliseconds = performance.now() - openStarted
    sessionId = opened.sessionId
    const largestSheet = opened.sheets.reduce((largest, sheet) =>
      sheet.rowCount * sheet.columnCount > largest.rowCount * largest.columnCount ? sheet : largest,
    )
    const endColumn = Math.min(49, largestSheet.columnCount - 1)
    const firstRange = {
      startRow: 0,
      endRow: Math.min(99, largestSheet.rowCount - 1),
      startColumn: 0,
      endColumn,
    }
    const firstViewportStarted = performance.now()
    const firstViewport = workbookRangeResultSchema.parse(
      await client.readRange({
        sessionId,
        sheetId: largestSheet.id,
        range: firstRange,
      }),
    )
    const firstViewportMilliseconds = performance.now() - firstViewportStarted

    const bottomRange = {
      startRow: Math.max(0, largestSheet.rowCount - 100),
      endRow: largestSheet.rowCount - 1,
      startColumn: 0,
      endColumn,
    }
    const fullIndexStarted = performance.now()
    let bottomResult = firstViewport
    for (let attempt = 0; attempt < 120; attempt += 1) {
      bottomResult = workbookRangeResultSchema.parse(
        await client.readRange({
          sessionId,
          sheetId: largestSheet.id,
          range: bottomRange,
        }),
      )
      if (
        bottomResult.indexingComplete ||
        (bottomResult.indexedThroughRow ?? -1) >= bottomRange.endRow
      ) {
        break
      }
    }
    const fullIndexMilliseconds = performance.now() - fullIndexStarted
    if ((bottomResult.indexedThroughRow ?? -1) < bottomRange.endRow) {
      throw new Error('Large worksheet did not finish indexing within the benchmark timeout.')
    }

    const report = {
      platform: platform(),
      nodeVersion: process.version,
      fixture: resolve(fixturePath),
      workbook: opened.name,
      entryCount: opened.entryCount,
      sheetCount: opened.sheets.length,
      largestSheet,
      openMilliseconds: round(openMilliseconds),
      firstViewportMilliseconds: round(firstViewportMilliseconds),
      firstViewportCells: firstViewport.cells.length,
      fullIndexMilliseconds: round(fullIndexMilliseconds),
      bottomViewportCells: bottomResult.cells.length,
      peakSidecarRssKilobytes,
    }
    await mkdir(resolve('reports'), { recursive: true })
    await writeFile(
      resolve('reports/large-workbook-benchmark.json'),
      `${JSON.stringify(report, null, 2)}\n`,
    )
    process.stdout.write(`${JSON.stringify(report, null, 2)}\n`)
  } finally {
    clearInterval(memoryTimer)
    if (sessionId) await client.close(sessionId)
    client.stop()
  }
}

async function readResidentMemory(pid: number): Promise<number | null> {
  try {
    const { stdout } = await execFileAsync('ps', ['-o', 'rss=', '-p', String(pid)])
    const value = Number(stdout.trim())
    return Number.isFinite(value) ? value : null
  } catch {
    return null
  }
}

function round(value: number): number {
  return Math.round(value * 100) / 100
}

void main()
