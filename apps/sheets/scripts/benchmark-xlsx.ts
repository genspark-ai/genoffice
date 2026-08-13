import { mkdir, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'

import JSZip from 'jszip'

import { applyPlanToXlsx } from '../src/gateway/xlsx-gateway'
import type { ChangePlan } from '../src/domain/workbook.types'
import { buildCompatibilityFixture } from '../tests/fixture-builder'

const ROW_COUNT = 10_000
const COLUMN_COUNT = 5

async function main(): Promise<void> {
  const source = await buildLargeFixture()
  const plan: ChangePlan = {
    transactionId: 'benchmark',
    baseRevision: 0,
    cellChanges: [
      {
        sheetId: 'sheet-1',
        address: 'A1',
        before: { value: 1 },
        after: { value: 2 },
      },
    ],
    sheetRenames: [],
    structuralChanges: [],
    formatChanges: [],
    warnings: [],
  }
  const started = performance.now()
  const mutation = await applyPlanToXlsx(source, plan, { 'sheet-1': 'Sheet1' })
  const elapsedMilliseconds = performance.now() - started
  const report = {
    platform: process.platform,
    architecture: process.arch,
    nodeVersion: process.version,
    cells: ROW_COUNT * COLUMN_COUNT,
    sourceBytes: source.length,
    outputBytes: mutation.buffer.length,
    elapsedMilliseconds: Math.round(elapsedMilliseconds * 100) / 100,
    maximumResidentSetKilobytes: process.resourceUsage().maxRSS,
  }

  await mkdir(resolve('reports'), { recursive: true })
  await writeFile(resolve('reports/benchmark.json'), `${JSON.stringify(report, null, 2)}\n`)
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`)
}

async function buildLargeFixture(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await buildCompatibilityFixture())
  const rows = Array.from({ length: ROW_COUNT }, (_, rowIndex) => {
    const rowNumber = rowIndex + 1
    const cells = Array.from({ length: COLUMN_COUNT }, (_, columnIndex) => {
      const address = `${String.fromCharCode(65 + columnIndex)}${rowNumber}`
      return `<c r="${address}"><v>${rowNumber + columnIndex}</v></c>`
    }).join('')
    return `<row r="${rowNumber}">${cells}</row>`
  }).join('')
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

void main()
