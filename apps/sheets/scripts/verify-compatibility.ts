import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'

import { applyPlanToXlsx } from '../src/gateway/xlsx-gateway'
import type { CellState, ChangePlan } from '../src/domain/workbook.types'

interface CorpusCase {
  fixture: string
  sheetName: string
  address: string
  before: CellState
  after: CellState
}

/// One surgical cell edit per fixture; every other package entry must
/// survive byte-identical, including binary and exotic parts.
const CORPUS: readonly CorpusCase[] = [
  {
    fixture: 'compatibility-basic.xlsx',
    sheetName: 'Sheet1',
    address: 'A1',
    before: { value: 'Old' },
    after: { value: 'Verified' },
  },
  {
    fixture: 'compatibility-edit.xlsx',
    sheetName: 'Data',
    address: 'C1',
    before: { value: 5 },
    after: { value: 6 },
  },
  {
    fixture: 'compatibility-structure.xlsx',
    sheetName: 'Data',
    address: 'A1',
    before: { value: 1 },
    after: { value: 99 },
  },
  {
    fixture: 'compatibility-sheets.xlsx',
    sheetName: 'Data',
    address: 'A1',
    before: { value: 1 },
    after: { value: 99 },
  },
  {
    fixture: 'compatibility-kitchen-sink.xlsx',
    sheetName: 'Data',
    address: 'A1',
    before: { value: 1 },
    after: { value: 99 },
  },
]

interface FixtureReport {
  fixture: string
  passed: boolean
  error?: string
  touchedEntries: string[]
  changedEntries: string[]
  removedEntries: string[]
  unexpectedChanges: string[]
  preservedEntryCount: number
}

async function verifyCase(entry: CorpusCase): Promise<FixtureReport> {
  const source = await readFile(resolve('fixtures/generated', entry.fixture))
  const plan: ChangePlan = {
    transactionId: `compatibility-check-${entry.fixture}`,
    baseRevision: 0,
    cellChanges: [
      {
        sheetId: 'sheet-under-test',
        address: entry.address,
        before: entry.before,
        after: entry.after,
      },
    ],
    sheetRenames: [],
    structuralChanges: [],
    formatChanges: [],
    warnings: [],
  }
  try {
    const mutation = await applyPlanToXlsx(source, plan, {
      'sheet-under-test': entry.sheetName,
    })
    const beforeByPath = new Map(mutation.beforeEntries.map((e) => [e.path, e.sha256]))
    const afterPaths = new Set(mutation.afterEntries.map((e) => e.path))
    const changedEntries = mutation.afterEntries
      .filter((e) => beforeByPath.get(e.path) !== e.sha256)
      .map((e) => e.path)
    // entries that vanished from the package are preservation failures too —
    // hash comparison alone never sees them because they have no after-entry
    const removedEntries = mutation.beforeEntries
      .filter((e) => !afterPaths.has(e.path))
      .map((e) => e.path)
    const unexpectedChanges = [...changedEntries, ...removedEntries].filter(
      (path) => !mutation.touchedEntries.includes(path),
    )
    return {
      fixture: entry.fixture,
      passed: unexpectedChanges.length === 0 && changedEntries.length > 0,
      touchedEntries: [...mutation.touchedEntries],
      changedEntries,
      removedEntries,
      unexpectedChanges,
      preservedEntryCount: mutation.afterEntries.length - changedEntries.length,
    }
  } catch (error) {
    return {
      fixture: entry.fixture,
      passed: false,
      error: error instanceof Error ? error.message : String(error),
      touchedEntries: [],
      changedEntries: [],
      removedEntries: [],
      unexpectedChanges: [],
      preservedEntryCount: 0,
    }
  }
}

async function main(): Promise<void> {
  const fixtures: FixtureReport[] = []
  for (const entry of CORPUS) fixtures.push(await verifyCase(entry))
  const report = {
    passed: fixtures.every((f) => f.passed),
    fixtureCount: fixtures.length,
    fixtures,
  }

  await mkdir(resolve('reports'), { recursive: true })
  await writeFile(resolve('reports/compatibility.json'), `${JSON.stringify(report, null, 2)}\n`)
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`)
  if (!report.passed) process.exitCode = 1
}

void main()
