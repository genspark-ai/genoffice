import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const sidecar = Boolean(xlsxSidecarPath())

async function sheetPart(path: string, name = 'xl/worksheets/sheet1.xml'): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file(name)!.async('string')
}

/** Rewrites one package part, so a fixture can carry a cache no engine wrote. */
async function rewritePart(
  path: string,
  name: string,
  mutate: (xml: string) => string,
): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  const xml = await zip.file(name)!.async('string')
  zip.file(name, mutate(xml))
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

/**
 * The issue's repro: A2=3 and B2="=A2*2" evaluated on create (cached 6), so a
 * value-only edit of A2 leaves B2's <v> at the old number.
 */
async function issueWorkbook(dir: string): Promise<string> {
  const table = join(dir, 'table.json')
  writeFileSync(
    table,
    JSON.stringify([
      [null, null],
      [3, '=A2*2'],
    ]),
  )
  const out = join(dir, 'book.xlsx')
  const r = await run(['create', '--type', 'xlsx', '--from', table, '--out', out, '--json'])
  expect(r.code).toBe(0)
  return out
}

async function createWorkbook(dir: string, name: string, table: unknown): Promise<string> {
  const spec = join(dir, `${name}.json`)
  writeFileSync(spec, JSON.stringify(table))
  const out = join(dir, `${name}.xlsx`)
  const r = await run(['create', '--type', 'xlsx', '--from', spec, '--out', out, '--json'])
  expect(r.code).toBe(0)
  return out
}

describe('sheet apply refreshes dependent formula caches (#2304)', () => {
  it.skipIf(!sidecar)('recalculates a dependent formula after a --cells value edit', async () => {
    const dir = tempDir()
    const out = await issueWorkbook(dir)
    expect(await sheetPart(out)).toContain('<v>6</v>')
    const cells = join(dir, 'cells.json')
    writeFileSync(cells, JSON.stringify([{ cell: 'A2', value: 5 }]))

    const applied = await run(['sheet', 'apply', out, '--cells', cells, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail).toMatchObject({ cells: 1, formulas: 0, cached_values: true })

    const read = await run(['sheet', 'read', out, '--range', 'A2:B2', '--json'])
    expect(read.code).toBe(0)
    expect(read.json().detail.rows[0]).toEqual([5, 10])
    const xml = await sheetPart(out)
    expect(xml).toMatch(/<c r="B2"[^>]*><f>A2\*2<\/f><v>10<\/v><\/c>/)
  })

  it.skipIf(!sidecar)('recalculates a dependent formula after an --ops set_cell edit', async () => {
    const dir = tempDir()
    const out = await issueWorkbook(dir)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_cell', address: 'A2', value: 5 }]))

    const applied = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail.cached_values).toBe(true)

    const read = await run(['sheet', 'read', out, '--range', 'A2:B2', '--json'])
    expect(read.json().detail.rows[0]).toEqual([5, 10])
    expect(await sheetPart(out)).toContain('<f>A2*2</f><v>10</v>')
  })

  it.skipIf(!sidecar)('recalculates across sheets and through a dependency chain', async () => {
    const dir = tempDir()
    const out = await createWorkbook(dir, 'chain', {
      sheets: [
        { name: 'Data', rows: [[3]] },
        // B1 depends on A1, which depends on another sheet
        { name: 'Calc', rows: [['=Data!A1*2', '=A1+1']] },
      ],
    })
    expect(await sheetPart(out, 'xl/worksheets/sheet2.xml')).toContain('<v>7</v>')

    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_cell', sheet: 'Data', address: 'A1', value: 5 }]))
    const applied = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail.cached_values).toBe(true)

    const calc = await sheetPart(out, 'xl/worksheets/sheet2.xml')
    expect(calc).toMatch(/<c r="A1"[^>]*><f>Data!A1\*2<\/f><v>10<\/v><\/c>/)
    expect(calc).toMatch(/<c r="B1"[^>]*><f>A1\+1<\/f><v>11<\/v><\/c>/)
    const read = await run(['sheet', 'read', out, '--sheet', 'Calc', '--range', 'A1:B1', '--json'])
    expect(read.json().detail.rows[0]).toEqual([10, 11])
  })

  it.skipIf(!sidecar)(
    'leaves unchanged formula caches byte-identical (changed-only churn)',
    async () => {
      const dir = tempDir()
      const out = await createWorkbook(dir, 'churn', [[3, '=A1*2', null, '=1+1']])
      // Rewrite the unrelated cache with extra precision: the engine computes 2,
      // so a changed-only refresh must leave the file's own spelling alone.
      await rewritePart(out, 'xl/worksheets/sheet1.xml', (xml) =>
        xml.replace(
          /(<c r="D1"[^>]*><f>1\+1<\/f><v>)2(<\/v>)/,
          (_match, before: string, after: string) => `${before}2.0${after}`,
        ),
      )
      expect(await sheetPart(out)).toContain('<v>2.0</v>')

      const cells = join(dir, 'cells.json')
      writeFileSync(cells, JSON.stringify([{ cell: 'A1', value: 5 }]))
      const applied = await run(['sheet', 'apply', out, '--cells', cells, '--json'])
      expect(applied.code).toBe(0)
      expect(applied.json().detail.cached_values).toBe(true)

      const xml = await sheetPart(out)
      // the dependent got the new value, the unrelated formula kept its bytes
      expect(xml).toContain('<f>A1*2</f><v>10</v>')
      expect(xml).toContain('<f>1+1</f><v>2.0</v>')
    },
  )

  it.skipIf(!sidecar)(
    'reports cached_values true when no formula depends on the edit',
    async () => {
      const dir = tempDir()
      const out = await createWorkbook(dir, 'unrelated', [[3, '=A1*2']])
      const cells = join(dir, 'cells.json')
      writeFileSync(cells, JSON.stringify([{ cell: 'C1', value: 9 }]))

      const applied = await run(['sheet', 'apply', out, '--cells', cells, '--json'])
      expect(applied.code).toBe(0)
      expect(applied.json().detail.cached_values).toBe(true)
      expect(await sheetPart(out)).toContain('<f>A1*2</f><v>6</v>')
    },
  )

  it.skipIf(!sidecar)('keeps cached_values true for a workbook without formulas', async () => {
    const dir = tempDir()
    const out = await createWorkbook(dir, 'plain', [[1, 2]])
    const cells = join(dir, 'cells.json')
    writeFileSync(cells, JSON.stringify([{ cell: 'A1', value: 9 }]))

    const applied = await run(['sheet', 'apply', out, '--cells', cells, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail.cached_values).toBe(true)
    expect(applied.json().warnings ?? []).toEqual([])
  })

  it('reports dependents_not_recalculated when the engine cannot start', async () => {
    const dir = tempDir()
    const previous = process.env.XLSX_SIDECAR_PATH
    process.env.XLSX_SIDECAR_PATH = join(dir, 'missing-sidecar')
    try {
      // create leaves the formula uncached; apply must then say so instead of
      // claiming cached_values: true while B2 keeps its stale (absent) cache.
      const table = join(dir, 'table.json')
      writeFileSync(
        table,
        JSON.stringify([
          [null, null],
          [3, '=A2*2'],
        ]),
      )
      const out = join(dir, 'book.xlsx')
      expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
      // --ops does not need the engine to inspect the workbook the way --cells does
      const ops = join(dir, 'ops.json')
      writeFileSync(ops, JSON.stringify([{ op: 'set_cell', address: 'A2', value: 5 }]))
      const applied = await run(['sheet', 'apply', out, '--ops', ops, '--json'])
      expect(applied.code).toBe(0)
      expect(applied.json().detail.cached_values).toBe(false)
      const warnings = (applied.json().warnings ?? []) as { code: string; message: string }[]
      const dependent = warnings.find((w) => w.code === 'dependents_not_recalculated')
      expect(dependent, JSON.stringify(warnings)).toBeTruthy()
      expect(dependent!.message).toMatch(/depend|stale|recalculat/i)
    } finally {
      if (previous === undefined) delete process.env.XLSX_SIDECAR_PATH
      else process.env.XLSX_SIDECAR_PATH = previous
    }
  })
})
