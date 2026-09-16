import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const sidecar = Boolean(xlsxSidecarPath())

function opsFile(dir: string, name: string, ops: unknown[]): string {
  const path = join(dir, name)
  writeFileSync(path, JSON.stringify(ops))
  return path
}

const textbox = (slide: number, text: string) => ({
  op: 'addElement',
  target: { slide },
  kind: 'textbox',
  offset: { x: 914400, y: 914400, cx: 3657600, cy: 914400 },
  paragraphs: [{ runs: [{ text }] }],
})

async function deck(dir: string): Promise<string> {
  const ops = opsFile(dir, 'create.json', [textbox(0, 'Hello')])
  const out = join(dir, 'deck.pptx')
  expect((await run(['create', '--type', 'pptx', '--ops', ops, '--out', out])).code).toBe(0)
  return out
}

async function book(dir: string): Promise<string> {
  const table = join(dir, 'table.json')
  writeFileSync(
    table,
    JSON.stringify([
      ['item', 'qty'],
      ['Apple', 2],
    ]),
  )
  const out = join(dir, 'book.xlsx')
  expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
  return out
}

async function document(dir: string): Promise<string> {
  const md = join(dir, 'a.md')
  writeFileSync(md, '# Title\n\nBody.\n\nMore.\n')
  const docx = join(dir, 'a.docx')
  expect((await run(['create', '--type', 'docx', '--from', md, '--out', docx])).code).toBe(0)
  return docx
}

describe('batch modes on slides apply', () => {
  // op 1 targets a slide that does not exist; op 2 is fine again
  const ops = (dir: string) =>
    opsFile(dir, 'ops.json', [textbox(0, 'one'), textbox(7, 'nope'), textbox(0, 'three')])

  it('atomic: writes nothing, reports the counts', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const before = readFileSync(pptx)
    const r = await run(['slides', 'apply', pptx, '--ops', ops(dir), '--json'])
    expect(r.code).toBe(1)
    expect(r.json().status).toBe('error')
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 0, failed: 1, skipped: 2 })
    expect(readFileSync(pptx).equals(before)).toBe(true)
    const text = await run(['slides', 'apply', pptx, '--ops', ops(dir)])
    expect(text.stderr).toContain('batch: 0 of 3 ops applied, 1 failed, 2 skipped')
  })

  it('best effort: applies ops 0 and 2, answers partial', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const r = await run(['slides', 'apply', pptx, '--ops', ops(dir), '--best-effort', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().status).toBe('partial')
    expect(r.json().summary).toBe('applied 2 of 3 ops to deck.pptx; 1 failed')
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 2, failed: 1, skipped: 0 })
    expect(r.json().detail.failures.map((f: { index: number }) => f.index)).toEqual([1])
    const read = await run(['slides', 'read', pptx, '--full', '--json'])
    const texts = JSON.stringify(read.json().detail.pages[0])
    expect(texts).toContain('one')
    expect(texts).toContain('three')
  })

  it('stop on error: applies op 0 only, skips op 2', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const r = await run(['slides', 'apply', pptx, '--ops', ops(dir), '--stop-on-error', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().status).toBe('partial')
    expect(r.json().summary).toBe('applied 1 of 3 ops to deck.pptx; 1 failed, 1 skipped')
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 1, failed: 1, skipped: 1 })
    const read = await run(['slides', 'read', pptx, '--full', '--json'])
    const texts = JSON.stringify(read.json().detail.pages[0])
    expect(texts).toContain('one')
    expect(texts).not.toContain('three')
  })

  it('a clean batch stays status ok with the counts', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const clean = opsFile(dir, 'clean.json', [textbox(0, 'a'), textbox(0, 'b')])
    const r = await run(['slides', 'apply', pptx, '--ops', clean, '--best-effort', '--json'])
    expect(r.json().status).toBe('ok')
    expect(r.json().detail.batch).toEqual({ total: 2, applied: 2, failed: 0, skipped: 0 })
    expect(r.json().detail.failures).toBeUndefined()
  })

  it('refuses both modes at once and keeps --isolation per_op as an alias', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const both = await run([
      'slides',
      'apply',
      pptx,
      '--ops',
      ops(dir),
      '--best-effort',
      '--stop-on-error',
      '--json',
    ])
    expect(both.code).toBe(1)
    expect(both.json().error).toBe('invalid_argument')
    const alias = await run([
      'slides',
      'apply',
      pptx,
      '--ops',
      ops(dir),
      '--isolation',
      'per_op',
      '--json',
    ])
    expect(alias.json().status).toBe('partial')
    expect(alias.json().detail.batch.applied).toBe(2)
  })

  it('every op rejected is an error even in best effort', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const bad = opsFile(dir, 'bad.json', [textbox(7, 'x'), textbox(8, 'y')])
    const r = await run(['slides', 'apply', pptx, '--ops', bad, '--best-effort', '--json'])
    expect(r.code).toBe(1)
    expect(r.json().detail.batch).toEqual({ total: 2, applied: 0, failed: 2, skipped: 0 })
    const dry = await run([
      'slides',
      'apply',
      pptx,
      '--ops',
      bad,
      '--best-effort',
      '--dry-run',
      '--json',
    ])
    expect(dry.code).toBe(1)
    expect(dry.json().detail.batch.applied).toBe(0)
  })
})

describe.skipIf(!sidecar)('batch modes on sheet apply', () => {
  const ops = (dir: string) =>
    opsFile(dir, 'ops.json', [
      { op: 'set_cell', sheet: 'table', address: 'C1', value: 'one' },
      { op: 'set_cell', sheet: 'Nope', address: 'A1', value: 'x' },
      { op: 'set_cell', sheet: 'table', address: 'C2', value: 'three' },
    ])
  const cells = async (xlsx: string) => {
    const r = await run(['sheet', 'read', xlsx, '--range', 'C1:C2', '--json'])
    return JSON.stringify(r.json().detail)
  }

  it('atomic: nothing written, counts in the error', async () => {
    const dir = tempDir()
    const xlsx = await book(dir)
    const before = readFileSync(xlsx)
    const r = await run(['sheet', 'apply', xlsx, '--ops', ops(dir), '--json'])
    expect(r.code).toBe(1)
    expect(r.json().error).toBe('sheet_not_found')
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 0, failed: 1, skipped: 2 })
    expect(r.json().detail.failures[0]).toMatchObject({ index: 1, op: 'set_cell' })
    expect(readFileSync(xlsx).equals(before)).toBe(true)
  })

  it('best effort: drops the bad op and writes the other two', async () => {
    const dir = tempDir()
    const xlsx = await book(dir)
    const r = await run(['sheet', 'apply', xlsx, '--ops', ops(dir), '--best-effort', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().status).toBe('partial')
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 2, failed: 1, skipped: 0 })
    expect(r.json().detail.failures[0].index).toBe(1)
    const text = await cells(xlsx)
    expect(text).toContain('one')
    expect(text).toContain('three')
  })

  it('stop on error: writes only what came before the rejection', async () => {
    const dir = tempDir()
    const xlsx = await book(dir)
    const r = await run(['sheet', 'apply', xlsx, '--ops', ops(dir), '--stop-on-error', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().status).toBe('partial')
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 1, failed: 1, skipped: 1 })
    const text = await cells(xlsx)
    expect(text).toContain('one')
    expect(text).not.toContain('three')
  })

  it("every op rejected keeps the first rejection's own reason, hints and detail", async () => {
    const dir = tempDir()
    const xlsx = await book(dir)
    const bad = opsFile(dir, 'bad.json', [
      { op: 'set_cell', sheet: 'Nope', address: 'A1', value: 'x' },
      { op: 'frobnicate' },
    ])
    const r = await run(['sheet', 'apply', xlsx, '--ops', bad, '--best-effort', '--json'])
    expect(r.code).toBe(1)
    expect(r.json().error).toBe('sheet_not_found')
    expect(r.json().message).toMatch(/^no ops were applied: ops\[0\]: sheet not found: Nope/)
    expect(r.json().suggestion).toBe('use one of the names in detail.sheets')
    expect(r.json().detail.sheets).toEqual(['table'])
    expect(r.json().detail.batch).toEqual({ total: 2, applied: 0, failed: 2, skipped: 0 })
    expect(
      r
        .json()
        .detail.failures.map((f: { index: number }) => f.index)
        .sort(),
    ).toEqual([0, 1])
  })

  it('a rejection the DSL cannot pin to one op stays fatal in best effort', async () => {
    const dir = tempDir()
    const xlsx = await book(dir)
    // structural and content ops mix since sheet apply replays shifts; defined names still refuse
    const mixed = opsFile(dir, 'mixed.json', [
      { op: 'insert_rows', sheet: 'table', row: 1, count: 1 },
      { op: 'add_defined_name', name: 'Later', ref: 'table!$B$2' },
    ])
    const r = await run(['sheet', 'apply', xlsx, '--ops', mixed, '--best-effort', '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('cannot share a batch')
  })

  it('dry run reports the counts without writing', async () => {
    const dir = tempDir()
    const xlsx = await book(dir)
    const before = readFileSync(xlsx)
    const r = await run([
      'sheet',
      'apply',
      xlsx,
      '--ops',
      ops(dir),
      '--best-effort',
      '--dry-run',
      '--json',
    ])
    expect(r.json().status).toBe('partial')
    expect(r.json().summary).toContain('2 of 3 ops validated')
    expect(readFileSync(xlsx).equals(before)).toBe(true)
  })
})

describe('batch modes on docs apply', () => {
  const ops = (dir: string) =>
    opsFile(dir, 'ops.json', [
      { op: 'findReplace', find: 'Body', replace: 'Text' },
      { op: 'deleteBlocks', target: { blockIndexes: [999] } },
      { op: 'findReplace', find: 'More', replace: 'Extra' },
    ])
  const text = async (docx: string) => {
    const r = await run(['docs', 'read', docx, '--full', '--json'])
    return JSON.stringify(r.json().detail.items)
  }

  it('atomic: nothing written, counts in the error', async () => {
    const dir = tempDir()
    const docx = await document(dir)
    const before = readFileSync(docx)
    const r = await run(['docs', 'apply', docx, '--ops', ops(dir), '--json'])
    expect(r.code).toBe(1)
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 0, failed: 1, skipped: 2 })
    expect(readFileSync(docx).equals(before)).toBe(true)
  })

  it('best effort: applies ops 0 and 2', async () => {
    const dir = tempDir()
    const docx = await document(dir)
    const r = await run(['docs', 'apply', docx, '--ops', ops(dir), '--best-effort', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().status).toBe('partial')
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 2, failed: 1, skipped: 0 })
    expect(r.json().detail.failures[0]).toMatchObject({ index: 1, op: 'deleteBlocks' })
    expect(r.json().detail.results.map((x: { index: number }) => x.index)).toEqual([0, 2])
    const t = await text(docx)
    expect(t).toContain('Text')
    expect(t).toContain('Extra')
  })

  it('stop on error: applies op 0 only', async () => {
    const dir = tempDir()
    const docx = await document(dir)
    const r = await run(['docs', 'apply', docx, '--ops', ops(dir), '--stop-on-error', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.batch).toEqual({ total: 3, applied: 1, failed: 1, skipped: 1 })
    const t = await text(docx)
    expect(t).toContain('Text')
    expect(t).toContain('More')
    expect(t).not.toContain('Extra')
  })
})
