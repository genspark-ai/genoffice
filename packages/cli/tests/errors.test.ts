import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { classifyOpError } from '../src/op-errors'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const REPO = resolve(__dirname, '../../..')
const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')
const PPTX = join(REPO, 'packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const XLSX = join(REPO, 'apps/sheets/fixtures/generated/compatibility-basic.xlsx')

const TEXT = [{ runs: [{ text: 'x' }] }]

function copyOf(src: string, name: string): { dir: string; file: string } {
  const dir = tempDir()
  const file = join(dir, name)
  writeFileSync(file, readFileSync(src))
  return { dir, file }
}

describe('structured errors', () => {
  it('names the reason for command-line mistakes', async () => {
    const unknown = (await run(['frobnicate', '--json'])).json()
    expect(unknown).toMatchObject({ status: 'error', code: 1, error: 'unknown_command' })
    expect(unknown.suggestion).toContain('genoffice help')
    expect(unknown.detail.commands).toContain('info')

    const option = (await run(['info', DOCX, '--dry-run', '--json'])).json()
    expect(option).toMatchObject({ error: 'unknown_option' })
    expect(option.suggestion).toContain('genoffice help info')

    expect((await run(['info', '--json'])).json().error).toBe('missing_argument')
    expect((await run(['convert', DOCX, '--json'])).json().error).toBe('missing_argument')
    expect((await run(['convert', DOCX, '--to', 'pptx', '--json'])).json()).toMatchObject({
      error: 'unsupported',
    })
    expect((await run(['guide', '--json'])).json().error).toBe('missing_argument')
    expect((await run(['guide', 'foo', '--json'])).json().error).toBe('invalid_argument')
    expect((await run(['docs', 'reed', DOCX, '--json'])).json().error).toBe('invalid_argument')
    expect((await run(['docs', '--json'])).json().error).toBe('missing_argument')
    expect((await run(['sheet', '--json'])).json().error).toBe('missing_argument')
  })

  it('separates file errors that share exit code 2', async () => {
    const missing = (await run(['info', '/nonexistent/x.docx', '--json'])).json()
    expect(missing).toMatchObject({ code: 2, error: 'file_not_found' })

    const { dir } = copyOf(DOCX, 'in.docx')
    const csv = join(dir, 'd.csv')
    writeFileSync(csv, 'a,b\n1,2\n')
    expect((await run(['convert', csv, '--to', 'xlsx', '--json'])).code).toBe(0)
    const exists = (await run(['convert', csv, '--to', 'xlsx', '--json'])).json()
    expect(exists).toMatchObject({ code: 2, error: 'output_exists' })
    expect(exists.suggestion).toContain('--force')
  })

  it('prints the suggestion as a hint line in human mode', async () => {
    const r = await run(['frobnicate'])
    expect(r.code).toBe(1)
    expect(r.stderr).toContain('genoffice: unknown command: frobnicate')
    expect(r.stderr).toContain('hint: run `genoffice help`')
  })

  it('lifts ranges, ids and usage lines out of guided op errors', () => {
    const range = classifyOpError(
      0,
      'setText',
      'op "setText": slide index 99 is out of range (0-4).',
    )
    expect(range).toMatchObject({ reason: 'out_of_range', valid_range: [0, 4] })

    const target = classifyOpError(
      2,
      'setText',
      'op "setText": no element "e_999" on slide 0. Available: [sp_0 (e_2), sp_1 (e_3)].\nUsage: setText {target, text} Nothing was applied (atomic) — fix this op and resend the whole transaction.',
    )
    expect(target).toMatchObject({
      reason: 'target_not_found',
      available: ['sp_0 (e_2)', 'sp_1 (e_3)'],
      usage: 'Usage: setText {target, text}',
    })

    const unknown = classifyOpError(
      1,
      'frob',
      'unknown op "frob". Supported ops: [setText, addElement].',
    )
    expect(unknown).toMatchObject({ reason: 'unknown_op', supported: ['setText', 'addElement'] })

    expect(classifyOpError(0, 'delete_blocks', 'No matching blocks for the target.').reason).toBe(
      'target_not_found',
    )
    expect(
      classifyOpError(
        0,
        'replace_blocks',
        'block index invalid or out of range (the document has 41 blocks); call get_document_context',
      ),
    ).toMatchObject({ reason: 'out_of_range', valid_range: [0, 40] })

    expect(classifyOpError(0, 'setFill', 'op "setFill": "color" must be a hex color.').reason).toBe(
      'op_rejected',
    )

    const group = classifyOpError(
      0,
      'ungroup',
      'op "ungroup": no group "g_9" on slide 1. Available groups: [g_1, g_2].',
    )
    expect(group).toMatchObject({ reason: 'target_not_found', available: ['g_1', 'g_2'] })
    expect(group.did_you_mean).toBe('g_1')
  })

  it('reports slides op failures with reason, range and suggestion', async () => {
    const { dir, file } = copyOf(PPTX, 'deck.pptx')
    const ops = join(dir, 'ops.json')

    writeFileSync(
      ops,
      JSON.stringify([{ op: 'setText', target: { slide: 99, el: 'e_1' }, paragraphs: TEXT }]),
    )
    const range = (await run(['slides', 'apply', file, '--ops', ops, '--json'])).json()
    expect(range).toMatchObject({ code: 1, error: 'out_of_range' })
    expect(range.detail.failures[0]).toMatchObject({
      index: 0,
      op: 'setText',
      reason: 'out_of_range',
    })
    expect(range.detail.failures[0].valid_range[0]).toBe(0)
    expect(range.suggestion).toContain('between 0 and')

    writeFileSync(
      ops,
      JSON.stringify([{ op: 'setText', target: { slide: 0, el: 'e_nope' }, paragraphs: TEXT }]),
    )
    const target = (await run(['slides', 'apply', file, '--ops', ops, '--json'])).json()
    expect(target).toMatchObject({ error: 'target_not_found' })
    expect(target.detail.failures[0].available.length).toBeGreaterThan(0)
    expect(target.suggestion).toContain('genoffice slides read')

    writeFileSync(ops, JSON.stringify([{ op: 'frobnicate', target: { slide: 0 } }]))
    const unknown = (await run(['slides', 'apply', file, '--ops', ops, '--json'])).json()
    expect(unknown).toMatchObject({ error: 'unknown_op' })
    expect(unknown.detail.failures[0].supported).toContain('setText')

    const flag = (await run(['slides', 'read', file, '--slide', '99', '--json'])).json()
    expect(flag).toMatchObject({ error: 'out_of_range' })
    expect(flag.detail.valid_range[0]).toBe(0)
    expect(readFileSync(file).equals(readFileSync(PPTX))).toBe(true)
  })

  it('reports docs op failures the same way', async () => {
    const { dir, file } = copyOf(DOCX, 'doc.docx')
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'frobnicate', target: { blockIndexes: [0] } }]))
    const unknown = (await run(['docs', 'apply', file, '--ops', ops, '--json'])).json()
    expect(unknown).toMatchObject({ code: 1, error: 'unknown_op' })
    expect(unknown.suggestion).toContain('genoffice guide docs')

    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'replace_blocks', startBlockIndex: 9999, endBlockIndex: 9999, html: '<p>x</p>' },
      ]),
    )
    const blocks = (await run(['docs', 'apply', file, '--ops', ops, '--json'])).json()
    expect(blocks.error).toBe('out_of_range')
    expect(blocks.detail.failures[0]).toMatchObject({
      reason: 'out_of_range',
      valid_range: [0, 40],
    })
    expect(blocks.suggestion).toContain('genoffice docs read')

    const range = (await run(['docs', 'read', file, '--range', '5000-6000', '--json'])).json()
    expect(range).toMatchObject({ error: 'out_of_range' })
    expect(range.detail.valid_range[0]).toBe(0)
    expect((await run(['docs', 'read', file, '--range', 'abc', '--json'])).json().error).toBe(
      'invalid_argument',
    )
  })

  it.skipIf(!xlsxSidecarPath())('reports sheet errors with the sheet list', async () => {
    const { dir, file } = copyOf(XLSX, 'book.xlsx')
    const missing = (await run(['sheet', 'read', file, '--sheet', 'Nope', '--json'])).json()
    expect(missing).toMatchObject({ error: 'sheet_not_found' })
    expect(missing.detail.sheets.length).toBeGreaterThan(0)

    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'frobnicate' }]))
    const unknown = (await run(['sheet', 'apply', file, '--ops', ops, '--json'])).json()
    expect(unknown).toMatchObject({ error: 'unknown_op' })
    expect(unknown.detail.supported).toContain('set_cell')
    expect(unknown.suggestion).toContain('genoffice guide sheets')
  })
})

describe('did-you-mean for ops', () => {
  it('suggests the nearest op and element id', () => {
    const op = classifyOpError(
      0,
      'setTxt',
      'unknown op "setTxt". Supported ops: [setText, setFill].',
    )
    expect(op.did_you_mean).toBe('setText')
    const el = classifyOpError(
      0,
      'setText',
      'op "setText": no element "e_22" on slide 0. Available: [sp_0 (e_2), sp_1 (e_3)].',
    )
    expect(el.did_you_mean).toBe('e_2')
  })
})
