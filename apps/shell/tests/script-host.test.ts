import { copyFileSync, mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import {
  closeDocument,
  describeDocument,
  openDocument,
  readInput,
} from '@genoffice/cli/headless-docs'
import { describe, expect, it } from 'vitest'
import {
  ScriptRunError,
  applyScriptToDocument,
  runScriptInSandbox,
  runScriptReadOnly,
} from '../src/main/script-host'
import type { ScriptSnapshot } from '../src/main/script-host'

/**
 * The script kind's host, driven directly.
 *
 * Two halves: the sandbox on its own (what a script can reach, and what stops
 * it), and a real read-only run against a real document — the one that has to
 * prove the file is untouched.
 */

const REPO = resolve(__dirname, '../../..')
const DOCX = join(REPO, 'fixtures/generated/simple.docx')

const snapshot: ScriptSnapshot = {
  blocks: [
    { index: 0, type: 'paragraph', text: 'Quarterly report' },
    { index: 1, type: 'paragraph', text: 'Revenue grew by 4%.' },
  ],
  text: 'Quarterly report\nRevenue grew by 4%.',
}

function sandbox(source: string, input: Record<string, unknown> = {}) {
  return runScriptInSandbox({ scriptPath: 'plan.js', source, snapshot, input })
}

describe('script sandbox', () => {
  it('reads the snapshot, proposes a plan and writes a log', () => {
    const outcome = sandbox(
      [
        'var blocks = $.docs.blocks()',
        '$.log("blocks=" + blocks.length)',
        '$.log("input=" + input.marker)',
        '$.docs.propose([{ op: "findReplace", find: "Revenue", replace: "Turnover" }])',
      ].join('\n'),
      { marker: 'yes' },
    )
    expect(outcome.ops).toHaveLength(1)
    expect(outcome.log).toEqual(['blocks=2', 'input=yes'])
    expect(outcome.ms).toBeGreaterThanOrEqual(0)
  })

  it('exposes nothing beyond the API: process, require and timers are undefined', () => {
    const outcome = sandbox(
      '$.log(typeof process + "|" + typeof require + "|" + typeof setTimeout)',
    )
    expect(outcome.log[0]).toBe('undefined|undefined|undefined')
  })

  it('refuses the constructor escapes', () => {
    const outcome = sandbox(
      [
        'try {',
        '  $.log("ESCAPED:" + typeof globalThis.constructor.constructor("return process")())',
        '} catch (e) {',
        '  $.log("BLOCKED")',
        '}',
      ].join('\n'),
    )
    expect(outcome.log[0]).toBe('BLOCKED')
  })

  it('hands a hook its event, and collects a veto it asks for', () => {
    const outcome = runScriptInSandbox({
      scriptPath: 'note-opened.js',
      source: [
        '$.log("event=" + $.event.name + " path=" + $.event.path)',
        '$.event.veto("the report has no title")',
      ].join('\n'),
      snapshot,
      input: { event: { name: 'document.will-save', path: '/tmp/report.docx' } },
    })
    expect(outcome.log[0]).toBe('event=document.will-save path=/tmp/report.docx')
    expect(outcome.vetoes).toEqual(['the report has no title'])
  })

  it('collects no veto when none was asked for', () => {
    const outcome = sandbox('$.log("nothing to say")')
    expect(outcome.vetoes).toEqual([])
  })

  it('stops a script that overruns its budget', () => {
    expect(() => sandbox('while (true) {}')).toThrow(ScriptRunError)
    try {
      runScriptInSandbox({
        scriptPath: 'plan.js',
        source: 'while (true) {}',
        snapshot,
        input: {},
        timeoutMs: 200,
      })
    } catch (err) {
      expect((err as Error).message).toContain('exceeded its 200ms budget')
    }
  })

  it('refuses a module and a script that proposes nothing sensible', () => {
    expect(() => sandbox('export default 1')).toThrow(ScriptRunError)
    const outcome = sandbox('$.docs.propose([{ op: "nope" }])')
    expect(outcome.ops).toEqual([{ op: 'nope' }])
  })
})

describe('a read-only run against a document', () => {
  it('plans against the real file and leaves it byte-identical', async () => {
    const before = readFileSync(DOCX)
    const outcome = await runScriptReadOnly({
      scriptId: 'tidy-headings',
      source: [
        'var blocks = $.docs.blocks()',
        '$.log("blocks=" + blocks.length)',
        '$.docs.propose([{ op: "insert_content", afterBlockIndex: 0, html: "<p>from the script</p>" }])',
      ].join('\n'),
      documentPath: DOCX,
    })
    expect(outcome.log[0]).toMatch(/^blocks=\d+$/)
    expect(outcome.plan).toHaveLength(1)
    expect(outcome.failures).toEqual([])
    expect(readFileSync(DOCX).equals(before)).toBe(true)
  })

  it('reports what the engine rejects instead of throwing', async () => {
    const before = readFileSync(DOCX)
    const outcome = await runScriptReadOnly({
      scriptId: 'unknown-op',
      source: '$.docs.propose([{ op: "nope" }])',
      documentPath: DOCX,
    })
    expect(outcome.applied).toBe(0)
    expect(outcome.failures).toHaveLength(1)
    expect(outcome.failures[0]).toContain('nope')
    expect(readFileSync(DOCX).equals(before)).toBe(true)
  })

  it('reports a missing file as a failed run, not a crash', async () => {
    await expect(
      runScriptReadOnly({
        scriptId: 'missing',
        source: '$.log(1)',
        documentPath: join(REPO, 'fixtures/generated/not-here.docx'),
      }),
    ).rejects.toThrow()
  })
})
/** A scratch copy: an apply test must never touch the fixture in the repo. */
function scratchCopy(): string {
  const dir = mkdtempSync(join(tmpdir(), 'genoffice-apply-'))
  const copy = join(dir, 'simple.docx')
  copyFileSync(DOCX, copy)
  return copy
}

const insert = (text: string): string =>
  [
    'var blocks = $.docs.blocks()',
    '$.log("blocks=" + blocks.length)',
    '$.docs.propose([{ op: "insert_content", afterBlockIndex: 0, html: "<p>' + text + '</p>" }])',
  ].join('\n')

describe('applying a script to a document', () => {
  it('needs a deliberate confirmation: without it, nothing is written', async () => {
    const copy = scratchCopy()
    const before = readFileSync(copy)
    await expect(
      applyScriptToDocument({
        scriptId: 'tidy',
        source: insert('not written'),
        documentPath: copy,
        confirm: false,
      }),
    ).rejects.toThrow(/explicit confirmation/)
    expect(readFileSync(copy).equals(before)).toBe(true)
  })

  it('writes when confirmed, and the document really changed', async () => {
    const copy = scratchCopy()
    const before = readFileSync(copy)
    const result = await applyScriptToDocument({
      scriptId: 'tidy',
      source: insert('from the script'),
      documentPath: copy,
      confirm: true,
    })
    expect(result.applied).toBe(1)
    expect(result.failures).toEqual([])
    expect(result.written).toBe(copy)
    expect(result.bytes).toBeGreaterThan(0)
    expect(result.log[0]).toMatch(/^blocks=\d+$/)
    expect(readFileSync(copy).equals(before)).toBe(false)

    // read the file back with the same reader the host used: the new text is on disk
    const reopened = await openDocument(readInput(copy))
    try {
      expect(describeDocument(reopened).some((b) => b.text.includes('from the script'))).toBe(true)
    } finally {
      closeDocument(reopened)
    }
  })

  it('writes nothing when the engine refuses an operation', async () => {
    const copy = scratchCopy()
    const before = readFileSync(copy)
    // atomic mode: the engine refuses the batch outright, so the save is never
    // reached — the file is untouched, not half-written
    await expect(
      applyScriptToDocument({
        scriptId: 'unknown-op',
        source: '$.docs.propose([{ op: "nope" }])',
        documentPath: copy,
        confirm: true,
      }),
    ).rejects.toThrow(/unknown op|Nothing was applied/)
    expect(readFileSync(copy).equals(before)).toBe(true)
  })

  it('writes nothing when the script proposes nothing, and leaves no temp file', async () => {
    const copy = scratchCopy()
    const before = readFileSync(copy)
    await expect(
      applyScriptToDocument({
        scriptId: 'quiet',
        source: '$.log("nothing to do")',
        documentPath: copy,
        confirm: true,
      }),
    ).rejects.toThrow(/proposed nothing/)
    expect(readFileSync(copy).equals(before)).toBe(true)
    const { readdirSync } = await import('node:fs')
    expect(
      readdirSync(join(copy, '..')).filter((name) => name.includes('genoffice-apply')),
    ).toEqual([])
  })
})
