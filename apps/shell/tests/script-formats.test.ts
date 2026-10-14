import { copyFileSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { readInput } from '@genoffice/cli/headless-docs'
import {
  blankDeck,
  blankWorkbook,
  cellEditsFromTable,
  openDeck,
  describeDeck,
  saveDeck,
  writeWorkbook,
} from '@genoffice/cli/headless-formats'
import { readBasicWorkbook } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { describe, expect, it } from 'vitest'
import {
  applyScriptToDocument,
  runScriptInSandbox,
  runScriptReadOnly,
  scriptFormatFor,
} from '../src/main/script-host'
import {
  ScriptRunError,
  type DeckSnapshot,
  type SheetsSnapshot,
  type TextSnapshot,
} from '../src/main/script-shared'
import { textDiffLines } from '../src/main/script-text'

/**
 * The script kind's per-format half, driven directly: what each namespace lets a
 * script reach, and a real run and apply against real files of every format —
 * each one proving the dry run's plan is what the write does, that a refusal or
 * a failed script leaves the bytes alone, and that the write is verified by
 * reading the file back with the same reader the host used.
 */

const REPO = resolve(__dirname, '../../..')

// ── format dispatch ────────────────────────────────────────────────────────────

describe('scriptFormatFor', () => {
  it('maps extensions to the five formats and nothing else', () => {
    expect(scriptFormatFor('/tmp/a.docx')).toBe('docs')
    expect(scriptFormatFor('/tmp/a.DOCX')).toBe('docs')
    expect(scriptFormatFor('/tmp/a.xlsx')).toBe('sheets')
    expect(scriptFormatFor('/tmp/a.md')).toBe('text')
    expect(scriptFormatFor('/tmp/a.markdown')).toBe('text')
    expect(scriptFormatFor('/tmp/a.html')).toBe('text')
    expect(scriptFormatFor('/tmp/a.pptx')).toBe('deck')
    expect(scriptFormatFor('/tmp/a.pdf')).toBe('pdf')
    expect(scriptFormatFor('/tmp/a.txt')).toBeUndefined()
    expect(scriptFormatFor('/tmp/a.csv')).toBeUndefined()
    expect(scriptFormatFor('/tmp/a')).toBeUndefined()
  })
})

// ── the sheets API inside the sandbox ──────────────────────────────────────────

const sheetSnapshot: SheetsSnapshot = {
  format: 'sheets',
  active: 'Sheet1',
  sheets: [
    {
      name: 'Sheet1',
      first: 'A1',
      last: 'B3',
      rows: 3,
      cols: 2,
      cells: { A1: 'name', B1: 'qty', A2: 'pen', B2: 3, A3: 'ink', B3: 4 },
      formulas: { B3: '=SUM(B2:B2)' },
      formats: { A1: { bold: true } },
    },
    {
      name: 'Extra',
      first: 'A1',
      last: 'A1',
      rows: 1,
      cols: 1,
      cells: { A1: 'x' },
      formulas: {},
      formats: {},
    },
  ],
}

function sheetsSandbox(source: string, input: Record<string, unknown> = {}) {
  return runScriptInSandbox({ scriptPath: 's.js', source, snapshot: sheetSnapshot, input })
}

describe('the $.sheets namespace', () => {
  it('lists sheets, the used range and the active one', () => {
    const outcome = sheetsSandbox(
      [
        'var list = $.sheets.sheets()',
        '$.log(list.map(function (s) { return s.name + (s.active ? "*" : "") }).join(","))',
        '$.log(JSON.stringify($.sheets.usedRange("Sheet1")))',
        '$.log(JSON.stringify($.sheets.usedRange()))',
      ].join('\n'),
    )
    expect(outcome.log[0]).toBe('Sheet1*,Extra')
    expect(JSON.parse(outcome.log[1]!)).toEqual({
      sheet: 'Sheet1',
      first: 'A1',
      last: 'B3',
      rows: 3,
      cols: 2,
    })
    expect(JSON.parse(outcome.log[2]!).sheet).toBe('Sheet1')
  })

  it('reads a range as a 2-D grid of values, formulas kept apart', () => {
    const outcome = sheetsSandbox(
      [
        'var r = $.sheets.range("Sheet1!A1:B3")',
        '$.log(JSON.stringify(r.values))',
        '$.log(JSON.stringify($.sheets.formulas()))',
      ].join('\n'),
    )
    expect(JSON.parse(outcome.log[0]!)).toEqual([
      ['name', 'qty'],
      ['pen', 3],
      ['ink', 4],
    ])
    expect(JSON.parse(outcome.log[1]!)).toEqual({ B3: '=SUM(B2:B2)' })
  })

  it('refuses a read over the cap with a reason, and an unknown sheet', () => {
    const outcome = sheetsSandbox(
      [
        'try { $.sheets.range("Sheet1!A1:ZZ2000") } catch (e) { $.log("cap: " + e.message.slice(0, 40)) }',
        'try { $.sheets.range("Nope!A1:B2") } catch (e) { $.log("sheet: " + e.message.slice(0, 30)) }',
      ].join('\n'),
    )
    expect(outcome.log[0]).toMatch(/^cap: range Sheet1!A1:ZZ2000 covers \d+ ce/)
    expect(outcome.log[1]).toMatch(/^sheet: no sheet named Nope/)
  })

  it('finds matches and echoes formats', () => {
    const outcome = sheetsSandbox(
      [
        'var hits = $.sheets.find("PEN")',
        '$.log(hits.length + " at " + hits[0].ref + " truncated=" + hits.truncated)',
        '$.log(JSON.stringify($.sheets.formats("Sheet1!A1:A1")))',
      ].join('\n'),
    )
    expect(outcome.log[0]).toBe('1 at A2 truncated=false')
    expect(JSON.parse(outcome.log[1]!)).toEqual({ A1: { bold: true } })
  })

  it('proposes ops for the engine to judge', () => {
    const outcome = sheetsSandbox(
      '$.sheets.propose([{ op: "set_cell", sheet: "Sheet1", address: "C1", value: 1 }])',
    )
    expect(outcome.ops).toEqual([{ op: 'set_cell', sheet: 'Sheet1', address: 'C1', value: 1 }])
  })
})

// ── a real workbook: dry run, apply, atomic refusal ────────────────────────────

async function makeWorkbook(path: string): Promise<void> {
  const blank = await blankWorkbook('Sheet1')
  await writeWorkbook(
    blank,
    cellEditsFromTable('Sheet1', [
      ['item', 'qty'],
      ['pen', 3],
      ['ink', 4],
    ]),
    path,
  )
}

function workbookScript(source: string): string {
  return source
}

describe('sheets runs and applies', () => {
  it('dry-runs a plan against the real file and leaves it byte-identical', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-sheets-'))
    const path = join(dir, 'book.xlsx')
    await makeWorkbook(path)
    const before = readFileSync(path)
    const outcome = await runScriptReadOnly({
      scriptId: 'add-total',
      source: workbookScript(`
        var range = $.sheets.range('Sheet1!A2:B3')
        var total = range.values.reduce(function (sum, row) { return sum + Number(row[1]) }, 0)
        $.log('total=' + total)
        $.sheets.propose([{ op: 'set_cell', sheet: 'Sheet1', address: 'A4', value: 'total' }, { op: 'set_cell', sheet: 'Sheet1', address: 'B4', value: total }])
      `),
      documentPath: path,
    })
    expect(outcome.log[0]).toBe('total=7')
    expect(outcome.failures).toEqual([])
    expect(outcome.plan[0]).toMatch(/^2 cell write\(s\)$/)
    expect(readFileSync(path).equals(before)).toBe(true)
  })

  it('reports a rejected op while the rest still validate', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-sheets-'))
    const path = join(dir, 'book.xlsx')
    await makeWorkbook(path)
    const outcome = await runScriptReadOnly({
      scriptId: 'one-bad-op',
      source: workbookScript(`
        $.sheets.propose([
          { op: 'set_cell', sheet: 'Sheet1', address: 'A4', value: 'ok' },
          { op: 'no_such_op', sheet: 'Sheet1', address: 'A5', value: 'nope' },
        ])
      `),
      documentPath: path,
    })
    expect(outcome.applied).toBe(1)
    expect(outcome.failures).toHaveLength(1)
    expect(outcome.failures[0]).toContain('no_such_op')
  })

  it('writes when confirmed, verified by reading the file back', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-sheets-'))
    const path = join(dir, 'book.xlsx')
    await makeWorkbook(path)
    const before = readFileSync(path)
    const result = await applyScriptToDocument({
      scriptId: 'add-total',
      source: workbookScript(`
        var range = $.sheets.range('Sheet1!A2:B3')
        var total = range.values.reduce(function (sum, row) { return sum + Number(row[1]) }, 0)
        $.sheets.propose([
          { op: 'set_cell', sheet: 'Sheet1', address: 'A4', value: 'total' },
          { op: 'set_cell', sheet: 'Sheet1', address: 'B4', value: total },
        ])
      `),
      documentPath: path,
      confirm: true,
    })
    expect(result.applied).toBe(2)
    expect(result.written).toBe(path)
    expect(result.bytes).toBeGreaterThan(0)
    expect(readFileSync(path).equals(before)).toBe(false)
    const reread = await readBasicWorkbook(readFileSync(path))
    expect(reread.snapshot.sheets[0]!.cells['A4']!.value).toBe('total')
    expect(reread.snapshot.sheets[0]!.cells['B4']!.value).toBe(7)
  })

  it('writes nothing when the engine refuses the batch (atomic)', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-sheets-'))
    const path = join(dir, 'book.xlsx')
    await makeWorkbook(path)
    const before = readFileSync(path)
    await expect(
      applyScriptToDocument({
        scriptId: 'bad-op',
        source: workbookScript(
          "$.sheets.propose([{ op: 'no_such_op', sheet: 'Sheet1', address: 'A4', value: 'x' }])",
        ),
        documentPath: path,
        confirm: true,
      }),
    ).rejects.toThrow(/refused the batch/)
    expect(readFileSync(path).equals(before)).toBe(true)
    expect(readdirSync(dir).filter((name) => name.includes('genoffice-apply'))).toEqual([])
  })

  it('refuses a script whose declared format does not match the file', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-sheets-'))
    const path = join(dir, 'book.xlsx')
    await makeWorkbook(path)
    await expect(
      runScriptReadOnly({
        scriptId: 'wrong-format',
        source: '$.log(1)',
        documentPath: path,
        format: 'docs',
      }),
    ).rejects.toThrow(/targets docs documents.*is a sheets file/)
  })
})

// ── the text API: markdown and html ────────────────────────────────────────────

describe('the $.text namespace and text runs', () => {
  const snapshot: TextSnapshot = { format: 'text', text: '# Title\n\nOne.\n\nTwo.\n' }

  it('hands the script the text, and the return value becomes the edit', () => {
    const outcome = runScriptInSandbox({
      scriptPath: 't.js',
      source: [
        '$.log("lines=" + $.text.lines().length)',
        'return $.text.get().replace("One.", "1.")',
      ].join('\n'),
      snapshot,
      input: {},
    })
    expect(outcome.log[0]).toBe('lines=6')
    expect(outcome.text).toBe('# Title\n\n1.\n\nTwo.\n')
    expect(outcome.ops).toEqual([])
  })

  it('accepts { text } as the return shape and nothing else', () => {
    const object = runScriptInSandbox({
      scriptPath: 't.js',
      source: 'return { text: "abc" }',
      snapshot,
      input: {},
    })
    expect(object.text).toBe('abc')
    const number = runScriptInSandbox({
      scriptPath: 't.js',
      source: 'return 42',
      snapshot,
      input: {},
    })
    expect(number.text).toBeUndefined()
  })

  it('dry-runs a markdown edit as a diff summary and leaves the file alone', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-text-'))
    const path = join(dir, 'note.md')
    writeFileSync(path, '# Title\n\nOne.\n\nTwo.\n', 'utf8')
    const before = readFileSync(path, 'utf8')
    const outcome = await runScriptReadOnly({
      scriptId: 'renumber',
      source: 'return $.text.get().replace("One.", "1.")',
      documentPath: path,
    })
    expect(outcome.applied).toBe(1)
    expect(outcome.plan[0]).toMatch(/^1 line\(s\) replaced by 1 at line 3$/)
    expect(outcome.plan).toContain('- One.')
    expect(outcome.plan).toContain('+ 1.')
    expect(readFileSync(path, 'utf8')).toBe(before)
  })

  it('applies the returned text to an html file, verified by reading it back', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-text-'))
    const path = join(dir, 'page.html')
    writeFileSync(path, '<p>alpha</p>\n<p>beta</p>\n', 'utf8')
    const result = await applyScriptToDocument({
      scriptId: 'sweep',
      source: 'return $.text.get().replace("beta", "beta2")',
      documentPath: path,
      confirm: true,
    })
    expect(result.applied).toBe(1)
    expect(readFileSync(path, 'utf8')).toBe('<p>alpha</p>\n<p>beta2</p>\n')
  })

  it('writes nothing when the script returns the text unchanged', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-text-'))
    const path = join(dir, 'note.md')
    writeFileSync(path, '# Title\n', 'utf8')
    const before = readFileSync(path, 'utf8')
    await expect(
      applyScriptToDocument({
        scriptId: 'same',
        source: 'return $.text.get()',
        documentPath: path,
        confirm: true,
      }),
    ).rejects.toThrow(/returned the text unchanged/)
    expect(readFileSync(path, 'utf8')).toBe(before)
  })

  it('writes nothing when the script throws or returns no text', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-text-'))
    const path = join(dir, 'note.md')
    writeFileSync(path, '# Title\n', 'utf8')
    const before = readFileSync(path, 'utf8')
    await expect(
      applyScriptToDocument({
        scriptId: 'throw',
        source: 'throw new Error("no")',
        documentPath: path,
        confirm: true,
      }),
    ).rejects.toThrow(/no/)
    await expect(
      applyScriptToDocument({
        scriptId: 'quiet',
        source: '$.log("nothing")',
        documentPath: path,
        confirm: true,
      }),
    ).rejects.toThrow(/returned no text/)
    expect(readFileSync(path, 'utf8')).toBe(before)
  })
})

describe('textDiffLines', () => {
  it('describes insertions, deletions and replacements', () => {
    expect(textDiffLines('a\nb\nc', 'a\nb\nc\nd')).toEqual([
      '1 line(s) inserted after line 3',
      '+ d',
    ])
    expect(textDiffLines('a\nb\nc', 'a')).toEqual(['2 line(s) deleted at line 2', '- b', '- c'])
    expect(textDiffLines('a\nb\nc', 'a\nB\nc')).toEqual([
      '1 line(s) replaced by 1 at line 2',
      '- b',
      '+ B',
    ])
  })
})

// ── the deck API: the slides app's own vocabulary ──────────────────────────────

async function makeDeck(path: string): Promise<void> {
  writeFileSync(path, Buffer.from(await saveDeck(await blankDeck())))
}

const deckSnapshot: DeckSnapshot = {
  format: 'deck',
  deck: {
    slides: 1,
    size: { cx: 9144000, cy: 6858000, inches: { width: 10, height: 7.5 } },
    emu_per_inch: 914400,
    pages: [
      {
        index: 0,
        id: 's_0',
        elements: [
          {
            id: 'e_1',
            type: 'textbox',
            box: { x: 914400, y: 685800, cx: 7315200, cy: 914400 },
            text: 'Old title',
          },
        ],
      },
    ],
  },
}

describe('the $.deck namespace and deck runs', () => {
  it('hands the script every element with its geometry and text', () => {
    const outcome = runScriptInSandbox({
      scriptPath: 'd.js',
      source: [
        'var els = $.deck.elements(0)',
        '$.log(els.length + " el at " + els[0].box.x + " text=" + els[0].text)',
        '$.deck.propose([{ op: "setText", target: { slide: 0, el: "e_1" }, text: "New title" }])',
      ].join('\n'),
      snapshot: deckSnapshot,
      input: {},
    })
    expect(outcome.log[0]).toBe('1 el at 914400 text=Old title')
    expect(outcome.ops).toHaveLength(1)
  })

  it('dry-runs a setText against a real deck and leaves it byte-identical', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-deck-'))
    const path = join(dir, 'deck.pptx')
    await makeDeck(path)
    const before = readFileSync(path)
    const outcome = await runScriptReadOnly({
      scriptId: 'retitle',
      source: [
        'var els = $.deck.elements(0)',
        '$.log("slide0 has " + els.length + " element(s)")',
        '$.deck.propose([{ op: "addElement", target: { slide: 0 }, kind: "textbox", offset: { x: 914400, y: 685800, cx: 7315200, cy: 914400 }, paragraphs: [{ runs: [{ text: "From the script" }] }] }])',
      ].join('\n'),
      documentPath: path,
    })
    expect(outcome.log[0]).toMatch(/^slide0 has \d+ element/)
    expect(outcome.failures).toEqual([])
    expect(outcome.plan.length).toBe(1)
    expect(readFileSync(path).equals(before)).toBe(true)
  })

  it('applies, audits the layout, and the deck reads back with the new text', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-deck-'))
    const path = join(dir, 'deck.pptx')
    await makeDeck(path)
    const result = await applyScriptToDocument({
      scriptId: 'retitle',
      source: [
        'var page = $.deck.slides()[0]',
        '$.deck.propose([{ op: "addElement", target: { slide: page.index }, kind: "textbox", offset: { x: 914400, y: 685800, cx: 7315200, cy: 914400 }, paragraphs: [{ runs: [{ text: "From the script" }] }] }])',
      ].join('\n'),
      documentPath: path,
      confirm: true,
    })
    expect(result.applied).toBe(1)
    expect(result.written).toBe(path)
    expect(result.plan.some((line) => line.startsWith('layout audit:'))).toBe(true)
    const reopened = await openDeck(readInput(path))
    const summary = describeDeck(reopened)
    const texts = JSON.stringify(summary)
    expect(texts).toContain('From the script')
  })

  it('writes nothing when the engine refuses an op (atomic rollback)', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-deck-'))
    const path = join(dir, 'deck.pptx')
    await makeDeck(path)
    const before = readFileSync(path)
    await expect(
      applyScriptToDocument({
        scriptId: 'bad-target',
        source:
          '$.deck.propose([{ op: "setText", target: { slide: 9, el: "e_missing" }, text: "x" }])',
        documentPath: path,
        confirm: true,
      }),
    ).rejects.toThrow(/refused .* nothing was written/)
    expect(readFileSync(path).equals(before)).toBe(true)
  })
})

// ── the pdf API: read-only, report as the result ───────────────────────────────

/** A tiny valid PDF with one page of Helvetica text per entry, xref offsets computed. */
function minimalPdf(texts: string[]): Buffer {
  let pdf = '%PDF-1.4\n'
  const offsets: number[] = []
  const add = (body: string): void => {
    offsets.push(pdf.length)
    pdf += body
  }
  const kids = texts.map((_, i) => 4 + i * 2 + ' 0 R').join(' ')
  add('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n')
  add('2 0 obj\n<< /Type /Pages /Kids [' + kids + '] /Count ' + texts.length + ' >>\nendobj\n')
  add('3 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n')
  texts.forEach((text, i) => {
    const content = 'BT /F1 24 Tf 72 700 Td (' + text.replace(/([()\\])/g, '\\$1') + ') Tj ET'
    add(
      4 +
        i * 2 +
        ' 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ' +
        (5 + i * 2) +
        ' 0 R >>\nendobj\n',
    )
    add(
      5 +
        i * 2 +
        ' 0 obj\n<< /Length ' +
        content.length +
        ' >>\nstream\n' +
        content +
        '\nendstream\nendobj\n',
    )
  })
  const size = 3 + texts.length * 2
  const startxref = pdf.length
  pdf += 'xref\n0 ' + (size + 1) + '\n0000000000 65535 f \n'
  for (const offset of offsets) pdf += String(offset).padStart(10, '0') + ' 00000 n \n'
  pdf +=
    'trailer\n<< /Size ' + (size + 1) + ' /Root 1 0 R >>\nstartxref\n' + startxref + '\n%%EOF\n'
  return Buffer.from(pdf, 'latin1')
}

describe('the $.pdf namespace and pdf runs', () => {
  it('serves info and bounded page texts, and the report comes back as the result', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-pdf-'))
    const path = join(dir, 'scan.pdf')
    writeFileSync(path, minimalPdf(['Invoice 2026-014 total 480', 'Page two mentions 480 again']))
    const before = readFileSync(path)
    const outcome = await runScriptReadOnly({
      scriptId: 'scan',
      source: [
        '$.log("pages=" + $.pdf.info().pages)',
        'var pages = $.pdf.pages(1, 2)',
        'var hits = pages.filter(function (p) { return p.text.indexOf("480") !== -1 })',
        '$.log("hit pages: " + hits.map(function (p) { return p.page }).join(","))',
        'return { matches: hits.length, pages: hits.map(function (p) { return p.page }) }',
      ].join('\n'),
      documentPath: path,
    })
    expect(outcome.log[0]).toBe('pages=2')
    expect(outcome.log[1]).toBe('hit pages: 1,2')
    expect(outcome.applied).toBe(0)
    expect(JSON.parse(outcome.result!)).toEqual({ matches: 2, pages: [1, 2] })
    expect(readFileSync(path).equals(before)).toBe(true)
  })

  it('clips page text to the maxChars argument', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-pdf-'))
    const path = join(dir, 'scan.pdf')
    writeFileSync(path, minimalPdf(['abcdefghij'.repeat(50)]))
    const outcome = await runScriptReadOnly({
      scriptId: 'clip',
      source: [
        'var page = $.pdf.pages(1, 1, 20)[0]',
        '$.log(page.text.length + " " + page.text.endsWith("...(+480 chars)"))',
      ].join('\n'),
      documentPath: path,
    })
    expect(outcome.log[0]).toBe('35 true')
  })

  it('refuses an apply: pdf scripts are read-only', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-pdf-'))
    const path = join(dir, 'scan.pdf')
    writeFileSync(path, minimalPdf(['hello']))
    await expect(
      applyScriptToDocument({
        scriptId: 'scan',
        source: 'return { ok: true }',
        documentPath: path,
        confirm: true,
      }),
    ).rejects.toThrow(/pdf scripts are read-only/)
  })

  it('fails with a readable message on a file that is not a pdf', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-pdf-'))
    const path = join(dir, 'scan.pdf')
    writeFileSync(path, 'not a pdf at all')
    await expect(
      runScriptReadOnly({ scriptId: 'scan', source: '$.log(1)', documentPath: path }),
    ).rejects.toThrow(ScriptRunError)
  })
})

// ── the shared gates hold for every format ─────────────────────────────────────

describe('the shared gates', () => {
  it('needs a deliberate confirmation for a text write, like every format', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-gates-'))
    const path = join(dir, 'note.md')
    writeFileSync(path, '# Title\n', 'utf8')
    const before = readFileSync(path, 'utf8')
    await expect(
      applyScriptToDocument({
        scriptId: 'edit',
        source: 'return "new text"',
        documentPath: path,
        confirm: false,
      }),
    ).rejects.toThrow(/explicit confirmation/)
    expect(readFileSync(path, 'utf8')).toBe(before)
  })

  it('names the unsupported types instead of failing deep in a sandbox', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-gates-'))
    const path = join(dir, 'notes.txt')
    writeFileSync(path, 'hello', 'utf8')
    await expect(
      runScriptReadOnly({ scriptId: 'edit', source: '$.log(1)', documentPath: path }),
    ).rejects.toThrow(/scripts do not support this file type \(\.txt\)/)
  })

  it('keeps the fixture deck of the repo out of the format map (sanity)', () => {
    expect(scriptFormatFor(join(REPO, 'fixtures/generated/simple.docx'))).toBe('docs')
  })
})
