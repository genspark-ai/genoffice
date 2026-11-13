/** Seeded round-trip fidelity fuzz for the XLSX save pipeline — the sheets
 *  sibling of the docs suite's roundtrip-fuzz.test.ts. The hostile-input
 *  suite (#2542) feeds corrupted packages to the reader and asserts bounded
 *  failure; this suite feeds valid packages to the *saver* and asserts the
 *  three properties that matter for a part-level archive editor (the
 *  save_archive contract: untouched entries ride through as raw copies):
 *
 *  1. untouched: a zero-edit save keeps every package part byte-identical
 *     (raw copies), the output is itself a fixed point (a second save is
 *     file-level identical — the zip central directory is deterministic), and
 *     the saved package still opens.
 *  2. edit: appending one row to the first sheet keeps every other part
 *     byte-identical and the replaced part verbatim, and the edited package
 *     still opens. This mirrors the app's save shape, which ships changed
 *     worksheet XML as `replacements` and leaves the rest to the archive.
 *
 *  The generator (helpers/fuzz-xlsx.ts) covers the save pipeline's part
 *  families; CI runs the default 24 seeds, `scripts/xlsx-roundtrip-sweep.mjs
 *  N` sweeps N seeds locally. A failing seed reproduces exactly via
 *  buildFuzzXlsx(seed). */
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { buildFuzzXlsx } from './helpers/fuzz-xlsx'

const SEEDS = Math.max(1, Number(process.env.FUZZ_SEEDS ?? 24))
const TEST_TIMEOUT_MS = 30_000

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}

async function partBytes(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes)
  const out = new Map<string, Uint8Array>()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (!entry.dir) out.set(name, await entry.async('uint8array'))
  }
  return out
}

async function openAndClose(path: string): Promise<void> {
  const opened = (await client.open(path)) as { sessionId?: string }
  expect(opened, `the saved package must still open: ${path}`).toBeTruthy()
  if (typeof opened?.sessionId === 'string') await client.close(opened.sessionId)
}

/** The edit transform: one deterministic row appended to the first sheet.
 *  Any seed's sheet1 has at least one row, so the appended index is fresh. */
function appendRow(sheetXml: Uint8Array, seed: number): Uint8Array {
  const text = new TextDecoder().decode(sheetXml)
  let maxRow = 0
  for (const match of text.matchAll(/<row r="(\d+)"/g)) {
    maxRow = Math.max(maxRow, Number(match[1]))
  }
  const newRow = maxRow + 1
  const appended = `<row r="${newRow}"><c r="A${newRow}"><v>${seed}</v></c></row>`
  const edited = text.replace('</sheetData>', `${appended}</sheetData>`)
  expect(edited).not.toBe(text)
  return new TextEncoder().encode(edited)
}

let directory: string
let client: XlsxSidecarClient

beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'xlsx-roundtrip-fuzz-'))
  client = new XlsxSidecarClient(sidecarBinaryPath())
})

afterAll(async () => {
  client?.stop()
  if (directory) await rm(directory, { recursive: true, force: true })
})

describe('seeded xlsx round-trip fidelity fuzz', () => {
  // no dead knobs: over the CI seed range every part family the generator
  // knows how to build must actually appear at least once — otherwise a
  // family can rot silently and the per-seed assertions stop covering it
  it('the knob families fire across the seed range', async () => {
    const seen = new Map<string, number>(
      Object.entries({
        styles: 0,
        sharedStrings: 0,
        formulas: 0,
        merges: 0,
        frozenPanes: 0,
        autoFilter: 0,
        conditionalFormatting: 0,
        dataValidation: 0,
        hyperlinks: 0,
        comments: 0,
        vmlDrawing: 0,
        tables: 0,
        definedNames: 0,
      }).map(([k, v]) => [k, v]),
    )
    for (let seed = 1; seed <= Math.max(SEEDS, 24); seed++) {
      const zip = await JSZip.loadAsync(await buildFuzzXlsx(seed))
      if (zip.file('xl/styles.xml')) seen.set('styles', seen.get('styles')! + 1)
      if (zip.file('xl/sharedStrings.xml'))
        seen.set('sharedStrings', seen.get('sharedStrings')! + 1)
      if (zip.file('xl/workbook.xml')) {
        const workbook = await zip.file('xl/workbook.xml')!.async('string')
        if (workbook.includes('<definedName '))
          seen.set('definedNames', seen.get('definedNames')! + 1)
      }
      for (const name of Object.keys(zip.files)) {
        if (!name.startsWith('xl/worksheets/sheet')) continue
        if (name.includes('_rels')) continue
        const sheet = await zip.file(name)!.async('string')
        if (sheet.includes('<f>')) seen.set('formulas', seen.get('formulas')! + 1)
        if (sheet.includes('<mergeCell ')) seen.set('merges', seen.get('merges')! + 1)
        if (sheet.includes('state="frozen"')) seen.set('frozenPanes', seen.get('frozenPanes')! + 1)
        if (sheet.includes('<autoFilter ')) seen.set('autoFilter', seen.get('autoFilter')! + 1)
        if (sheet.includes('<conditionalFormatting '))
          seen.set('conditionalFormatting', seen.get('conditionalFormatting')! + 1)
        if (sheet.includes('<dataValidation '))
          seen.set('dataValidation', seen.get('dataValidation')! + 1)
        if (sheet.includes('<hyperlink ')) seen.set('hyperlinks', seen.get('hyperlinks')! + 1)
        if (sheet.includes('<legacyDrawing ')) seen.set('vmlDrawing', seen.get('vmlDrawing')! + 1)
        if (sheet.includes('<tablePart ')) seen.set('tables', seen.get('tables')! + 1)
      }
      for (const name of Object.keys(zip.files)) {
        if (name.startsWith('xl/comments')) seen.set('comments', seen.get('comments')! + 1)
      }
    }
    for (const [family, count] of seen) {
      expect(count, `${family} must fire across seeds 1..${Math.max(SEEDS, 24)}`).toBeGreaterThan(0)
    }
  }, 60_000)

  for (let seed = 1; seed <= SEEDS; seed++) {
    it(
      `seed ${seed}: untouched save keeps every part byte-identical and is a fixed point`,
      async () => {
        const bytes = await buildFuzzXlsx(seed)
        const source = join(directory, `in-${seed}.xlsx`)
        const first = join(directory, `out1-${seed}.xlsx`)
        const second = join(directory, `out2-${seed}.xlsx`)
        await writeFile(source, bytes)

        await client.saveArchive({
          sourcePath: source,
          targetPath: first,
          replacements: [],
          removals: [],
          additions: [],
        })

        // raw copies: every part rides through unchanged
        const before = await partBytes(await readFile(source))
        const after = await partBytes(await readFile(first))
        expect([...after.keys()].sort()).toEqual([...before.keys()].sort())
        for (const [name, content] of before) {
          expect(Buffer.from(after.get(name)!).equals(Buffer.from(content)), name).toBe(true)
        }

        // the zip container itself is deterministic: saving the save is
        // file-level identical
        await client.saveArchive({
          sourcePath: first,
          targetPath: second,
          replacements: [],
          removals: [],
          additions: [],
        })
        expect(Buffer.from(await readFile(second)).equals(Buffer.from(await readFile(first)))).toBe(
          true,
        )

        await openAndClose(first)
      },
      TEST_TIMEOUT_MS,
    )

    it(
      `seed ${seed}: appending one row keeps every other part byte-identical`,
      async () => {
        const bytes = await buildFuzzXlsx(seed)
        const source = join(directory, `edit-in-${seed}.xlsx`)
        const target = join(directory, `edit-out-${seed}.xlsx`)
        await writeFile(source, bytes)

        const before = await partBytes(bytes)
        const sheetName = 'xl/worksheets/sheet1.xml'
        const edited = appendRow(before.get(sheetName)!, seed)
        const patched = join(directory, `sheet1-patched-${seed}.xml`)
        await writeFile(patched, edited)

        await client.saveArchive({
          sourcePath: source,
          targetPath: target,
          replacements: [{ name: sheetName, contentPath: patched }],
          removals: [],
          additions: [],
        })

        const after = await partBytes(await readFile(target))
        expect([...after.keys()].sort()).toEqual([...before.keys()].sort())
        for (const [name, content] of before) {
          if (name === sheetName) continue
          expect(Buffer.from(after.get(name)!).equals(Buffer.from(content)), name).toBe(true)
        }
        // the replaced part is exactly the bytes the caller shipped
        expect(Buffer.from(after.get(sheetName)!).equals(Buffer.from(edited))).toBe(true)

        await openAndClose(target)
      },
      TEST_TIMEOUT_MS,
    )
  }
})
