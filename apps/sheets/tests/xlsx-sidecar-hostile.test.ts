/** Hostile-input fuzz for the XLSX import surface — the sheets sibling of the
 *  docs suite's roundtrip-hostile.test.ts. The regular sidecar tests feed it
 *  valid workbooks; this file feeds corrupted ones and asserts the one
 *  property that matters for an importer of untrusted bytes: **bounded, clean
 *  failure** — an open either answers or refuses, the sidecar process survives
 *  every case (an exit listener records a crash), and a known-good workbook
 *  still opens afterwards, so a wedged reader fails the test on its own. No
 *  fidelity assertions here: a corrupted package has no "right" output.
 *
 *  Corruption is seeded (same seed → same bytes, failing cases reproduce
 *  exactly) over the generator below. CI runs the default counts;
 *  HOSTILE_SEEDS sweeps larger spaces. */
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import JSZip from 'jszip'
import { afterAll, afterEach, beforeAll, describe, expect, it } from 'vitest'

import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'

const SEEDS = Math.max(1, Number(process.env.HOSTILE_SEEDS ?? 10))
const CASE_TIMEOUT_MS = 20_000
const BOMB_TIMEOUT_MS = 60_000

function sidecarBinaryPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  return fileURLToPath(
    new URL(`../native/xlsx-engine/target/release/${executable}`, import.meta.url),
  )
}

/** mulberry32: small, fast, fully deterministic across runs and platforms. */
function rng(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a |= 0
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const CJK_NAMES = ['季度回顾', '营收明细', '附录 甲', '风险登记']
const LATIN_NAMES = ['Summary', 'data 2024', 'Q3', 'notes']
const TEXTS = ['季度回顾', 'Quarterly review', '风险登记', 'milestone 12', 'draft v2', '附录 甲']

/** A seeded generator of varied-but-valid XLSX packages: one to three sheets
 *  with CJK and latin names, cells mixing shared strings, inline strings,
 *  numbers, booleans, errors and formulas with cached values, a stylesheet
 *  with a date format, merges, one hyperlink and one defined name. Parts the
 *  generator omits stay omitted, so corruption also hits packages that never
 *  had a sharedStrings or a styles part at all. */
async function buildFuzzXlsx(seed: number): Promise<Uint8Array> {
  const r = rng(seed * 7919 + 17)
  const pick = <T>(list: readonly T[]): T => list[Math.floor(r() * list.length)]!
  const zip = new JSZip()

  const sheetCount = 1 + Math.floor(r() * 3)
  const names = Array.from({ length: sheetCount }, (_v, i) =>
    r() < 0.5
      ? CJK_NAMES[(seed + i) % CJK_NAMES.length]!
      : LATIN_NAMES[(seed * 3 + i) % LATIN_NAMES.length]!,
  )
  const sheetIds = names.map((_n, i) => i + 1)

  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheetIds
        .map(
          (id) =>
            `<Override PartName="/xl/worksheets/sheet${id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join('') +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    `${XML_DECL}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<definedNames><definedName name="Scope_${seed}">'${names[0]}'!$A$1:$B$2</definedName></definedNames>` +
      '<sheets>' +
      names
        .map(
          (name, i) =>
            `<sheet name="${name.replace(/&/g, '&amp;')}" sheetId="${sheetIds[i]}" r:id="rId${i + 1}"/>`,
        )
        .join('') +
      '</sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      sheetIds
        .map(
          (_id, i) =>
            `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join('') +
      '</Relationships>',
  )

  const withStyles = r() < 0.85
  if (withStyles) {
    zip.file(
      'xl/styles.xml',
      `${XML_DECL}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy&quot;-&quot;mm&quot;-&quot;dd"/></numFmts>' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><name val="Arial"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills>' +
        '<borders count="1"><border/></borders>' +
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
        '<cellXfs count="2"><xf xfId="0"/><xf xfId="0" fontId="1" fillId="2" numFmtId="164" applyNumberFormat="1"/></cellXfs>' +
        '</styleSheet>',
    )
  }

  const withShared = r() < 0.8
  const shared: string[] = []
  if (withShared) {
    const count = 3 + Math.floor(r() * 5)
    for (let i = 0; i < count; i++) shared.push(pick(TEXTS))
    zip.file(
      'xl/sharedStrings.xml',
      `${XML_DECL}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${shared.length}" uniqueCount="${shared.length}">` +
        shared.map((text) => `<si><t>${text}</t></si>`).join('') +
        '</sst>',
    )
  }

  names.forEach((name, sheetAt) => {
    const rows = 1 + Math.floor(r() * 12)
    const cols = 1 + Math.floor(r() * 6)
    const cells: string[] = []
    for (let row = 1; row <= rows; row++) {
      const rowCells: string[] = []
      for (let col = 0; col < cols; col++) {
        const ref = `${String.fromCharCode(65 + col)}${row}`
        const kind = r()
        if (kind < 0.3 && shared.length > 0) {
          rowCells.push(`<c r="${ref}" t="s"><v>${Math.floor(r() * shared.length)}</v></c>`)
        } else if (kind < 0.45) {
          rowCells.push(`<c r="${ref}" t="inlineStr"><is><t>${pick(TEXTS)}</t></is></c>`)
        } else if (kind < 0.7) {
          rowCells.push(`<c r="${ref}"><v>${Math.floor(r() * 100000) / 100}</v></c>`)
        } else if (kind < 0.8) {
          rowCells.push(`<c r="${ref}" t="b"><v>${r() < 0.5 ? 1 : 0}</v></c>`)
        } else if (kind < 0.9) {
          rowCells.push(`<c r="${ref}" t="e"><v>#N/A</v></c>`)
        } else {
          rowCells.push(`<c r="${ref}"><f>SUM(A1:A${row})</f><v>${row}</v></c>`)
        }
      }
      cells.push(`<row r="${row}">${rowCells.join('')}</row>`)
    }
    const merges =
      rows > 2 && r() < 0.5 ? `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>` : ''
    const hyperlink =
      r() < 0.4
        ? {
            rels: true,
            xml: `<hyperlinks><hyperlink ref="A1" r:id="rIdH1"/></hyperlinks>`,
          }
        : null
    zip.file(
      `xl/worksheets/sheet${sheetAt + 1}.xml`,
      `${XML_DECL}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        `<dimension ref="A1:${String.fromCharCode(64 + Math.max(1, cols))}${rows}"/>` +
        `<sheetData>${cells.join('')}</sheetData>${merges}${hyperlink?.xml ?? ''}</worksheet>`,
    )
    if (hyperlink) {
      zip.file(
        `xl/worksheets/_rels/sheet${sheetAt + 1}.xml.rels`,
        `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>' +
          '</Relationships>',
      )
    }
  })

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

type Corrupt = (bytes: Uint8Array, r: () => number) => Promise<Uint8Array>

/** Flip a handful of random bytes anywhere in the package. */
const bitflip: Corrupt = async (bytes, r) => {
  const out = Uint8Array.from(bytes)
  const flips = 1 + Math.floor(r() * 8)
  for (let i = 0; i < flips; i++) {
    const at = Math.floor(r() * out.length)
    out[at] = (out[at] ?? 0) ^ (1 << Math.floor(r() * 8))
  }
  return out
}

/** Cut the file at a random offset. */
const truncate: Corrupt = async (bytes, r) => {
  const at = 100 + Math.floor(r() * Math.max(1, bytes.length - 100))
  return bytes.slice(0, at)
}

/** Remove a random package entry (a missing sheet or styles part). */
const entryDrop: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir)
  if (names.length === 0) return bytes
  zip.remove(names[Math.floor(r() * names.length)]!)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** Replace a random entry's content with garbage bytes. */
const entryGarbage: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const names = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir)
  if (names.length === 0) return bytes
  const victim = names[Math.floor(r() * names.length)]!
  const junk = new Uint8Array(64 + Math.floor(r() * 512))
  for (let i = 0; i < junk.length; i++) junk[i] = Math.floor(r() * 256)
  zip.file(victim, junk)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** Trailing junk after the end of central directory. */
const trailingJunk: Corrupt = async (bytes, r) => {
  const junk = new Uint8Array(16 + Math.floor(r() * 256))
  for (let i = 0; i < junk.length; i++) junk[i] = Math.floor(r() * 256)
  const out = new Uint8Array(bytes.length + junk.length)
  out.set(bytes)
  out.set(junk, bytes.length)
  return out
}

/** A shared-string index past the end of the table (a structural lie generic
 *  byte corruption rarely lands). */
const sstIndex: Corrupt = async (bytes, r) => {
  const zip = await JSZip.loadAsync(bytes)
  const sheet = Object.keys(zip.files).find((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name))
  if (!sheet || !zip.file(sheet)) return bytes
  const xml = await zip.file(sheet)!.async('string')
  const patched = xml.replace(/<c r="([A-Z]+\d+)" t="s"><v>(\d+)<\/v><\/c>/, (_m, ref: string) => {
    return `<c r="${ref}" t="s"><v>${100000 + Math.floor(r() * 1000)}</v></c>`
  })
  if (patched === xml) return bytes
  zip.file(sheet, patched)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

const CORRUPTIONS: Array<[string, Corrupt]> = [
  ['bitflip', bitflip],
  ['truncate', truncate],
  ['entrydrop', entryDrop],
  ['entrygarbage', entryGarbage],
  ['trailingjunk', trailingJunk],
  ['sstindex', sstIndex],
]

/** The smallest workbook the sidecar must open after every hostile case. */
async function writeHealthyFixture(directory: string): Promise<string> {
  const path = join(directory, `healthy-${Math.random().toString(36).slice(2)}.xlsx`)
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    `${XML_DECL}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Healthy" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${XML_DECL}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
  return path
}

describe('hostile XLSX import fuzz', () => {
  const client = new XlsxSidecarClient(sidecarBinaryPath())
  const crashes: Error[] = []
  const unsubscribeExit = client.onProcessExit(() => {
    crashes.push(new Error('the sidecar process died while handling hostile input'))
  })
  let directory: string
  let healthy: string | null = null

  afterAll(async () => {
    unsubscribeExit()
    client.stop()
    if (directory) await rm(directory, { recursive: true, force: true })
  })

  afterEach(async () => {
    // the crash verdict comes after the liveness probe so a wedged reader is
    // the reported failure, not a slow timeout
    expect(crashes, 'sidecar process exits recorded').toEqual([])
    if (!healthy) return
    const opened = (await client.open(healthy)) as { sessionId?: string } | undefined
    expect(
      opened,
      'the sidecar must still open a healthy workbook after hostile input',
    ).toBeTruthy()
    if (typeof opened?.sessionId === 'string') await client.close(opened.sessionId)
  })

  /** Run one hostile open. Both outcomes are legitimate: a parsed session or a
      refused open — what must not happen is a hang (the timeout), a process
      death (the exit listener) or a wedged reader (the afterEach probe). */
  async function attemptOpen(bytes: Uint8Array): Promise<'opened' | 'refused'> {
    const caseDir = directory
    const path = join(caseDir, `hostile-${Math.random().toString(36).slice(2)}.xlsx`)
    await writeFile(path, bytes)
    try {
      const opened = (await client.open(path)) as { sessionId?: string }
      if (typeof opened.sessionId === 'string') await client.close(opened.sessionId)
      return 'opened'
    } catch {
      return 'refused'
    }
  }

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'xlsx-hostile-fuzz-'))
    healthy = await writeHealthyFixture(directory)
  })

  for (const [mode, corrupt] of CORRUPTIONS) {
    for (let seed = 1; seed <= SEEDS; seed++) {
      it(
        `seed ${seed} ${mode}: the sidecar refuses cleanly or opens, never hangs or dies`,
        async () => {
          const valid = await buildFuzzXlsx(seed)
          const bytes = await corrupt(valid, rng(seed * 7919 + 31))
          await attemptOpen(bytes)
        },
        CASE_TIMEOUT_MS,
      )
    }
  }

  // Resource classes, as fixed constructions: a reader of untrusted packages
  // must answer each of these within its normal request budget.
  describe('resource bombs', () => {
    it(
      'an entry flood opens or refuses without hanging',
      async () => {
        const zip = new JSZip()
        for (let i = 0; i < 5000; i++) zip.file(`flood/${i}.xml`, '<a/>')
        await attemptOpen(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
      },
      BOMB_TIMEOUT_MS,
    )

    it(
      'deeply nested worksheet XML opens or refuses without hanging',
      async () => {
        const depth = 20_000
        const zip = new JSZip()
        addMinimalSheet(zip, `${'<x:evil>'.repeat(depth)}<sheetData/>${'</x:evil>'.repeat(depth)}`)
        await attemptOpen(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
      },
      BOMB_TIMEOUT_MS,
    )

    it(
      'an attribute bomb opens or refuses without hanging',
      async () => {
        const big = 'A'.repeat(8 * 1024 * 1024)
        const zip = new JSZip()
        addMinimalSheet(zip, `<row r="1"><c r="A1" evil="${big}"><v>1</v></c></row>`)
        await attemptOpen(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
      },
      BOMB_TIMEOUT_MS,
    )

    it(
      'a full-grid dimension claim opens or refuses without hanging',
      async () => {
        const zip = new JSZip()
        addMinimalSheet(
          zip,
          '<row r="1"><c r="A1"><v>1</v></c></row>',
          '<dimension ref="A1:XFD1048576"/>',
        )
        await attemptOpen(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
      },
      BOMB_TIMEOUT_MS,
    )
  })
})

/** A workbook whose only sheet body is `bodyXml` (plus an optional leading
    element such as a lying dimension), valid enough to reach the XML reader. */
function addMinimalSheet(zip: JSZip, bodyXml: string, leading = ''): void {
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
  zip.file(
    '[Content_Types].xml',
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Hostile" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
      ` xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${leading}${bodyXml}</worksheet>`,
  )
}
