import { describe, expect, it } from 'vitest'
import { convertXlsToXlsx, parseXlsWorkbook } from '../src/index'
import {
  buildSstRecord,
  buildWorkbookStream,
  buildXlsContainer,
  type SheetSpec,
} from './synthetic-xls'
import * as CFB from 'cfb'

const tree: SheetSpec[] = [
  {
    name: 'Data',
    dimensions: { lastRow: 2, lastCol: 3 },
    colWidths: [{ from: 0, to: 1, width: 12.5 }],
    cells: [
      { kind: 'label', row: 0, col: 0, stringIndex: 0 },
      { kind: 'label', row: 0, col: 1, stringIndex: 1 },
      { kind: 'number', row: 1, col: 0, value: 42.5 },
      { kind: 'rk', row: 1, col: 1, rk: 0x0211_0002 | 2 }, // integer 4, div100 off? see expectations
      { kind: 'bool', row: 1, col: 2, value: true },
      { kind: 'error', row: 1, col: 3, code: 0x07 },
    ],
    merges: [[0, 0, 0, 1]],
  },
  { name: 'Hidden', hidden: true, cells: [{ kind: 'number', row: 0, col: 0, value: 7 }] },
]

const strings = ['Name', 'Value', '中文标签']

const bytesOf = (sheets: SheetSpec[], sst?: Uint8Array): Uint8Array =>
  buildXlsContainer(buildWorkbookStream(sheets, strings, sst))

interface CellRead {
  kind: string
  value: string
}

async function cellsOf(bytes: Uint8Array): Promise<Map<string, CellRead>> {
  const out = await convertXlsToXlsx(bytes)
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(out.bytes)
  const cells = new Map<string, CellRead>()
  for (let s = 0; s < out.sheetCount; s++) {
    const xml = await zip.file(`xl/worksheets/sheet${s + 1}.xml`)!.async('string')
    for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*)><v>([^<]*)<\/v><\/c>/g)) {
      const kind = /t="(\w+)"/.exec(m[2]!)?.[1] ?? 'n'
      cells.set(`sheet${s + 1}!${m[1]}`, { kind, value: m[3] })
    }
  }
  return cells
}

describe('convertXlsToXlsx: the values foundation', () => {
  it('carries strings, numbers, booleans and errors into the sheet XML', async () => {
    const bytes = bytesOf(tree)
    const cells = await cellsOf(bytes)
    expect(cells.get('sheet1!A1')).toEqual({ kind: 's', value: '0' })
    expect(cells.get('sheet1!B1')).toEqual({ kind: 's', value: '1' })
    expect(cells.get('sheet1!A2')).toEqual({ kind: 'n', value: '42.5' })
    expect(cells.get('sheet1!C2')).toEqual({ kind: 'b', value: '1' })
    expect(cells.get('sheet1!D2')).toEqual({ kind: 'e', value: '#DIV/0!' })
  })

  it('decodes the RK compressed form', async () => {
    // RK 0x10 = int flag (2) | 0 (value 0 << 2): 0
    // RK 0x1000 0012 = int 4 (4 << 2 = 0x10) with the int flag: RK bits = 2 | (4 << 2)
    const sheets: SheetSpec[] = [
      {
        name: 'S',
        cells: [
          { kind: 'rk', row: 0, col: 0, rk: 2 | (4 << 2) }, // integer 4
          { kind: 'rk', row: 0, col: 1, rk: 3 | (150 << 2) }, // 150 / 100 = 1.5
        ],
      },
    ]
    const cells = await cellsOf(buildXlsContainer(buildWorkbookStream(sheets, [])))
    expect(cells.get('sheet1!A1')).toEqual({ kind: 'n', value: '4' })
    expect(Number(cells.get('sheet1!B1')!.value)).toBeCloseTo(1.5, 6)
  })

  it('keeps the sheet order, names and the hidden state', async () => {
    const out = await convertXlsToXlsx(bytesOf(tree))
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(out.bytes)
    const wb = await zip.file('xl/workbook.xml')!.async('string')
    expect(wb).toContain('name="Data"')
    expect(wb).toContain('name="Hidden"')
    expect(wb).toContain('state="hidden"')
    expect(out.sheetCount).toBe(2)
  })

  it('carries merges and column widths', async () => {
    const out = await convertXlsToXlsx(bytesOf(tree))
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(out.bytes)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<mergeCell ref="A1:B1"/>')
    expect(sheet).toContain('width="12.5"')
  })

  it('carries CJK labels and formula cached values', async () => {
    const sheets: SheetSpec[] = [
      {
        name: '中文',
        cells: [
          { kind: 'label', row: 0, col: 0, stringIndex: 0 },
          { kind: 'formulaNumber', row: 1, col: 0, value: 1234.5 },
          { kind: 'formulaString', row: 2, col: 0, text: '结果' },
        ],
      },
    ]
    const cells = await cellsOf(buildXlsContainer(buildWorkbookStream(sheets, strings)))
    expect(cells.get('sheet1!A1')).toEqual({ kind: 's', value: '0' })
    expect(cells.get('sheet1!A2')).toEqual({ kind: 'n', value: '1234.5' })
    const out = await convertXlsToXlsx(buildXlsContainer(buildWorkbookStream(sheets, strings)))
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(out.bytes)
    const shared = await zip.file('xl/sharedStrings.xml')!.async('string')
    expect(shared).toContain('结果')
    // the unreferenced CJK label below lives in the source SST but no cell
    // points at it, so the converted workbook carries only what the sheets show
    expect(shared).not.toContain('中文标签')
    // the CJK sheet name survived the round trip
    const wbx = await zip.file('xl/workbook.xml')!.async('string')
    expect(wbx).toContain('name="中文"')
  })

  it('decodes a shared string table split across CONTINUE records', async () => {
    // the split sits at a whole-string boundary; each CONTINUE restarts the
    // compression flag, and the reader walks the segments
    // the split variant: SST carries 'Name'; CONTINUE carries the rest
    const sst = (() => {
      const record = (id: number, payload: Uint8Array): Uint8Array => {
        const out = new Uint8Array(4 + payload.length)
        const v = new DataView(out.buffer)
        v.setUint16(0, id, true)
        v.setUint16(2, payload.length, true)
        out.set(payload, 4)
        return out
      }
      const u16b = (v: number): Uint8Array => {
        const out = new Uint8Array(2)
        new DataView(out.buffer).setUint16(0, v, true)
        return out
      }
      const chars = (text: string): Uint8Array =>
        new Uint8Array(
          [...text].some((c) => c.charCodeAt(0) > 0xff)
            ? Buffer.from(text, 'utf16le')
            : Buffer.from(text, 'latin1'),
        )
      const one = (text: string): Uint8Array =>
        (() => {
          const body = concat2(
            u16b(text.length),
            u8b([...text].some((c) => c.charCodeAt(0) > 0xff) ? 1 : 0),
            chars(text),
          )
          return body
        })()
      const concat2 = (a: Uint8Array, ...rest: Uint8Array[]): Uint8Array => {
        const total = a.length + rest.reduce((s, p) => s + p.length, 0)
        const out = new Uint8Array(total)
        out.set(a, 0)
        let at = a.length
        for (const p of rest) {
          out.set(p, at)
          at += p.length
        }
        return out
      }
      const u8b = (v: number): Uint8Array => new Uint8Array([v & 0xff])
      const sstPart = record(0x00fc, concat2(u16b(3), u16b(3), one('Name')))
      const contPart = record(0x003c, concat2(one('Value'), one('中文标签')))
      return concat2(sstPart, contPart)
    })()
    const splitBytes = buildXlsContainer(buildWorkbookStream(tree, strings, sst))
    const cells = await cellsOf(splitBytes)
    expect(cells.get('sheet1!A1')).toEqual({ kind: 's', value: '0' })
    expect(cells.get('sheet1!B1')).toEqual({ kind: 's', value: '1' })
    // the CONTINUE's strings decoded: 'Value' is the label in B1's SST slot
    expect(cells.get('sheet1!B1')).toEqual({ kind: 's', value: '1' })
  })
})

describe('convertXlsToXlsx: the failure model', () => {
  it('refuses a file that is not a compound container', async () => {
    await expect(convertXlsToXlsx(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))).rejects.toMatchObject({
      code: 'not_cfb',
    })
  })

  it('refuses an encrypted workbook through its FilePass record', async () => {
    const stream = buildWorkbookStream(tree, strings)
    const withFilePass = new Uint8Array(stream.length + 10)
    // the globals BOF in full (header + its 16-byte payload), then FilePass
    withFilePass.set(stream.subarray(0, 20), 0)
    const fp = new Uint8Array(10)
    const v = new DataView(fp.buffer)
    v.setUint16(0, 0x002f, true)
    v.setUint16(2, 6, true)
    v.setUint16(4, 2, true) // RC4 encryption
    withFilePass.set(fp, 20)
    withFilePass.set(stream.subarray(20), 30)
    // lbPlyPos offsets shifted by 10 would need rebuilding; the FilePass fires
    // before any sheet lookup, so the error surfaces first
    const cfb = CFB.utils.cfb_new({ root: 'Root Entry' })
    CFB.utils.cfb_add(cfb, '/Workbook', Buffer.from(withFilePass))
    const out = CFB.write(cfb, { type: 'buffer' }) as Buffer
    await expect(
      convertXlsToXlsx(new Uint8Array(out.buffer, out.byteOffset, out.byteLength)),
    ).rejects.toMatchObject({ code: 'encrypted' })
  })

  it('refuses a BIFF5 stream (unsupported version)', async () => {
    const stream = buildWorkbookStream(tree, strings)
    // BIFF5's BOF carries version 0x0500 in the same position
    const patched = new Uint8Array(stream)
    patched[4] = 0x00
    patched[5] = 0x05
    const cfb = CFB.utils.cfb_new({ root: 'Root Entry' })
    CFB.utils.cfb_add(cfb, '/Workbook', Buffer.from(patched))
    const out = CFB.write(cfb, { type: 'buffer' }) as Buffer
    await expect(
      convertXlsToXlsx(new Uint8Array(out.buffer, out.byteOffset, out.byteLength)),
    ).rejects.toMatchObject({ code: 'unsupported_version' })
  })
})

describe('parseXlsWorkbook: the model underneath', () => {
  it('exposes the parsed values for tests and later stages', () => {
    const workbook = parseXlsWorkbook(buildWorkbookStream(tree, strings))
    expect(workbook.sheets.map((s) => s.name)).toEqual(['Data', 'Hidden'])
    expect(workbook.sheets[0]!.visible).toBe(true)
    expect(workbook.sheets[1]!.visible).toBe(false)
    const a1 = workbook.sheets[0]!.rows.get(0)!.get(0)!
    expect(a1).toEqual({ row: 0, col: 0, type: 'string', value: 'Name' })
  })
})
