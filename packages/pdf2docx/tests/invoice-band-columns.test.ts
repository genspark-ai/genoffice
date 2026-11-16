/**
 * #2669 regression: an e-invoice whose form is an outer frame plus
 * full-width band rules, with the borderless detail table (item name /
 * spec / unit / quantity / unit price / amount / tax rate / tax amount)
 * living in one band that has no interior rules of its own. The lattice
 * solved the whole form into ONE grid and the
 * band merged into a single giant cell: every line item's columns collapsed
 * into one text stream, the buyer tax ID split in two (…FL0000X losing its
 * leading 9 into another fragment), and the four remark lines interleaved
 * with stray row text.
 *
 * `augmentGridStreamColumns` (table.ts) refines such bands with text-implied
 * column boundaries; this pins the converted result on the issue's own
 * redacted sample (see fixtures/README.md).
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx } from '../src'
import { loadPdfium } from './helpers/wasm'

const SAMPLE = join(import.meta.dirname, 'fixtures', 'invoice-text-merge.pdf')

const docXml = async (docx: Uint8Array): Promise<string> => {
  const { default: JSZip } = await import('jszip')
  const zip = await JSZip.loadAsync(docx)
  return zip.file('word/document.xml')!.async('string')
}

/** the first w:tbl's rows as cell-text arrays (w:t concatenation per w:tc) */
function tableRows(xml: string): string[][] {
  const table = /<w:tbl>[\s\S]*?<\/w:tbl>/.exec(xml)?.[0] ?? ''
  return [...table.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((row) =>
    [...row[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((cell) =>
      [...cell[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]!).join(''),
    ),
  )
}

describe('invoice band columns (#2669)', () => {
  it('keeps the detail band columns, the tax ID and the remark lines apart', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(new Uint8Array(readFileSync(SAMPLE)), { pdfium })
    const xml = await docXml(result.docx)

    // the buyer tax ID survives as one contiguous value, not 9 | 1310000…
    expect(xml).toContain('>91310000MA1FL0000X<')

    const rows = tableRows(xml)
    expect(rows.length).toBeGreaterThanOrEqual(4)
    // the detail band's row holds the real columns, not one glued stream
    const detail = rows.find((r) => r.some((c) => c.includes('项目名称')))!
    expect(detail).toBeDefined()
    const texts = detail.join('\u0001')
    expect(texts).toContain('规格型号')
    expect(texts).toContain('单')
    expect(texts).toContain('税率')
    // the item line's own columns: description | quantity | unit price | tax
    expect(texts).toContain('*生活服务*样例服务')
    expect(texts).toContain('100.000000000000')
    expect(texts).toContain('1%')
    // the description cell no longer swallows the header labels
    const descCell = detail.find((c) => c.includes('*生活服务*样例服务'))!
    expect(descCell).not.toContain('规格型号')
    expect(descCell).not.toContain('数量')
  })

  it('keeps each remark line its own paragraph instead of interleaving rows', async () => {
    const pdfium = await loadPdfium()
    const result = await convertPdfToDocx(new Uint8Array(readFileSync(SAMPLE)), { pdfium })
    const xml = await docXml(result.docx)
    const rows = tableRows(xml)
    const remarks = rows.find((r) => r.some((c) => c.includes('备注')))!
    const remarkCell = remarks.find((c) => c.includes('购买方地址'))!
    // each semicolon-terminated field stays whole and in order
    expect(remarkCell).toContain('购买方地址:示例市示例区示例路1号;')
    expect(remarkCell).toContain('购方开户银行:示例银行示例支行;')
    expect(remarkCell).toContain('销售方地址:示例市示例区示例街2号;')
    // the bank account digits are one run, not interleaved with stray 111s
    expect(remarkCell).toContain('111111111111111111111')
    expect(remarkCell).not.toMatch(/电话:\d+;111/)
  })
})
