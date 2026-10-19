/** Synthetic PDF font metadata through real PDFium extraction and DOCX rebuild. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertPdfToDocx, extractIrDocument } from '../src'
import { loadPdfium } from './helpers/wasm'

// Self-generated rectangle glyphs; only the name table models the reported font.
// This exercises naming, not DFKai-SB glyph metrics or visual fidelity.
const FONT_BASE64 =
  'AAEAAAAKAIAAAwAgT1MvMqDI1w8AAAEoAAAAYGNtYXARf515AAABlAAAAExnbHlmyLZS5wAAAewAAACCaGVhZDDoAZoAAACsAAAANmhoZWEGQgPqAAAA5AAAACRobXR4A+gAAAAAAYgAAAAMbG9jYQBOAHUAAAHgAAAADG1heHAABwAGAAABCAAAACBuYW1lN6pKYwAAAnAAAAEXcG9zdPnGrqIAAAOIAAAATAABAAAAAQAA1esMgl8PPPUAAQPoAAAAAOb13v8AAAAA5vXe/wBkAAADhAK8AAAAAwACAAAAAAAAAAEAAAMg/zgAAAPoAAAAyAMgAAEAAAAAAAAAAAAAAAAAAAABAAEAAAAFAAQAAQAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAwPoAZAABQAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAPz8/PwAAW+mJiwMg/zgAAAMgAMgAAAAAAAAAAAAAAAAAAAAgAAAD6AAAAAAAAAAAAAAAAAACAAAAAwAAABQAAwABAAAAFAAEADgAAAAKAAgAAgACW+lhD2fliYv//wAAW+lhD2fliYv//6QYnvSYHXZ5AAEAAAAAAAAAAAAAAAAADQAaACcANABBAAEAZAAAA4QCvAADAAAzIREhZAMg/OACvAAAAQBkAAADhAK8AAMAADMhESFkAyD84AK8AAABAGQAAAOEArwAAwAAMyERIWQDIPzgArwAAAEAZAAAA4QCvAADAAAzIREhZAMg/OACvAAAAQBkAAADhAK8AAMAADMhESFkAyD84AK8AAAAAAAACgB+AAEAAAAAAAEAEwAAAAEAAAAAAAIABwATAAEAAAAAAAMAGQAaAAEAAAAAAAQAEwAAAAEAAAAAAAYAEwAAAAMAAQQJAAEAJgAzAAMAAQQJAAIADgBZAAMAAQQJAAMAMgBnAAMAAQQJAAQAJgAzAAMAAQQJAAYAJgAzREZLYWlTaHUtU0ItRXN0ZC1CRlJlZ3VsYXJTeW50aGV0aWMgZm9udCBhbGlhcyB0ZXN0AEQARgBLAGEAaQBTAGgAdQAtAFMAQgAtAEUAcwB0AGQALQBCAEYAUgBlAGcAdQBsAGEAcgBTAHkAbgB0AGgAZQB0AGkAYwAgAGYAbwBuAHQAIABhAGwAaQBhAHMAIAB0AGUAcwB0AAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAECAQMBBAEFB3VuaTVCRTkHdW5pNjdFNQd1bmk2MTBGB3VuaTg5OEI='
function fontMetadataPdf(psName: string): Uint8Array {
  const font = Buffer.from(FONT_BASE64, 'base64')
  const content = 'BT /F1 12 Tf 72 700 Td <0001000200030004> Tj ET'
  const cmap = `/CIDInit /ProcSet findresource begin
12 dict begin begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /Synthetic def /CMapType 2 def
1 begincodespacerange <0000> <FFFF> endcodespacerange
4 beginbfchar <0001> <5BE9> <0002> <67E5> <0003> <610F> <0004> <898B> endbfchar
endcmap CMapName currentdict /CMap defineresource pop end end`
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
    `<< /Type /Font /Subtype /Type0 /BaseFont /${psName} /Encoding /Identity-H /DescendantFonts [6 0 R] /ToUnicode 8 0 R >>`,
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${psName} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 7 0 R /DW 1000 >>`,
    `<< /Type /FontDescriptor /FontName /${psName} /Flags 4 /FontBBox [0 -200 1000 1000] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 9 0 R >>`,
    `<< /Length ${cmap.length} >>\nstream\n${cmap}\nendstream`,
    `<< /Length ${font.length} /Length1 ${font.length} >>\nstream\n${font.toString('latin1')}\nendstream`,
  ]
  let pdf = '%PDF-1.7\n'
  const offsets = objects.map((obj, i) => {
    const offset = pdf.length
    pdf += `${i + 1} 0 obj\n${obj}\nendobj\n`
    return offset
  })
  const xref = pdf.length
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const offset of offsets) pdf += `${String(offset).padStart(10, '0')} 00000 n \n`
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Uint8Array.from(Buffer.from(pdf, 'latin1'))
}

describe('PDF font aliases in DOCX', () => {
  it.each([
    ['DFKaiShu-SB-Estd-BF', 'DFKai-SB'],
    ['XLRIPU+DFKaiShu-SB-Estd-BF', 'DFKai-SB'],
  ])('keeps %s resolved as %s from font metadata to rFonts', async (psName, family) => {
    const pdfium = await loadPdfium()
    const pdf = fontMetadataPdf(psName!)
    const ir = extractIrDocument(pdf, { pdfium })
    const spans = ir.irPages.flatMap((p) =>
      p.blocks.flatMap((b) => (b.kind === 'text' ? b.lines.flatMap((l) => l.spans) : [])),
    )
    expect(spans.map((s) => s.text).join('')).toBe('\u5be9\u67e5\u610f\u898b')
    const result = await convertPdfToDocx(pdf, { pdfium })
    const zip = await JSZip.loadAsync(result.docx)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('\u5be9\u67e5\u610f\u898b')
    expect(xml).toContain(`w:eastAsia="${family}"`)
    expect(xml).not.toContain('DFKai Shu SB Estd BF')
    expect(spans.map((s) => s.fontFamily)).toEqual([family])
  })
})
