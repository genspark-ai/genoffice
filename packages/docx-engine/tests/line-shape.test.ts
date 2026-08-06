import { describe, expect, it } from 'vitest'
import { buildLineParagraphXml, patchTextboxSizes } from '../src/generate'
import { parseDocx } from '../src/parse'
import { buildDocx } from './helpers/build-docx'

describe('line shapes', () => {
  it('builds stroke-only connector xml with arrow ends', () => {
    const xml = buildLineParagraphXml({ kind: 'lineArrowDouble', id: 7 })
    expect(xml).toContain('prst="straightConnector1"')
    expect(xml).toContain('<a:headEnd type="triangle"/>')
    expect(xml).toContain('<a:tailEnd type="triangle"/>')
    expect(xml).toContain('<a:noFill/>')
    expect(xml).not.toContain('<w:txbxContent')
  })

  it('patches extents of line drawings (no text body)', () => {
    const xml = buildLineParagraphXml({ kind: 'line', widthEmu: 1800000, heightEmu: 114300 })
    const out = patchTextboxSizes(xml, [{ wPx: 300, hPx: 20 }])
    expect(out).toContain(`cx="${300 * 9525}"`)
    expect(out).toContain(`cy="${20 * 9525}"`)
  })

  it('round-trips through parse as a readOnly display box', async () => {
    const paragraph = buildLineParagraphXml({
      kind: 'lineArrow',
      widthEmu: 1905000,
      heightEmu: 114300,
    })
    const doc = await parseDocx(await buildDocx({ bodyXml: paragraph }))
    const box = doc.blocks.find((b) => b.textboxes?.length)?.textboxes?.[0]
    expect(box).toBeTruthy()
    expect(box!.prst).toBe('lineArrow')
    expect(box!.readOnly).toBe(true)
    expect(box!.widthPx).toBe(200)
    expect(box!.heightPx).toBe(12)
    expect(box!.paras).toEqual([])
  })
})
