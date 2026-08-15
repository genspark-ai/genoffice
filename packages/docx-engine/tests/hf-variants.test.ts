import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const BODY = '<w:p><w:r><w:t>正文</w:t></w:r></w:p>'

async function zipText(bytes: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(path)
  return file ? file.async('string') : null
}

describe('first-page / even-page headers & footers', () => {
  it('creates typed parts, references and flags; round-trips through parse', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    expect(parsed.titlePg).toBe(false)
    expect(parsed.evenAndOddHeaders).toBe(false)
    expect(parsed.headerFirst).toBeNull()

    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: '默认页眉' },
      headerFirst: { text: '首页页眉' },
      headerEven: { text: '偶数页页眉' },
      footerEven: { text: '偶 #', pageNumber: true },
      titlePg: true,
      evenAndOddHeaders: true,
    })

    const docXml = (await zipText(saved, 'word/document.xml'))!
    expect(docXml).toContain('<w:titlePg/>')
    expect(docXml).toMatch(/<w:headerReference w:type="first"/)
    expect(docXml).toMatch(/<w:headerReference w:type="even"/)
    expect(docXml).toMatch(/<w:footerReference w:type="even"/)
    const settingsXml = (await zipText(saved, 'word/settings.xml'))!
    expect(settingsXml).toContain('<w:evenAndOddHeaders/>')

    const reparsed = await parseDocx(saved)
    expect(reparsed.titlePg).toBe(true)
    expect(reparsed.evenAndOddHeaders).toBe(true)
    expect(reparsed.headerText).toBe('默认页眉')
    expect(reparsed.headerFirst?.text).toBe('首页页眉')
    expect(reparsed.headerEven?.text).toBe('偶数页页眉')
    expect(reparsed.footerEven?.text).toBe(`偶 ${PAGE_MARK}`)
    expect(reparsed.footerEven?.hasPageNumber).toBe(true)
    expect(reparsed.footerFirst).toBeNull()
  })

  it('gives each new part a distinct filename in a single save', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: 'a' },
      headerFirst: { text: 'b' },
      headerEven: { text: 'c' },
    })
    const zip = await JSZip.loadAsync(saved)
    const headers = Object.keys(zip.files).filter((n) => /^word\/header\d+\.xml$/.test(n))
    expect(headers.sort()).toEqual(['word/header1.xml', 'word/header2.xml', 'word/header3.xml'])
  })

  it('rewrites only the edited variant; the others stay byte-identical', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    const withBoth = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: '默认' },
      headerFirst: { text: '首页' },
      titlePg: true,
    })
    const p2 = await parseDocx(withBoth)
    const firstPartBefore = await zipText(withBoth, 'word/header2.xml')
    const edited = await saveDocx(p2, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: '默认已改' },
    })
    expect(await zipText(edited, 'word/header2.xml')).toBe(firstPartBefore)
    const p3 = await parseDocx(edited)
    expect(p3.headerText).toBe('默认已改')
    expect(p3.headerFirst?.text).toBe('首页')
    expect(p3.titlePg).toBe(true)
  })

  it('removes titlePg and evenAndOddHeaders when toggled off', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    const parsed = await parseDocx(source)
    const on = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      titlePg: true,
      evenAndOddHeaders: true,
    })
    const p2 = await parseDocx(on)
    const off = await saveDocx(p2, [{ kind: 'original', docxIndex: 0 }], {
      titlePg: false,
      evenAndOddHeaders: false,
    })
    const p3 = await parseDocx(off)
    expect(p3.titlePg).toBe(false)
    expect(p3.evenAndOddHeaders).toBe(false)
    expect(await zipText(off, 'word/document.xml')).not.toContain('<w:titlePg/>')
  })

  it('keeps titlePg before w:docGrid in schema order', async () => {
    const source = await buildDocx({ bodyXml: BODY })
    // inject a docGrid into the trailing sectPr
    const zip = await JSZip.loadAsync(source)
    const docXml = (await zip.file('word/document.xml')!.async('string')).replace(
      '</w:sectPr>',
      '<w:docGrid w:linePitch="360"/></w:sectPr>',
    )
    zip.file('word/document.xml', docXml)
    const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { titlePg: true })
    const out = (await zipText(saved, 'word/document.xml'))!
    expect(out.indexOf('<w:titlePg/>')).toBeGreaterThan(-1)
    expect(out.indexOf('<w:titlePg/>')).toBeLessThan(out.indexOf('<w:docGrid'))
  })
})
