import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

async function zipText(bytes: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return (await zip.file(path)?.async('string')) ?? ''
}

describe('text wrapping break (w:br w:type="textWrapping" w:clear="all")', () => {
  it('parses to its own control character, distinct from soft and column breaks', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t>' +
        '<w:br w:type="textWrapping" w:clear="all"/><w:t>c</w:t>' +
        '<w:br w:type="column"/><w:t>d</w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(source)
    const para = parsed.blocks.find((b) => b.type === 'paragraph')!
    expect(para.runs!.map((r) => r.text).join('')).toBe('a\nb\u001ec\vd')
  })

  it('writes a generated paragraph back as a textWrapping break that clears floats', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'a\u001eb' }] } },
    ])
    const xml = await zipText(saved, 'word/document.xml')
    expect(xml).toContain(
      '<w:t xml:space="preserve">a</w:t><w:br w:type="textWrapping" w:clear="all"/><w:t xml:space="preserve">b</w:t>',
    )
    expect(xml).not.toContain('\u001e')
  })
})

describe('mirrorMargins settings flag', () => {
  it('is off by default and round-trips through settings.xml', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    expect(parsed.mirrorMargins).toBeUndefined()

    const on = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      mirrorMargins: true,
    })
    expect(await zipText(on, 'word/settings.xml')).toContain('<w:mirrorMargins/>')
    const reparsed = await parseDocx(on)
    expect(reparsed.mirrorMargins).toBe(true)

    const off = await saveDocx(reparsed, [{ kind: 'original', docxIndex: 0 }], {
      mirrorMargins: false,
    })
    expect(await zipText(off, 'word/settings.xml')).not.toContain('mirrorMargins')
    expect((await parseDocx(off)).mirrorMargins).toBeUndefined()
  })
})
