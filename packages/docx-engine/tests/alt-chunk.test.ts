import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'
import { saveDocx } from '../src/patch'
import { decodeMhtToHtml, decodeQuotedPrintable, setAltChunkHtmlConverter } from '../src/alt-chunk'
import { buildDocx } from './helpers/build-docx'

const CHUNK_REL =
  '<Relationship Id="rIdChunk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="/word/afchunk.htm"/>'

const HTML =
  '<!DOCTYPE html><html><body><p>Simple paragraph with a <strong>bold</strong> word.</p>' +
  '<table><tr><th>Col 1</th><th>Col 2</th></tr><tr><td>ROW 1</td><td>ROW 1</td></tr></table>' +
  '</body></html>'

/** what the html2docx chain would hand back for HTML: a paragraph and a table */
const CONVERTED_BODY =
  '<w:p><w:r><w:t xml:space="preserve">Simple paragraph with a </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve"> word.</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>Col 1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Col 2</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>ROW 1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>ROW 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

function hostDocx(chunk: { path: string; body: string | Uint8Array; contentType: string }) {
  return buildDocx({
    bodyXml: '<w:altChunk r:id="rIdChunk"/><w:p><w:r><w:t>After the chunk</w:t></w:r></w:p>',
    extraRels: CHUNK_REL.replace('/word/afchunk.htm', `/${chunk.path}`),
    extraParts: [{ path: chunk.path, xml: chunk.body as string, contentType: chunk.contentType }],
  })
}

const received: string[] = []
function installStubConverter(): void {
  setAltChunkHtmlConverter(async (html) => {
    received.push(html)
    return buildDocx({ bodyXml: CONVERTED_BODY })
  })
}

afterEach(() => {
  setAltChunkHtmlConverter(null)
  received.length = 0
})

describe('w:altChunk expansion', () => {
  it('shows the converted HTML as paragraph + table blocks, keeping the chunk in the body', async () => {
    installStubConverter()
    const bytes = await hostDocx({ path: 'word/afchunk.htm', body: HTML, contentType: 'text/html' })
    const parsed = await parseDocx(bytes)

    expect(received).toEqual([HTML])
    const visible = parsed.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => b.type)).toEqual(['paragraph', 'table', 'paragraph'])
    expect(visible[0].runs?.map((r) => r.text).join('')).toBe('Simple paragraph with a bold word.')
    expect(visible[1].table?.rows).toHaveLength(2)
    expect(visible.map((b) => b.docxIndex)).toEqual([0, 1, 2])
    // chunk blocks are laid out by the chunk's own conventions, not the host's compat mode
    expect(visible.map((b) => b.altChunk)).toEqual([true, true, undefined])

    // the chunk element owns the first range; the second block has an empty one
    const { elements } = parsed.extras
    const docXml = parsed.internal.documentXml
    expect(docXml.slice(elements[0].start, elements[0].end)).toBe('<w:altChunk r:id="rIdChunk"/>')
    expect(elements[1].start).toBe(elements[1].end)
    expect(elements[1].start).toBe(elements[0].end)

    // untouched save is byte-identical
    const unchanged = await saveDocx(
      parsed,
      visible.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(Buffer.from(unchanged).equals(Buffer.from(bytes))).toBe(true)

    // editing another paragraph re-emits the chunk once and none of its display blocks
    const edited = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { kind: 'original', docxIndex: 1 },
      { kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'Edited' }] } },
    ])
    const outXml = await (await JSZip.loadAsync(edited)).file('word/document.xml')!.async('string')
    expect(outXml.match(/<w:altChunk/g)).toHaveLength(1)
    expect(outXml).not.toContain('Col 1')
    expect(outXml).toContain('Edited')
  })

  it('adopts chunk numbering under fresh ids, in body and table-cell lists alike', async () => {
    const listPara = (numId: string, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    setAltChunkHtmlConverter(() =>
      buildDocx({
        withNumbering: true,
        bodyXml:
          listPara('1', 'chunk bullet') +
          `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${listPara('1', 'cell bullet')}</w:tc></w:tr></w:tbl>`,
      }),
    )
    const bytes = await buildDocx({
      withNumbering: true,
      bodyXml: '<w:altChunk r:id="rIdChunk"/>' + listPara('2', 'host decimal'),
      extraRels: CHUNK_REL,
      extraParts: [{ path: 'word/afchunk.htm', xml: HTML, contentType: 'text/html' }],
    })
    const parsed = await parseDocx(bytes)
    const [bullet, table, host] = parsed.blocks.filter((b) => !b.hidden)
    expect(host.list?.numId).toBe('2')
    const mapped = bullet.list?.numId
    expect(mapped).toBe('3')
    expect(table.table?.rows[0][0].richParas?.[0].list?.numId).toBe(mapped)
    const def = parsed.numbering.get(mapped!)
    expect(def?.numId).toBe(mapped)
    expect(def?.levels[0].numFmt).toBe('bullet')
    expect(def?.abstractNumId).not.toBe(parsed.numbering.get('1')?.abstractNumId)
  })

  it('decodes an MHT chunk (quoted-printable, charset) before converting', async () => {
    installStubConverter()
    const mht =
      'MIME-Version: 1.0\r\n' +
      'Content-Type: multipart/related;\r\n\tboundary="----=_NextPart_01"\r\n' +
      '\r\n' +
      '------=_NextPart_01\r\n' +
      'Content-Location: file:///C:/doc.htm\r\n' +
      'Content-Transfer-Encoding: quoted-printable\r\n' +
      'Content-Type: text/html; charset="utf-8"\r\n' +
      '\r\n' +
      '<html><body><p>Caf=C3=A9 =\r\nrow</p><img src=3D"image001.png"></body></html>\r\n' +
      '------=_NextPart_01\r\n' +
      'Content-Location: file:///C:/image001.png\r\n' +
      'Content-Transfer-Encoding: base64\r\n' +
      'Content-Type: image/png\r\n' +
      '\r\n' +
      'iVBORw0KGgo=\r\n' +
      '------=_NextPart_01--\r\n'
    const bytes = await hostDocx({
      path: 'word/afchunk.mht',
      body: mht,
      contentType: 'message/rfc822',
    })
    const parsed = await parseDocx(bytes)

    expect(received).toHaveLength(1)
    expect(received[0]).toContain('<p>Caf\u00e9 row</p>')
    expect(received[0]).toContain('src="data:image/png;base64,iVBORw0KGgo="')
    expect(parsed.blocks.filter((b) => !b.hidden).map((b) => b.type)).toEqual([
      'paragraph',
      'table',
      'paragraph',
    ])
  })

  it('parses a .docx chunk directly, without the HTML converter', async () => {
    const inner = await buildDocx({ bodyXml: CONVERTED_BODY })
    const zip = await JSZip.loadAsync(
      await hostDocx({
        path: 'word/afchunk.dat',
        body: '',
        contentType: 'application/octet-stream',
      }),
    )
    zip.file('word/afchunk.dat', inner)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const parsed = await parseDocx(bytes)
    expect(parsed.blocks.filter((b) => !b.hidden).map((b) => b.type)).toEqual([
      'paragraph',
      'table',
      'paragraph',
    ])
  })

  it('leaves the chunk as a passthrough when no converter is installed', async () => {
    const bytes = await hostDocx({ path: 'word/afchunk.htm', body: HTML, contentType: 'text/html' })
    const parsed = await parseDocx(bytes)
    const visible = parsed.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => b.type)).toEqual(['passthrough', 'paragraph'])
    expect(visible[0].label).toBe('w:altChunk')
    // a host that can convert (the UI thread) reparses when it sees this
    expect(parsed.extras.altChunksNeedConverter).toBe(1)
  })

  it('does not flag chunks when a converter is installed', async () => {
    installStubConverter()
    const bytes = await hostDocx({ path: 'word/afchunk.htm', body: HTML, contentType: 'text/html' })
    const parsed = await parseDocx(bytes)
    expect(parsed.extras.altChunksNeedConverter).toBeUndefined()
  })
})

describe('MIME decoding helpers', () => {
  it('quoted-printable: soft breaks, escapes, literal bytes', () => {
    const out = decodeQuotedPrintable('a=3Db=\r\nc=E9')
    expect([...out]).toEqual([0x61, 0x3d, 0x62, 0x63, 0xe9])
  })

  it('single-part MHT without a boundary is the HTML itself', () => {
    const mht =
      'Content-Type: text/html; charset=windows-1252\r\n' +
      'Content-Transfer-Encoding: 8bit\r\n\r\n<html><body>\u00e9</body></html>'
    const bytes = Uint8Array.from(mht, (c) => c.charCodeAt(0) & 0xff)
    expect(decodeMhtToHtml(bytes)).toBe('<html><body>\u00e9</body></html>')
  })

  it('inlines thousands of image parts in one pass over the html', () => {
    const boundary = 'bomb'
    let mht =
      `Content-Type: multipart/related; boundary="${boundary}"\r\n\r\n` +
      `--${boundary}\r\nContent-Type: text/html\r\n\r\n<html><body><p>hi</p><img src="F1999.PNG"><img src=f0.png></body></html>\r\n`
    for (let i = 0; i < 2000; i++) {
      mht += `--${boundary}\r\nContent-Type: image/png\r\nContent-Location: f${i}.png\r\n\r\nx\r\n`
    }
    mht += `--${boundary}--\r\n`
    const bytes = Uint8Array.from(mht, (c) => c.charCodeAt(0) & 0xff)
    const html = decodeMhtToHtml(bytes)
    expect(html).toContain('<p>hi</p>')
    expect(html).toContain(
      '<img src="data:image/png;base64,eA=="><img src="data:image/png;base64,eA==">',
    )
  })
})
