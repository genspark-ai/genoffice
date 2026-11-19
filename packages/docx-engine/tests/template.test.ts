/** A template is the document package with its main part declared as one, and back. */
import { mkdtemp, writeFile } from 'node:fs/promises'
import { crc32, deflateRawSync } from 'node:zlib'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  docxAsTemplate,
  parseDocx,
  templateAsDocx,
  templateThumbnail,
  withTemplateThumbnail,
} from '../src/index'
import {
  lazyMediaHashesIn,
  materializeDocx,
  openZipFile,
  slimDocx,
  writeZip,
  type ZipEntry,
} from '../src/zip-splice'
import { IMAGE_PARAGRAPH_XML, buildDocx } from './helpers/build-docx'

const types = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('[Content_Types].xml')!.async('string')

describe('document and template packages', () => {
  it('declares the main part a template, and a document again, leaving the content alone', async () => {
    const docx = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Dear reader</w:t></w:r></w:p>' })
    expect(await types(docx)).toContain('wordprocessingml.document.main+xml')

    const dotx = await docxAsTemplate(docx)
    expect(await types(dotx)).toContain('wordprocessingml.template.main+xml')
    expect(await types(dotx)).not.toContain('wordprocessingml.document.main+xml')

    const back = await templateAsDocx(dotx)
    expect(await types(back)).toContain('wordprocessingml.document.main+xml')
    const parsed = await parseDocx(back)
    expect((parsed.blocks[0] as { runs: Array<{ text: string }> }).runs[0].text).toBe('Dear reader')
  })

  it('hands back a package that is already what was asked for', async () => {
    const docx = await buildDocx({ bodyXml: '<w:p/>' })
    expect(await templateAsDocx(docx)).toBe(docx)
    const dotx = await docxAsTemplate(docx)
    expect(await docxAsTemplate(dotx)).toBe(dotx)
  })

  it('keeps lazily served pictures restorable: a slimmed template turned document still names its source', async () => {
    const hash = 'b'.repeat(64)
    const big = Buffer.alloc(4096, 7)
    const template = await docxAsTemplate(
      await buildDocx({
        bodyXml: IMAGE_PARAGRAPH_XML,
        withImage: true,
        binaryParts: [
          {
            path: 'word/media/image2.png',
            base64: big.toString('base64'),
            extension: 'png',
            contentType: 'image/png',
          },
        ],
      }),
    )
    const path = join(await mkdtemp(join(tmpdir(), 'template-lazy-')), 'big.dotx')
    await writeFile(path, template)
    const file = await openZipFile(path)
    try {
      const slim = await slimDocx(file, hash, 1024)
      expect(slim).not.toBeNull()
      const asDocument = Buffer.from(await templateAsDocx(slim!.bytes))
      expect(await types(asDocument)).toContain('wordprocessingml.document.main+xml')
      expect(await lazyMediaHashesIn(asDocument)).toEqual(new Set([hash]))
      const full = await materializeDocx(asDocument, async (h) =>
        h === hash ? await openZipFile(path) : null,
      )
      const media = await (
        await JSZip.loadAsync(full)
      )
        .file('word/media/image2.png')!
        .async('uint8array')
      expect(Buffer.from(media)).toEqual(big)
    } finally {
      await file.close()
    }
  })
  it('makes a document without the macros out of a macro-enabled template', async () => {
    const zip = await JSZip.loadAsync(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Macro letter</w:t></w:r></w:p>',
        extraRels:
          '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>',
        extraParts: [
          {
            path: 'word/vbaData.xml',
            xml: '<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>',
            contentType: 'application/vnd.ms-word.vbaData+xml',
          },
        ],
        binaryParts: [
          {
            path: 'word/vbaProject.bin',
            base64: Buffer.from('not a real project').toString('base64'),
            extension: 'bin',
            contentType: 'application/vnd.ms-office.vbaProject',
          },
        ],
      }),
    )
    zip.file(
      'word/_rels/vbaProject.bin.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/>' +
        '</Relationships>',
    )
    zip.file(
      '[Content_Types].xml',
      (await zip.file('[Content_Types].xml')!.async('string')).replace(
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
      ),
    )
    const dotm = await zip.generateAsync({ type: 'uint8array' })

    const docx = await templateAsDocx(dotm)
    const out = await JSZip.loadAsync(docx)
    expect(Object.keys(out.files).filter((name) => /vba/i.test(name))).toEqual([])
    const outTypes = await types(docx)
    expect(outTypes).toContain('wordprocessingml.document.main+xml')
    expect(outTypes).not.toMatch(/macroEnabled|vbaData/)
    expect(await out.file('word/_rels/document.xml.rels')!.async('string')).not.toContain('vba')
    const parsed = await parseDocx(docx)
    expect((parsed.blocks[0] as { runs: Array<{ text: string }> }).runs[0].text).toBe(
      'Macro letter',
    )
  })
  it('carries a preview picture as a template and leaves it behind as a document', async () => {
    const docx = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Preview me</w:t></w:r></w:p>' })
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xd9])
    const dotx = await withTemplateThumbnail(await docxAsTemplate(docx), jpeg)
    expect(await templateThumbnail(dotx)).toEqual({ mime: 'image/jpeg', data: jpeg })
    expect(await types(dotx)).toContain('Extension="jpeg"')
    // a second picture takes the place of the first
    const again = await withTemplateThumbnail(dotx, new Uint8Array([0xff, 0xd8, 1, 0xff, 0xd9]))
    const rels = await (await JSZip.loadAsync(again)).file('_rels/.rels')!.async('string')
    expect(rels.match(/metadata\/thumbnail/g)).toHaveLength(1)
    expect((await templateThumbnail(again))?.data.length).toBe(5)

    const made = await templateAsDocx(again)
    expect(await templateThumbnail(made)).toBeNull()
    expect(Object.keys((await JSZip.loadAsync(made)).files)).not.toContain(
      'docProps/thumbnail.jpeg',
    )
    expect(await templateThumbnail(docx)).toBeNull()
  })

  it('drops the declaration of the preview picture with the picture', async () => {
    const OVERRIDE =
      '<Override PartName="/docProps/thumbnail.jpeg" ContentType="image/jpeg"/></Types>'
    const docx = await buildDocx({ bodyXml: '<w:p/>' })
    const zip = await JSZip.loadAsync(
      await withTemplateThumbnail(await docxAsTemplate(docx), new Uint8Array([0xff, 0xd8])),
    )
    zip.file(
      '[Content_Types].xml',
      (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', OVERRIDE),
    )
    const dotx = await zip.generateAsync({ type: 'uint8array' })
    expect(await types(dotx)).toContain('/docProps/thumbnail.jpeg')

    const made = await templateAsDocx(dotx)
    expect(await types(made)).not.toContain('thumbnail')
    expect(await types(made)).toContain('/word/document.xml')
    await expect(parseDocx(made)).resolves.toBeTruthy()
  })

  it('makes a document without the preview picture from a .docx used as a template', async () => {
    const docx = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Kept</w:t></w:r></w:p>' })
    // nothing to drop: the same bytes
    expect(await templateAsDocx(docx)).toBe(docx)
    const withPicture = await withTemplateThumbnail(docx, new Uint8Array([0xff, 0xd8]))
    expect(await types(withPicture)).toContain('wordprocessingml.document.main+xml')
    expect(await templateThumbnail(withPicture)).not.toBeNull()

    const made = await templateAsDocx(withPicture)
    expect(await templateThumbnail(made)).toBeNull()
    expect(Object.keys((await JSZip.loadAsync(made)).files)).not.toContain(
      'docProps/thumbnail.jpeg',
    )
    const parsed = await parseDocx(made)
    expect((parsed.blocks[0] as { runs: Array<{ text: string }> }).runs[0].text).toBe('Kept')
  })

  it('refuses a part that inflates past what the package declares, for the preview and for a new document', async () => {
    const part = (name: string, truth: Buffer, declares = truth.length) => {
      const data = deflateRawSync(truth)
      const meta: ZipEntry = {
        name,
        nameBytes: Buffer.from(name),
        flags: 0,
        method: 8,
        time: 0,
        date: 0,
        crc: crc32(truth),
        csize: data.length,
        usize: declares,
        verMade: 20,
        verNeed: 20,
        intAttr: 0,
        extAttr: 0,
        dataOffset: 0,
      }
      return { meta, data }
    }
    const rels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>' +
      '</Relationships>'
    const honest = writeZip([
      part('_rels/.rels', Buffer.from(rels)),
      part('docProps/thumbnail.jpeg', Buffer.from([0xff, 0xd8, 0xff, 0xd9])),
    ])
    expect((await templateThumbnail(honest))?.data.length).toBe(4)
    const bomb = writeZip([
      part('_rels/.rels', Buffer.from(rels)),
      part('docProps/thumbnail.jpeg', Buffer.alloc(64 * 1024 * 1024), 300),
    ])
    await expect(templateThumbnail(bomb)).rejects.toThrow(/inflates past/)
    await expect(templateAsDocx(bomb)).rejects.toThrow(/inflates past/)
  })
})
