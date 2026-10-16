import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { run, tempDir } from './helpers'
import { parseDocx, saveDocx } from '../../docx-engine/src/index'
import { TINY_PNG_BASE64 } from '../../docx-engine/tests/helpers/build-docx'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const hdr = (text: string) =>
  `<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`
const ref = (type: string, id: string) => `<w:headerReference w:type="${type}" r:id="${id}"/>`

async function fixture(linked = false, missingFirst = false) {
  return buildDocx({
    bodyXml: `<w:p><w:pPr><w:sectPr>${ref('default', 'rH1')}${missingFirst ? '' : ref('first', 'rF1')}${ref('even', 'rE1')}<w:titlePg/></w:sectPr></w:pPr><w:r><w:t>Caption page.</w:t></w:r></w:p><w:p><w:r><w:t>Section 2 body.</w:t></w:r></w:p>`,
    sectPrExtra: linked ? '' : ref('default', 'rH2'),
    extraRels: ['H1', 'F1', 'E1', 'H2']
      .map((id) => `<Relationship Id="r${id}" Type="${R}/header" Target="header${id}.xml"/>`)
      .join(''),
    extraParts: ['H1', 'F1', 'E1', 'H2'].map((id) => ({
      path: `word/header${id}.xml`,
      xml: hdr(`Keep ${id}`),
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    })),
  })
}

it.each([false, true])(
  'CLI watermarks every active header, preserving linked=%s references and content',
  async (linked) => {
    const dir = tempDir()
    const input = join(dir, 'in.docx')
    const ops = join(dir, 'ops.json')
    const bytes = await fixture(linked)
    const zip = await JSZip.loadAsync(bytes)
    zip.file('word/settings.xml', `<w:settings xmlns:w="${W}"><w:evenAndOddHeaders/></w:settings>`)
    writeFileSync(input, await zip.generateAsync({ type: 'nodebuffer' }))
    writeFileSync(ops, JSON.stringify([{ op: 'set_watermark', text: 'DRAFT' }]))
    const result = await run(['docs', 'apply', input, '--ops', ops, '--json'])
    expect(result.code, result.stderr).toBe(0)
    const saved = await JSZip.loadAsync(readFileSync(input))
    expect(await saved.file('word/document.xml')!.async('string')).toBe(
      await zip.file('word/document.xml')!.async('string'),
    )
    for (const id of linked ? ['H1', 'F1', 'E1'] : ['H1', 'F1', 'E1', 'H2']) {
      const xml = await saved.file(`word/header${id}.xml`)!.async('string')
      expect(xml, id).toContain('string="DRAFT"')
      expect(xml, id).toContain(`<w:p><w:r><w:t>Keep ${id}</w:t></w:r></w:p>`)
    }
    expect(await saved.file('word/_rels/document.xml.rels')!.async('string')).toBe(
      await zip.file('word/_rels/document.xml.rels')!.async('string'),
    )
    writeFileSync(ops, JSON.stringify([{ op: 'set_watermark', text: null }]))
    expect((await run(['docs', 'apply', input, '--ops', ops, '--json'])).code).toBe(0)
    const removed = await JSZip.loadAsync(readFileSync(input))
    for (const name of Object.keys(removed.files).filter((n) => /^word\/header.*\.xml$/.test(n))) {
      expect(await removed.file(name)!.async('string')).not.toContain('PowerPlusWaterMarkObject')
    }
  },
)

async function engineSave(bytes: Uint8Array, options: Parameters<typeof saveDocx>[2]) {
  const parsed = await parseDocx(bytes)
  return saveDocx(
    parsed,
    parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    options,
  )
}

it('creates an absent active first-page header and leaves disabled even headers untouched', async () => {
  const bytes = await fixture(false, true)
  const original = await JSZip.loadAsync(bytes)
  const out = await JSZip.loadAsync(await engineSave(bytes, { watermark: 'DRAFT' }))
  expect(await out.file('word/headerE1.xml')!.async('string')).toBe(
    await original.file('word/headerE1.xml')!.async('string'),
  )
  expect(await out.file('word/headerF1.xml')!.async('string')).toBe(
    await original.file('word/headerF1.xml')!.async('string'),
  )
  const doc = await out.file('word/document.xml')!.async('string')
  const first = /<w:headerReference w:type="first" r:id="([^"]+)"/.exec(doc)![1]
  const rels = await out.file('word/_rels/document.xml.rels')!.async('string')
  const target = new RegExp(`<Relationship Id="${first}"[^>]*Target="([^"]+)"`).exec(rels)![1]
  expect(await out.file(`word/${target}`)!.async('string')).toContain('string="DRAFT"')
  expect(await out.file('[Content_Types].xml')!.async('string')).toContain(`/word/${target}`)
})

it('keeps an inherited default header and does not create extra parts on removal', async () => {
  const bytes = await fixture(true)
  const out = await JSZip.loadAsync(await engineSave(bytes, { watermark: null }))
  const original = await JSZip.loadAsync(bytes)
  expect(Object.keys(out.files)).toEqual(Object.keys(original.files))
  expect(await out.file('word/document.xml')!.async('string')).toBe(
    await original.file('word/document.xml')!.async('string'),
  )
  expect(await out.file('word/_rels/document.xml.rels')!.async('string')).toBe(
    await original.file('word/_rels/document.xml.rels')!.async('string'),
  )
})

it('adds and removes picture watermarks in each active part without losing other relationships', async () => {
  const original = await JSZip.loadAsync(await fixture(false))
  original.file(
    'word/_rels/headerH1.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdKeep" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
  )
  const bytes = await original.generateAsync({ type: 'uint8array' })
  const withPicture = await engineSave(bytes, {
    watermark: { image: { base64: TINY_PNG_BASE64, mime: 'image/png', widthPx: 1, heightPx: 1 } },
    evenAndOddHeaders: true,
  })
  const out = await JSZip.loadAsync(withPicture)
  for (const id of ['H1', 'F1', 'E1', 'H2']) {
    const part = await out.file(`word/header${id}.xml`)!.async('string')
    const imageId = /<v:imagedata r:id="([^"]+)"/.exec(part)![1]
    const rels = await out.file(`word/_rels/header${id}.xml.rels`)!.async('string')
    const target = new RegExp(`<Relationship Id="${imageId}"[^>]*Target="([^"]+)"`).exec(rels)![1]
    expect(out.file(`word/${target}`)).not.toBeNull()
    expect(part).toContain(`Keep ${id}`)
  }
  const removed = await JSZip.loadAsync(await engineSave(withPicture, { watermark: null }))
  const rels = await removed.file('word/_rels/headerH1.xml.rels')!.async('string')
  expect(rels).toContain('Id="rIdKeep"')
  expect(rels).not.toContain(`${R}/image`)
  for (const id of ['H1', 'F1', 'E1', 'H2']) {
    expect(await removed.file(`word/header${id}.xml`)!.async('string')).not.toContain(
      '<v:imagedata',
    )
  }
})

it('keeps a header text edit and the watermark in the same save', async () => {
  const out = await JSZip.loadAsync(
    await engineSave(await fixture(), { header: { text: 'Updated header' }, watermark: 'DRAFT' }),
  )
  const xml = await out.file('word/headerH2.xml')!.async('string')
  expect(xml).toContain('Updated header')
  expect(xml).toContain('string="DRAFT"')
})
