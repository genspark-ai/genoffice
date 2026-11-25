import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

it('keeps adjacent links with the same target and different history settings separate', async () => {
  const target = 'https://example.com/doc.html#section2'
  const bytes = await buildDocx({
    bodyXml:
      '<w:p><w:hyperlink r:id="rId11" w:history="0"><w:r><w:t>first</w:t></w:r></w:hyperlink>' +
      '<w:hyperlink r:id="rId11"><w:r><w:t>second</w:t></w:r></w:hyperlink></w:p>',
    extraRels:
      '<Relationship Id="rId11" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
      `Target="${target}" TargetMode="External"/>`,
  })
  const parsed = await parseDocx(bytes)
  const paragraph = parsed.blocks[0]
  if (paragraph.type !== 'paragraph' || !paragraph.runs) {
    throw new Error('Expected a paragraph with runs')
  }
  const output = await saveDocx(parsed, [
    {
      kind: 'generated',
      docxIndex: paragraph.docxIndex ?? undefined,
      block: { type: 'paragraph', runs: paragraph.runs },
    },
  ])
  const zip = await JSZip.loadAsync(output)
  const xml = await zip.file('word/document.xml')!.async('string')
  expect(xml.match(/<w:hyperlink\b[^>]*>/g)).toEqual([
    '<w:hyperlink r:id="rId11" w:history="0">',
    '<w:hyperlink r:id="rId11">',
  ])

  const reopened = (await parseDocx(output)).blocks[0]
  if (reopened.type !== 'paragraph' || !reopened.runs) {
    throw new Error('Expected a paragraph with runs')
  }
  expect(reopened.runs.map((run) => run.link?.href)).toEqual([target, target])
})

it('keeps w:tgtFrame and w:docLocation when a nearby run is edited', async () => {
  const bytes = await buildDocx({
    bodyXml:
      '<w:p><w:hyperlink r:id="rId11" w:anchor="ch2" w:tgtFrame="_blank" ' +
      'w:docLocation="page-two"><w:r><w:t>link</w:t></w:r></w:hyperlink>' +
      '<w:r><w:t>after</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId11" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
      'Target="https://example.org/doc" TargetMode="External"/>',
  })
  const parsed = await parseDocx(bytes)
  const paragraph = parsed.blocks[0]
  if (paragraph.type !== 'paragraph' || !paragraph.runs) throw new Error('Expected paragraph')

  expect(paragraph.runs[0].link).toMatchObject({
    rId: 'rId11',
    anchor: 'ch2',
    tgtFrame: '_blank',
    docLocation: 'page-two',
  })

  const runs = paragraph.runs.map((r) => (r.text === 'after' ? { ...r, text: 'end' } : r))
  const saved = await saveDocx(parsed, [
    {
      kind: 'generated',
      docxIndex: paragraph.docxIndex ?? undefined,
      block: { type: 'paragraph', runs },
    },
  ])
  const zip = await JSZip.loadAsync(saved)
  const xml = await zip.file('word/document.xml')!.async('string')
  expect(xml).toContain('w:tgtFrame="_blank"')
  expect(xml).toContain('w:docLocation="page-two"')
  expect(xml).toContain('w:anchor="ch2"')
  expect(xml).toContain('r:id="rId11"')

  const reopened = await parseDocx(saved)
  const link = (reopened.blocks[0] as { runs?: Array<{ link?: Record<string, string> }> }).runs?.[0]
    .link
  expect(link).toMatchObject({ tgtFrame: '_blank', docLocation: 'page-two' })
})
