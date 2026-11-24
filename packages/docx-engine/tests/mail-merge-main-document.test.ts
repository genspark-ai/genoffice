import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { parseDocx, saveDocx } from '../src/index'
import type { SaveOptions } from '../src/patch'
import type { Run } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const withSettings = (inner: string, bodyXml = '<w:p><w:r><w:t>body</w:t></w:r></w:p>') =>
  buildDocx({
    bodyXml,
    extraParts: [
      {
        path: 'word/settings.xml',
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}>${inner}</w:settings>`,
        contentType: SETTINGS_TYPE,
      },
    ],
  })

async function saved(inner: string, mailMergeMainType: SaveOptions['mailMergeMainType']) {
  const parsed = await parseDocx(await withSettings(inner))
  const bytes = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { mailMergeMainType })
  const zip = await JSZip.loadAsync(bytes)
  return {
    xml: await zip.file('word/settings.xml')!.async('string'),
    reparsed: await parseDocx(bytes),
  }
}

describe('the mail-merge kind of a document', () => {
  it('is read from settings.xml', async () => {
    const labels = await parseDocx(
      await withSettings(
        '<w:mailMerge><w:mainDocumentType w:val="mailingLabels"/><w:dataType w:val="textFile"/></w:mailMerge>',
      ),
    )
    expect(labels.mailMergeMainType).toBe('mailingLabels')
    expect(
      (await parseDocx(await withSettings('<w:zoom w:percent="100"/>'))).mailMergeMainType,
    ).toBe(undefined)
  })

  it('is written where CT_Settings has it: after the view settings, before the revision ones', async () => {
    const { xml, reparsed } = await saved(
      '<w:zoom w:percent="100"/><w:trackRevisions/><w:defaultTabStop w:val="720"/>',
      'mailingLabels',
    )
    expect(xml).toContain(
      '<w:zoom w:percent="100"/><w:mailMerge><w:mainDocumentType w:val="mailingLabels"/>' +
        '<w:dataType w:val="textFile"/></w:mailMerge><w:trackRevisions/>',
    )
    expect(reparsed.mailMergeMainType).toBe('mailingLabels')
  })

  it('changes the kind of a main document that has a data source and keeps the source', async () => {
    const source =
      '<w:linkToQuery/><w:dataType w:val="textFile"/><w:query w:val="SELECT * FROM list"/>'
    const { xml } = await saved(
      `<w:mailMerge><w:mainDocumentType w:val="formLetters"/>${source}</w:mailMerge>`,
      'mailingLabels',
    )
    expect(xml).toContain(
      `<w:mailMerge><w:mainDocumentType w:val="mailingLabels"/>${source}</w:mailMerge>`,
    )
    expect(xml.match(/<w:mainDocumentType/g)).toHaveLength(1)
  })

  it('null takes the main document out of the file', async () => {
    const { xml, reparsed } = await saved(
      '<w:zoom w:percent="100"/><w:mailMerge><w:mainDocumentType w:val="mailingLabels"/><w:dataType w:val="textFile"/></w:mailMerge>',
      null,
    )
    expect(xml).not.toContain('mailMerge')
    expect(xml).toContain('<w:zoom w:percent="100"/>')
    expect(reparsed.mailMergeMainType).toBeUndefined()
  })
})

const field = (instr: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const held = (id: number, name: string, text: string) =>
  `<w:bookmarkStart w:id="${id}" w:name="${name}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="${id}"/>`

describe('a SET or an ASK that Word has run', () => {
  const body =
    '<w:p><w:r><w:t xml:space="preserve">Terms: </w:t></w:r>' +
    field('SET Rate "5%"', held(3, 'Rate', '5%')) +
    field('ASK Clerk "Who signs?" \\d "Ada"', held(4, 'Clerk', 'Grace')) +
    '</w:p>'

  it('is a field of an editable paragraph: it shows nothing and keeps the text of its bookmark', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const [block] = parsed.blocks
    expect(block.type).toBe('paragraph')
    const runs = (block.runs ?? []).filter((run: Run) => run.instrField)
    expect(runs.map((run: Run) => [run.instrField, run.text, run.fldBookmarkText])).toEqual([
      ['SET Rate "5%"', '​', '5%'],
      ['ASK Clerk "Who signs?" \\d "Ada"', '​', 'Grace'],
    ])
    // the bookmarks belong to the fields, not to the paragraph
    expect(block.bookmarks).toBeUndefined()
  })

  it('is written back with its bookmark around the result when the paragraph is rebuilt', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const [block] = parsed.blocks
    const runs = [{ text: 'New terms: ' }, ...(block.runs ?? []).slice(1)]
    const bytes = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs } },
    ])
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
    expect(xml).toContain('New terms: ')
    expect(xml).toMatch(
      /> SET Rate "5%" <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:bookmarkStart w:id="(\d+)" w:name="Rate"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>5%<\/w:t><\/w:r><w:bookmarkEnd w:id="\1"\/><w:r><w:fldChar w:fldCharType="end"\/>/,
    )
    expect(xml).toMatch(/w:name="Clerk"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>Grace<\/w:t>/)
    expect(xml.match(/<w:bookmarkStart /g)).toHaveLength(2)
    const again = await parseDocx(bytes)
    expect(
      (again.blocks[0].runs ?? [])
        .filter((run: Run) => run.instrField)
        .map((r) => r.fldBookmarkText),
    ).toEqual(['5%', 'Grace'])
  })

  it('only the SET that holds the bookmark carries it, when two name the same one', async () => {
    const twice =
      `<w:p>${field('SET Rate "4%"', '<w:r><w:t>4%</w:t></w:r>')}` +
      `${field('SET Rate "5%"', held(3, 'Rate', '5%'))}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: twice }))
    expect(
      (parsed.blocks[0].runs ?? []).map((run: Run) => [run.instrField, run.fldBookmarkText]),
    ).toEqual([
      ['SET Rate "4%"', undefined],
      ['SET Rate "5%"', '5%'],
    ])
  })

  it('stays as the file has it when the bookmark is not the one the field names', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: `<w:p>${field('SET Rate "5%"', held(3, 'Other', '5%'))}</w:p>` }),
    )
    expect(parsed.blocks[0].type).not.toBe('paragraph')
  })
})
