/** Insert > File with a range: the content a bookmark covers, as a document of its own. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { docxBookmarkNames, docxBookmarkRange, parseDocx } from '../src/index'
import { buildDocx, TABLE_XML } from './helpers/build-docx'

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const para = (inner: string) => `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${inner}</w:p>`
const start = (id: number, name: string) => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`
const end = (id: number) => `<w:bookmarkEnd w:id="${id}"/>`

async function textsOf(bytes: Uint8Array): Promise<string[]> {
  const parsed = await parseDocx(bytes)
  return parsed.blocks
    .filter((b) => !b.hidden)
    .map((b) => (b.type === 'table' ? '[table]' : (b.runs ?? []).map((r) => r.text).join('')))
}

describe('the range of a bookmark', () => {
  it('lists the bookmarks a user made, not the ones Word keeps for itself', async () => {
    const docx = await buildDocx({
      bodyXml:
        para(start(1, 'Intro') + run('one') + end(1)) +
        para(start(2, '_Toc123') + run('two') + end(2)) +
        para(start(3, 'Terms &amp; Conditions') + run('three') + end(3)),
    })
    expect(await docxBookmarkNames(docx)).toEqual(['Intro', 'Terms & Conditions'])
  })

  it('takes the part of one paragraph the bookmark covers, without its paragraph mark', async () => {
    const docx = await buildDocx({
      bodyXml:
        para(run('before ') + start(1, 'Mid') + run('inside') + end(1) + run(' after')) +
        para(run('next')),
    })
    const range = await docxBookmarkRange(docx, 'Mid')
    expect(range?.inline).toBe(true)
    expect(await textsOf(range!.bytes)).toEqual(['inside'])
    const xml = await (
      await JSZip.loadAsync(range!.bytes)
    )
      .file('word/document.xml')!
      .async('string')
    expect(xml).toContain('<w:jc w:val="center"/>')
    expect(xml).toContain('<w:sectPr')
  })

  it('takes whole paragraphs and a table between the two markers', async () => {
    const docx = await buildDocx({
      bodyXml:
        para(run('zero')) +
        para(run('skipped ') + start(4, 'Span') + run('first')) +
        TABLE_XML +
        para(run('second')) +
        para(end(4) + run('past the end')) +
        para(run('tail')),
    })
    const range = await docxBookmarkRange(docx, 'Span')
    expect(range?.inline).toBe(false)
    expect(await textsOf(range!.bytes)).toEqual(['first', '[table]', 'second'])
  })

  it('ends inside the last paragraph when the end marker sits there', async () => {
    const docx = await buildDocx({
      bodyXml: para(start(1, 'Two') + run('alpha')) + para(run('beta') + end(1) + run(' gamma')),
    })
    const range = await docxBookmarkRange(docx, 'Two')
    expect(range?.inline).toBe(false)
    expect(await textsOf(range!.bytes)).toEqual(['alpha', 'beta'])
  })

  it('takes the whole table a bookmark starts in', async () => {
    const cell = `<w:tbl><w:tr><w:tc><w:p>${start(9, 'Cell')}${run('in a cell')}${end(9)}</w:p></w:tc></w:tr></w:tbl>`
    const docx = await buildDocx({ bodyXml: para(run('top')) + cell + para(run('bottom')) })
    expect(await textsOf((await docxBookmarkRange(docx, 'Cell'))!.bytes)).toEqual(['[table]'])
  })

  it('has nothing for a bookmark that is not there or covers nothing', async () => {
    const docx = await buildDocx({
      bodyXml: para(run('a') + start(1, 'Empty') + end(1) + run('b')),
    })
    expect(await docxBookmarkRange(docx, 'Missing')).toBeNull()
    expect(await docxBookmarkRange(docx, 'Empty')).toBeNull()
  })

  it('reads an id as text, whatever characters the file put in it', async () => {
    const id = '(x+x+)+y'
    const docx = await buildDocx({
      bodyXml: para(
        `<w:bookmarkStart w:id="${id}" w:name="Odd"/>` +
          run('kept') +
          `<w:bookmarkEnd w:id="${id}"/>`,
      ),
    })
    const range = await docxBookmarkRange(docx, 'Odd')
    expect(await textsOf(range!.bytes)).toEqual(['kept'])
  })
})
