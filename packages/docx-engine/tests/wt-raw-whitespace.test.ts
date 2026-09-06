/**
 * Literal whitespace inside w:t (Word probe 2026-09-03): every raw newline
 * renders as one space, never a line break; without xml:space="preserve" the
 * element's leading/trailing whitespace is dropped and a literal tab becomes a
 * space, while a preserved tab stays a tab.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

function para(inner: string, preserve = true): string {
  const attr = preserve ? ' xml:space="preserve"' : ''
  return `<w:p><w:r><w:t${attr}>${inner}</w:t></w:r></w:p>`
}

const BODY =
  para('abc\ndef') +
  para('abc\n\ndef') +
  para('abc\r\ndef') +
  para('思います。\n\nただ、') +
  para('abc\tdef') +
  para('\n\tabc\tdef\n', false) +
  para('abc\ndef', false)

describe('raw whitespace inside w:t', () => {
  it('maps newlines to spaces and keeps preserved tabs', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const texts = doc.blocks.filter((b) => !b.hidden).map((b) => b.runs![0].text)
    expect(texts).toEqual([
      'abc def',
      'abc  def',
      'abc def',
      '思います。  ただ、',
      'abc\tdef',
      'abc def',
      'abc def',
    ])
  })

  it('never yields a line break from a w:t newline', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY }))
    for (const b of doc.blocks) for (const r of b.runs ?? []) expect(r.text).not.toMatch(/[\r\n]/)
  })

  it('saving untouched paragraphs keeps the raw newline bytes', async () => {
    const bytes = await buildDocx({ bodyXml: BODY })
    const doc = await parseDocx(bytes)
    const blocks: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toBe(bytes)
  })
})
