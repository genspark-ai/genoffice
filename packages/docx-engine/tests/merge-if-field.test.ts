import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { ifBranch, parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * A mail-merge IF as Word writes it: the operand is a nested MERGEFIELD, not
 * text. The editor carries the whole field as one instruction with the nested
 * field as «Name», so the merge can evaluate it, and writes the nesting back.
 */

const fld = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
const instr = (text: string): string =>
  `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`
const nested = (code: string, shown: string): string =>
  fld('begin') +
  instr(` ${code} `) +
  fld('separate') +
  `<w:r><w:t>${shown}</w:t></w:r>` +
  fld('end')

/** `IF { MERGEFIELD Remote } = "Yes" "At home." "On site."` with a stale result */
const ifParagraph = (operand = nested('MERGEFIELD Remote', '«Remote»'), result = ''): string =>
  '<w:p><w:r><w:t xml:space="preserve">Place: </w:t></w:r>' +
  fld('begin') +
  instr(' IF ') +
  operand +
  instr(' = "Yes" "At home." "On site." ') +
  fld('separate') +
  (result || '<w:r><w:t>On site.</w:t></w:r>') +
  fld('end') +
  '</w:p>'

const FOLDED = 'IF «Remote» = "Yes" "At home." "On site."'

const documentXmlOf = async (bytes: Uint8Array): Promise<string> =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

describe('a mail-merge IF with a nested MERGEFIELD', () => {
  it('is one field run the merge can evaluate', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: ifParagraph() }))
    expect(doc.blocks[0]!.type).toBe('paragraph')
    expect(doc.blocks[0]!.runs).toEqual([
      { text: 'Place: ' },
      { text: 'On site.', instrField: FOLDED },
    ])
    expect(ifBranch(FOLDED, { Remote: 'Yes' })).toBe('At home.')
    expect(ifBranch(FOLDED, { Remote: 'No' })).toBe('On site.')
  })

  it("reads a quoted name and Word's MERGEFORMAT switch", async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: ifParagraph(nested('MERGEFIELD "Office City" \\* MERGEFORMAT', '«Office City»')),
      }),
    )
    expect(doc.blocks[0]!.runs![1]!.instrField).toBe(
      'IF «Office City» = "Yes" "At home." "On site."',
    )
  })

  it('a SKIPIF written by the editor nests its data column too, or Word would compare the text', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Dear</w:t></w:r></w:p>' }),
    )
    const runs = [{ text: 'Dear' }, { text: ' ', instrField: 'SKIPIF «City» = "London"' }]
    const xml = await documentXmlOf(
      await saveDocx(parsed, [{ kind: 'generated', block: { type: 'paragraph', runs } }]),
    )
    expect(xml).toMatch(/<w:instrText[^>]*> SKIPIF <\/w:instrText>/)
    expect(xml).toMatch(/<w:instrText[^>]*> MERGEFIELD City <\/w:instrText>/)
    expect(xml).toMatch(/<w:instrText[^>]*> = "London" <\/w:instrText>/)
    expect(xml).not.toMatch(/<w:instrText[^>]*>[^<]*«/)
  })

  it('writes the nested field back when the paragraph is regenerated', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: ifParagraph() }))
    const block = parsed.blocks[0]!
    const runs = [{ text: 'Workplace: ' }, ...block.runs!.slice(1)]
    const bytes = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs } },
    ])
    const xml = await documentXmlOf(bytes)
    expect(xml).toContain('Workplace: ')
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(2)
    expect(xml).toMatch(/<w:instrText[^>]*> IF <\/w:instrText>/)
    expect(xml).toMatch(/<w:instrText[^>]*> MERGEFIELD Remote <\/w:instrText>/)
    const again = await parseDocx(bytes)
    expect(again.blocks[0]!.runs).toEqual([
      { text: 'Workplace: ' },
      { text: 'On site.', instrField: FOLDED },
    ])
  })

  it.each([
    [
      'a nested field with a formatting switch',
      ifParagraph(nested('MERGEFIELD Remote \\* Upper', '«Remote»')),
    ],
    ['a nested field that is no merge field', ifParagraph(nested('PAGE', '1'))],
    ['a field in the result', ifParagraph(undefined, nested('MERGEFIELD Name', '«Name»'))],
    ['an IF with no nested field', ifParagraph(instr('1'))],
  ])('stays byte-preserved with %s', async (_label, bodyXml) => {
    const doc = await parseDocx(await buildDocx({ bodyXml }))
    expect(doc.blocks[0]!.type).toBe('passthrough')
  })
})
