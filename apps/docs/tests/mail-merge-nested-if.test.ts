import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { documentText, substituteMergeFields } from '../src/renderer/mail-merge/template'

/**
 * The IF a Word template carries (`IF { MERGEFIELD Remote } = "Yes" …`), from
 * the file to the merged letter: each record gets the branch its own value
 * selects, not the result the template was last saved with.
 */

const fld = (type: string): string => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
const instr = (text: string): string =>
  `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`

const TEMPLATE =
  '<w:p>' +
  fld('begin') +
  instr(' IF ') +
  fld('begin') +
  instr(' MERGEFIELD Remote ') +
  fld('separate') +
  '<w:r><w:t>«Remote»</w:t></w:r>' +
  fld('end') +
  instr(' = "Yes" "You will work remotely." "You will work on site." ') +
  fld('separate') +
  '<w:r><w:t>You will work on site.</w:t></w:r>' +
  fld('end') +
  '</w:p>'

async function merged(record: Record<string, string>): Promise<string> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TEMPLATE }))
  const template = blocksToPmDoc(parsed.blocks, readSections(parsed)) as PmNode
  const letter = substituteMergeFields(template, record)
  const bytes = await saveDocx(parsed, pmDocToSavePlan(letter, parsed.blocks).saveBlocks)
  const out = await parseDocx(bytes)
  expect(out.blocks[0]!.runs!.some((run) => run.instrField !== undefined)).toBe(false)
  expect(documentText(letter)).toBe(out.blocks[0]!.runs!.map((run) => run.text).join(''))
  return documentText(letter)
}

describe('a Word IF field in a merged letter', () => {
  it('takes the branch the record selects', async () => {
    expect(await merged({ Remote: 'Yes' })).toBe('You will work remotely.')
    expect(await merged({ Remote: 'No' })).toBe('You will work on site.')
  })
})
