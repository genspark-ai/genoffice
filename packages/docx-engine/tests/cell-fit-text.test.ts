import { describe, expect, it } from 'vitest'
import { generateTableModelXml, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const cell = (tcPr: string, text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('<w:tcFitText/>', 'fitted')}${cell('<w:tcFitText w:val="0"/>', 'plain')}</w:tr></w:tbl>`

describe('w:tcFitText', () => {
  it('is read from the cell properties', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
    const [fitted, plain] = parsed.blocks[0]!.table!.rows[0]!
    expect(fitted!.fitText).toBe(true)
    expect(plain!.fitText).toBeUndefined()
  })

  it('an edited cell writes it from the model', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
    const block = parsed.blocks[0]!
    const model = structuredClone(block.table!)
    const [fitted, plain] = model.rows[0]!
    Object.assign(fitted!, { tcPrEdited: true, fitText: false })
    Object.assign(plain!, { tcPrEdited: true, fitText: true })
    const cells = generateTableModelXml(model, block.originalXml ?? undefined).match(
      /<w:tc>[\s\S]*?<\/w:tc>/g,
    )!
    expect(cells[0]).not.toContain('w:tcFitText')
    expect(cells[1]).toContain('<w:tcFitText/>')
    expect(cells[1]).not.toContain('w:val="0"')
  })

  it('an untouched cell keeps its bytes', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
    const block = parsed.blocks[0]!
    const xml = generateTableModelXml(block.table!, block.originalXml ?? undefined)
    expect(xml).toContain('<w:tcFitText/>')
    expect(xml).toContain('<w:tcFitText w:val="0"/>')
  })
})
