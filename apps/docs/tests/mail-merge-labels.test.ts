import { describe, expect, it } from 'vitest'
import {
  buildBlankDocx,
  mergeFieldInstruction,
  parseDocx,
  readSections,
  saveDocx,
} from '@genoffice/docx-engine'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { substituteMergeFields } from '../src/renderer/mail-merge/template'
import {
  AVERY_5160,
  cellContentOf,
  labelSheetDoc,
  layoutSection,
} from '../src/renderer/mail-merge/layout'

/**
 * The label sheet, end to end: the records laid into the grid, the grid saved
 * through the same plan the editor saves with, and the .docx that comes out —
 * one table, a cell per record, the stock's paper. genoffice#2097.
 */

const template = (): PmNode => ({
  type: 'doc',
  content: [
    {
      type: 'docParagraph',
      content: [
        {
          type: 'text',
          text: '«FirstName»',
          marks: [{ type: 'instrField', attrs: { instr: mergeFieldInstruction('FirstName') } }],
        },
      ],
    },
  ],
})

describe('the label sheet document', () => {
  it('saves as one grid table with a record per cell on the stock paper', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const records = [{ FirstName: 'Ada' }, { FirstName: 'Ben' }]
    const cells = records.map((record) => cellContentOf(substituteMergeFields(template(), record)))
    const label = labelSheetDoc(cells, AVERY_5160)
    const plan = pmDocToSavePlan(label, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks, {
      section: layoutSection(AVERY_5160, null),
    })
    const out = await parseDocx(bytes)
    const table = out.blocks.find((block) => block.type === 'table')
    expect(table).toBeTruthy()
    const rows = table!.table!.rows
    expect(rows).toHaveLength(1)
    // label, gutter, label, gutter, label: the 5160 stock's 3780/180/… grid
    expect(rows[0]).toHaveLength(5)
    expect(rows[0]![0]!.paras.join('')).toBe('Ada')
    expect(rows[0]![1]!.paras.join('')).toBe('')
    expect(rows[0]![2]!.paras.join('')).toBe('Ben')
    expect(rows[0]![3]!.paras.join('')).toBe('')
    expect(rows[0]![4]!.paras.join('')).toBe('')
    const sections = readSections(out)
    expect(sections.at(-1)?.settings.pageWidth).toBe(AVERY_5160.pageWidth)
    expect(sections.at(-1)?.settings.pageHeight).toBe(AVERY_5160.pageHeight)
  })

  it('saves the stock grid, borderless cells and a trailing paragraph', async () => {
    const parsed = await parseDocx(await buildBlankDocx())
    const cells = cellContentOf(substituteMergeFields(template(), { FirstName: 'Ada' }))
    const label = labelSheetDoc([cells], AVERY_5160)
    const plan = pmDocToSavePlan(label, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks, {
      section: layoutSection(AVERY_5160, null),
    })
    const xml = (await parseDocx(bytes)).internal.documentXml
    // the table spans the 11700-twip text column with the stock's 3780/180 grid
    expect(xml).toContain('<w:tblW w:w="11700" w:type="dxa"/>')
    expect(xml).toMatch(
      /<w:gridCol w:w="3780"\/><w:gridCol w:w="180"\/><w:gridCol w:w="3780"\/><w:gridCol w:w="180"\/><w:gridCol w:w="3780"\/>/,
    )
    // no printed frame: every cell declares all four sides none
    expect(xml).toMatch(
      /<w:tcBorders><w:top w:val="none"\/><w:left w:val="none"\/><w:bottom w:val="none"\/><w:right w:val="none"\/><\/w:tcBorders>/,
    )
    // the label rows carry the exact 1in height
    expect(xml).toMatch(/<w:trHeight w:val="1440" w:hRule="exact"\/>/)
    // Word wants a paragraph between the table and the section properties
    expect(xml).toMatch(/<\/w:tbl><w:p[\s/>][\s\S]*<w:sectPr/)
  })
})
