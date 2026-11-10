// @vitest-environment jsdom
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'

const NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const ID = '{01234567-89AB-CDEF-0123-456789ABCDEF}'
const CF = `<conditionalFormatting sqref="B2:B4"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${NS}"><x14:id>${ID}</x14:id></ext></extLst></cfRule></conditionalFormatting>`
const EXT = `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${NS}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="${ID}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:fillColor rgb="FF638EC6"/></x14:dataBar></x14:cfRule><xm:sqref xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">B2:B4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
const NAME =
  '<definedNames><definedName name="SalesData">Data!$A$1:$B$4</definedName></definedNames>'

async function fixture(tail = EXT, decoration = '', cf = CF) {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace('</worksheet>', `${decoration}${cf}${tail}</worksheet>`),
  )
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  zip.file('xl/workbook.xml', workbook.replace('</workbook>', `${NAME}</workbook>`))
  return zip.generateAsync({ type: 'nodebuffer' })
}
async function add(buffer: Buffer) {
  return planCellEditsToXlsx(
    await createBufferEntrySource(buffer),
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    [
      {
        sheetName: 'Data',
        name: 'Users',
        area: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
        columnNames: ['Item', 'Amount'],
        bandedRows: true,
      },
    ],
  )
}

describe('worksheet table parts beside nested extensions', () => {
  it.each([true, false])(
    'round-trips a table beside linked x14 data (root extLst: %s)',
    async (rootExtension) => {
      const bytes = await fixture(rootExtension ? EXT : '')
      const mutation = await assembleWithJsZip(bytes, await add(bytes))
      const zip = await JSZip.loadAsync(mutation.buffer)
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      const doc = new DOMParser().parseFromString(xml, 'application/xml')
      expect(doc.querySelector('parsererror')).toBeNull()
      const children = [...doc.documentElement.children]
      const tables = children.filter((e) => e.localName === 'tableParts')
      expect(tables).toHaveLength(1)
      expect(tables[0]!.children).toHaveLength(1)
      expect(doc.querySelector('cfRule tableParts')).toBeNull()
      expect(xml).toContain(CF)
      if (rootExtension) {
        expect(xml).toContain(EXT)
        expect(children.indexOf(tables[0]!)).toBeLessThan(
          children.findIndex((e) => e.localName === 'extLst'),
        )
      } else expect(children.at(-1)).toBe(tables[0])
      expect(await zip.file('xl/workbook.xml')!.async('string')).toContain(NAME)
      const original = await JSZip.loadAsync(bytes)
      expect(await zip.file('customXml/item1.xml')!.async('uint8array')).toEqual(
        await original.file('customXml/item1.xml')!.async('uint8array'),
      )
      expect(await zip.file('xl/tables/table1.xml')!.async('string')).toContain('ref="A1:B4"')
    },
  )
  it('ignores markup-looking comments, CDATA and quoted greater-than signs', async () => {
    const comment = '<!-- <extLst>not an element</extLst> --><?probe ext="<extLst/>"?>'
    const extension = `<ext uri="synthetic"><sample xmlns="urn:synthetic" label="a > b" note='>'><![CDATA[<extLst/></worksheet>]]></sample></ext>`
    const decoratedCf = CF.replace('</extLst>', `${extension}</extLst>`)
    const plan = await add(await fixture(EXT, comment, decoratedCf))
    const xml = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(xml).toContain(comment)
    expect(xml).toContain(decoratedCf)
    expect(xml).toContain(EXT)
    const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement
    expect([...root.children].map((e) => e.localName).slice(-2)).toEqual(['tableParts', 'extLst'])
  })
  it('removes a linked x14 rule when the snapshot drops it (issue #2510)', async () => {
    const plan = await planCellEditsToXlsx(
      await createBufferEntrySource(await fixture()),
      [],
      [],
      [],
      undefined,
      [],
      [],
      [{ sheetName: 'Data', rules: [] }],
    )
    const xml = plan.replaced.get('xl/worksheets/sheet1.xml')!
    expect(xml).not.toContain('<conditionalFormatting')
    expect(xml).not.toContain(ID)
    expect(xml).not.toContain('78C0D931')
    expect(xml).toContain('<sheetData>')
  })
  it('still refuses deleting a sheet referenced by a defined name', async () => {
    await expect(
      planCellEditsToXlsx(await createBufferEntrySource(await fixture()), [], [], [], {
        renames: [],
        additions: [{ name: 'Keep' }],
        removals: ['Data'],
        order: ['Keep'],
      }),
    ).rejects.toThrow(/defined name/i)
  })
})
