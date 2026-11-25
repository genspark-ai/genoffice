import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { generateTableModelXml } from '../src/generate'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const custom = (kind: string, inner: string) =>
  `<w:customXml w:uri="urn:example" w:element="${kind}">` +
  '<w:customXmlPr><w:attr w:name="source" w:val="metadata-only"/></w:customXmlPr>' +
  inner +
  '</w:customXml>'

/** one-cell table; optionally wrap the row, the cell and/or the cell paragraph */
const tableWith = (level: 'row' | 'cell' | 'paragraph' | 'none') => {
  const p =
    level === 'paragraph'
      ? custom('paragraph', '<w:p><w:r><w:t>Sample</w:t></w:r></w:p>')
      : '<w:p><w:r><w:t>Sample</w:t></w:r></w:p>'
  const tc =
    level === 'cell'
      ? custom('cell', '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>' + p + '</w:tc>')
      : '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>' + p + '</w:tc>'
  const tr = level === 'row' ? custom('row', '<w:tr>' + tc + '</w:tr>') : '<w:tr>' + tc + '</w:tr>'
  return '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' + tr + '</w:tbl>'
}

const parasOf = (docx: Awaited<ReturnType<typeof parseDocx>>) =>
  docx.blocks.find((block) => block.table)?.table?.rows[0]?.[0]?.paras

describe('table-level custom XML visible content', () => {
  it.each(['row', 'cell', 'paragraph'] as const)(
    'reads the cell text through a wrapped %s',
    async (level) => {
      const parsed = await parseDocx(await buildDocx({ bodyXml: tableWith(level) }))
      expect(parasOf(parsed)).toEqual(['Sample'])
    },
  )

  it('reads text through nested custom XML and sdt wrappers', async () => {
    const p = custom(
      'outer',
      custom(
        'inner',
        '<w:sdt><w:sdtContent><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:sdtContent></w:sdt>',
      ),
    )
    const bodyXml =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      custom(
        'row',
        '<w:tr>' +
          custom(
            'cell',
            '<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>' + p + '</w:tc>',
          ) +
          '</w:tr>',
      ) +
      '</w:tbl>'
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    expect(parasOf(parsed)).toEqual(['Nested'])
  })

  it('saves an untouched wrapped table byte-for-byte', async () => {
    const bytes = await buildDocx({ bodyXml: tableWith('row') })
    const parsed = await parseDocx(bytes)
    expect(parsed.blocks.find((block) => block.table)).toMatchObject({ type: 'table' })
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: parsed.blocks[0].docxIndex! },
    ])
    const originalXml = await (
      await JSZip.loadAsync(bytes)
    )
      .file('word/document.xml')!
      .async('string')
    const savedXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(savedXml).toContain(tableWith('row'))
    expect(savedXml).toContain('metadata-only')
    expect(originalXml).toContain(tableWith('row'))
  })

  it('regenerates a wrapped cell after a text edit', async () => {
    const bytes = await buildDocx({ bodyXml: tableWith('cell') })
    const parsed = await parseDocx(bytes)
    const block = parsed.blocks.find((b) => b.table)!
    const xml = generateTableModelXml({
      ...block.table!,
      rows: [[{ paras: ['Changed'] }]],
    })
    const reopened = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(parasOf(reopened)).toEqual(['Changed'])
  })
})
