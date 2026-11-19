import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx, REVISION_TABLE_XML } from './helpers/build-docx'

describe('table row / cell revisions (trPr w:ins/w:del, tcPr w:cellIns/w:cellDel)', () => {
  it('parses row and cell revisions into the table model', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: REVISION_TABLE_XML }))
    const model = doc.blocks[0].table!
    expect(doc.blocks[0].type).toBe('table')
    expect(model.rowRevisions).toEqual([
      null,
      { kind: 'del', author: 'Alice', date: '2026-01-01T00:00:00Z', id: '11' },
      null,
    ])
    expect(model.rows[2][0].cellRevision).toEqual({
      kind: 'ins',
      author: 'Alice',
      date: '2026-01-01T00:00:00Z',
      id: '15',
    })
    // deleted-row text arrives as del-marked runs, still part of the editable model
    expect(model.rows[1][0].richParas?.[0]?.runs?.[0]?.del?.author).toBe('Alice')
  })

  it('keeps an untouched revision table byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: REVISION_TABLE_XML })
    const doc = await parseDocx(bytes)
    expect(await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }])).toEqual(bytes)
  })

  it('parses row insertion (trPr w:ins)', async () => {
    const xml = REVISION_TABLE_XML.replace('<w:del w:id="11"', '<w:ins w:id="11"')
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].table!.rowRevisions?.[1]?.kind).toBe('ins')
  })
})
