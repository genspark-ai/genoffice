/** Table of contents entries (TC fields) ride with their text, as index entries and citations do. */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import type { Run } from '../src/types'
import { buildDocx } from './helpers/build-docx'

const field = (instr: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const ENTRY = 'TC "Appendix A" \\f F \\l 2'

describe('table of contents entry fields', () => {
  it('reads a TC field as a marker after its text and keeps the paragraph editable', async () => {
    const body = `<w:p><w:r><w:t>Appendix</w:t></w:r>${field(ENTRY)}<w:r><w:t xml:space="preserve"> text.</w:t></w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('paragraph')
    const runs = (block as { runs: Run[] }).runs
    const marker = runs.find((r) => r.tcInstr !== undefined)!
    expect(marker.tcInstr).toBe('"Appendix A" \\f F \\l 2')
    expect(marker.xeTerm).toBe('')
    expect(marker.taInstr).toBeUndefined()
    expect(runs.map((r) => r.text).join('')).toBe('Appendix text.')
  })

  it('reads the w:fldSimple form too', async () => {
    const body = `<w:p><w:r><w:t>Appendix</w:t></w:r><w:fldSimple w:instr=" TC &quot;Appendix A&quot; "/></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const runs = (parsed.blocks[0] as { runs: Run[] }).runs
    expect(runs.find((r) => r.tcInstr !== undefined)?.tcInstr).toBe('"Appendix A"')
  })

  it('writes the field back with its instruction when the paragraph is edited', async () => {
    const body = `<w:p><w:r><w:t>Appendix</w:t></w:r>${field(ENTRY)}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const runs = (parsed.blocks[0] as { runs: Run[] }).runs
    const saved = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'Edited. ' }, ...runs] },
        docxIndex: 0,
      },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain(` ${ENTRY} `)
    expect(xml).not.toContain(' XE ')
    expect(xml).not.toContain(' TA ')
    const reparsed = (await parseDocx(saved)).blocks[0] as { runs: Run[] }
    expect(reparsed.runs.filter((r) => r.tcInstr !== undefined)).toHaveLength(1)
  })
})
