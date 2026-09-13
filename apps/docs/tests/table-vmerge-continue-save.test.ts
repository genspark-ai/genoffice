import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

// row 2 continuation cell carries its own text ("dup"), as some producers
// duplicate the merged cell's content into every continuation cell
const VMERGE_TABLE_XML =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>dup</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>' +
  '</w:tbl>'

const BODY_XML = '<w:p><w:r><w:t>Intro</w:t></w:r></w:p>' + VMERGE_TABLE_XML

async function openDoc() {
  const bytes = await buildDocx({ bodyXml: BODY_XML })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed, bytes }
}

function cellTexts(xml: string): string[] {
  return [...xml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((m) =>
    [...m[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
  )
}

describe('vMerge continuation cells survive saving', () => {
  it('an untouched table with a text-bearing continuation cell keeps its bytes', async () => {
    const { editor, parsed, bytes } = await openDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(plan.saveBlocks[1]).toEqual({ kind: 'original', docxIndex: 1 })
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(bytes)
    editor.destroy()
  })

  it('editing a sibling paragraph leaves the continuation cell text in place', async () => {
    const { editor, parsed } = await openDoc()
    editor.commands.insertContentAt(1, 'X')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    expect(plan.saveBlocks[1]).toEqual({ kind: 'original', docxIndex: 1 })
    const saved = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const table = saved.blocks[1].table!
    expect(table.rows[1][0].vMerge).toBe('continue')
    expect(table.rows[1][0].paras).toEqual(['dup'])
    editor.destroy()
  })

  it('editing another cell patches only that cell and keeps the continuation text', async () => {
    const { editor, parsed } = await openDoc()
    let done = false
    editor.state.doc.descendants((node, pos) => {
      if (done || node.type.name !== 'docTableCell' || node.textContent !== 'B2') return !done
      editor.commands.insertContentAt(pos + 2, 'Z')
      done = true
      return false
    })
    expect(done).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const block = plan.saveBlocks[1]
    expect(block.kind).toBe('xml')
    const xml = (block as { xml: string }).xml
    expect(cellTexts(xml)).toEqual(['Name', 'B1', 'dup', 'ZB2'])
    expect(xml).toContain('<w:vMerge/>')
    editor.destroy()
  })
})
