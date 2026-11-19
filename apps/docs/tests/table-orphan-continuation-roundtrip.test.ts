import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  tableStructureSignature,
  type PmNode,
} from '../src/renderer/editor/convert'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

// The issue #2381 shape: an orphan vMerge continuation (no restart above it in
// its column) that carries its own tcPr and hidden text next to a normal cell,
// plus a rowspan-covered continuation with its own shading. Editors that drop
// the merged-away cells shift the row's later cells and the docx round trip
// regenerates the table, losing the hidden text (benchmark check text.exact).
const TABLE_XML =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="7200" w:type="dxa"/>' +
  '<w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  '<w:tr>' +
  '<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>H1</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>H2</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>H3</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '<w:tr>' +
  '<w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>' +
  '<w:p><w:r><w:t>Top</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>Other</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '<w:tr>' +
  '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>' +
  '<w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>' +
  // the orphan continuation: its own btLr tcPr and hidden text
  '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:textDirection w:val="btLr"/><w:vMerge/></w:tcPr>' +
  '<w:p><w:r><w:t>second</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>' +
  '</w:tr>' +
  '<w:tr>' +
  // the rowspan-covered continuation: its own shading
  '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/><w:vMerge/></w:tcPr>' +
  '<w:p><w:r><w:t>continuation</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc>' +
  '</w:tr></w:tbl>'

const BODY_XML = '<w:p><w:r><w:t>Intro</w:t></w:r></w:p>' + TABLE_XML

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

function bodyTable(editor: Editor): PmNode {
  return (editor.getJSON() as PmNode).content!.find((node) => node.type === 'docTable')!
}

/// The saved document's own text stream (all w:t contents, like the trial's text check).
async function savedText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml')!.async('string')
  return [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join('')
}

function cellXml(xml: string, text: string): string {
  const cell = [...xml.matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].find((m) =>
    m[0].includes('>' + text + '<'),
  )
  if (!cell) throw new Error('cell not found: ' + text)
  return cell[0]
}

describe('orphan vMerge continuation round trip (issue #2381)', () => {
  it('an untouched table keeps its block original and its bytes', async () => {
    const { editor, parsed, bytes } = await openDoc()
    // the editor must not "repair" the grid by appending a cell after load
    const table = bodyTable(editor)
    expect(tableStructureSignature(table)).toBe(table.attrs!.originalStructure)

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(plan.saveBlocks[1]).toEqual({ kind: 'original', docxIndex: 1 })
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(saved).toEqual(bytes)
    const text = await savedText(saved)
    // the merged-away cells' hidden text survives the untouched round trip
    expect(text).toContain('second')
    expect(text).toContain('continuation')
    editor.destroy()
  })

  it('rebuilding the table (a real edit) keeps the merged-away cells own bytes and text', async () => {
    const { editor, parsed } = await openDoc()
    let edited = false
    editor.state.doc.descendants((node, pos) => {
      if (
        edited ||
        (node.type.name !== 'docTableCell' && node.type.name !== 'docTableHeader') ||
        node.textContent !== 'H1'
      )
        return !edited
      editor.commands.insertContentAt(pos + 2, ' edited')
      edited = true
      return false
    })
    expect(edited).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const block = plan.saveBlocks[1]
    expect(block.kind).toBe('xml')
    const xml = (block as { xml: string }).xml
    // the edit reached the regenerated table (inserted at the text start)
    expect(xml).toContain('editedH1')
    // the orphan continuation keeps its own tcPr (btLr) and hidden text
    const orphan = cellXml(xml, 'second')
    expect(orphan).toContain('<w:vMerge/>')
    expect(orphan).toContain('<w:textDirection w:val="btLr"/>')
    // the rowspan-covered continuation keeps its own shading and text
    const covered = cellXml(xml, 'continuation')
    expect(covered).toContain('<w:vMerge/>')
    expect(covered).toContain('w:fill="EEEEEE"')
    editor.destroy()
  })

  it('a text edit keeps the continuation cells through the patch path', async () => {
    const { editor, parsed } = await openDoc()
    let edited = false
    editor.state.doc.descendants((node, pos) => {
      if (
        edited ||
        (node.type.name !== 'docTableCell' && node.type.name !== 'docTableHeader') ||
        node.textContent !== 'Other'
      )
        return !edited
      editor.commands.insertContentAt(pos + 2, 'X')
      edited = true
      return false
    })
    expect(edited).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const table = saved.blocks[1].table!
    expect(table.rows[2][0].paras).toEqual(['A'])
    expect(table.rows[2][1].vMerge).toBe('continue')
    expect(table.rows[2][1].paras).toEqual(['second'])
    expect(table.rows[3][0].vMerge).toBe('continue')
    expect(table.rows[3][0].paras).toEqual(['continuation'])
    // the edit landed (insertContentAt at the text start prepends)
    expect(table.rows[1][1].paras).toEqual(['XOther'])
    editor.destroy()
  })

  it('a formatting edit keeps the orphan column instead of shifting the later cell', async () => {
    const { editor, parsed } = await openDoc()
    // A fill edit changes tableStructureSignature, so pmTableToModel clears the
    // positional continuation records (their row:column keys are stale after a
    // real edit). The mergeGap placeholder must still emit a vMerge
    // continuation cell for its column, or the row loses it and "C" shifts left.
    let filled = false
    editor.state.doc.descendants((node, pos) => {
      if (
        filled ||
        (node.type.name !== 'docTableCell' && node.type.name !== 'docTableHeader') ||
        node.textContent !== 'C'
      )
        return !filled
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, fill: 'FF0000' }),
      )
      filled = true
      return false
    })
    expect(filled).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.saveBlocks[1].kind).toBe('xml')
    const saved = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const row = saved.blocks[1].table!.rows[2]!
    expect(row).toHaveLength(3)
    expect(row[0]!.vMerge).toBe('restart')
    expect(row[1]!.vMerge).toBe('continue')
    expect(row[2]!.paras).toEqual(['C'])
    editor.destroy()
  })
})
