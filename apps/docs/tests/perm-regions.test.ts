import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setPermRegionsEnabled } from '../src/renderer/editor/perm-regions'

/** a read-only protected fill-in template: two open cells, two open top-level paragraphs */
const BODY = [
  '<w:p><w:r><w:t>Locked intro</w:t></w:r></w:p>',
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>',
  '<w:tr><w:tc><w:p><w:r><w:t>A. TITLE</w:t></w:r></w:p></w:tc></w:tr>',
  '<w:tr><w:tc><w:p><w:permStart w:id="11" w:edGrp="everyone"/><w:r><w:t>[title]</w:t></w:r><w:permEnd w:id="11"/></w:p></w:tc></w:tr>',
  '<w:tr><w:tc><w:p><w:r><w:t>B. SUMMARY</w:t></w:r></w:p></w:tc></w:tr>',
  '<w:tr><w:tc><w:p><w:permStart w:id="22" w:edGrp="everyone" w:colFirst="0" w:colLast="0"/><w:r><w:t>[first]</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p></w:tc></w:tr>',
  '<w:permEnd w:id="22"/>',
  '</w:tbl>',
  '<w:permStart w:id="33" w:edGrp="everyone"/>',
  '<w:p><w:r><w:t>open paragraph</w:t></w:r></w:p>',
  '<w:permEnd w:id="33"/>',
  '<w:p><w:r><w:t>Locked outro</w:t></w:r></w:p>',
].join('')

const SETTINGS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>'

async function open(enabled = true, bodyXml = BODY) {
  const source = await buildDocx({
    bodyXml,
    extraParts: [
      {
        path: 'word/settings.xml',
        xml: SETTINGS,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
      },
    ],
  })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  setPermRegionsEnabled(editor, enabled)
  return { editor, parsed, source }
}

/** document position of the first character of `text` */
function posOf(editor: Editor, text: string): number {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text))
      found = pos + node.text.indexOf(text)
    return found === -1
  })
  if (found === -1) throw new Error(`no text ${text}`)
  return found
}

function typeAt(editor: Editor, text: string, insert: string): boolean {
  const before = editor.state.doc
  editor.view.dispatch(editor.state.tr.insertText(insert, posOf(editor, text)))
  return !editor.state.doc.eq(before)
}

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file(path)!.async('string')
}

const count = (xml: string, re: RegExp) => [...xml.matchAll(re)].length

describe('editing permission regions in the editor', () => {
  it('highlights the text of the open regions, with no brackets of its own', async () => {
    const { editor } = await open()
    const dom = editor.view.dom
    expect([...dom.querySelectorAll('.perm-region')].map((el) => el.textContent)).toEqual([
      '[title]',
      '[first]',
      'second',
      'open paragraph',
    ])
    // Word draws no brackets; a template's own "[" / "]" are its text
    expect(dom.textContent).not.toContain('[[')
    editor.destroy()
  })

  it('highlights a table nested in an open cell and marks empty paragraphs', async () => {
    const nested = [
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>',
      '<w:tr><w:tc><w:p><w:permStart w:id="44" w:edGrp="everyone" w:colFirst="0" w:colLast="0"/><w:r><w:t>Year 1</w:t></w:r></w:p>',
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>',
      '<w:tr><w:tc><w:p><w:r><w:t>Task</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>',
      '</w:tbl><w:p/></w:tc></w:tr><w:permEnd w:id="44"/>',
      '<w:tr><w:tc><w:p><w:r><w:t>Locked</w:t></w:r></w:p></w:tc></w:tr>',
      '</w:tbl>',
    ].join('')
    const { editor } = await open(true, nested)
    const dom = editor.view.dom
    expect([...dom.querySelectorAll('.perm-region')].map((el) => el.textContent)).toEqual([
      'Year 1',
    ])
    // the nested table is an atom: its cells are highlighted through its class
    const nestedTables = dom.querySelectorAll('.perm-region-nested')
    expect(nestedTables).toHaveLength(1)
    expect(nestedTables[0]!.textContent).toContain('Task')
    // the empty paragraph after the nested table
    expect(dom.querySelectorAll('.perm-region-empty')).toHaveLength(1)
    expect(dom.textContent).toContain('Locked')
    expect(
      [...dom.querySelectorAll('.perm-region')].some((el) => el.textContent === 'Locked'),
    ).toBe(false)
    editor.destroy()
  })

  it('shows and guards nothing while disabled', async () => {
    const { editor } = await open(false)
    expect(
      editor.view.dom.querySelectorAll('.perm-region, .perm-region-empty, .perm-region-nested'),
    ).toHaveLength(0)
    expect(typeAt(editor, 'Locked intro', 'x')).toBe(true)
    editor.destroy()
  })

  it('lets typing into open cells and paragraphs and refuses it elsewhere', async () => {
    const { editor } = await open()
    expect(typeAt(editor, '[title]', 'x')).toBe(true)
    expect(typeAt(editor, 'second', 'x')).toBe(true)
    expect(typeAt(editor, 'open paragraph', 'x')).toBe(true)
    expect(typeAt(editor, 'A. TITLE', 'x')).toBe(false)
    expect(typeAt(editor, 'Locked intro', 'x')).toBe(false)
    expect(typeAt(editor, 'Locked outro', 'x')).toBe(false)
    editor.destroy()
  })

  it('refuses edits that reach out of a region and cell attribute changes', async () => {
    const { editor } = await open()
    const before = editor.state.doc
    // a deletion from the open cell into the locked header cell above it
    editor.view.dispatch(
      editor.state.tr.delete(posOf(editor, 'A. TITLE'), posOf(editor, '[title]') + 2),
    )
    expect(editor.state.doc.eq(before)).toBe(true)
    // the open cell's own attributes would regenerate the whole table
    const $cell = editor.state.doc.resolve(posOf(editor, '[title]'))
    const cellDepth = [...Array($cell.depth + 1).keys()].find(
      (d) => $cell.node(d).type.name === 'docTableCell',
    )!
    editor.view.dispatch(
      editor.state.tr.setNodeAttribute($cell.before(cellDepth), 'fill', 'FF0000'),
    )
    expect(editor.state.doc.eq(before)).toBe(true)
    editor.destroy()
  })

  it('a new top-level paragraph split off an open one stays open', async () => {
    const { editor } = await open()
    const at = posOf(editor, 'open paragraph') + 'open'.length
    editor.view.dispatch(editor.state.tr.split(at))
    expect(typeAt(editor, ' paragraph', 'y')).toBe(true)
    const opened: string[] = []
    editor.state.doc.forEach((node) => {
      if (node.attrs.permEditable) opened.push(node.textContent)
    })
    expect(opened).toEqual(['open', 'y paragraph'])
    editor.destroy()
  })

  it('saves edits in open regions with every marker and the protection intact', async () => {
    const { editor, parsed } = await open()
    typeAt(editor, '[title]', 'New ')
    // Enter in the column-range cell, then type into the new paragraph
    editor.view.dispatch(editor.state.tr.split(posOf(editor, 'second') + 'second'.length))
    editor.view.dispatch(
      editor.state.tr.insertText('third', posOf(editor, 'second') + 'second'.length + 2),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const out = await part(saved, 'word/document.xml')
    for (const id of ['11', '22', '33']) {
      expect(count(out, new RegExp(`<w:permStart w:id="${id}"`, 'g'))).toBe(1)
      expect(count(out, new RegExp(`<w:permEnd w:id="${id}"`, 'g'))).toBe(1)
    }
    expect(out).toMatch(/<\/w:tr><w:permEnd w:id="22"\/><\/w:tbl>/)
    expect(await part(saved, 'word/settings.xml')).toBe(SETTINGS)

    const again = await parseDocx(saved)
    const rows = again.blocks.find((b) => b.type === 'table')!.table!.rows
    expect(rows.map((row) => row[0]!.permEditable === true)).toEqual([false, true, false, true])
    expect(rows[1]![0]!.paras).toEqual(['New [title]'])
    expect(rows[3]![0]!.paras).toEqual(['[first]', 'second', 'third'])
    editor.destroy()
  })

  it('keeps an untouched protected template byte-identical', async () => {
    const { editor, parsed, source } = await open()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })
})
