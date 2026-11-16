import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function openBlank() {
  const source = await buildBlankDocx()
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

/** caret into the first paragraph of the first table cell */
function selectFirstCellParagraph(editor: Editor) {
  let target = -1
  editor.state.doc.descendants((node, pos) => {
    if (target >= 0) return false
    if (node.type.name === 'docTableCell') {
      target = pos + 2
      return false
    }
    return true
  })
  expect(target).toBeGreaterThan(0)
  editor.commands.setTextSelection(target)
}

async function saveAndUnzip(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml')!.async('string')
  return { plan, xml }
}

describe('table cell paragraph style and bookmarks survive the save', () => {
  // issue #2803: a style-only edit inside a table reported success but was lost
  // on save (the formatting signature had no styleId, and the regenerated cell
  // XML never wrote w:pStyle)
  it('applyStyle on a cell paragraph reaches the saved DOCX', async () => {
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 2, 2)
    selectFirstCellParagraph(editor)
    const ok = editor.commands.updateAttributes('docParagraph', { styleId: 'Heading1' })
    expect(ok).toBe(true)

    const { plan, xml } = await saveAndUnzip(editor, parsed)
    expect(plan.changedCount).toBeGreaterThan(0)
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>')
  })

  // issue #2804: bookmark targets vanished when the cell rebuilt for an
  // unrelated formatting edit; the rebuilt XML now re-emits the anchors
  it('bookmark names on a cell paragraph come back as w:bookmarkStart/End', async () => {
    const { editor, parsed } = await openBlank()
    insertTableAt(editor, 2, 2)
    selectFirstCellParagraph(editor)
    const ok = editor.commands.updateAttributes('docParagraph', {
      bookmarks: ['KeepAnchor'],
    })
    expect(ok).toBe(true)

    const { xml } = await saveAndUnzip(editor, parsed)
    expect(xml).toContain('w:name="KeepAnchor"')
    expect(xml).toContain('<w:bookmarkStart ')
    expect(xml).toContain('<w:bookmarkEnd ')
  })
})
