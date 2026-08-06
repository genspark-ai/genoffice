import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const TXBX_CONTENT =
  '<w:txbxContent>' +
  '<w:p><w:pPr><w:spacing w:line="312" w:lineRule="auto" w:before="211"/>' +
  '<w:ind w:left="180" w:right="90"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Courier New"/><w:color w:val="415461"/></w:rPr>' +
  '<w:t>**Current date:** {Month}</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>CRITICAL</w:t></w:r></w:p>' +
  '</w:txbxContent>'

const TEXTBOX_PARAGRAPH =
  '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor behindDoc="1" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="6038850" cy="1704975"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6038850" cy="1704975"/></a:xfrm>' +
  '<a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill></wps:spPr>' +
  '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>' +
  `<wps:txbx>${TXBX_CONTENT}</wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
  '<mc:Fallback><w:pict>' +
  '<v:rect xmlns:v="urn:schemas-microsoft-com:vml" id="box1" style="position:absolute">' +
  `<v:textbox>${TXBX_CONTENT}</v:textbox></v:rect>` +
  '</w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>'

async function openTextboxDoc() {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml: `<w:p><w:r><w:t>Leading text</w:t></w:r></w:p>${TEXTBOX_PARAGRAPH}`,
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

/** locate the protected textbox node in the outer document */
function textboxNodePos(editor: Editor): number {
  let pos = -1
  editor.state.doc.descendants((node, p) => {
    if (pos !== -1) return false
    if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = p
    return pos === -1
  })
  expect(pos).toBeGreaterThanOrEqual(0)
  return pos
}

/** simulate a rich sub-editor commit: rewrite one paragraph's runs in node attrs */
function editTextboxPara(
  editor: Editor,
  paraIndex: number,
  runs: TextboxDisplay['paras'][number]['runs'],
) {
  const pos = textboxNodePos(editor)
  const node = editor.state.doc.nodeAt(pos)!
  const boxes = node.attrs.textboxes as TextboxDisplay[]
  const next = boxes.map((box) => ({
    ...box,
    paras: box.paras.map((para, i) => (i === paraIndex ? { ...para, runs } : para)),
  }))
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, textboxes: next }),
  )
}

describe('text box rich-text edit and save', () => {
  it('mounts a rich sub-editor inside each rendered box', async () => {
    const { editor } = await openTextboxDoc()
    const subDom = editor.view.dom.querySelector('.doc-textbox .doc-textbox-editor')
    expect(subDom).not.toBeNull()
    expect(subDom?.getAttribute('contenteditable')).toBe('false')
    subDom
      ?.closest('.doc-textbox')
      ?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    expect(subDom?.getAttribute('contenteditable')).toBe('true')
    expect(subDom?.textContent).toContain('**Current date:** {Month}')
    expect(subDom?.textContent).toContain('CRITICAL')
    const box = subDom?.closest('.doc-textbox') as HTMLElement
    const firstPara = box.querySelector('.doc-textbox-para') as HTMLElement
    expect(box.getAttribute('style')).toContain('padding:0px 0px 0px 0px')
    expect(firstPara.style.lineHeight).toBe('1.56')
    expect(firstPara.style.marginTop).toBe('10.55pt')
    expect(firstPara.style.marginLeft).toBe('9pt')
    expect(firstPara.style.marginRight).toBe('4.5pt')
    editor.destroy()
  })

  it('untouched textbox saves as original (byte-stable)', async () => {
    const { editor, parsed } = await openTextboxDoc()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(plan.saveBlocks.every((b) => b.kind === 'original')).toBe(true)
    editor.destroy()
  })

  it('persists an auto-grown textbox height without changing its width', async () => {
    const { editor, parsed } = await openTextboxDoc()
    const pos = textboxNodePos(editor)
    const node = editor.state.doc.nodeAt(pos)!
    const boxes = node.attrs.textboxes as TextboxDisplay[]
    const originalHeight = boxes[0].heightPx!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        textboxes: boxes.map((box, index) =>
          index === 0 ? { ...box, heightPx: originalHeight + 80 } : box,
        ),
      }),
    )

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const resized = reparsed.blocks.find((block) => block.textboxes)?.textboxes?.[0]
    expect(resized?.heightPx).toBe(originalHeight + 80)
    expect(resized?.widthPx).toBe(boxes[0].widthPx)
    editor.destroy()
  })

  it('persists a resize of an auto-fit textbox as a fixed height', async () => {
    const autoFitXml = TEXTBOX_PARAGRAPH.replace(
      '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>',
      '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></wps:bodyPr>',
    )
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: `<w:p><w:r><w:t>Leading text</w:t></w:r></w:p>${autoFitXml}` }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const pos = textboxNodePos(editor)
    const node = editor.state.doc.nodeAt(pos)!
    const boxes = node.attrs.textboxes as TextboxDisplay[]
    expect(boxes[0].heightPx).toBeUndefined() // autofit carries no fixed height
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        textboxes: [{ ...boxes[0], heightPx: 240, minHeightPx: 240 }],
      }),
    )

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const resized = reparsed.blocks.find((block) => block.textboxes)?.textboxes?.[0]
    expect(resized?.heightPx).toBe(240)
    expect(resized?.widthPx).toBe(boxes[0].widthPx)
    editor.destroy()
  })

  it('edited textbox paragraph saves as a patched XML block and re-parses with the new text', async () => {
    const { editor, parsed } = await openTextboxDoc()
    editTextboxPara(editor, 0, [
      { text: '**Current date:** July 21, 2026', font: 'Courier New', color: '415461' },
    ])

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    expect(plan.saveBlocks.find((b) => b.kind === 'xml')).toBeDefined()

    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const box = reparsed.blocks.find((b) => b.textboxes)?.textboxes?.[0]
    expect(box?.paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual([
      '**Current date:** July 21, 2026',
      'CRITICAL',
    ])
    // styling survives the patch
    expect(box?.fill).toBe('F5F5F5')
    expect(box?.paras[0].runs[0]).toMatchObject({ font: 'Courier New', color: '415461' })
    expect(box?.paras[1].runs[0]).toMatchObject({ bold: true })
    editor.destroy()
  })

  it('rich formatting changes (color, bold, size) round-trip through the save patch', async () => {
    const { editor, parsed } = await openTextboxDoc()
    editTextboxPara(editor, 1, [
      { text: 'CRIT', bold: true, color: 'FF0000', sizeHalfPoints: 28 },
      { text: 'ICAL', italic: true },
    ])

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)

    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const box = reparsed.blocks.find((b) => b.textboxes)?.textboxes?.[0]
    // untouched first paragraph keeps its original styling
    expect(box?.paras[0].runs[0]).toMatchObject({
      text: '**Current date:** {Month}',
      font: 'Courier New',
      color: '415461',
    })
    expect(box?.paras[1].runs).toEqual([
      expect.objectContaining({ text: 'CRIT', bold: true, color: 'FF0000', sizeHalfPoints: 28 }),
      expect.objectContaining({ text: 'ICAL', italic: true }),
    ])
    // centered alignment survives (align is re-applied from the display model)
    expect(box?.paras[1].align).toBe('center')
    editor.destroy()
  })

  it('sub-editor edits commit into the model on the pre-save broadcast', async () => {
    const { editor, parsed } = await openTextboxDoc()
    const subDom = editor.view.dom.querySelector('.doc-textbox .doc-textbox-editor')
    // tiptap exposes the Editor instance on its root DOM element
    const sub = (subDom as (HTMLElement & { editor?: Editor }) | null)?.editor
    expect(sub).toBeDefined()

    // apply a rich edit exactly like the ribbon would (color the whole box)
    sub!.chain().selectAll().setMark('docTextStyle', { color: 'FF0000' }).run()
    window.dispatchEvent(new Event('ai-docs-commit-tables'))

    const pos = textboxNodePos(editor)
    const boxes = editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
    expect(boxes[0].paras[0].runs[0]).toMatchObject({ color: 'FF0000' })
    expect(boxes[0].paras[1].runs[0]).toMatchObject({ color: 'FF0000', bold: true })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const box = reparsed.blocks.find((b) => b.textboxes)?.textboxes?.[0]
    expect(box?.paras[0].runs[0]).toMatchObject({ color: 'FF0000', font: 'Courier New' })
    expect(box?.paras[1].runs[0]).toMatchObject({ color: 'FF0000', bold: true })
    editor.destroy()
  })
})
