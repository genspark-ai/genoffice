import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type Run } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { MENU_COMMAND_HANDLERS, type CommandHost } from '../src/renderer/commands/menu-commands'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  applyPictureGeometry,
  selectedPictureAlt,
  selectedPictureGeometry,
  setSelectedPictureAlt,
} from '../src/renderer/editor/picture-props'

const picture = (docPr: string) =>
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  `<wp:extent cx="952500" cy="476250"/>${docPr}` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing>'

async function open(bodyXml: string) {
  const source = await buildDocx({ bodyXml, withImage: true })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (pos === -1 && (node.type.name === 'docInlineImage' || node.attrs.blockType === 'image'))
      pos = at
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return { editor, parsed, source, pos }
}

const amongText = (docPr: string) =>
  `<w:p><w:r><w:t>Before </w:t></w:r><w:r>${picture(docPr)}</w:r><w:r><w:t> after.</w:t></w:r></w:p>`

async function saved(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const reparsed = await parseDocx(bytes)
  const runs = reparsed.blocks.flatMap((block) => ('runs' in block ? (block.runs ?? []) : []))
  return { bytes, plan, image: runs.find((run: Run) => run.image)?.image }
}

describe('alternative text of a picture in a line of text', () => {
  it('reads what the file holds and saves what is typed, the rest of the file untouched', async () => {
    const { editor, parsed, source } = await open(
      amongText('<wp:docPr id="1" name="Picture 1" descr="A &amp; B"/>'),
    )
    expect(selectedPictureAlt(editor)).toBe('A & B')
    // nothing typed: nothing to save
    expect(setSelectedPictureAlt(editor, ' A & B ')).toBe(false)
    expect(
      await saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks),
    ).toEqual(source)

    expect(setSelectedPictureAlt(editor, 'Quarterly "sales" chart')).toBe(true)
    expect(selectedPictureAlt(editor)).toBe('Quarterly "sales" chart')
    const { image, plan } = await saved(editor, parsed)
    expect(plan.changedCount).toBe(1)
    expect(image?.xml).toContain('descr="Quarterly &quot;sales&quot; chart"')
    expect(image?.xml.match(/descr=/g)).toHaveLength(1)
    expect([image?.widthPx, image?.heightPx]).toEqual([100, 50])
    editor.destroy()
  })

  it('removes the text when it is cleared', async () => {
    const { editor, parsed } = await open(
      amongText('<wp:docPr id="1" name="Picture 1" descr="Old"/>'),
    )
    expect(setSelectedPictureAlt(editor, '')).toBe(true)
    expect(selectedPictureAlt(editor)).toBe('')
    const { image } = await saved(editor, parsed)
    expect(image?.xml).not.toContain('descr=')
    expect(image?.xml).toContain('<wp:docPr id="1" name="Picture 1"/>')
    editor.destroy()
  })

  it('undoes in one step', async () => {
    const { editor } = await open(amongText('<wp:docPr id="1" name="Picture 1" descr="Old"/>'))
    setSelectedPictureAlt(editor, 'New')
    editor.commands.undo()
    expect(selectedPictureAlt(editor)).toBe('Old')
    editor.destroy()
  })
})

describe('Size and Position on a picture alone in its paragraph', () => {
  it('says where a preset can go and moves the picture to one', async () => {
    const { editor, pos } = await open(
      `<w:p><w:r>${picture('<wp:docPr id="1" name="Picture 1"/>')}</w:r></w:p><w:p><w:r><w:t>Text</w:t></w:r></w:p>`,
    )
    const geometry = selectedPictureGeometry(editor)!
    expect(geometry).toMatchObject({
      canPreset: true,
      preset: null,
      floating: false,
      offsetPx: null,
    })
    applyPictureGeometry(editor, {
      ...geometry,
      widthPx: 200,
      heightPx: 100,
      preset: { h: 'right', v: 'bottom' },
    })
    expect(editor.state.doc.nodeAt(pos)?.attrs).toMatchObject({
      imageWidthPx: 200,
      imageHeightPx: 100,
      imageWrap: 'square-right',
      imagePosH: 'right',
      imagePosV: 'bottom',
      imageRelH: 'margin',
      imageRelV: 'margin',
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })
    expect(selectedPictureGeometry(editor)).toMatchObject({
      preset: { h: 'right', v: 'bottom' },
      floating: true,
    })
    editor.destroy()
  })

  it('takes the same place from the Position gallery, against the margin on the page as in the file', async () => {
    const { editor, parsed, pos } = await open(
      `<w:p><w:r>${picture('<wp:docPr id="1" name="Picture 1"/>')}</w:r></w:p><w:p><w:r><w:t>Text</w:t></w:r></w:p>`,
    )
    MENU_COMMAND_HANDLERS['picture-position'](
      { editor, hasDoc: true } as unknown as CommandHost,
      'center,bottom',
    )
    expect(editor.state.doc.nodeAt(pos)?.attrs).toMatchObject({
      imageWrap: 'square-left',
      imagePosH: 'center',
      imagePosV: 'bottom',
      imageRelH: 'margin',
      imageRelV: 'margin',
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })
    expect(selectedPictureGeometry(editor)).toMatchObject({ preset: { h: 'center', v: 'bottom' } })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const block = reparsed.blocks.find((b) => b.type === 'image')
    // what the file reads back as is what the editor already holds
    expect(block).toMatchObject({
      imagePosH: 'center',
      imagePosV: 'bottom',
      imageRelH: 'margin',
      imageRelV: 'margin',
    })
    editor.destroy()
  })

  it('offers no preset to a picture in a line of text', async () => {
    const { editor } = await open(amongText('<wp:docPr id="1" name="Picture 1"/>'))
    expect(selectedPictureGeometry(editor)).toMatchObject({ canPreset: false, preset: null })
    editor.destroy()
  })
})
