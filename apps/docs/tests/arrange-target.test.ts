/**
 * What the Arrange group can do with each kind of selected object, and that
 * Rotate and Flip reach a picture in a line of text, the page and the file.
 */
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type Run, type TextboxDisplay } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { insertShapeAt } from '../src/renderer/components/ribbon-tabs'
import { arrangeStateOf } from '../src/renderer/editor/arrange-target'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { selectBoxAt } from '../src/renderer/editor/selected-box'
import { commandRunner } from './helpers/ribbon-props'

const graphic =
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic>'

const inline =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Picture 1"/>' +
  `${graphic}</wp:inline></w:drawing>`

const anchored =
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
  'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  '<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>' +
  `${graphic}</wp:anchor></w:drawing>`

const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`
const amongText = (drawing: string) =>
  `<w:p>${text('Before ')}<w:r>${drawing}</w:r>${text(' after.')}</w:p>`
const alone = (drawing: string) => `<w:p><w:r>${drawing}</w:r></w:p><w:p>${text('Text')}</w:p>`

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
  if (pos >= 0)
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return { editor, parsed, source, pos }
}

function selectNewShape(editor: Editor): number {
  insertShapeAt(editor, 'rect')
  let pos = -1
  editor.state.doc.descendants((node, at) => {
    if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
  return pos
}

async function savedPicture(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const reparsed = await parseDocx(bytes)
  const runs = reparsed.blocks.flatMap((block) => ('runs' in block ? (block.runs ?? []) : []))
  return { bytes, image: runs.find((run: Run) => run.image)?.image }
}

const can = (editor: Editor) => {
  const { kind, canWrap, canPosition, canOrder, canRotate, canFlip } = arrangeStateOf(editor, null)
  return { kind, canWrap, canPosition, canOrder, canRotate, canFlip }
}

describe('what Arrange can do per kind of object', () => {
  it('offers nothing with the caret in text', async () => {
    const { editor } = await open(`<w:p>${text('Only text')}</w:p>`)
    expect(can(editor)).toEqual({
      kind: null,
      canWrap: false,
      canPosition: false,
      canOrder: false,
      canRotate: false,
      canFlip: false,
    })
    editor.destroy()
  })

  it('wraps, places, rotates and flips a picture in a line of text, which has no rank', async () => {
    const { editor } = await open(amongText(inline))
    expect(can(editor)).toEqual({
      kind: 'runPicture',
      canWrap: true,
      canPosition: true,
      canOrder: false,
      canRotate: true,
      canFlip: true,
    })
    expect(arrangeStateOf(editor, null).wrap).toBeNull()
    editor.destroy()
  })

  it('reads the wrap of a floating picture among text', async () => {
    const { editor } = await open(amongText(anchored))
    expect(arrangeStateOf(editor, null)).toMatchObject({
      kind: 'runPicture',
      wrap: expect.stringMatching(/^square-/),
      canPosition: true,
      canOrder: true,
    })
    editor.destroy()
  })

  it('does not restack a picture in a line of text', async () => {
    const { editor } = await open(amongText(inline))
    expect(arrangeStateOf(editor, null)).toMatchObject({ kind: 'runPicture', canOrder: false })
    editor.destroy()
  })

  it('offers no place to a picture among text that a picture block cannot be written from', async () => {
    const cell = (body: string) =>
      `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${body}</w:tc></w:tr></w:tbl><w:p/>`
    const inCell = await open(cell(amongText(inline)))
    expect(inCell.editor.state.selection).toBeInstanceOf(NodeSelection)
    expect(can(inCell.editor)).toMatchObject({ kind: 'runPicture', canPosition: false })
    inCell.editor.destroy()

    const cropped = inline.replace('<a:stretch>', '<a:srcRect l="10000"/><a:stretch>')
    for (const drawing of [
      cropped,
      inline.replace(
        '<pic:cNvPr id="0" name="Picture 1"/>',
        '<pic:cNvPr id="0" name="Picture 1"><a:hlinkClick r:id="rId10"/></pic:cNvPr>',
      ),
    ]) {
      const { editor } = await open(amongText(drawing))
      expect(can(editor)).toMatchObject({ kind: 'runPicture', canPosition: false })
      commandRunner(editor)('picture-position', 'left,top')
      expect(editor.state.doc.firstChild?.type.name).toBe('docParagraph')
      editor.destroy()
    }
  })

  it('does everything but restack with a picture block in line, and restacks it once it floats', async () => {
    const { editor } = await open(alone(inline))
    expect(can(editor)).toEqual({
      kind: 'blockPicture',
      canWrap: true,
      canPosition: true,
      canOrder: false,
      canRotate: true,
      canFlip: true,
    })
    commandRunner(editor)('picture-wrap', 'front')
    expect(arrangeStateOf(editor, null)).toMatchObject({ wrap: 'front', canOrder: true })
    commandRunner(editor)('picture-position', 'right,bottom')
    expect(arrangeStateOf(editor, null)).toMatchObject({ posH: 'right', posV: 'bottom' })
    editor.destroy()
  })

  it('wraps, restacks, rotates and flips a shape', async () => {
    const { editor } = await open(`<w:p>${text('Body')}</w:p>`)
    selectNewShape(editor)
    expect(can(editor)).toEqual({
      kind: 'shape',
      canWrap: true,
      canPosition: false,
      canOrder: true,
      canRotate: true,
      canFlip: true,
    })
    editor.destroy()
  })

  it('wraps and restacks a picture frame; one with no drawing of its own keeps its turn', async () => {
    const { editor } = await open(`<w:p>${text('Body')}</w:p>`)
    const pos = selectNewShape(editor)
    const node = editor.state.doc.nodeAt(pos)!
    const [box] = node.attrs.textboxes as TextboxDisplay[]
    const tr = editor.state.tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      textboxes: [{ ...box, picture: true, readOnly: true }],
    })
    editor.view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)))
    expect(can(editor)).toEqual({
      kind: 'pictureFrame',
      canWrap: true,
      canPosition: false,
      canOrder: true,
      canRotate: false,
      canFlip: false,
    })
    commandRunner(editor)('picture-rotate', '90')
    expect((editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[])[0].rotDeg).toBe(
      undefined,
    )
    editor.destroy()
  })
})

describe('Rotate and Flip through the one pair of commands', () => {
  it('turns a picture that shares its paragraph with a shape, and saves the turn on the picture', async () => {
    const shape =
      '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
      'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>1905000</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
      '<wp:docPr id="2" name="Text Box 2"/><wp:cNvGraphicFramePr/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp><wps:cNvPr id="2" name="Text Box 2"/><wps:cNvSpPr txBox="1"/><wps:spPr>' +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
      '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
      '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'
    const body = `<w:p><w:r>${shape}</w:r><w:r>${anchored}</w:r></w:p><w:p>${text('Text')}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body, withImage: true }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    let pos = -1
    editor.state.doc.forEach((node, at) => {
      if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = at
    })
    const boxes = () => editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[]
    const frame = boxes().findIndex((box) => box.picture)
    expect(frame).toBeGreaterThanOrEqual(0)
    editor.view.dispatch(selectBoxAt(editor.state.tr, pos, frame))
    expect(can(editor)).toMatchObject({ kind: 'pictureFrame', canRotate: true, canFlip: true })

    commandRunner(editor)('picture-rotate', '90')
    expect(boxes()[frame].rotDeg).toBe(90)
    expect(boxes()[1 - frame].rotDeg).toBeUndefined()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toMatch(/<pic:spPr><a:xfrm rot="5400000">/)
    expect(xml).not.toMatch(/<wps:spPr><a:xfrm rot=/)
    const again = (await parseDocx(saved)).blocks.flatMap((block) => block.textboxes ?? [])
    expect(again.find((box) => box.picture)?.rotDeg).toBe(90)
    expect(again.find((box) => !box.picture)?.rotDeg).toBeUndefined()

    // turned back, the attribute is gone again
    commandRunner(editor)('picture-rotate', '-90')
    const back = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const plain = await saveDocx(parsed, back.saveBlocks)
    expect(
      await (await JSZip.loadAsync(plain)).file('word/document.xml')!.async('string'),
    ).not.toContain(' rot="')

    // mirrored: on the picture's own xfrm, read back on the frame
    commandRunner(editor)('picture-flip', 'h')
    expect(boxes()[frame].flipH).toBe(true)
    expect(boxes()[1 - frame].flipH).toBeUndefined()
    const flipped = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const flippedXml = await (
      await JSZip.loadAsync(flipped)
    )
      .file('word/document.xml')!
      .async('string')
    expect(flippedXml).toMatch(/<pic:spPr><a:xfrm flipH="1">/)
    expect(flippedXml).not.toMatch(/<wps:spPr><a:xfrm[^>]*flip/)
    const mirrored = (await parseDocx(flipped)).blocks.flatMap((block) => block.textboxes ?? [])
    expect(mirrored.find((box) => box.picture)?.flipH).toBe(true)
    editor.destroy()
  })

  it('turns and mirrors a picture in a line of text, on the node and in the file', async () => {
    const { editor, parsed, source, pos } = await open(amongText(inline))
    // untouched: the file is the one that was opened
    expect((await savedPicture(editor, parsed)).bytes).toEqual(source)

    const run = commandRunner(editor)
    run('picture-rotate', '90')
    expect(editor.state.doc.nodeAt(pos)?.attrs.rotDeg).toBe(90)
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)
    run('picture-flip', 'h')
    expect(editor.state.doc.nodeAt(pos)?.attrs).toMatchObject({ rotDeg: 90, flipH: true })

    const { image } = await savedPicture(editor, parsed)
    expect(image?.xml).toContain('<a:xfrm rot="5400000" flipH="1">')
    // Word lays a line out by the extent plus what the turn adds on each side
    expect(image?.xml).toContain('<wp:effectExtent l="0" t="238125" r="0" b="238125"/>')
    expect(image?.xml).toContain('<wp:extent cx="952500" cy="476250"/>')
    expect(image).toMatchObject({ rotDeg: 90, flipH: true, widthPx: 100, heightPx: 50 })
    expect(image?.flipV).toBeUndefined()
    editor.destroy()
  })

  it('takes the turn back out of the file after a full circle', async () => {
    const { editor, parsed, pos } = await open(amongText(inline))
    const run = commandRunner(editor)
    for (let i = 0; i < 3; i++) run('picture-rotate', '90')
    run('picture-flip', 'v')
    expect(editor.state.doc.nodeAt(pos)?.attrs).toMatchObject({ rotDeg: 270, flipV: true })
    run('picture-rotate', '90')
    run('picture-flip', 'v')
    expect(editor.state.doc.nodeAt(pos)?.attrs).toMatchObject({ rotDeg: null, flipV: false })
    const { image } = await savedPicture(editor, parsed)
    expect(image?.xml).toContain('<a:xfrm>')
    expect(image?.xml).not.toMatch(/\brot=|flipV=|flipH=/)
    editor.destroy()
  })

  it('turns a floating picture among text and leaves where it stands alone', async () => {
    const { editor, parsed } = await open(amongText(anchored))
    commandRunner(editor)('picture-rotate', '-90')
    const { image } = await savedPicture(editor, parsed)
    expect(image?.xml).toContain('<a:xfrm rot="16200000">')
    expect(image?.xml).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>',
    )
    expect(image?.xml).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(image?.rotDeg).toBe(270)
    editor.destroy()
  })

  it('undoes a turn in one step', async () => {
    const { editor, pos } = await open(amongText(inline))
    commandRunner(editor)('picture-rotate', '90')
    editor.commands.undo()
    expect(editor.state.doc.nodeAt(pos)?.attrs.rotDeg).toBeNull()
    editor.destroy()
  })

  it('still turns and mirrors a picture block', async () => {
    const { editor, pos } = await open(alone(inline))
    const run = commandRunner(editor)
    run('picture-rotate', '90')
    run('picture-flip', 'v')
    expect(editor.state.doc.nodeAt(pos)?.attrs).toMatchObject({
      imageRotDeg: 90,
      imageFlipV: true,
    })
    editor.destroy()
  })

  it('turns and mirrors a shape, and saves both', async () => {
    const { editor, parsed } = await open(`<w:p>${text('Body')}</w:p>`)
    const pos = selectNewShape(editor)
    const run = commandRunner(editor)
    const box = () => (editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[])[0]
    run('picture-rotate', '-90')
    expect(box().rotDeg).toBe(270)
    run('picture-flip', 'h')
    expect(box().flipH).toBe(true)
    run('picture-flip', 'v')
    run('picture-flip', 'h')
    expect(box()).not.toHaveProperty('flipH')
    expect(box().flipV).toBe(true)
    run('picture-rotate', '90')
    expect(box()).not.toHaveProperty('rotDeg')

    const saved = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toMatch(/<wps:spPr><a:xfrm flipV="1">/)
    expect(xml).not.toContain('flipH')
    const again = (await parseDocx(saved)).blocks.flatMap((block) => block.textboxes ?? [])
    expect(again[0]).toMatchObject({ flipV: true })
    expect(again[0]).not.toHaveProperty('flipH')
    editor.destroy()
  })

  it('a file shape opened mirrored is mirrored back, and untouched it saves as it came', async () => {
    const shape = (flip: string) =>
      '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
      'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="952500" cy="476250"/><wp:wrapNone/><wp:docPr id="7" name="Shape 7"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm${flip}><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>` +
      '<a:prstGeom prst="rtTriangle"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>' +
      '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
    const { editor, parsed, source } = await open(
      shape(' flipH="1"') + `<w:p>${text('Body')}</w:p>`,
    )
    let pos = -1
    editor.state.doc.forEach((node, at) => {
      if (node.attrs.textboxes) pos = at
    })
    const box = () => (editor.state.doc.nodeAt(pos)!.attrs.textboxes as TextboxDisplay[])[0]
    expect(box()).toMatchObject({ prst: 'rtTriangle', flipH: true })
    const untouched = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    expect(untouched).toEqual(source)

    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    commandRunner(editor)('picture-flip', 'h')
    expect(box()).not.toHaveProperty('flipH')
    const saved = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain('<wps:spPr><a:xfrm><a:off x="0" y="0"/>')
    editor.destroy()
  })

  it('restacks a floating picture among text past the nearest rank, and saves the rank', async () => {
    const over = anchored
      .replace('<wp:wrapSquare wrapText="bothSides"/>', '<wp:wrapNone/>')
      .replace('relativeHeight="251659264"', 'relativeHeight="251661312"')
      .replace('id="1" name="Picture 1"', 'id="2" name="Picture 2"')
    const under = anchored.replace('<wp:wrapSquare wrapText="bothSides"/>', '<wp:wrapNone/>')
    const { editor, parsed, source } = await open(
      `<w:p>${text('Before ')}<w:r>${under}</w:r><w:r>${over}</w:r>${text(' after.')}</w:p>`,
    )
    const ranks = () => {
      const out: (number | null)[] = []
      editor.state.doc.descendants((node) => {
        if (node.type.name === 'docInlineImage') out.push(node.attrs.zOrder as number | null)
      })
      return out
    }
    expect(ranks()).toEqual([1024, 3072])
    expect((await savedPicture(editor, parsed)).bytes).toEqual(source)
    expect(arrangeStateOf(editor, null)).toMatchObject({ wrap: 'front', canOrder: true })

    const run = commandRunner(editor)
    run('arrange-order', 'forward')
    expect(ranks()).toEqual([3073, 3072])
    run('arrange-order', 'backward')
    expect(ranks()).toEqual([3071, 3072])
    run('arrange-order', 'back')
    expect(ranks()).toEqual([3070, 3072])
    run('arrange-order', 'front')
    expect(ranks()).toEqual([3073, 3072])

    const { bytes } = await savedPicture(editor, parsed)
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
    expect(
      [...xml.matchAll(/relativeHeight="(\d+)"/g)].map((m) => Number(m[1]) - 251658240),
    ).toEqual([3073, 3072])
    const again = (await parseDocx(bytes)).blocks.flatMap((block) =>
      (block.runs ?? []).flatMap((r) => (r.image ? [r.image.zOrder] : [])),
    )
    expect(again).toEqual([3073, 3072])
    editor.destroy()
  })

  it('does nothing in a read-only document', async () => {
    const { editor, pos } = await open(amongText(inline))
    editor.setEditable(false)
    commandRunner(editor)('picture-rotate', '90')
    expect(editor.state.doc.nodeAt(pos)?.attrs.rotDeg).toBeNull()
    editor.destroy()
  })
})

describe('Position on a picture in a line of text', () => {
  const blocksOf = (editor: Editor) => {
    const out: string[] = []
    editor.state.doc.forEach((node) =>
      out.push(node.type.name === 'docProtected' ? 'picture' : node.textContent),
    )
    return out
  }

  it('takes the picture out of its line into a picture block at the preset place', async () => {
    const { editor } = await open(amongText(inline))
    commandRunner(editor)('picture-rotate', '90')
    commandRunner(editor)('picture-position', 'right,bottom')
    expect(blocksOf(editor)).toEqual(['picture', 'Before  after.'])
    const sel = editor.state.selection
    expect(sel).toBeInstanceOf(NodeSelection)
    expect((sel as NodeSelection).node.attrs).toMatchObject({
      blockType: 'image',
      imageWrap: 'square-right',
      imagePosH: 'right',
      imagePosV: 'bottom',
      imageRelH: 'margin',
      imageRelV: 'margin',
      imageWidthPx: 100,
      imageHeightPx: 50,
      imageRotDeg: 90,
      anchorLine: { styleId: null },
    })
    // it is a picture block from here on: the group reads its place and offers the rank
    expect(arrangeStateOf(editor, null)).toMatchObject({
      kind: 'blockPicture',
      posH: 'right',
      posV: 'bottom',
      canOrder: true,
    })
    commandRunner(editor)('picture-position', 'center,top')
    expect(arrangeStateOf(editor, null)).toMatchObject({ posH: 'center', posV: 'top' })
    expect(blocksOf(editor)).toEqual(['picture', 'Before  after.'])
    editor.destroy()
  })

  it('writes the anchor with the preset pair, and the file reads back as the page shows', async () => {
    const { editor, parsed } = await open(amongText(inline))
    commandRunner(editor)('picture-flip', 'h')
    commandRunner(editor)('picture-position', 'center,bottom')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reparsed.blocks.slice(0, 2).map((b) => b.type)).toEqual(['image', 'paragraph'])
    const [picture, para] = reparsed.blocks
    expect(picture.originalXml).toContain(
      '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
    )
    expect(picture.originalXml).toContain(
      '<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align></wp:positionV>',
    )
    expect(picture.originalXml).toMatch(/<wp:anchor\b[\s\S]*<wp:wrapSquare\b/)
    expect(picture.originalXml).toContain('flipH="1"')
    expect(para.originalXml).not.toContain('<w:drawing>')
    // the anchor's own paragraph keeps a line in the file, and so on the page
    expect(picture.anchorLine).toBeDefined()
    const shown = (editor.state.selection as NodeSelection).node.attrs
    expect(picture).toMatchObject({
      imageWrap: shown.imageWrap,
      imagePosH: shown.imagePosH,
      imagePosV: shown.imagePosV,
      imageRelH: shown.imageRelH,
      imageRelV: shown.imageRelV,
      imageWidthPx: shown.imageWidthPx,
      imageHeightPx: shown.imageHeightPx,
      imageWrapDistLeftEmu: shown.imageWrapDistLeftEmu,
      imageWrapDistRightEmu: shown.imageWrapDistRightEmu,
      imageFlipH: true,
    })
    editor.destroy()
  })

  it('places a picture already floating among text the same way', async () => {
    const { editor } = await open(amongText(anchored))
    commandRunner(editor)('picture-position', 'left,center')
    expect(blocksOf(editor)).toEqual(['picture', 'Before  after.'])
    expect(arrangeStateOf(editor, null)).toMatchObject({
      kind: 'blockPicture',
      wrap: 'square-left',
      posH: 'left',
      posV: 'center',
    })
    editor.destroy()
  })

  it('puts the picture back in its line in one undo step', async () => {
    const { editor, pos } = await open(amongText(inline))
    const before = editor.getJSON()
    commandRunner(editor)('picture-position', 'left,top')
    expect(blocksOf(editor)).toEqual(['picture', 'Before  after.'])
    editor.commands.undo()
    expect(editor.getJSON()).toEqual(before)
    expect(editor.state.doc.nodeAt(pos)?.type.name).toBe('docInlineImage')
    editor.commands.redo()
    expect(blocksOf(editor)).toEqual(['picture', 'Before  after.'])
    editor.destroy()
  })

  it('leaves a read-only document alone', async () => {
    const { editor } = await open(amongText(inline))
    editor.setEditable(false)
    commandRunner(editor)('picture-position', 'left,top')
    expect(blocksOf(editor)).toEqual(['Before  after.'])
    editor.destroy()
  })
})
