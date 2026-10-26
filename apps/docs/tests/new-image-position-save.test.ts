/**
 * Newly inserted pictures (genImage payload, no docxIndex): drag offsets and
 * Layout ▸ Position presets must survive save + reopen. The genImage save
 * branch re-derives the floating position from the live node attrs into
 * NewImage.posOffsetEmu / posH+posV (the insert-time payload is stale after
 * any edit); the parse side reads the anchor back into imageOffsetXEmu /
 * imageRelH/imageRelV / imagePosH+imagePosV, so the reopened picture edits
 * like an original one.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import type { Node as ProsemirrorNode } from '@tiptap/pm/model'
import JSZip from 'jszip'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { pictureNode } from '../src/renderer/ai/floating-ops'
import { insertPictureAtSelection } from '../src/renderer/editor/insert-picture'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

/** a floating picture (square wrap, column/paragraph posOffset) for paste tests */
const FLOATING_PIC_PARA =
  '<w:p><w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1143000" cy="762000"/>' +
  '<wp:wrapSquare wrapText="bothSides"/>' +
  '<wp:docPr id="1" name="Image 1"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="762000"/></a:xfrm></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic>' +
  '</wp:anchor></w:drawing></w:r></w:p>'

const picture = {
  dataUrl: `data:image/png;base64,${TINY_PNG_BASE64}`,
  base64: TINY_PNG_BASE64,
  mime: 'image/png',
  widthPx: 120,
  heightPx: 80,
  label: 'Picture',
}

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

function selectPicture(editor: Editor) {
  let target = -1
  editor.state.doc.descendants((node, pos) => {
    if (target >= 0) return false
    if (node.type.name === 'docProtected' && node.attrs.blockType === 'image') {
      target = pos
      return false
    }
    return true
  })
  expect(target).toBeGreaterThanOrEqual(0)
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, target)))
}

async function save(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks, { newPictures: plan.newPictures })
  const zip = await JSZip.loadAsync(bytes)
  return { plan, zip, xml: await zip.file('word/document.xml')!.async('string') }
}

async function reopen(zip: JSZip) {
  const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { parsed, editor }
}

function savedImage(plan: Awaited<ReturnType<typeof pmDocToSavePlan>>) {
  const block = plan.saveBlocks.find((b) => b.kind === 'image')
  if (!block || block.kind !== 'image') throw new Error('no image save block')
  return block.image
}

function pmImage(editor: Editor): ProsemirrorNode | undefined {
  let found: ProsemirrorNode | undefined
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docProtected' && node.attrs.blockType === 'image') {
      found = found ?? node
      return false
    }
    return true
  })
  return found
}

describe('new image floating position on save', () => {
  it('keeps drag offsets of a newly inserted floating picture', async () => {
    const { editor, parsed } = await openBlank()
    insertPictureAtSelection(editor, picture)
    selectPicture(editor)
    // what a drag commits: wrap keeps its kind, offsets replace the align pair
    editor.commands.updateAttributes('docProtected', {
      imageWrap: 'square-left',
      imageOffsetXEmu: 914400,
      imageOffsetYEmu: 457200,
      imagePosH: null,
      imagePosV: null,
    })

    const { plan, zip, xml } = await save(editor, parsed)
    const image = savedImage(plan)
    expect(image.wrap).toBe('square-left')
    expect(image.posOffsetEmu).toEqual({ x: 914400, y: 457200 })
    expect(xml).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>',
    )
    expect(xml).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>',
    )

    const { parsed: reparsed, editor: reopened } = await reopen(zip)
    const block = reparsed.blocks.find((b) => b.type === 'image')
    expect(block?.imageWrap).toBe('square-left')
    expect(block?.imageOffsetXEmu).toBe(914400)
    expect(block?.imageOffsetYEmu).toBe(457200)
    // the reopened picture is an ordinary anchored original: later edits patch it
    expect(pmImage(reopened)?.attrs.docxIndex).not.toBeNull()
    reopened.destroy()
  })

  it('keeps a Layout position preset of a newly inserted picture (margin align)', async () => {
    const { editor, parsed } = await openBlank()
    insertPictureAtSelection(editor, picture)
    selectPicture(editor)
    // what applyPositionPreset writes (center/bottom cell of the nine-grid)
    editor.commands.updateAttributes('docProtected', {
      imageWrap: 'square-left',
      imagePosH: 'center',
      imagePosV: 'bottom',
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })

    const { plan, zip, xml } = await save(editor, parsed)
    const image = savedImage(plan)
    expect(image.posH).toBe('center')
    expect(image.posV).toBe('bottom')
    expect(image.posOffsetEmu).toBeUndefined()
    // same bytes the original-image patch path writes for presets
    expect(xml).toContain(
      '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
    )
    expect(xml).toContain(
      '<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align></wp:positionV>',
    )

    const { parsed: reparsed } = await reopen(zip)
    const block = reparsed.blocks.find((b) => b.type === 'image')
    expect(block?.imageWrap).toBe('square-left')
    expect(block?.imagePosH).toBe('center')
    expect(block?.imagePosV).toBe('bottom')
    expect(block?.imageOffsetXEmu).toBeUndefined()
  })

  it('keeps page-relative drag offsets of an AI-inserted picture', async () => {
    const inserted = pictureNode({
      base64: TINY_PNG_BASE64,
      mime: 'image/png',
      naturalWidth: 120,
      naturalHeight: 80,
      float: { anchor: 'page', xEmu: 914400, yEmu: 457200, wrap: 'behind' },
    })
    if ('error' in inserted) throw new Error(inserted.error)
    const { editor, parsed } = await openBlank()
    editor.chain().insertContentAt(editor.state.doc.content.size, inserted.node).run()
    selectPicture(editor)
    // drag moves the picture: the attrs change, the stale insert payload does not
    editor.commands.updateAttributes('docProtected', {
      imageOffsetXEmu: 1828800,
      imageOffsetYEmu: 914400,
      imagePosH: null,
      imagePosV: null,
    })

    const { plan, zip, xml } = await save(editor, parsed)
    const image = savedImage(plan)
    expect(image.posOffsetEmu).toEqual({ x: 1828800, y: 914400, relativeTo: 'page' })
    expect(xml).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionH>',
    )

    const { parsed: reparsed } = await reopen(zip)
    const block = reparsed.blocks.find((b) => b.type === 'image')
    expect(block?.imageOffsetXEmu).toBe(1828800)
    expect(block?.imageOffsetYEmu).toBe(914400)
    expect(block?.imageRelH).toBe('page')
    expect(block?.imageRelV).toBe('page')
  })

  it('drops the anchor when a floating new picture goes back inline', async () => {
    const inserted = pictureNode({
      base64: TINY_PNG_BASE64,
      mime: 'image/png',
      naturalWidth: 120,
      naturalHeight: 80,
      float: { anchor: 'page', xEmu: 914400, yEmu: 457200, wrap: 'behind' },
    })
    if ('error' in inserted) throw new Error(inserted.error)
    const { editor, parsed } = await openBlank()
    editor.chain().insertContentAt(editor.state.doc.content.size, inserted.node).run()
    selectPicture(editor)
    // what setFloatingWrap(editor, null) (Wrap Text ▸ Inline) writes: the
    // genImage payload still carries wrap + posOffsetEmu and must lose them
    editor.commands.updateAttributes('docProtected', {
      imageWrap: null,
      imagePosH: null,
      imagePosV: null,
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })

    const { plan, zip, xml } = await save(editor, parsed)
    const image = savedImage(plan)
    expect(image.wrap).toBeUndefined()
    expect(image.posOffsetEmu).toBeUndefined()
    expect(xml).not.toContain('<wp:anchor')

    const { parsed: reparsed } = await reopen(zip)
    const block = reparsed.blocks.find((b) => b.type === 'image')
    expect(block?.imageWrap ?? null).toBeNull()
  })

  it('keeps the position of a pasted copy of an original floating picture', async () => {
    const source = await buildDocx({ bodyXml: FLOATING_PIC_PARA, withImage: true })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    // a duplicate shares the source's docxIndex: the anchor is consumed by the
    // first occurrence, the copy rebuilds from its attrs via imageFromProtectedAttrs
    let node: ProsemirrorNode | undefined
    editor.state.doc.descendants((n) => {
      if (n.type.name === 'docProtected' && n.attrs.blockType === 'image') {
        node = node ?? n
        return false
      }
      return true
    })
    const origin = node
    if (!origin) throw new Error('no image node')
    editor.view.dispatch(editor.state.tr.insert(editor.state.doc.content.size, origin))

    const { plan } = await save(editor, parsed)
    const clone = plan.saveBlocks.find(
      (b) => b.kind === 'image' && b.image.posOffsetEmu !== undefined,
    )
    if (!clone || clone.kind !== 'image') throw new Error('no cloned image save block')
    expect(clone.image.wrap).toBe('square-left')
    expect(clone.image.posOffsetEmu).toEqual({ x: 914400, y: 457200 })
  })
})
