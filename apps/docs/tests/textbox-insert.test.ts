/**
 * Textbox insertion tests - Item 1: create textbox → save → re-parse
 *
 * Uses buildTextboxParagraphXml() to produce genXml nodes, verifying:
 *  1. The saved genXml contains a valid wps:wsp/txbxContent structure
 *  2. The textboxes array is non-empty after re-parsing
 *  3. Text edits can be patched in
 *  4. The wrap mode can be updated (applyImageWrap)
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  buildTextboxParagraphXml,
  parseDocx,
  saveDocx,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const WIDTH_EMU = 1800000
const HEIGHT_EMU = 1080000

function makeEmptyTextbox(): TextboxDisplay {
  return {
    fill: 'FFFFFF',
    borderColor: '000000',
    widthPx: Math.round(WIDTH_EMU / 9525),
    heightPx: Math.round(HEIGHT_EMU / 9525),
    paras: [{ runs: [{ text: '' }] }],
  }
}

async function openBlankDoc() {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

/** locate the first docProtected node with textboxes in the outer doc */
function textboxNodePos(editor: Editor): number {
  let pos = -1
  editor.state.doc.descendants((node, p) => {
    if (pos !== -1) return false
    if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = p
    return pos === -1
  })
  return pos
}

describe('text box insertion', () => {
  it('buildTextboxParagraphXml generates paragraph XML containing wps:wsp', () => {
    const xml = buildTextboxParagraphXml({ widthEmu: WIDTH_EMU, heightEmu: HEIGHT_EMU, id: 1 })
    expect(xml).toContain('wps:wsp')
    expect(xml).toContain('w:txbxContent')
    expect(xml).toContain('wp:anchor')
    expect(xml).toContain('prstGeom prst="rect"')
    expect(xml).toContain(`cx="${WIDTH_EMU}"`)
    expect(xml).toContain(`cy="${HEIGHT_EMU}"`)
    expect(xml).toContain('FFFFFF')
    expect(xml).toContain('000000')
    expect(xml).toContain('wp:wrapSquare')
    expect(xml).toContain('mc:AlternateContent')
  })

  it('genXml textbox node saves as kind:xml with the XML containing the wps:wsp structure', async () => {
    const { editor, parsed } = await openBlankDoc()
    const xml = buildTextboxParagraphXml({ widthEmu: WIDTH_EMU, heightEmu: HEIGHT_EMU, id: 1 })
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Text box',
          genXml: xml,
          textboxes: [makeEmptyTextbox()],
        },
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBeGreaterThanOrEqual(1)
    const xmlBlock = plan.saveBlocks.find((b) => b.kind === 'xml')
    expect(xmlBlock).toBeDefined()
    expect((xmlBlock as { kind: 'xml'; xml: string }).xml).toContain('wps:wsp')
    expect((xmlBlock as { kind: 'xml'; xml: string }).xml).toContain('w:txbxContent')
    editor.destroy()
  })

  it('genXml textbox is discoverable as textboxes after save and reparse', async () => {
    const { editor, parsed } = await openBlankDoc()
    const xml = buildTextboxParagraphXml({ widthEmu: WIDTH_EMU, heightEmu: HEIGHT_EMU, id: 1 })
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Text box',
          genXml: xml,
          textboxes: [makeEmptyTextbox()],
        },
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.textboxes)
    expect(block).toBeDefined()
    expect(block?.textboxes?.length).toBe(1)
    expect(block?.textboxes?.[0].fill).toBe('FFFFFF')
    expect(block?.textboxes?.[0].borderColor).toBe('000000')
    editor.destroy()
  })

  it('reparse returns the correct text after inserting and editing text', async () => {
    const { editor, parsed } = await openBlankDoc()
    const xml = buildTextboxParagraphXml({ widthEmu: WIDTH_EMU, heightEmu: HEIGHT_EMU, id: 2 })
    const textbox: TextboxDisplay = {
      fill: 'FFFFFF',
      borderColor: '000000',
      widthPx: Math.round(WIDTH_EMU / 9525),
      heightPx: Math.round(HEIGHT_EMU / 9525),
      paras: [{ runs: [{ text: 'Hello, text box!' }] }],
    }
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Text box',
          genXml: xml,
          textboxes: [textbox],
        },
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks.find((b) => b.textboxes)
    expect(block?.textboxes?.[0].paras[0].runs[0].text).toBe('Hello, text box!')
    editor.destroy()
  })

  it('patchTextboxParas updates the content when the genXml node already has textboxes', async () => {
    const { editor, parsed } = await openBlankDoc()
    const xml = buildTextboxParagraphXml({ widthEmu: WIDTH_EMU, heightEmu: HEIGHT_EMU, id: 3 })
    // Insert with empty text first
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Text box',
          genXml: xml,
          textboxes: [makeEmptyTextbox()],
        },
      })
      .run()
    // Then update the textboxes attr to simulate sub-editor commit
    const pos = textboxNodePos(editor)
    expect(pos).toBeGreaterThanOrEqual(0)
    const node = editor.state.doc.nodeAt(pos)!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        textboxes: [
          {
            ...makeEmptyTextbox(),
            paras: [{ runs: [{ text: 'Edited content' }] }],
          },
        ],
      }),
    )

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xmlBlock = plan.saveBlocks.find((b) => b.kind === 'xml') as
      { kind: 'xml'; xml: string } | undefined
    expect(xmlBlock).toBeDefined()
    expect(xmlBlock?.xml).toContain('Edited content')
    editor.destroy()
  })

  it('genXml textbox with a wrap mode set saves with wp:wrapSquare', async () => {
    const { editor, parsed } = await openBlankDoc()
    const xml = buildTextboxParagraphXml({ widthEmu: WIDTH_EMU, heightEmu: HEIGHT_EMU, id: 4 })
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Text box',
          genXml: xml,
          textboxes: [makeEmptyTextbox()],
          imageWrap: 'square-right',
        },
      })
      .run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xmlBlock = plan.saveBlocks.find((b) => b.kind === 'xml') as
      { kind: 'xml'; xml: string } | undefined
    expect(xmlBlock?.xml).toContain('wp:wrapSquare')
    expect(xmlBlock?.xml).toContain('wp:anchor')
    editor.destroy()
  })
})
