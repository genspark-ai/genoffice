/**
 * Font Substitution's Convert Permanently on the text the editor holds, a
 * picked stand-in in the list of missing fonts, Graphics Fill on an SVG, and
 * the direction of a selected text box.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { buildTextboxParagraphXml, parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { convertFontsInEditor } from '../src/renderer/editor/convert-fonts'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { recolorSvg } from '../src/renderer/editor/graphics-fill'
import {
  selectedBoxDirection,
  setSelectedBoxDirection,
  textboxDirectionOf,
} from '../src/renderer/editor/textbox-direction'
import { checkMissingFonts } from '../src/renderer/font-check'

const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
  vi.restoreAllMocks()
})

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  liveEditors.push(editor)
  return { editor, parsed }
}

describe('Convert Permanently in the editor', () => {
  const run = (font: string, text: string) =>
    `<w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`

  it('renames the font on the runs that name it, in one undo step, and only on those', async () => {
    const { editor, parsed } = await open(
      `<w:p>${run('Old Sans', 'one ')}${run('Courier New', 'two ')}${run('old sans', 'three')}</w:p>` +
        '<w:p><w:r><w:t>untouched</w:t></w:r></w:p>',
    )
    expect(convertFontsInEditor(editor, { 'Old Sans': 'Arial' })).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const first = plan.saveBlocks[0]
    const fonts =
      first.kind === 'generated' ? (first.block.runs ?? []).map((r) => r.fontAscii ?? r.font) : []
    expect(fonts).toEqual(['Arial', 'Courier New', 'Arial'])
    // the paragraph that named no such font is still saved from the file's bytes
    expect(plan.saveBlocks[1].kind).toBe('original')

    editor.commands.undo()
    const back = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(back.saveBlocks[0].kind).toBe('original')
  })

  it('changes nothing when no run names the font', async () => {
    const { editor } = await open(`<w:p>${run('Courier New', 'two')}</w:p>`)
    const before = JSON.stringify(editor.getJSON())
    expect(convertFontsInEditor(editor, { 'Old Sans': 'Arial' })).toBe(false)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
  })
})

describe('a picked stand-in in the list of missing fonts', () => {
  it('lists the font with the pick even though the probe finds its name installed', () => {
    // every family measures unlike the fallback: the probe calls all of them installed
    let current = 0
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
      set font(spec: string) {
        current = /"/.test(spec) ? 100 : 50
      },
      measureText: () => ({ width: current }),
    } as unknown as CanvasRenderingContext2D)
    expect(checkMissingFonts(['Old Sans', 'Arial'])).toEqual([])
    expect(
      checkMissingFonts(['Old Sans', 'Arial'], (name) => (name === 'Old Sans' ? 'Verdana' : null)),
    ).toEqual([{ name: 'Old Sans', substitute: 'Verdana', chosen: true }])
  })
})

describe('Graphics Fill', () => {
  const ICON =
    '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24"><path d="M2 2h20v20H2z"/></svg>'

  it('fills an icon in the colour, and takes the fill off for none', () => {
    const red = recolorSvg(ICON, 'FF0000')!
    expect(red).toContain('<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96"')
    expect(red).toContain('fill="#FF0000"')
    expect(red).toContain('<path d="M2 2h20v20H2z"/>')
    // again: the colour is replaced, not added a second time
    const blue = recolorSvg(red, '0000FF')!
    expect(blue.match(/fill=/g)).toHaveLength(1)
    expect(blue).toContain('fill="#0000FF"')
    expect(recolorSvg(ICON, null)).toContain('fill="none"')
  })

  it('recolours shapes that carry a fill of their own, leaves unfilled ones and strokes alone', () => {
    const drawing =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">' +
      '<rect fill="#000" width="4" height="4"/>' +
      '<circle fill="none" stroke="#123456" r="2"/>' +
      '<path style="stroke:#111;fill:blue" d="M0 0h1"/></svg>'
    const out = recolorSvg(drawing, '00AA00')!
    expect(out).toContain('<rect fill="#00AA00"')
    expect(out).toContain('<circle fill="none" stroke="#123456"')
    expect(out).toContain('style="stroke:#111;fill:#00AA00"')
  })

  it('refuses what is not an SVG', () => {
    expect(recolorSvg('<html><body/></html>', 'FF0000')).toBeNull()
    expect(recolorSvg('<svg><path', 'FF0000')).toBeNull()
  })
})

describe('the direction of a selected text box', () => {
  const selectBox = (editor: Editor) => {
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (at < 0 && node.type.name === 'docProtected' && node.attrs.textboxes) at = pos
      return at < 0
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, at)))
  }
  const savedXml = (editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) => {
    const block = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks.find(
      (b) => b.kind === 'xml',
    )
    return block?.kind === 'xml' ? block.xml : ''
  }

  it('turns the text of a box from a file and is saved as its bodyPr vert', async () => {
    const { editor, parsed } = await open(
      buildTextboxParagraphXml({ id: 11 }) + '<w:p><w:r><w:t>after</w:t></w:r></w:p>',
    )
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(selectedBoxDirection(editor)).toBeUndefined()
    selectBox(editor)
    expect(selectedBoxDirection(editor)).toBeNull()
    expect(setSelectedBoxDirection(editor, 'vert270')).toBe(true)
    expect(selectedBoxDirection(editor)).toBe('vert270')
    const box = () => editor.view.dom.querySelector('.doc-textbox')!
    expect([...box().classList]).toEqual(
      expect.arrayContaining(['doc-textbox-vert', 'doc-textbox-vert270']),
    )
    expect(savedXml(editor, parsed)).toContain('<wps:bodyPr vert="vert270"/>')

    setSelectedBoxDirection(editor, 'wordArtVert')
    expect(box().classList.contains('doc-textbox-wordartvert')).toBe(true)
    expect(box().classList.contains('doc-textbox-vert270')).toBe(false)
    expect(savedXml(editor, parsed)).toContain('<wps:bodyPr vert="wordArtVert"/>')

    // horizontal again: nothing of the box is left changed
    setSelectedBoxDirection(editor, null)
    expect(box().className).not.toContain('vert')
    expect(savedXml(editor, parsed)).toBe('')
  })

  it('makes a vertical box from a file horizontal', async () => {
    const { editor, parsed } = await open(buildTextboxParagraphXml({ id: 12, vert: 'eaVert' }))
    selectBox(editor)
    expect(selectedBoxDirection(editor)).toBe('eaVert')
    setSelectedBoxDirection(editor, null)
    const xml = savedXml(editor, parsed)
    expect(xml).toContain('<wps:bodyPr/>')
    expect(xml).not.toContain('vert=')
  })

  it('reads a command payload as a direction', () => {
    expect(textboxDirectionOf('horz')).toBeNull()
    expect(textboxDirectionOf('wordArtVert')).toBe('wordArtVert')
    expect(textboxDirectionOf('sideways')).toBeUndefined()
    expect(textboxDirectionOf(undefined)).toBeUndefined()
  })
})
