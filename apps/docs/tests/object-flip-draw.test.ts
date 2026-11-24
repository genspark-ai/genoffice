// @vitest-environment jsdom
/**
 * A mirrored shape on the page: the outline is mirrored in its box and the
 * text is left upright; a picture box turns over whole; a straight line's
 * flips only choose its ends. And Table to Text takes the shapes anchored in
 * the cells along.
 */
import { afterEach, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  tableModelToPmNode,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'
import { shapeBackgroundCss } from '../src/renderer/editor/shape-svg'
import { convertTableToText } from '../src/renderer/editor/table-convert'

const svgOf = (css: string | null) =>
  decodeURIComponent(/url\("data:image\/svg\+xml,([^"]+)"\)/.exec(css ?? '')?.[1] ?? '')

describe('a mirrored outline', () => {
  it('is the same path drawn through a mirror about the middle of its box', () => {
    const plain = svgOf(shapeBackgroundCss('rtTriangle', 102, 52, '4472C4', '000000'))
    const flipped = svgOf(
      shapeBackgroundCss('rtTriangle', 102, 52, '4472C4', '000000', undefined, undefined, {
        h: true,
      }),
    )
    const path = /<path [^>]*\/>/.exec(plain)![0]
    expect(plain).not.toContain('<g ')
    expect(flipped).toContain(`<g transform="translate(100 0) scale(-1 1)">${path}</g>`)
    const both = svgOf(
      shapeBackgroundCss('rtTriangle', 102, 52, '4472C4', '000000', undefined, undefined, {
        h: true,
        v: true,
      }),
    )
    expect(both).toContain('<g transform="translate(100 50) scale(-1 -1)">')
  })

  it('leaves a straight line to its own end points', () => {
    const line = { diag: true, flipH: true }
    expect(
      shapeBackgroundCss('lineArrow', 102, 52, undefined, '000000', line, undefined, { h: true }),
    ).toBe(shapeBackgroundCss('lineArrow', 102, 52, undefined, '000000', line))
  })
})

describe('a mirrored box on the page', () => {
  const base: TextboxDisplay = { paras: [], prst: 'rtTriangle', widthPx: 102, heightPx: 52 }
  const worded: TextboxDisplay = {
    ...base,
    paras: [{ runs: [{ text: 'Go' }] }] as TextboxDisplay['paras'],
  }

  it('mirrors the outline and not the box, so its text stays upright', () => {
    const style = textboxBoxStyle({ ...worded, fill: '4472C4', flipH: true, rotDeg: 90 })
    expect(svgOf(style)).toContain('scale(-1 1)')
    expect(style).toContain('transform:rotate(90deg)')
    expect(style).not.toContain('scaleX')
    expect(svgOf(textboxBoxStyle({ ...worded, fill: '4472C4' }))).not.toContain('scale(')
  })

  it('turns a picture box over whole, inside its rotation', () => {
    const frame: TextboxDisplay = {
      paras: [],
      readOnly: true,
      picture: true,
      fillImageDataUrl: 'data:image/png;base64,AAAA',
      widthPx: 100,
      heightPx: 50,
      flipH: true,
      flipV: true,
      rotDeg: 30,
    }
    expect(textboxBoxStyle(frame)).toContain('transform:rotate(30deg) scaleX(-1) scaleY(-1)')
    // the editor's rotator carries the turn; the flip stays on the box
    expect(textboxBoxStyle(frame, { editorChrome: true })).toContain(
      'transform:scaleX(-1) scaleY(-1)',
    )
    expect(textboxBoxStyle({ ...frame, flipH: false, flipV: false, rotDeg: 0 })).not.toContain(
      'transform',
    )
  })

  it('does not mirror a straight line a second time', () => {
    const line: TextboxDisplay = {
      paras: [],
      readOnly: true,
      prst: 'lineArrow',
      borderColor: '000000',
      widthPx: 100,
      heightPx: 50,
      lineDiag: true,
      flipV: true,
    }
    const style = textboxBoxStyle(line)
    expect(style).not.toContain('scaleY')
    expect(svgOf(style)).not.toContain('<g ')
  })
})

describe('Convert Table to Text with shapes anchored in the cells', () => {
  const editors = new Set<Editor>()
  afterEach(() => {
    for (const editor of editors) editor.destroy()
    editors.clear()
  })

  const EMU = 9525
  const shape = (id: number) =>
    '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" ' +
    'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${30 * EMU}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${10 * EMU}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${60 * EMU}" cy="${40 * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="${id}" name="Shape ${id}"/><wp:cNvGraphicFramePr/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${60 * EMU}" cy="${40 * EMU}"/></a:xfrm>` +
    '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>' +
    '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'
  const cell = (inner: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
  const p = (text: string, extra = '') => `<w:p><w:r><w:t>${text}</w:t></w:r>${extra}</w:p>`
  const body =
    '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(p('A1', `<w:r>${shape(11)}</w:r>`))}${cell(p('B1'))}</w:tr>` +
    `<w:tr>${cell(p('A2'))}${cell(p('B2', `<w:r>${shape(12)}</w:r>`))}</w:tr></w:tbl>` +
    p('After')

  async function open(tracked = false) {
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    editors.add(editor)
    if (tracked) {
      const storage = editor.storage.trackChanges as { enabled: boolean; author: string }
      storage.enabled = true
      storage.author = 'Reviewer'
    }
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 4)))
    return { editor, parsed }
  }

  const outline = (editor: Editor) => {
    const out: string[] = []
    editor.state.doc.forEach((node) =>
      out.push(
        node.type.name === 'docProtected'
          ? `shape${node.attrs.carried ? ' carried' : ''}`
          : node.type.name === 'docTable'
            ? 'table'
            : node.textContent,
      ),
    )
    return out
  }

  async function savedBody(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks, {
      carried: plan.carried.length > 0 ? plan.carried : undefined,
    })
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
    return { xml: /<w:body>([\s\S]*)<w:sectPr/.exec(xml)![1], bytes }
  }

  it('a paragraph per cell: each shape follows the paragraph it was anchored in and is saved in it', async () => {
    const { editor, parsed } = await open()
    expect(convertTableToText(editor, { separator: { kind: 'paragraph' }, nested: true })).toBe(
      true,
    )
    expect(outline(editor)).toEqual([
      'A1',
      'shape carried',
      'B1',
      'A2',
      'B2',
      'shape carried',
      'After',
    ])
    const { xml, bytes } = await savedBody(editor, parsed)
    const paragraphs = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g)!
    expect(paragraphs).toHaveLength(5)
    expect(paragraphs[0]).toContain('>A1</w:t>')
    expect(paragraphs[0]).toContain('name="Shape 11"')
    expect(paragraphs[3]).toContain('>B2</w:t>')
    expect(paragraphs[3]).toContain('name="Shape 12"')
    expect(xml).not.toContain('<w:tbl>')

    const again = await parseDocx(bytes)
    const boxes = again.blocks.flatMap((block) => block.textboxes ?? [])
    expect(boxes.map((box) => [box.prst, box.offsetXEmu, box.offsetYEmu])).toEqual([
      ['ellipse', 30 * EMU, 10 * EMU],
      ['ellipse', 30 * EMU, 10 * EMU],
    ])
  })

  it('cells joined by tabs: a row is one paragraph and its shapes come behind it', async () => {
    const { editor, parsed } = await open()
    convertTableToText(editor, { separator: { kind: 'text', text: '\t' }, nested: true })
    expect(outline(editor)).toEqual(['A1\tB1', 'shape carried', 'A2\tB2', 'shape carried', 'After'])
    const { xml } = await savedBody(editor, parsed)
    const paragraphs = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g)!
    expect(paragraphs).toHaveLength(3)
    expect(paragraphs[0]).toContain('name="Shape 11"')
    expect(paragraphs[1]).toContain('name="Shape 12"')
  })

  it('boxes of one cell paragraph that are drawn apart are not carried, so no drawing is saved twice', () => {
    const source = '<w:p><w:r><w:drawing/></w:r><w:r><w:drawing/></w:r></w:p>'
    const onPage = { paras: [], outsideCell: true, floating: true, pageRelV: true }
    const inCell = { paras: [], floating: true }
    const cellOf = (anchoredBoxes: TextboxDisplay[]) =>
      tableModelToPmNode({
        rows: [
          [
            {
              paras: ['A'],
              anchoredBoxes,
              anchoredBoxAnchors: anchoredBoxes.map(() => 0),
              anchoredBoxXml: anchoredBoxes.map(() => source),
            },
          ],
        ],
      } as never).content![0].content![0]
    const boxNodes = (cell: PmNode) => (cell.content ?? []).filter((n) => n.type === 'docCellBoxes')
    expect(boxNodes(cellOf([inCell, inCell])).map((n) => n.attrs?.xml)).toEqual([[source, source]])
    const parted = boxNodes(cellOf([onPage, inCell]))
    expect(parted).toHaveLength(2)
    expect(parted.map((n) => n.attrs?.xml ?? null)).toEqual([null, null])
  })

  it('a tracked conversion leaves the shapes with the struck table', async () => {
    const { editor } = await open(true)
    convertTableToText(editor, { separator: { kind: 'paragraph' }, nested: true })
    expect(outline(editor).filter((line) => line.startsWith('shape'))).toEqual([])
    expect(outline(editor)[0]).toBe('table')
  })
})
