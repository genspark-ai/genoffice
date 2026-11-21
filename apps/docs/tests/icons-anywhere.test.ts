/**
 * Icons outside the body: one put into a header or a text box carries its
 * drawing to the save, which writes the SVG and its PNG fallback into the
 * part that holds it; Graphics Fill reaches a floating picture run and reads
 * the colour a graphic has.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import {
  buildBlankDocx,
  buildTextboxParagraphXml,
  parseDocx,
  saveDocx,
  type HeaderFooter,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import { buildDocx, TINY_PNG_BASE64 } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { setActiveSubEditor } from '../src/renderer/editor/active-editor'
import {
  blocksToPmDoc,
  inlineToRuns,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions, textboxSubExtensions } from '../src/renderer/editor/extensions'
import {
  graphicSelected,
  retargetGraphicXml,
  selectedGraphicFill,
  setSelectedGraphicFill,
  svgFillOf,
} from '../src/renderer/editor/graphics-fill'
import { mountHfEditor, type HfEditorHandle } from '../src/renderer/editor/hf-editor'
import { insertPreparedIcons, prepareLibraryIcons } from '../src/renderer/editor/insert-icons'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'

// jsdom draws nothing: each SVG gets a bitmap of its own, as a real drawing would
const drawn = vi.hoisted(() => ({ count: 0 }))
vi.mock('../src/renderer/editor/svg-raster', async (original) => ({
  ...(await original<typeof import('../src/renderer/editor/svg-raster')>()),
  rasteriseSvg: async () => {
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
    return `data:image/png;base64,${png}${btoa(`pad${++drawn.count}`)}`
  },
}))

const editors: Editor[] = []
const handles: HfEditorHandle[] = []
afterEach(() => {
  setActiveSubEditor(null)
  for (const handle of handles.splice(0)) if (!handle.editor.isDestroyed) handle.discard()
  for (const editor of editors.splice(0)) editor.destroy()
})

async function openBlank() {
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.push(editor)
  return { editor, parsed }
}

const text = (zip: JSZip, name: string) => zip.file(name)!.async('string')

/** the media a part's drawing names: its bitmap and its SVG, through the part's own relationships */
async function mediaOf(zip: JSZip, part: string, drawing: string) {
  const rels = await text(zip, `word/_rels/${part}.rels`)
  const target = (rId: string | undefined) =>
    rId ? (new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ?? null) : null
  return {
    png: target(/<a:blip r:embed="([^"]+)"/.exec(drawing)?.[1]),
    svg: target(/<asvg:svgBlip [^>]*r:embed="([^"]+)"/.exec(drawing)?.[1]),
  }
}

const svgOf = async (zip: JSZip, target: string | null) =>
  target ? text(zip, `word/${target}`) : ''

describe('the fill of a graphic', () => {
  it('is the root fill, else the first shape that names one, else black', () => {
    expect(svgFillOf('<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0"/></svg>')).toBe(
      '000000',
    )
    expect(
      svgFillOf('<svg xmlns="http://www.w3.org/2000/svg" fill="#ff0000"><path d="M0 0"/></svg>'),
    ).toBe('FF0000')
    expect(
      svgFillOf('<svg xmlns="http://www.w3.org/2000/svg"><path fill="#0a0" d="M0 0"/></svg>'),
    ).toBe('00AA00')
    expect(
      svgFillOf(
        '<svg xmlns="http://www.w3.org/2000/svg"><g fill="none"/><path style="stroke:#000;fill:#123456" d="M0 0"/></svg>',
      ),
    ).toBe('123456')
  })

  it('is none for a root without fill, and unknown for a paint that is no plain colour', () => {
    expect(
      svgFillOf('<svg xmlns="http://www.w3.org/2000/svg" fill="none"><path d="M0 0"/></svg>'),
    ).toBeNull()
    expect(
      svgFillOf('<svg xmlns="http://www.w3.org/2000/svg" fill="url(#g)"><path d="M0 0"/></svg>'),
    ).toBeUndefined()
    expect(svgFillOf('<html/>')).toBeUndefined()
  })
})

describe('an icon put into a header', () => {
  it('saves into the header part with its SVG and its PNG fallback', async () => {
    const { editor, parsed } = await openBlank()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    let committed: HeaderFooter | null = null
    const host = document.createElement('div')
    document.body.append(host)
    const handle = mountHfEditor(host, {
      value: { text: '' },
      pageNo: '1',
      pageTotal: '1',
      spellcheck: false,
      onCommit: (next) => (committed = next),
      onExit: () => host.remove(),
    })
    handles.push(handle)
    handle.editor.commands.focus('end')
    setActiveSubEditor(handle.editor)

    await prepareLibraryIcons(['heart', 'rocket'])
    expect(insertPreparedIcons(editor, ['heart', 'rocket'])).toBe(true)
    // the header took them; the body did not
    const inHeader: PmNode[] = []
    handle.editor.state.doc.descendants((node) => {
      if (node.type.name === 'docInlineImage') inHeader.push(node.toJSON() as PmNode)
    })
    expect(inHeader).toHaveLength(2)
    expect(JSON.stringify(editor.getJSON())).not.toContain('docInlineImage')
    expect(inHeader[0].attrs?.vector).toMatchObject({ for: inHeader[0].attrs?.dataUrl })

    handle.exit()
    const header = committed as HeaderFooter | null
    expect(Object.values(header?.newPictures ?? {})).toHaveLength(2)
    for (const picture of Object.values(header!.newPictures!))
      expect(atob(picture.svgBase64 ?? '')).toMatch(/^<svg /)
    expect(header!.paras![0].runs.map((run) => !!run.image?.vectorUrl)).toEqual([true, true])

    const saved = await saveDocx(parsed, plan.saveBlocks, { header: header! })
    const zip = await JSZip.loadAsync(saved)
    const xml = await text(zip, 'word/header1.xml')
    expect(xml).not.toContain('aidocs-new-pic')
    const drawings = [...xml.matchAll(/<w:drawing>.*?<\/w:drawing>/g)].map(([drawing]) => drawing)
    expect(drawings).toHaveLength(2)
    expect(drawings.map((d) => /descr="([^"]*)"/.exec(d)?.[1])).toEqual(['Heart', 'Rocket'])
    const media = await Promise.all(drawings.map((d) => mediaOf(zip, 'header1.xml', d)))
    for (const { png, svg } of media) {
      expect(png).toMatch(/^media\/[^/]+\.png$/)
      expect(svg).toMatch(/^media\/[^/]+\.svg$/)
      expect(zip.file(`word/${png}`)).toBeTruthy()
      expect(await svgOf(zip, svg)).toMatch(/^<svg /)
    }
    expect(new Set(media.flatMap((m) => [m.png, m.svg])).size).toBe(4)
    expect(new Set(drawings.map((d) => /<wp:docPr id="(\d+)"/.exec(d)?.[1])).size).toBe(2)
    expect(await text(zip, '[Content_Types].xml')).toContain(
      '<Default Extension="svg" ContentType="image/svg+xml"/>',
    )
    // the document's own relationships gained no picture
    expect(await text(zip, 'word/_rels/document.xml.rels')).not.toContain('media/')

    const reparsed = await parseDocx(saved)
    // a header paragraph of pictures alone reads back as the header's pictures
    expect(reparsed.headerImages?.map((image) => [image.widthPx, image.heightPx])).toEqual([
      [96, 96],
      [96, 96],
    ])
  })
})

describe('an icon put into a text box', () => {
  it('keeps its bytes for the save, apart from a cell picture numbered by the same save', async () => {
    const { editor, parsed } = await openBlank()
    // a picture in a table cell, which this save numbers from 1
    insertTableAt(editor, 1, 1)
    let cell = -1
    editor.state.doc.descendants((n, pos) => {
      if (cell < 0 && n.type.name === 'docTableCell') cell = pos + 2
    })
    editor.commands.setTextSelection(cell)
    await prepareLibraryIcons(['heart'])
    expect(insertPreparedIcons(editor, ['heart'])).toBe(true)

    const box: TextboxDisplay = {
      fill: 'FFFFFF',
      borderColor: '000000',
      widthPx: 189,
      heightPx: 113,
      paras: [{ runs: [{ text: '' }] }],
    }
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'Text box',
          genXml: buildTextboxParagraphXml({ widthEmu: 1800000, heightEmu: 1080000, id: 3 }),
          textboxes: [box],
        },
      })
      .run()

    // the box's own editor, as the page mounts it
    const sub = new Editor({
      element: document.createElement('div'),
      extensions: textboxSubExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    })
    editors.push(sub)
    sub.commands.focus('end')
    setActiveSubEditor(sub)
    await prepareLibraryIcons(['star'])
    expect(insertPreparedIcons(editor, ['star'])).toBe(true)
    const paras = ((sub.getJSON().content ?? []) as PmNode[]).map((p) => ({
      runs: inlineToRuns(p.content ?? []),
    }))
    const held = /r:embed="(aidocs-new-pic-\d+)"/.exec(paras[0].runs[0].image!.xml)?.[1]
    expect(held).toBeTruthy()
    // a second conversion of the same content names the same placeholder
    expect(
      inlineToRuns(((sub.getJSON().content ?? []) as PmNode[])[0].content ?? [])[0].image!.xml,
    ).toBe(paras[0].runs[0].image!.xml)
    setActiveSubEditor(null)

    let boxPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docProtected' && node.attrs.textboxes) boxPos = pos
    })
    const node = editor.state.doc.nodeAt(boxPos)!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(boxPos, undefined, {
        ...node.attrs,
        textboxes: [{ ...box, paras }],
      }),
    )

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(Object.keys(plan.newPictures).sort()).toEqual(['aidocs-new-pic-1', held].sort())
    const saved = await saveDocx(parsed, plan.saveBlocks, { newPictures: plan.newPictures })
    const zip = await JSZip.loadAsync(saved)
    const xml = await text(zip, 'word/document.xml')
    expect(xml).not.toContain('aidocs-new-pic')
    const inBox = /<w:txbxContent>.*?<\/w:txbxContent>/.exec(xml)?.[0] ?? ''
    const boxDrawing = /<w:drawing>.*?<\/w:drawing>/.exec(inBox)?.[0] ?? ''
    expect(boxDrawing).toContain('descr="Star"')
    const cellDrawing = /<w:tc>.*?(<w:drawing>.*?<\/w:drawing>).*?<\/w:tc>/.exec(xml)?.[1] ?? ''
    expect(cellDrawing).toContain('descr="Heart"')
    const inBoxMedia = await mediaOf(zip, 'document.xml', boxDrawing)
    const inCellMedia = await mediaOf(zip, 'document.xml', cellDrawing)
    expect(inBoxMedia.svg).toMatch(/\.svg$/)
    expect(inCellMedia.svg).toMatch(/\.svg$/)
    expect(inBoxMedia.svg).not.toBe(inCellMedia.svg)
    expect(inBoxMedia.png).not.toBe(inCellMedia.png)
    expect(await svgOf(zip, inBoxMedia.svg)).not.toBe(await svgOf(zip, inCellMedia.svg))

    const reparsed = await parseDocx(saved)
    const boxes = reparsed.blocks.flatMap((block) => block.textboxes ?? [])
    expect(boxes[0].paras[0].runs.some((run) => run.image?.vectorUrl)).toBe(true)
  })
})

const ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24"><path d="M2 2h20v20H2z"/></svg>'

/** an icon as Word writes it floating beside the text of its paragraph */
const anchoredIcon =
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
  'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>476250</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="914400"/><wp:wrapSquare wrapText="bothSides"/>' +
  '<wp:docPr id="1" name="Graphic 1" descr="Square"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Graphic 1"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rId10"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
  '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId11"/>' +
  '</a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing>'

describe('Graphics Fill on a floating picture', () => {
  it('points a drawing at new bytes and leaves the rest of it alone', () => {
    const out = retargetGraphicXml(anchoredIcon, 'aidocs-new-pic-7')!
    expect(out).toContain('<a:blip r:embed="aidocs-new-pic-7">')
    expect(out).toContain('r:embed="aidocs-new-pic-7-svg"/>')
    expect(out.replace(/aidocs-new-pic-7-svg/, 'rId11').replace(/aidocs-new-pic-7/, 'rId10')).toBe(
      anchoredIcon,
    )
    // a picture without an SVG, or more than one picture, is not a graphic to refill
    expect(retargetGraphicXml(anchoredIcon.replace(/<a:extLst>.*<\/a:extLst>/, ''), 't')).toBeNull()
    expect(retargetGraphicXml(anchoredIcon + anchoredIcon, 't')).toBeNull()
  })

  it('recolours the picture where it floats, shows the colour, and saves the SVG and its fallback', async () => {
    const source = await buildDocx({
      bodyXml:
        `<w:p><w:r><w:t>Text beside the icon.</w:t></w:r><w:r>${anchoredIcon}</w:r></w:p>` +
        '<w:p><w:r><w:t>Next paragraph.</w:t></w:r></w:p>',
      withImage: true,
      extraRels:
        '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.svg"/>',
      binaryParts: [
        {
          path: 'word/media/image2.svg',
          base64: btoa(ICON_SVG),
          extension: 'svg',
          contentType: 'image/svg+xml',
        },
      ],
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.push(editor)
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docInlineImage') at = pos
    })
    expect(editor.state.doc.nodeAt(at)?.attrs.wrap).toBeTruthy()
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, at)))
    expect(graphicSelected(editor)).toBe(true)
    expect(await selectedGraphicFill(editor)).toBe('000000')

    expect(await setSelectedGraphicFill(editor, 'FF0000')).toBe(true)
    expect(await selectedGraphicFill(editor)).toBe('FF0000')
    const after = editor.state.doc.nodeAt(at)!
    expect(after.attrs.wrap).toBe(editor.state.doc.nodeAt(at)!.attrs.wrap)
    expect(after.attrs.xml).toContain('<wp:wrapSquare wrapText="bothSides"/>')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(Object.keys(plan.newPictures)).toHaveLength(1)
    const saved = await saveDocx(parsed, plan.saveBlocks, { newPictures: plan.newPictures })
    const zip = await JSZip.loadAsync(saved)
    const xml = await text(zip, 'word/document.xml')
    expect(xml).not.toContain('aidocs-new-pic')
    const drawing = /<w:drawing>.*?<\/w:drawing>/.exec(xml)?.[0] ?? ''
    // still the anchored drawing the file had, at its place
    expect(drawing).toContain('<wp:posOffset>476250</wp:posOffset>')
    expect(drawing).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(drawing).toContain('descr="Square"')
    const media = await mediaOf(zip, 'document.xml', drawing)
    expect(media.png).not.toBe('media/image1.png')
    expect(await svgOf(zip, media.svg)).toContain('fill="#FF0000"')
    const png = await zip.file(`word/${media.png}`)!.async('base64')
    expect(png.startsWith(TINY_PNG_BASE64.slice(0, 40))).toBe(true)

    // no fill, then read back from the saved file
    expect(await setSelectedGraphicFill(editor, null)).toBe(true)
    expect(await selectedGraphicFill(editor)).toBeNull()
    const reparsed = await parseDocx(saved)
    const run = reparsed.blocks
      .flatMap((block) => ('runs' in block ? (block.runs ?? []) : []))
      .find((r) => r.image)
    expect(run?.image?.wrap).toBeTruthy()
    expect(run?.image?.vectorUrl).toBeTruthy()
  })
})
