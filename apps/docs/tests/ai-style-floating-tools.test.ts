import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { PictureWatermarkSpec, StyleUpsert, WatermarkSpec } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { executeTool, type AiDocExtras } from '../src/renderer/ai/tools'
import type { AiStyleInfo } from '../src/renderer/ai/style-ops'
import { pictureNode, textBoxNode } from '../src/renderer/ai/floating-ops'
import { MAX_FONT_SIZE_PT, MIN_FONT_SIZE_PT } from '../src/renderer/ai/ops'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const heading = (t: string, level = 1): JsonNode => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(t)],
})
const para = (t: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: [text(t)],
})

const liveEditors: Editor[] = []
function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

const NUM_IDS = { bullet: null, ordered: null }

function catalog(): { styles: AiStyleInfo[]; upserts: StyleUpsert[]; extras: AiDocExtras } {
  const styles: AiStyleInfo[] = [
    { styleId: 'Normal', name: 'Normal', type: 'paragraph' },
    {
      styleId: 'Heading1',
      name: 'heading 1',
      type: 'paragraph',
      headingLevel: 1,
      basedOn: 'Normal',
    },
    {
      styleId: 'Heading2',
      name: 'heading 2',
      type: 'paragraph',
      headingLevel: 2,
      basedOn: 'Normal',
    },
    { styleId: 'Quote', name: 'Quote', type: 'paragraph', basedOn: 'Normal' },
    { styleId: 'Strong', name: 'Strong', type: 'character' },
  ]
  const upserts: StyleUpsert[] = []
  const extras: AiDocExtras = {
    styles: {
      list: () => [
        ...styles,
        ...upserts
          .filter((u) => !styles.some((s) => s.styleId === u.styleId))
          .map((u) => ({
            styleId: u.styleId,
            name: u.name ?? u.styleId,
            type: u.type ?? ('paragraph' as const),
            ...(u.pPr?.outlineLevel ? { headingLevel: u.pPr.outlineLevel } : {}),
            pending: true,
          })),
      ],
      upsert: (up) => {
        upserts.push(up)
        return null
      },
    },
  }
  return { styles, upserts, extras }
}

const call = (name: string, input: Record<string, unknown>) => ({ id: 't', name, input })

async function run(
  editor: Editor,
  name: string,
  input: Record<string, unknown>,
  extras?: AiDocExtras,
) {
  return executeTool(
    editor,
    call(name, input),
    NUM_IDS,
    undefined,
    undefined,
    null,
    undefined,
    undefined,
    undefined,
    undefined,
    extras,
  )
}

describe('define_style', () => {
  it('validates and hands the engine upsert to the catalog', async () => {
    const editor = createEditor([para('x')])
    const { upserts, extras } = catalog()
    const exec = await run(
      editor,
      'define_style',
      {
        styleId: 'Callout',
        basedOn: 'Normal',
        paragraph: {
          spaceBefore: 6,
          spaceAfter: '0.5cm',
          indentLeft: '1in',
          firstLineIndent: -18,
          align: 'justify',
        },
        run: { italic: true, color: '#1A73E8', fontSize: 11, fontFamily: 'Georgia' },
      },
      extras,
    )
    expect(exec.isError).toBeUndefined()
    expect(exec.output).toContain('Created style Callout')
    expect(upserts[0]).toEqual({
      styleId: 'Callout',
      basedOn: 'Normal',
      pPr: {
        spaceBeforeTwips: 120,
        spaceAfterTwips: 283,
        indentLeftTwips: 1440,
        firstLineTwips: -360,
        align: 'justify',
      },
      rPr: { italic: true, color: '1A73E8', sizeHalfPoints: 22, font: 'Georgia' },
    })

    const again = await run(
      editor,
      'define_style',
      { styleId: 'Heading1', run: { color: '#FF0000' } },
      extras,
    )
    expect(again.output).toContain('Updated style Heading1')
    expect(upserts[1]).toEqual({ styleId: 'Heading1', rPr: { color: 'FF0000' } })
  })

  it('refuses unknown parents, type changes and empty definitions', async () => {
    const editor = createEditor([para('x')])
    const { extras } = catalog()
    const parent = await run(
      editor,
      'define_style',
      { styleId: 'New', basedOn: 'Nope', run: { bold: true } },
      extras,
    )
    expect(parent.isError).toBe(true)
    expect(parent.output).toContain('available paragraph styles: Normal, Heading1')
    const type = await run(
      editor,
      'define_style',
      { styleId: 'Strong', type: 'paragraph', run: { bold: true } },
      extras,
    )
    expect(type.output).toContain('is a character style')
    const empty = await run(editor, 'define_style', { styleId: 'Empty' }, extras)
    expect(empty.output).toContain('needs paragraph or run formatting')
    const charPara = await run(
      editor,
      'define_style',
      { styleId: 'Strong', paragraph: { align: 'center' } },
      extras,
    )
    expect(charPara.output).toContain('character style has no paragraph formatting')
    const bad = await run(
      editor,
      'define_style',
      { styleId: 'bad id!', run: { bold: true } },
      extras,
    )
    expect(bad.output).toContain('styleId must be')
  })
})

describe('applyStyle op', () => {
  it('switches block types with heading styles and keeps the styleId on save', async () => {
    const editor = createEditor([
      heading('Title', 1),
      para('Body one'),
      para('Body two'),
      heading('Sub', 2),
    ])
    const { extras } = catalog()
    const exec = await run(
      editor,
      'apply_ops',
      {
        ops: [
          { op: 'applyStyle', target: { blockIndexes: [1] }, styleId: 'Heading2' },
          { op: 'applyStyle', target: { blockIndexes: [3] }, styleId: 'Quote' },
          { op: 'applyStyle', target: { containsText: 'two' }, styleId: 'Quote' },
        ],
      },
      extras,
    )
    expect(exec.isError).toBeUndefined()
    const doc = editor.getJSON() as PmNode
    expect(doc.content![1]).toMatchObject({
      type: 'docHeading',
      attrs: { level: 2, styleId: 'Heading2' },
    })
    expect(doc.content![2]).toMatchObject({ type: 'docParagraph', attrs: { styleId: 'Quote' } })
    expect(doc.content![3]).toMatchObject({ type: 'docParagraph', attrs: { styleId: 'Quote' } })
    const plan = pmDocToSavePlan(doc, [])
    const styleIds = plan.saveBlocks.map((b) => (b.kind === 'generated' ? b.block.styleId : null))
    expect(styleIds).toEqual([undefined, 'Heading2', 'Quote', 'Quote'])
  })

  it('keeps table cell paragraphs as paragraphs under a heading style', async () => {
    const cell = (t: string): JsonNode => ({ type: 'docTableCell', content: [para(t)] })
    const editor = createEditor([
      para('Intro'),
      {
        type: 'docTable',
        content: [{ type: 'docTableRow', content: [cell('A1'), cell('B1')] }],
      },
    ])
    const { extras } = catalog()
    const exec = await run(
      editor,
      'apply_ops',
      { ops: [{ op: 'applyStyle', target: { blockIndexes: [0, 1] }, styleId: 'Heading2' }] },
      extras,
    )
    expect(exec.isError).toBeUndefined()
    expect(() => editor.state.doc.check()).not.toThrow()
    const doc = editor.getJSON() as PmNode
    expect(doc.content![0]).toMatchObject({ type: 'docHeading', attrs: { level: 2 } })
    const cells = doc.content![1].content![0].content!
    for (const c of cells)
      expect(c.content![0]).toMatchObject({ type: 'docParagraph', attrs: { styleId: 'Heading2' } })

    const again = await run(
      editor,
      'apply_ops',
      { ops: [{ op: 'applyStyle', target: { blockIndexes: [1] }, styleId: 'Heading2' }] },
      extras,
    )
    expect(again.isError).toBeUndefined()
    expect(editor.getJSON()).toEqual(doc)
  })

  it('accepts a pending define_style id and rejects unknown or character styles atomically', async () => {
    const editor = createEditor([para('a'), para('b')])
    const { extras } = catalog()
    await run(editor, 'define_style', { styleId: 'Callout', run: { italic: true } }, extras)
    const ok = await run(
      editor,
      'apply_ops',
      { ops: [{ op: 'applyStyle', target: { blockIndexes: [0] }, styleId: 'Callout' }] },
      extras,
    )
    expect(ok.isError).toBeUndefined()
    expect((editor.getJSON() as PmNode).content![0].attrs?.styleId).toBe('Callout')

    const unknown = await run(
      editor,
      'apply_ops',
      {
        ops: [
          { op: 'setFont', target: { blockIndexes: [1] }, bold: true },
          { op: 'applyStyle', target: { blockIndexes: [1] }, styleId: 'Missing' },
        ],
      },
      extras,
    )
    expect(unknown.isError).toBe(true)
    expect(unknown.output).toContain('no style "Missing"')
    expect((editor.getJSON() as PmNode).content![1].content![0].marks).toBeUndefined()

    const character = await run(
      editor,
      'apply_ops',
      { ops: [{ op: 'applyStyle', target: { blockIndexes: [1] }, styleId: 'Strong' }] },
      extras,
    )
    expect(character.output).toContain('is a character style')
  })
})

const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

describe('set_watermark', () => {
  it('resolves the spec and removes on null', async () => {
    const editor = createEditor([para('x')])
    const seen: Array<WatermarkSpec | PictureWatermarkSpec | null> = []
    const extras: AiDocExtras = {
      watermark: {
        current: () => null,
        set: (spec) => {
          seen.push(spec)
          return null
        },
      },
    }
    const set = await run(
      editor,
      'set_watermark',
      { text: 'DRAFT', color: '#FF0000', opacity: 0.3, diagonal: false, bold: true },
      extras,
    )
    expect(set.isError).toBeUndefined()
    expect(seen[0]).toEqual({
      text: 'DRAFT',
      colorHex: 'FF0000',
      opacity: 0.3,
      diagonal: false,
      bold: true,
    })
    const removed = await run(editor, 'set_watermark', { text: null }, extras)
    expect(removed.output).toBe('Watermark removed.')
    expect(seen[1]).toBeNull()
    const bad = await run(editor, 'set_watermark', { text: 'x', color: 'red' }, extras)
    expect(bad.isError).toBe(true)
    const missing = await run(editor, 'set_watermark', { text: 'x' })
    expect(missing.output).toContain('not available here')
  })

  it('image sets a picture watermark from a data URL; image null removes; text and image together are rejected', async () => {
    // jsdom never decodes images: stand in for <img> with the fixture's size
    const RealImage = globalThis.Image
    class FakeImage {
      naturalWidth = 400
      naturalHeight = 200
      onload: (() => void) | null = null
      onerror: (() => void) | null = null
      set src(_v: string) {
        queueMicrotask(() => this.onload?.())
      }
    }
    globalThis.Image = FakeImage as unknown as typeof Image
    try {
      const editor = createEditor([para('x')])
      const seen: Array<WatermarkSpec | PictureWatermarkSpec | null> = []
      const extras: AiDocExtras = {
        watermark: { current: () => null, set: (spec) => (seen.push(spec), null) },
      }
      const set = await run(
        editor,
        'set_watermark',
        { image: `data:image/png;base64,${PNG}`, scale: 40, washout: false },
        extras,
      )
      expect(set.isError).toBeUndefined()
      expect(set.output).toContain('Picture watermark set (400x200px source)')
      expect(seen[0]).toEqual({
        image: { base64: PNG, mime: 'image/png', widthPx: 400, heightPx: 200 },
        scale: 40,
        washout: false,
      })
      const both = await run(
        editor,
        'set_watermark',
        { text: 'DRAFT', image: `data:image/png;base64,${PNG}` },
        extras,
      )
      expect(both.isError).toBe(true)
      expect(both.output).toContain('not both')
      const badUrl = await run(editor, 'set_watermark', { image: 'file:///tmp/x.png' }, extras)
      expect(badUrl.isError).toBe(true)
      const removed = await run(editor, 'set_watermark', { image: null }, extras)
      expect(removed.output).toBe('Watermark removed.')
      expect(seen[1]).toBeNull()
    } finally {
      globalThis.Image = RealImage
    }
  })
})

describe('insert_text_box', () => {
  it('anchors a floating wps text box after the block and reports the index shift', async () => {
    const editor = createEditor([para('a'), para('b')])
    const exec = await run(editor, 'insert_text_box', {
      afterBlockIndex: 0,
      text: 'Note\nSecond line',
      width: '5cm',
      height: '2cm',
      x: '1in',
      y: 12,
      fill: '#FFF2CC',
      borderColor: null,
      fontSize: 9,
      align: 'center',
    })
    expect(exec.isError).toBeUndefined()
    expect(exec.output).toContain('it is block 1')
    const doc = editor.getJSON() as PmNode
    expect(doc.content).toHaveLength(3)
    const box = doc.content![1]
    expect(box.type).toBe('docProtected')
    const xml = String(box.attrs?.genXml)
    expect(xml).toContain('<wp:anchor')
    expect(xml).toContain('<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset>')
    expect(xml).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>152400</wp:posOffset>',
    )
    expect(xml).toContain('<wp:extent cx="1800000" cy="720000"/>')
    expect(xml).toContain('<wp:wrapTopAndBottom/>')
    expect(xml).toContain('<a:srgbClr val="FFF2CC"/>')
    expect(xml).toContain('<a:ln><a:noFill/></a:ln>')
    expect(xml).toContain('<w:jc w:val="center"/>')
    expect(xml).toContain('<w:sz w:val="18"/>')
    // wps choice + VML fallback each carry the two paragraphs
    expect((xml.match(/<w:txbxContent>/g) ?? []).length).toBe(2)
    expect((xml.match(/<w:t xml:space="preserve">/g) ?? []).length).toBe(4)
    const display = (box.attrs?.textboxes as Array<Record<string, unknown>>)[0]
    expect(display).toMatchObject({
      fill: 'FFF2CC',
      widthPx: 189,
      heightPx: 76,
      offsetXEmu: 914400,
    })
    expect(display.floating).toBeUndefined()
    expect(display.bandBottomPx).toBe(16 + 76)
  })

  it('page anchors float free of the text and validate their inputs', async () => {
    const editor = createEditor([para('a')])
    const exec = await run(editor, 'insert_text_box', {
      text: 'x',
      width: 100,
      height: 50,
      x: 0,
      y: 0,
      anchor: 'page',
    })
    expect(exec.isError).toBeUndefined()
    const box = (editor.getJSON() as PmNode).content![1]
    expect(String(box.attrs?.genXml)).toContain('<wp:positionH relativeFrom="page">')
    expect(String(box.attrs?.genXml)).toContain('<wp:wrapNone/>')
    expect((box.attrs?.textboxes as Array<Record<string, unknown>>)[0]).toMatchObject({
      floating: true,
      pagePinned: true,
    })
    const bad = await run(editor, 'insert_text_box', {
      text: 'x',
      width: '5 furlongs',
      height: 50,
      x: 0,
      y: 0,
    })
    expect(bad.isError).toBe(true)
    const range = await run(editor, 'insert_text_box', {
      afterBlockIndex: 9,
      text: 'x',
      width: 10,
      height: 10,
      x: 0,
      y: 0,
    })
    expect(range.output).toContain('afterBlockIndex must be -1..1')
  })
})

describe('pictureNode', () => {
  const base = { base64: PNG, mime: 'image/png' as const, naturalWidth: 400, naturalHeight: 300 }

  it('keeps the aspect ratio and caps inline pictures at 480px', () => {
    const inline = pictureNode({ ...base, naturalWidth: 960, naturalHeight: 480 })
    expect(inline).toMatchObject({ widthPx: 480, heightPx: 240 })
    const sized = pictureNode({ ...base, width: '2in' })
    expect(sized).toMatchObject({ widthPx: 192, heightPx: 144 })
    if ('error' in sized) throw new Error(sized.error)
    expect(sized.node.attrs).toMatchObject({
      blockType: 'image',
      imageWidthPx: 192,
      imageHeightPx: 144,
    })
    expect(sized.node.attrs!.imageWrap).toBeUndefined()
    expect(pictureNode({ ...base, width: 'wide' })).toHaveProperty('error')
  })

  it('floating pictures carry the anchor into genImage and the display attrs', () => {
    const r = pictureNode({
      ...base,
      width: '96px',
      float: { anchor: 'page', xEmu: 914400, yEmu: 457200, wrap: 'behind' },
      altText: 'Logo',
    })
    if ('error' in r) throw new Error(r.error)
    expect(r.node.attrs).toMatchObject({
      imageWrap: 'behind',
      imageOffsetXEmu: 914400,
      imageOffsetYEmu: 457200,
      imageRelV: 'page',
    })
    expect(r.node.attrs!.genImage).toMatchObject({
      widthPx: 96,
      heightPx: 72,
      wrap: 'behind',
      posOffsetEmu: { x: 914400, y: 457200, relativeTo: 'page' },
      altText: 'Logo',
    })
    const para = pictureNode({
      ...base,
      float: { anchor: 'paragraph', xEmu: 0, yEmu: 0, wrap: 'square' },
    })
    if ('error' in para) throw new Error(para.error)
    expect(para.node.attrs!.genImage).toMatchObject({
      wrap: 'square-left',
      posOffsetEmu: { x: 0, y: 0 },
    })
    expect(para.node.attrs!.imageRelV).toBeUndefined()
  })

  it('textBoxNode rejects empty text and unknown anchors', () => {
    expect(textBoxNode({ text: '  ', width: 10, height: 10, x: 0, y: 0 })).toHaveProperty('error')
    expect(
      textBoxNode({ text: 'x', width: 10, height: 10, x: 0, y: 0, anchor: 'margin' }),
    ).toHaveProperty('error')
  })

  it('font sizes are bounded in setFont and textBoxNode', async () => {
    const editor = createEditor([para('x')])
    const { extras } = catalog()
    try {
      const bad = await run(
        editor,
        'apply_ops',
        { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, fontSize: 1e9 }] },
        extras,
      )
      expect(bad.isError).toBe(true)
      expect(String(bad.output)).toContain('fontSize')
      const tiny = await run(
        editor,
        'apply_ops',
        { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, fontSize: 0.1 }] },
        extras,
      )
      expect(tiny.isError).toBe(true)
      const ok = await run(
        editor,
        'apply_ops',
        { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, fontSize: 18 }] },
        extras,
      )
      expect(ok.isError).toBeUndefined()
      const setFontSize = (fontSize: number) =>
        run(
          editor,
          'apply_ops',
          { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, fontSize }] },
          extras,
        )
      expect((await setFontSize(MAX_FONT_SIZE_PT)).isError).toBeUndefined()
      expect((await setFontSize(MIN_FONT_SIZE_PT)).isError).toBeUndefined()
      const over = await setFontSize(MAX_FONT_SIZE_PT + 1)
      expect(over.isError).toBe(true)
      expect(String(over.output)).toContain('1-1638pt')
      expect((await setFontSize(0.5)).isError).toBe(true)
    } finally {
      editor.destroy()
    }
    expect(
      textBoxNode({ text: 'x', width: 10, height: 10, x: 0, y: 0, fontSize: 1e9 }),
    ).toHaveProperty('error')
    expect(
      textBoxNode({ text: 'x', width: 10, height: 10, x: 0, y: 0, fontSize: 1639 }),
    ).toHaveProperty('error')
    expect(
      textBoxNode({ text: 'x', width: 10, height: 10, x: 0, y: 0, fontSize: 0.5 }),
    ).toHaveProperty('error')
    expect(
      textBoxNode({ text: 'x', width: 10, height: 10, x: 0, y: 0, fontSize: 1638 }),
    ).not.toHaveProperty('error')
  })
})
