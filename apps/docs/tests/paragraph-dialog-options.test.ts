/**
 * Paragraph dialog options beyond the pagination flags: Mirror indents, Don't
 * hyphenate and Snap to grid round-trip as pPr flags, mirrored indents trade
 * sides on even pages, and Set As Default hands the default paragraph style
 * only what the paragraph declares or the dialog changed.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateParagraphXml, mergePPrFormat, parseDocx } from '@genoffice/docx-engine'
import { upsertStyleXml } from '../../../packages/docx-engine/src/style-upsert'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmNodeToGeneratedBlock,
  signatureOfBlock,
  signatureOfGenerated,
} from '../src/renderer/editor/convert'
import { fromPmNode } from '../src/renderer/editor/para-layout-spec'
import { clearParagraphFormatting, setParaAttrs } from '../src/renderer/components/ribbon-tabs'
import { paragraphDefaultProps } from '../src/renderer/components/paragraph-default-props'
import { mirrorIndentSpecs, type BlockBox, type PageSlice } from '../src/renderer/pagination'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

async function openDoc(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return { editor, parsed }
}

const CTX = { headingStyleIds: new Map<number, string>(), allocateHyperlinkRel: () => 'rId1' }
const xmlOf = (editor: Editor, index = 0) =>
  generateParagraphXml(pmNodeToGeneratedBlock(editor.state.doc.child(index).toJSON()), CTX)

describe('mirror indents / hyphenation / grid flags: parse -> node attrs', () => {
  it('direct flags land on the node, absent ones stay null', async () => {
    const { editor } = await openDoc(
      p(
        '<w:suppressAutoHyphens/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/><w:mirrorIndents/>',
        'a',
      ) +
        p('<w:suppressAutoHyphens w:val="0"/>', 'b') +
        p('', 'c'),
    )
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      suppressAutoHyphens: true,
      mirrorIndents: true,
      snapToGrid: false,
      adjustRightInd: false,
    })
    expect(editor.state.doc.child(1).attrs).toMatchObject({
      suppressAutoHyphens: false,
      mirrorIndents: null,
    })
    expect(editor.state.doc.child(2).attrs).toMatchObject({
      suppressAutoHyphens: null,
      mirrorIndents: null,
      snapToGrid: null,
      adjustRightInd: null,
    })
  })

  it('an untouched paragraph signs the same on both sides', async () => {
    const { editor, parsed } = await openDoc(
      p(
        '<w:suppressAutoHyphens/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/><w:mirrorIndents/>',
        'a',
      ),
    )
    expect(signatureOfGenerated(pmNodeToGeneratedBlock(editor.state.doc.child(0).toJSON()))).toBe(
      signatureOfBlock(parsed.blocks[0]),
    )
  })
})

describe('mirror indents / hyphenation / grid flags: node attrs -> saved pPr', () => {
  it('writes the flags in schema order', async () => {
    const { editor } = await openDoc(p('', 'body'))
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, {
      suppressAutoHyphens: true,
      mirrorIndents: true,
      snapToGrid: false,
      indentLeft: 720,
    })
    const xml = xmlOf(editor)
    expect(xml).toContain('<w:suppressAutoHyphens/>')
    expect(xml).toContain('<w:snapToGrid w:val="0"/>')
    expect(xml).toContain('<w:mirrorIndents/>')
    const order = ['<w:suppressAutoHyphens/>', '<w:snapToGrid', '<w:ind ', '<w:mirrorIndents/>']
    const at = order.map((tag) => xml.indexOf(tag))
    expect(at).toEqual([...at].sort((a, b) => a - b))
  })

  it('an untouched paragraph writes none of them', async () => {
    const { editor } = await openDoc(p('', 'body'))
    expect(xmlOf(editor)).not.toMatch(/suppressAutoHyphens|snapToGrid|mirrorIndents|adjustRightInd/)
  })

  it('clearing direct formatting drops them', async () => {
    const { editor } = await openDoc(
      p('<w:suppressAutoHyphens/><w:snapToGrid w:val="0"/><w:mirrorIndents/>', 'body'),
    )
    editor.commands.setTextSelection(2)
    clearParagraphFormatting(editor)
    expect(editor.state.doc.child(0).attrs).toMatchObject({
      suppressAutoHyphens: null,
      mirrorIndents: null,
      snapToGrid: null,
    })
  })
})

describe('mergePPrFormat keeps what was not edited', () => {
  const RAW =
    '<w:pPr><w:suppressAutoHyphens/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/><w:mirrorIndents/></w:pPr>'
  const SAME = { suppressAutoHyphens: true, snapToGrid: false, mirrorIndents: true }

  it('same values keep the raw bytes', () => {
    expect(mergePPrFormat(RAW, SAME, SAME)).toBe(RAW)
  })

  it('a flag back on the style is removed, the others stay', () => {
    expect(mergePPrFormat(RAW, { snapToGrid: false, mirrorIndents: true }, SAME)).toBe(
      '<w:pPr><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/><w:mirrorIndents/></w:pPr>',
    )
    expect(mergePPrFormat(RAW, { suppressAutoHyphens: true, mirrorIndents: true }, SAME)).toBe(
      '<w:pPr><w:suppressAutoHyphens/><w:adjustRightInd w:val="0"/><w:mirrorIndents/></w:pPr>',
    )
  })

  it('an explicit off overriding the style is written as w:val="0"', () => {
    expect(
      mergePPrFormat('<w:pPr><w:pStyle w:val="A"/></w:pPr>', { suppressAutoHyphens: false }),
    ).toBe('<w:pPr><w:pStyle w:val="A"/><w:suppressAutoHyphens w:val="0"/></w:pPr>')
  })
})

describe('page effect', () => {
  const specOf = (attrs: Record<string, unknown>) =>
    fromPmNode({ attrs: { docxIndex: null, ...attrs }, content: [] } as never)

  it("Don't hyphenate turns hyphenation off on the paragraph, an explicit off opts back in", () => {
    expect(JSON.stringify(specOf({ suppressAutoHyphens: true }))).toContain('hyphens:manual')
    expect(JSON.stringify(specOf({ suppressAutoHyphens: false }))).toContain('doc-hyphens-on')
    expect(JSON.stringify(specOf({}))).not.toMatch(/hyphens/)
  })

  it('a mirrored paragraph is marked for the layout pass', () => {
    expect(JSON.stringify(specOf({ mirrorIndents: true }))).toContain('doc-mirror-ind')
    expect(JSON.stringify(specOf({}))).not.toContain('doc-mirror-ind')
  })

  const slices: PageSlice[] = [
    { start: 0, end: 500, section: 0 },
    { start: 500, end: 1000, section: 0 },
    { start: 1000, end: 1500, section: 0 },
  ]
  const block = (top: number, mirrored: boolean): BlockBox => {
    const el = document.createElement('p')
    if (mirrored) el.className = 'doc-mirror-ind'
    return { el, top, height: 100 } as BlockBox
  }

  it('on even pages a mirrored paragraph moves by right minus left', () => {
    const blocks = [block(0, true), block(500, true), block(700, false), block(1000, true)]
    const specs = mirrorIndentSpecs(blocks, slices, [1, 2, 3], () => ({ left: 96, right: 24 }))
    expect(specs).toEqual([{ el: blocks[1].el, dx: -72, dy: 0, mirrorDx: -72 }])
  })

  it('follows displayed page numbers and skips equal indents', () => {
    const blocks = [block(0, true), block(500, true)]
    expect(
      mirrorIndentSpecs(blocks, slices, [2, 3, 4], () => ({ left: 0, right: 48 })).map((s) => s.el),
    ).toEqual([blocks[0].el])
    expect(mirrorIndentSpecs(blocks, slices, [2, 3, 4], () => ({ left: 48, right: 48 }))).toEqual(
      [],
    )
  })
})

describe('Set As Default', () => {
  const base = {
    attrs: {},
    align: 'left' as const,
    line: { lineSpacing: null, lineRule: null, lineRawTwips: null },
    indentLeft: 0,
    indentRight: 0,
    firstLine: null,
    spaceBefore: 0,
    spaceAfter: 0,
    flags: { widowControl: true, keepNext: false, suppressAutoHyphens: false },
    initialFlags: { widowControl: true, keepNext: false, suppressAutoHyphens: false },
  }

  it('an untouched dialog on a plain paragraph changes nothing in the style', () => {
    expect(paragraphDefaultProps(base)).toEqual({})
  })

  it('takes what the dialog changed', () => {
    expect(
      paragraphDefaultProps({
        ...base,
        align: 'justify',
        line: { lineSpacing: 1.5, lineRule: null, lineRawTwips: null },
        indentLeft: 720,
        firstLine: -360,
        spaceAfter: 240,
        flags: { ...base.flags, suppressAutoHyphens: true },
      }),
    ).toEqual({
      align: 'justify',
      lineSpacing: 1.5,
      indentLeftTwips: 720,
      firstLineTwips: -360,
      spaceAfterTwips: 240,
      suppressAutoHyphens: true,
    })
  })

  it("takes the paragraph's own values, a cleared one as an explicit zero", () => {
    expect(
      paragraphDefaultProps({
        ...base,
        attrs: { spaceAfter: 160, lineRule: 'exact', lineRawTwips: 360, keepNext: true },
        line: { lineSpacing: null, lineRule: 'exact', lineRawTwips: 360 },
        flags: { ...base.flags, keepNext: true },
        initialFlags: { ...base.initialFlags, keepNext: true },
      }),
    ).toEqual({
      spaceAfterTwips: 0,
      lineFixed: { rule: 'exact', twips: 360 },
      keepNext: true,
    })
  })

  it('the style definition takes the new properties', () => {
    const styles =
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
      '<w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:style></w:styles>'
    const out = upsertStyleXml(styles, {
      styleId: 'Normal',
      pPr: {
        lineFixed: { rule: 'exact', twips: 360 },
        widowControl: false,
        suppressAutoHyphens: true,
        contextualSpacing: true,
      },
    })
    expect(out).toContain(
      '<w:pPr><w:widowControl w:val="0"/><w:suppressAutoHyphens/>' +
        '<w:spacing w:after="160" w:line="360" w:lineRule="exact"/><w:contextualSpacing/></w:pPr>',
    )
  })
})
