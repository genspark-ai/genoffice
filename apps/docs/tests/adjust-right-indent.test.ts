/**
 * w:adjustRightInd ("Automatically adjust right indent when document grid is
 * defined"). Word probes 2026-11-22 (A4, 415.3pt of text, Normal 10.5pt):
 * below compatibilityMode 15 a paragraph with the setting on ends its lines
 * short of the right indent by what a whole number of character cells leaves
 * over in the column; a Word 2013 document lays out as if it were off.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { ParsedDocFull, SectionInfo, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { generateParagraphXml, mergePPrFormat, parseDocx } from '@genoffice/docx-engine'
import { upsertStyleXml } from '../../../packages/docx-engine/src/style-upsert'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { gridRightIndentPt } from '../src/renderer/char-grid'
import { docStyleCss, GRID_RIGHT_INDENT_RULE } from '../src/renderer/doc-style-css'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  pmNodeToGeneratedBlock,
  signatureOfBlock,
  signatureOfGenerated,
} from '../src/renderer/editor/convert'
import { endIndentDecl, fromPmNode } from '../src/renderer/editor/para-layout-spec'
import { setParaAttrs } from '../src/renderer/components/ribbon-tabs'
import { paragraphDefaultProps } from '../src/renderer/components/paragraph-default-props'
import {
  docRightIndentPt,
  sectionRightIndentPt,
  sectionRightIndentSpecs,
  type BlockBox,
} from '../src/renderer/pagination'
import { setOpenDocumentText } from '../src/renderer/text-layout/open-document'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
  setOpenDocumentText(null)
})

const TEXT_WIDTH = 415.3
const PITCH = 11

describe('the share a character grid adds to the right indent', () => {
  it('is what a whole number of cells leaves over in the column, less a hair', () => {
    expect(gridRightIndentPt(TEXT_WIDTH, PITCH)).toBeLessThan(TEXT_WIDTH - 37 * PITCH)
    expect(gridRightIndentPt(TEXT_WIDTH, PITCH)).toBe(8.28)
    // 37.75 cells of 11pt: Word ends a justified line 8.25pt before the margin
    expect(gridRightIndentPt(TEXT_WIDTH, PITCH)).toBeCloseTo(8.3, 1)
    // 39.55 cells of 10.5pt (measured 5.7), 40.52 cells of 10.25pt (measured 5.0)
    expect(gridRightIndentPt(TEXT_WIDTH, 10.5)).toBeCloseTo(5.8, 1)
    expect(gridRightIndentPt(TEXT_WIDTH, 10.25)).toBeCloseTo(5.3, 1)
    // one of two equal columns (measured 10.3), a 671.8pt landscape section (measured 0.7)
    expect(gridRightIndentPt(197.03, PITCH)).toBeCloseTo(10.03, 1)
    expect(gridRightIndentPt(671.8, PITCH)).toBeCloseTo(0.8, 1)
  })

  it('is nothing when the cells fill the column', () => {
    expect(gridRightIndentPt(420, 10.5)).toBe(0)
    expect(gridRightIndentPt(0, 10.5)).toBe(0)
    expect(gridRightIndentPt(420, 0)).toBe(0)
  })

  // characters per line Word fitted, setting on / off, for each right indent
  it.each([
    { name: 'no indent', indent: 0, size: PITCH, on: 37, off: 37 },
    { name: '0.4 cell', indent: 4.4, size: PITCH, on: 36, off: 37 },
    { name: '1 cell', indent: 11, size: PITCH, on: 36, off: 36 },
    { name: '2.5 cells', indent: 27.5, size: PITCH, on: 34, off: 35 },
    { name: '3 cells', indent: 33, size: PITCH, on: 34, off: 34 },
    { name: '16pt text', indent: 0, size: 16.5, on: 24, off: 25 },
  ])('fits the characters Word fits: $name', ({ indent, size, on, off }) => {
    const fit = (lineWidth: number) => Math.floor(lineWidth / size + 1e-6)
    expect(fit(TEXT_WIDTH - indent - gridRightIndentPt(TEXT_WIDTH, PITCH))).toBe(on)
    expect(fit(TEXT_WIDTH - indent)).toBe(off)
  })
})

const sec = (over: Partial<SectionInfo['settings']>): SectionInfo =>
  ({
    settings: {
      pageWidth: 11906,
      pageHeight: 16838,
      orientation: 'portrait',
      marginTop: 1440,
      marginRight: 1800,
      marginBottom: 1440,
      marginLeft: 1800,
      pageBorder: false,
      columns: 1,
      ...over,
    },
    startType: 'nextPage',
    firstBlockIndex: 0,
    lastBlockIndex: 0,
    sectPrXml: '',
    titlePg: false,
    headerRefs: {},
    footerRefs: {},
  }) as SectionInfo

const grid = (type: 'lines' | 'linesAndChars' | 'snapToChars', charSpace?: number) =>
  sec({ docGrid: { type, linePitch: 312, ...(charSpace !== undefined ? { charSpace } : {}) } })

describe('which sections adjust', () => {
  it('a character grid below compatibilityMode 15, of either kind', () => {
    expect(sectionRightIndentPt(grid('linesAndChars', 2048), 10.5, true)).toBeCloseTo(8.3, 1)
    expect(sectionRightIndentPt(grid('snapToChars', 2048), 10.5, true)).toBeCloseTo(8.3, 1)
    expect(sectionRightIndentPt(grid('linesAndChars'), 10.5, true)).toBeCloseTo(5.8, 1)
  })

  it('no Word 2013 document, and no section without a character grid', () => {
    expect(sectionRightIndentPt(grid('linesAndChars', 2048), 10.5, false)).toBe(0)
    expect(sectionRightIndentPt(grid('snapToChars', 2048), 10.5, false)).toBe(0)
    expect(sectionRightIndentPt(grid('lines'), 10.5, true)).toBe(0)
    expect(sectionRightIndentPt(sec({}), 10.5, true)).toBe(0)
  })

  it('counts the cells of one column', () => {
    const two = sec({
      columns: 2,
      colSpace: 425,
      docGrid: { type: 'linesAndChars', linePitch: 312, charSpace: 2048 },
    })
    expect(sectionRightIndentPt(two, 10.5, true)).toBeCloseTo(10.03, 1)
  })

  it('one value for the page when the sections agree, one per block when they do not', () => {
    const narrow = grid('linesAndChars', 2048)
    const wide = sec({
      pageWidth: 16838,
      pageHeight: 11906,
      marginLeft: 1701,
      marginRight: 1701,
      docGrid: { type: 'linesAndChars', linePitch: 312, charSpace: 2048 },
    })
    const el = () => document.createElement('p')
    const blocks: BlockBox[] = [
      { top: 0, height: 100, section: 0, el: el() },
      { top: 100, height: 100, section: 1, el: el() },
      { top: 200, height: 100, section: 2, el: el() },
    ]
    expect(docRightIndentPt([narrow, narrow], 10.5, true)).toBeCloseTo(8.3, 1)
    expect(sectionRightIndentSpecs(blocks, [narrow, narrow], 10.5, true)).toEqual([])
    expect(docRightIndentPt([narrow, wide], 10.5, true)).toBeNull()
    expect(docRightIndentPt([narrow, narrow], 10.5, false)).toBeNull()
    const specs = sectionRightIndentSpecs(blocks, [narrow, wide, sec({})], 10.5, true)
    expect(specs.map((s) => s.el)).toEqual([blocks[0].el, blocks[1].el])
    expect(specs[0].rightIndentPt).toBeCloseTo(8.3, 1)
    expect(specs[1].rightIndentPt).toBeCloseTo(0.8, 1)
    expect(sectionRightIndentSpecs(blocks, [narrow, wide], 10.5, false)).toEqual([])
  })
})

function parsedWith(
  display: StyleDisplay,
  over: Partial<ParsedDocFull> = {},
  normal?: StyleDisplay,
): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  styles.set('Body', { styleId: 'Body', name: 'Body', type: 'paragraph', display } as StyleInfo)
  if (normal)
    styles.set('Normal', {
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: normal,
    } as StyleInfo)
  return { styles, docDefaults: {}, blocks: [], ...over } as unknown as ParsedDocFull
}

describe('the document style sheet', () => {
  const BODY = 'p, h1, h2, h3, h4, h5, h6, .doc-li, .doc-li-stray'

  it('below compatibilityMode 15 body paragraphs take the share of their section', () => {
    const css = docStyleCss(parsedWith({}))
    expect(css).toContain(`.doc-page > :where(${BODY}) { --doc-rind:var(--doc-rind-sect, 0pt) }`)
    // the rule that applies the share is added only while a section has one to give
    expect(css).not.toContain('margin-inline-end:calc(var(--pind-e)')
    expect(GRID_RIGHT_INDENT_RULE).toBe(
      `.doc-page > :is(${BODY}):not([style*="margin-inline-end"]) { margin-inline-end:calc(var(--pind-e) + var(--doc-rind, 0pt)) !important }`,
    )
    expect(css).toContain(`.doc-page .doc-table :is(${BODY}) { --doc-rind:0pt !important }`)
  })

  it('a Word 2013 document gets none of it', () => {
    const css = docStyleCss(parsedWith({ adjustRightInd: false }, { compatibilityMode: 15 }))
    expect(css).not.toContain('--doc-rind')
  })

  it('a style, the default paragraph style and the document defaults switch it', () => {
    expect(docStyleCss(parsedWith({ adjustRightInd: false }))).toMatch(
      /\.doc-page \[data-style="Body"\] \{[^}]*--doc-rind:0pt/,
    )
    expect(docStyleCss(parsedWith({ adjustRightInd: true }))).toMatch(
      /\.doc-page \[data-style="Body"\] \{[^}]*--doc-rind:var\(--doc-rind-sect, 0pt\)/,
    )
    expect(docStyleCss(parsedWith({}))).not.toMatch(/\[data-style="Body"\] \{[^}]*--doc-rind/)
    expect(docStyleCss(parsedWith({}, {}, { adjustRightInd: false }))).toContain(
      `.doc-page > :where(${BODY}) { --doc-rind:0pt }`,
    )
    expect(
      docStyleCss(parsedWith({}, { docDefaults: { adjustRightInd: false } } as never)),
    ).toContain(`.doc-page > :where(${BODY}) { --doc-rind:0pt }`)
  })
})

describe('the paragraph on the page', () => {
  const specOf = (attrs: Record<string, unknown>) =>
    JSON.stringify(fromPmNode({ attrs: { docxIndex: null, ...attrs }, content: [] } as never))

  it('marks its own setting either way, and nothing when it has none', () => {
    expect(specOf({ adjustRightInd: false })).toContain('doc-rind-off')
    expect(specOf({ adjustRightInd: true })).toContain('doc-rind-on')
    expect(specOf({})).not.toContain('doc-rind')
  })

  it('adds the share to an indent of its own only below compatibilityMode 15', () => {
    expect(specOf({ indentRight: 880 })).toContain('margin-inline-end:44pt')
    expect(specOf({ indentRight: 880 })).not.toContain('--doc-rind')
    expect(endIndentDecl('44pt')).toBe('margin-inline-end:44pt')
    setOpenDocumentText({ legacyLayout: true })
    expect(specOf({ indentRight: 880 })).toContain(
      'margin-inline-end:calc(44pt + var(--doc-rind, 0pt))',
    )
    expect(specOf({})).not.toContain('--doc-rind')
  })
})

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

async function openDoc(bodyXml: string, stylesXml?: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, ...(stylesXml ? { stylesXml } : {}) }))
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

describe('read and write', () => {
  it('the paragraph keeps on, off and unset apart', async () => {
    const { editor, parsed } = await openDoc(
      p('<w:adjustRightInd w:val="0"/>', 'off') + p('<w:adjustRightInd/>', 'on') + p('', 'unset'),
    )
    expect(editor.state.doc.child(0).attrs.adjustRightInd).toBe(false)
    expect(editor.state.doc.child(1).attrs.adjustRightInd).toBe(true)
    expect(editor.state.doc.child(2).attrs.adjustRightInd).toBeNull()
    for (const i of [0, 1, 2])
      expect(signatureOfGenerated(pmNodeToGeneratedBlock(editor.state.doc.child(i).toJSON()))).toBe(
        signatureOfBlock(parsed.blocks[i]),
      )
  })

  it('writes the flag before snapToGrid and the indent', async () => {
    const { editor } = await openDoc(p('<w:snapToGrid w:val="0"/><w:ind w:right="880"/>', 'body'))
    editor.commands.setTextSelection(2)
    setParaAttrs(editor, { adjustRightInd: false })
    const xml = xmlOf(editor)
    const at = ['<w:adjustRightInd w:val="0"/>', '<w:snapToGrid', '<w:ind '].map((tag) =>
      xml.indexOf(tag),
    )
    expect(at.every((i) => i >= 0)).toBe(true)
    expect(at).toEqual([...at].sort((a, b) => a - b))
    setParaAttrs(editor, { adjustRightInd: true })
    expect(xmlOf(editor)).toContain('<w:adjustRightInd/>')
    setParaAttrs(editor, { adjustRightInd: null })
    expect(xmlOf(editor)).not.toContain('adjustRightInd')
  })

  it('an edit of something else keeps the bytes, a reset to the style drops the tag', () => {
    const RAW = '<w:pPr><w:adjustRightInd w:val="0"/><w:jc w:val="both"/></w:pPr>'
    const was = { adjustRightInd: false, align: 'justify' as const }
    expect(mergePPrFormat(RAW, was, was)).toBe(RAW)
    expect(mergePPrFormat(RAW, { align: 'justify' }, was)).toBe(
      '<w:pPr><w:jc w:val="both"/></w:pPr>',
    )
    expect(mergePPrFormat(RAW, { adjustRightInd: true, align: 'justify' }, was)).toBe(
      '<w:pPr><w:adjustRightInd/><w:jc w:val="both"/></w:pPr>',
    )
  })

  it('reads the style chain and the document defaults', async () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:pPrDefault><w:pPr><w:adjustRightInd w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Off"><w:name w:val="Off"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:adjustRightInd w:val="0"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Off"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Back"><w:name w:val="Back"/><w:basedOn w:val="Off"/>' +
      '<w:pPr><w:adjustRightInd/></w:pPr></w:style>' +
      '</w:styles>'
    const { parsed } = await openDoc(p('<w:pStyle w:val="Child"/>', 'body'), stylesXml)
    expect(parsed.docDefaults?.adjustRightInd).toBe(false)
    expect(parsed.styles.get('Normal')?.display?.adjustRightInd).toBeUndefined()
    expect(parsed.styles.get('Off')?.display?.adjustRightInd).toBe(false)
    expect(parsed.styles.get('Child')?.display?.adjustRightInd).toBe(false)
    expect(parsed.styles.get('Back')?.display?.adjustRightInd).toBe(true)
  })

  it('Set As Default hands the style the box when it was changed or is the paragraph own', () => {
    const base = {
      attrs: {},
      align: 'left' as const,
      line: { lineSpacing: null, lineRule: null, lineRawTwips: null },
      indentLeft: 0,
      indentRight: 0,
      firstLine: null,
      spaceBefore: 0,
      spaceAfter: 0,
      flags: { adjustRightInd: true },
      initialFlags: { adjustRightInd: true },
    }
    expect(paragraphDefaultProps(base)).toEqual({})
    expect(paragraphDefaultProps({ ...base, flags: { adjustRightInd: false } })).toEqual({
      adjustRightInd: false,
    })
    expect(
      paragraphDefaultProps({
        ...base,
        attrs: { adjustRightInd: false },
        flags: { adjustRightInd: false },
        initialFlags: { adjustRightInd: false },
      }),
    ).toEqual({ adjustRightInd: false })
    const style =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    const off = upsertStyleXml(style, { styleId: 'Normal', pPr: { adjustRightInd: false } })
    expect(off).toContain('<w:adjustRightInd w:val="0"/>')
  })
})
