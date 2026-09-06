import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { afterAll, describe, expect, it, vi } from 'vitest'
import { parseDocx, saveDocx, type NumberingDef, type StyleInfo } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  bulletMarkerScale,
  computeListMarkerInfos,
  computeListMarkers,
  formatNumber,
} from '../src/renderer/editor/numbering'

function def(
  partial: Partial<NumberingDef> & Pick<NumberingDef, 'numId' | 'levels'>,
): NumberingDef {
  return { abstractNumId: partial.numId, startOverrides: {}, ...partial }
}

const DECIMAL_3LVL = def({
  numId: '1',
  levels: {
    0: { numFmt: 'decimal', lvlText: '%1.', start: 1 },
    1: { numFmt: 'decimal', lvlText: '%1.%2', start: 1 },
    2: { numFmt: 'lowerLetter', lvlText: '%3)', start: 1 },
  },
})

const defs = (...list: NumberingDef[]) => new Map(list.map((d) => [d.numId, d]))

describe('formatNumber', () => {
  it('covers Word number formats', () => {
    expect(formatNumber(3, 'decimal')).toBe('3')
    expect(formatNumber(3, 'decimalZero')).toBe('03')
    expect(formatNumber(27, 'lowerLetter')).toBe('aa')
    expect(formatNumber(2, 'upperLetter')).toBe('B')
    expect(formatNumber(4, 'lowerRoman')).toBe('iv')
    expect(formatNumber(1949, 'upperRoman')).toBe('MCMXLIX')
    expect(formatNumber(12, 'chineseCountingThousand')).toBe('十二')
    expect(formatNumber(305, 'chineseCounting')).toBe('三百零五')
    expect(formatNumber(3, 'decimalEnclosedCircle')).toBe('③')
    expect(formatNumber(9, 'none')).toBe('')
  })
})

describe('computeListMarkers', () => {
  it('numbers multilevel lists and resets deeper levels', () => {
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 2 },
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1 },
      ],
      defs(DECIMAL_3LVL),
    )
    expect(markers).toEqual(['1.', '1.1', '1.2', 'a)', '2.', '2.1'])
  })

  it('a sub-level item instantiates untouched shallower levels at their start', () => {
    // Word: "1.1"-style items before any level-0 item consume the level-0
    // start value, so the first explicit level-0 item numbers as 2.
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 1 },
        { numId: '1', ilvl: 0 },
      ],
      defs(DECIMAL_3LVL),
    )
    expect(markers).toEqual(['1.1', '1.2', '2.', '2.1', '3.'])
  })

  it('continues numbering across interleaved plain paragraphs', () => {
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '1', ilvl: 0 },
      ],
      defs(DECIMAL_3LVL),
    )
    expect(markers).toEqual(['1.', '2.'])
  })

  it('applies start values and startOverride restarts', () => {
    const five = def({
      numId: '5',
      levels: { 0: { numFmt: 'decimal', lvlText: '%1.', start: 5 } },
    })
    const restart = def({
      numId: '6',
      abstractNumId: '5',
      levels: { 0: { numFmt: 'decimal', lvlText: '%1.', start: 5 } },
      startOverrides: { 0: 1 },
    })
    const markers = computeListMarkers(
      [
        { numId: '5', ilvl: 0 },
        { numId: '5', ilvl: 0 },
        { numId: '6', ilvl: 0 },
        { numId: '6', ilvl: 0 },
      ],
      defs(five, restart),
    )
    expect(markers).toEqual(['5.', '6.', '1.', '2.'])
  })

  it('shares counters across numIds of the same abstractNum', () => {
    const a = def({ numId: '1', abstractNumId: '9', levels: DECIMAL_3LVL.levels })
    const b = def({ numId: '2', abstractNumId: '9', levels: DECIMAL_3LVL.levels })
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 },
        { numId: '2', ilvl: 0 },
      ],
      defs(a, b),
    )
    expect(markers).toEqual(['1.', '2.'])
  })

  it('maps bullet glyphs and falls back per level', () => {
    const bullets = def({
      numId: '3',
      levels: {
        0: { numFmt: 'bullet', lvlText: '', start: 1 },
        1: { numFmt: 'bullet', lvlText: 'o', start: 1 },
        2: { numFmt: 'bullet', lvlText: '', start: 1 },
      },
    })
    const markers = computeListMarkers(
      [
        { numId: '3', ilvl: 0 },
        { numId: '3', ilvl: 1 },
        { numId: '3', ilvl: 2 },
      ],
      defs(bullets),
    )
    expect(markers).toEqual(['•', '◦', '▪'])
  })

  it('returns null for unknown numIds and undefined levels', () => {
    const markers = computeListMarkers(
      [
        { numId: '99', ilvl: 0 },
        { numId: '1', ilvl: 7 },
        { numId: null, ilvl: 0 },
      ],
      defs(DECIMAL_3LVL),
    )
    expect(markers).toEqual([null, null, null])
  })

  it('renders data-marker on list items in the editor', async () => {
    const numberingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const li = (ilvl: number, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    const source = await buildDocx({
      bodyXml: li(0, 'level one') + li(1, 'level two') + li(0, 'level one again'),
      numberingXml,
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const markers = Array.from(editor.view.dom.querySelectorAll('.doc-li')).map((el) =>
      el.getAttribute('data-marker'),
    )
    expect(markers).toEqual(['1.', '1.1', '2.'])
    editor.destroy()
  })

  it('counts a numbered textbox anchor paragraph in the sequence and marks its stray line', async () => {
    const numberingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/>' +
      '<w:pPr><w:ind w:left="564" w:hanging="328"/></w:pPr></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const numPr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    const li = (text: string) => `<w:p><w:pPr>${numPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
    const host =
      `<w:p><w:pPr>${numPr}<w:ind w:left="537" w:hanging="301"/></w:pPr>` +
      '<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="s1" type="#_x0000_t202" ' +
      'style="position:absolute;margin-left:300pt;margin-top:1pt;width:77pt;height:15pt">' +
      '<v:textbox><w:txbxContent><w:p><w:r><w:t>$918,600</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
      '</v:shape></w:pict></w:r><w:r><w:t>Total Amount Requested:</w:t></w:r></w:p>'
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: li('Duration') + host + li('Nature'), numberingXml }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const markers = Array.from(editor.view.dom.querySelectorAll('.doc-li')).map((el) =>
      el.getAttribute('data-marker'),
    )
    expect(markers).toEqual(['(a)', '(c)'])
    const wrapper = editor.view.dom.querySelector<HTMLElement>('[data-stray-marker]')
    expect(wrapper?.style.getPropertyValue('--li-marker')).toBe('"(b)"')
    const stray = wrapper?.querySelector<HTMLElement>('.doc-textbox-stray')
    expect(stray?.classList.contains('doc-li-stray')).toBe(true)
    expect(stray?.style.getPropertyValue('--li-left')).toBe('26.85pt')
    expect(stray?.style.getPropertyValue('--li-hang')).toBe('15.05pt')
    editor.destroy()
  })

  it('the stray marker honors the wrapper flags a .doc-li marker honors (suff, clip, marker line height)', () => {
    const css = readFileSync(
      join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
      'utf8',
    )
    const marker = /\[data-stray-marker\] \.doc-li-stray::before \{([^}]*)\}/.exec(css)!
    expect(marker[1]).toContain('line-height: var(--li-marker-lh, inherit)')
    expect(css).toMatch(
      /\[data-stray-marker\]\[data-marker-clip\] \.doc-li-stray::before \{\s*visibility: hidden;/,
    )
    expect(css).toMatch(
      /\[data-stray-marker\]\[data-suff\] \.doc-li-stray::before \{\s*min-width: 0;/,
    )
    expect(css).toMatch(
      /\[data-stray-marker\]\[data-suff='space'\] \.doc-li-stray::before \{\s*content: var\(--li-marker\) '\\00a0';/,
    )
  })

  it('decodes decimal and hexadecimal numeric references without changing source OOXML', async () => {
    const numberingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/></w:lvl>' +
      '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>',
      numberingXml,
    })
    const parsed = await parseDocx(source)
    const parsedDef = parsed.numbering.get('1')!

    expect(parsedDef.levels[0].lvlText).toBe(String.fromCodePoint(61623))
    expect(parsedDef.levels[1].lvlText).toBe(String.fromCodePoint(0xf0a7))
    expect(
      computeListMarkers(
        [
          { numId: '1', ilvl: 0 },
          { numId: '1', ilvl: 1 },
        ],
        parsed.numbering,
      ),
    ).toEqual(['•', '▪'])

    const originals = parsed.blocks
      .filter((block) => !block.hidden && block.docxIndex !== null)
      .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! }))
    expect(await saveDocx(parsed, originals)).toEqual(source)
  })

  it('Chinese numbering and bracketed/circled lvlText', () => {
    const cn = def({
      numId: '8',
      levels: { 0: { numFmt: 'chineseCountingThousand', lvlText: '%1、', start: 1 } },
    })
    const markers = computeListMarkers(
      Array.from({ length: 11 }, () => ({ numId: '8', ilvl: 0 })),
      defs(cn),
    )
    expect(markers[0]).toBe('一、')
    expect(markers[9]).toBe('十、')
    expect(markers[10]).toBe('十一、')
  })

  it('symbol-font bullets decode by font; undecodable ones fall back to the default symbol', () => {
    const mk = (numId: string, lvlText: string, font?: string) =>
      def({
        numId,
        levels: { 0: { numFmt: 'bullet', lvlText, start: 1, ...(font ? { font } : {}) } },
      })
    const markers = computeListMarkers(
      [
        { numId: '1', ilvl: 0 }, // Wingdings 'l' = ●
        { numId: '2', ilvl: 0 }, // Wingdings PUA F0A7 = ▪
        { numId: '3', ilvl: 0 }, // Webdings has no mapping table → default
        { numId: '4', ilvl: 0 }, // glyph not in the Wingdings table → default
        { numId: '5', ilvl: 0 }, // PUA without a font → legacy font-agnostic mapping
      ],
      defs(
        mk('1', 'l', 'Wingdings'),
        mk('2', '', 'Wingdings'),
        mk('3', '', 'Webdings'),
        mk('4', 'e', 'Wingdings'),
        mk('5', ''),
      ),
    )
    expect(markers).toEqual(['●', '▪', '•', '•', '•'])
  })

  it('bullet infos carry the original glyph (U+F0xx form) and the declared symbol font', () => {
    const mk = (numId: string, lvlText: string, font?: string) =>
      def({
        numId,
        levels: { 0: { numFmt: 'bullet', lvlText, start: 1, ...(font ? { font } : {}) } },
      })
    const infos = computeListMarkerInfos(
      [
        { numId: '1', ilvl: 0 },
        { numId: '2', ilvl: 0 },
        { numId: '3', ilvl: 0 },
        { numId: '4', ilvl: 0 },
        { numId: '5', ilvl: 0 },
      ],
      defs(
        mk('1', '\uF0B7', 'Symbol'),
        mk('2', 'l', 'Wingdings'), // raw byte normalized to U+F0xx
        mk('3', '\uF0C1', 'Symbol'), // undecodable → default text, glyph still kept
        mk('4', '\uF0B7'), // no symbol font → plain substitute
        def({ numId: '5', levels: DECIMAL_3LVL.levels }),
      ),
    )
    expect(infos[0]).toEqual({ text: '•', symbolChar: '\uF0B7', symbolFont: 'Symbol' })
    expect(infos[1]).toEqual({ text: '●', symbolChar: '\uF06C', symbolFont: 'Wingdings' })
    expect(infos[2]).toEqual({ text: '•', symbolChar: '\uF0C1', symbolFont: 'Symbol' })
    expect(infos[3]).toEqual({ text: '•' })
    expect(infos[4]).toEqual({ text: '1.' })
  })

  it('upscales only solid round substitute glyphs', () => {
    expect(bulletMarkerScale('•')).toBe(1.25)
    expect(bulletMarkerScale('●')).toBe(1.35)
    expect(bulletMarkerScale('▪')).toBe(1)
    expect(bulletMarkerScale('1.')).toBe(1)
  })

  it('substitutes uncovered symbol bullets with a pinned font + scale and follows the first run size', async () => {
    const numberingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/>' +
      '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>' +
      '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
        '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>item</w:t></w:r></w:p>',
      numberingXml,
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const el = editor.view.dom.querySelector('.doc-li')!
    // no canvas in the test DOM → coverage probe fails → substitution path
    expect(el.getAttribute('data-marker')).toBe('•')
    const style = el.getAttribute('style') ?? ''
    expect(style).toContain("--li-marker-font: Arial,'Helvetica Neue',sans-serif")
    expect(style).toContain('--li-marker-scale: 1.25')
    expect(style).toContain('--li-marker-lh: 0')
    expect(style).toContain('--li-marker-size: 12pt')
    editor.destroy()
  })
})

describe('list marker decorations', () => {
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
  const numberingXml = (lvl: string) =>
    XML_DECL +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">${lvl}</w:lvl></w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '</w:numbering>'
  const li =
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    '<w:r><w:t>item</w:t></w:r></w:p>'

  async function renderLi(
    lvlXml: string,
    styles?: Map<string, StyleInfo>,
  ): Promise<{ el: Element; destroy: () => void }> {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: li, numberingXml: numberingXml(lvlXml) }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    if (styles) editor.storage.listNumbering.styles = styles
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    return { el: editor.view.dom.querySelector('.doc-li')!, destroy: () => editor.destroy() }
  }

  // 8px per char: predictable widths for the tab-advance assertions;
  // measuredFonts records the ctx font of each measurement
  const measuredFonts: string[] = []
  const fakeCtx = {
    font: '',
    measureText(s: string) {
      measuredFonts.push(this.font)
      return { width: s.length * 8 }
    },
  }
  const measureStub = vi
    .spyOn(HTMLCanvasElement.prototype, 'getContext')
    .mockReturnValue(fakeCtx as never)
  afterAll(() => measureStub.mockRestore())

  it('flags markers hanging left of a table cell text area as clipped (Word clips there)', async () => {
    const lvl =
      '<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/>' +
      '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>'
    const item = (ind: string, text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${ind}</w:pPr>` +
      `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`
    const cell = (paras: string) =>
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${paras}</w:tc></w:tr></w:tbl>`
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          item('<w:ind w:left="195"/>', 'body: marker in the margin stays visible') +
          cell(
            item('<w:ind w:left="195"/>', 'cell: marker box ends 3pt left of the text area') +
              item('<w:ind w:left="300"/>', 'cell: marker still reaches into the text area') +
              item('<w:ind w:left="720" w:hanging="360"/>', 'cell: marker inside') +
              item(
                '<w:ind w:left="-186" w:firstLine="366"/>',
                'cell: first line pushes the marker in',
              ) +
              item(
                '<w:bidi/><w:ind w:left="-186" w:firstLine="0"/>',
                'cell: RTL start indent past the text area clips too',
              ) +
              item(
                '<w:ind w:left="195" w:firstLine="0"/>',
                'cell: explicit zero cancels the level hanging',
              ),
          ),
        numberingXml: numberingXml(lvl),
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const items = Array.from(editor.view.dom.querySelectorAll('.doc-li'))
    expect(items.map((el) => el.getAttribute('data-marker'))).toEqual(Array(7).fill('•'))
    expect(items.map((el) => el.hasAttribute('data-marker-clip'))).toEqual([
      false,
      true,
      false,
      false,
      false,
      true,
      false,
    ])
    editor.destroy()
  })

  it('numFmt "none" emits an empty data-marker (suppresses the CSS counter fallback)', async () => {
    const { el, destroy } = await renderLi(
      '<w:start w:val="1"/><w:numFmt w:val="none"/><w:suff w:val="nothing"/><w:lvlText w:val=""/>' +
        '<w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr>',
    )
    expect(el.getAttribute('data-marker')).toBe('')
    expect(el.getAttribute('data-suff')).toBe('nothing')
    destroy()
  })

  it('a level with only a left indent has no hanging area: its cell marker is not clipped', async () => {
    const lvl =
      '<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/>' +
      '<w:pPr><w:ind w:left="195"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
          '<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>' +
          '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
          '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>marker at the first-line position</w:t></w:r></w:p>' +
          '</w:tc></w:tr></w:tbl>',
        numberingXml: numberingXml(lvl),
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const item = editor.view.dom.querySelector('.doc-li') as HTMLElement
    expect(item.getAttribute('style')).toContain('--li-hang: 0pt')
    expect(item.hasAttribute('data-marker-clip')).toBe(false)
    editor.destroy()
  })

  it('markers overflowing the hanging area advance to the next default tab stop', async () => {
    const { el, destroy } = await renderLi(
      '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="NEW-%1-FORMAT"/>' +
        '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>',
    )
    expect(el.getAttribute('data-marker')).toBe('NEW-1-FORMAT')
    // 12 chars * 8px = 96px = 1440 twips from marker start 0 -> stop 2160 twips
    expect(el.getAttribute('style')).toContain('--li-tab: 108pt')
    destroy()
  })

  it('level positive firstLine shifts the marker right of the text indent', async () => {
    const { el, destroy } = await renderLi(
      '<w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1"/>' +
        '<w:pPr><w:ind w:left="432" w:firstLine="135"/></w:pPr>',
    )
    expect(el.getAttribute('data-marker')).toBe('A')
    const style = el.getAttribute('style') ?? ''
    // marker at 432+135=567, width 120 twips -> stop 720: no hang, box 7.65pt
    expect(style).toContain('--li-hang: 0pt')
    expect(style).toContain('text-indent: 6.75pt')
    expect(style).toContain('--li-tab: 7.65pt')
    destroy()
  })

  it('a paragraph firstLine (0 included) cancels the level hanging: marker at the first-line position, tab to the next default stop', async () => {
    const lvl =
      '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      '<w:pPr><w:ind w:left="1032" w:hanging="360"/></w:pPr>'
    const item = (pPr: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${pPr}</w:pPr>` +
      '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r></w:p>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="780" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="780"/></w:tblGrid>' +
          '<w:tr><w:tc><w:tcPr><w:tcW w:w="780" w:type="dxa"/></w:tcPr>' +
          item(
            '<w:bidi/><w:ind w:left="-186" w:right="-390" w:firstLine="0"/><w:jc w:val="right"/>',
          ) +
          item(
            '<w:bidi/><w:ind w:left="-186" w:right="-390" w:firstLine="366"/><w:jc w:val="center"/>',
          ) +
          item('<w:ind w:left="720" w:firstLine="0"/>') +
          '</w:tc></w:tr></w:tbl>',
        numberingXml: numberingXml(lvl),
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.storage.listNumbering.defs = parsed.numbering
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const styles = Array.from(editor.view.dom.querySelectorAll('.doc-li')).map((el) => ({
      style: el.getAttribute('style') ?? '',
      clip: el.hasAttribute('data-marker-clip'),
    }))
    // RTL end-aligned "1." (16px = 240 twips) from -186 ends at 54 -> stop 720: box 906 twips,
    // no default hang; the negative start indent moves the box instead of a padding
    expect(styles[0].style).toContain('margin-inline-start: -9.3pt')
    expect(styles[0].style).toContain('--li-hang: 0pt')
    expect(styles[0].style).toContain('--li-tab: 45.3pt')
    expect(styles[0].style).not.toContain('--li-hang: 18pt')
    expect(styles[0].clip).toBe(false)
    // centered with firstLine 366: marker at 180, ends 420 -> stop 720
    expect(styles[1].style).toContain('text-indent: 18.3pt')
    expect(styles[1].style).toContain('--li-tab: 27pt')
    // LTR explicit zero: marker at 720, ends 960 -> stop 1440
    expect(styles[2].style).toContain('--li-hang: 0pt')
    expect(styles[2].style).toContain('--li-tab: 36pt')
    editor.destroy()
  })

  it('measures with the font the ::before inherits (Normal style chain, not Calibri 11pt)', async () => {
    const styles = new Map<string, StyleInfo>([
      [
        'Normal',
        {
          styleId: 'Normal',
          name: 'Normal',
          type: 'paragraph',
          isDefault: true,
          display: { font: 'Times New Roman', fontAscii: 'Times New Roman', sizeHalfPoints: 24 },
        },
      ],
    ])
    measuredFonts.length = 0
    const { destroy } = await renderLi(
      '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="NEW-%1-FORMAT"/>' +
        '<w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>',
      styles,
    )
    // 12pt Normal -> 16px, family from the default paragraph style
    expect(measuredFonts[0]).toBe('16px "Times New Roman", Calibri, sans-serif')
    destroy()
  })
})
