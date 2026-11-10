import { describe, expect, it } from 'vitest'
import {
  parseDocx,
  resolveCtxOf,
  resolveParagraphPropsXml,
  resolveRunProps,
  type EffectivePPr,
  type ResolveCtx,
} from '../src/index'
import { findChild, findChildren, nameOf, xmlParser, type XNode } from '../src/xml-utils'
import { buildDocx } from './helpers/build-docx'

/**
 * The resolver re-derives the truths the legacy paths pin in the sibling test
 * files (char-unit-indents, indent-explicit-zero, keep-flags-tristate, rtl-runs,
 * run-vanish, style-widow-control, docdefaults-widowcontrol, table-style,
 * character-styles, link-styles-normal-template, run-size-clamp-...). Each case
 * here is the resolver-side twin of an assertion there.
 */

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const wrap = (styles: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>${styles}</w:styles>`

async function load(bodyXml: string, stylesXml?: string, extra?: { settings?: string }) {
  const doc = await parseDocx(
    await buildDocx({
      bodyXml,
      ...(stylesXml ? { stylesXml } : {}),
      ...(extra?.settings
        ? {
            extraParts: [
              {
                path: 'word/settings.xml',
                xml: `<w:settings ${NS}>${extra.settings}</w:settings>`,
                contentType:
                  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
              },
            ],
          }
        : {}),
    }),
    { shadowResolve: true },
  )
  const ctx = resolveCtxOf(doc)
  const blocks = doc.blocks.filter((b) => !b.hidden)
  const paras = blocks.map((b) => resolveParagraphPropsXml(b.originalXml ?? '', ctx))
  return { doc: { ...doc, blocks }, ctx, paras }
}

function runRPrs(xml: string): Array<XNode | undefined> {
  const parsed = xmlParser.parse(xml) as XNode[]
  const p = parsed.find((n) => nameOf(n) === 'w:p')!
  return findChildren(p, 'w:r').map((r) => findChild(r, 'w:rPr'))
}

function runs(xml: string, para: EffectivePPr, ctx: ResolveCtx) {
  return runRPrs(xml).map((rPr) => resolveRunProps(rPr, para, ctx))
}

describe('resolveParagraphProps: indents', () => {
  const STYLES = wrap(
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
      '<w:pPr><w:ind w:left="454" w:right="120"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:ind w:firstLine="400"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Hang"><w:name w:val="Hang"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:ind w:hanging="300"/></w:pPr></w:style>',
  )
  const para = (pPr: string) => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>text</w:t></w:r></w:p>`

  it('merges w:ind per attribute, an explicit direct 0 cancels the style (indent-explicit-zero)', async () => {
    const { paras } = await load(
      para('<w:pStyle w:val="Body"/><w:ind w:left="0" w:right="0" w:firstLine="0"/>') +
        para('<w:pStyle w:val="Hang"/><w:ind w:hanging="0"/>') +
        para('<w:pStyle w:val="Body"/><w:ind w:right="0"/>') +
        para('<w:pStyle w:val="Body"/>') +
        para(''),
      STYLES,
    )
    expect(paras[0].indent).toEqual({ left: 0, right: 0, firstLine: 0 })
    expect(paras[0].source.indentLeft).toBe('direct')
    expect(paras[1].indent).toEqual({ left: 454, right: 120, firstLine: 0 })
    expect(paras[1].source.indentLeft).toBe('paraStyle')
    expect(paras[2].indent).toEqual({ left: 454, right: 0, firstLine: 400 })
    expect(paras[3].indent).toEqual({ left: 454, right: 120, firstLine: 400 })
    expect(paras[3].source.indentFirstLine).toBe('paraStyle')
    expect(paras[4].indent).toEqual({ left: 454, right: 120, firstLine: 0 })
    expect(paras[4].source.indentFirstLine).toBe('builtin')
  })

  it('numbering level ind sits between style and direct (Word probe 2026-10-20)', async () => {
    const numbering =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          para(
            '<w:pStyle w:val="Body"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>',
          ) +
          para(
            '<w:pStyle w:val="Body"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="1000"/>',
          ),
        stylesXml: STYLES,
        numberingXml: numbering,
      }),
    )
    const ctx = resolveCtxOf(doc)
    const [a, b] = doc.blocks.map((blk) => resolveParagraphPropsXml(blk.originalXml ?? '', ctx))
    expect(a.indent).toEqual({ left: 720, right: 120, firstLine: -360 })
    expect(a.source).toMatchObject({
      indentLeft: 'numbering',
      indentFirstLine: 'numbering',
      indentRight: 'paraStyle',
    })
    expect(b.indent).toEqual({ left: 1000, right: 120, firstLine: -360 })
    expect(b.source.indentLeft).toBe('direct')
    expect(b.list).toMatchObject({ numId: '1', ilvl: 0 })
  })

  it('character-unit indents resolve with the given units (char-unit-indents)', async () => {
    const styles = wrap(
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Body16"><w:name w:val="Body16"/><w:basedOn w:val="Normal"/>' +
        '<w:pPr><w:ind w:firstLineChars="200"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>',
    )
    const { doc, ctx } = await load(
      para('<w:ind w:firstLineChars="200"/>') +
        para('<w:pStyle w:val="Body16"/>') +
        para('<w:pStyle w:val="Body16"/><w:ind w:firstLine="420"/>') +
        para('<w:pStyle w:val="Body16"/><w:ind w:firstLineChars="0" w:firstLine="420"/>') +
        para('<w:ind w:leftChars="200"/>'),
      styles,
    )
    const at = (i: number, run: number) =>
      resolveParagraphPropsXml(doc.blocks[i].originalXml ?? '', ctx, {
        charUnits: { run, normal: 210 },
      })
    expect(at(0, 240).indent).toMatchObject({ firstLine: 480, chars: { firstLine: 200 } })
    expect(at(1, 320).indent.firstLine).toBe(640)
    expect(at(1, 320).source.indentFirstLine).toBe('paraStyle')
    // a twips-only direct firstLine leaves the style's firstLineChars in force
    expect(at(2, 320).indent.firstLine).toBe(640)
    // an explicit zero *Chars makes the twips value absolute
    expect(at(3, 320).indent.firstLine).toBe(420)
    expect(at(3, 320).indent.chars).toBeUndefined()
    expect(at(4, 240).indent.left).toBe(420)
    // the legacy path agrees on every one of these paragraphs
    expect(doc.extras.resolveDiff).toEqual([])
  })
})

describe('resolveParagraphProps: flags and spacing', () => {
  it('keepNext / keepLines tri-state (keep-flags-tristate)', async () => {
    const styles = wrap(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="KeepHead"><w:name w:val="Keep Head"/>' +
        '<w:pPr><w:keepNext w:val="1"/><w:keepLines w:val="1"/></w:pPr></w:style>',
    )
    const p = (pPr: string) =>
      `<w:p><w:pPr><w:pStyle w:val="KeepHead"/>${pPr}</w:pPr><w:r><w:t>t</w:t></w:r></w:p>`
    const { paras, doc } = await load(
      p('<w:keepNext w:val="0"/><w:keepLines w:val="0"/>') +
        p('') +
        p('<w:keepNext/><w:keepLines/>'),
      styles,
    )
    expect(paras.map((e) => [e.keepNext, e.keepLines])).toEqual([
      [false, false],
      [true, true],
      [true, true],
    ])
    expect(paras.map((e) => e.source.keepNext)).toEqual(['direct', 'paraStyle', 'direct'])
    expect(doc.extras.resolveDiff).toEqual([])
  })

  it('widowControl: builtin on, style chain off, explicit on overrides (style-widow-control)', async () => {
    const styles = wrap(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
        '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Strict"><w:name w:val="Strict"/><w:basedOn w:val="Normal"/>' +
        '<w:pPr><w:widowControl/></w:pPr></w:style>',
    )
    const { paras } = await load(
      '<w:p><w:r><w:t>text</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:widowControl/></w:pPr><w:r><w:t>on</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:widowControl w:val="off"/></w:pPr><w:r><w:t>off</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="Strict"/></w:pPr><w:r><w:t>strict</w:t></w:r></w:p>',
      styles,
    )
    expect(paras.map((e) => [e.widowControl, e.source.widowControl])).toEqual([
      [false, 'paraStyle'],
      [true, 'direct'],
      [false, 'direct'],
      [true, 'paraStyle'],
    ])
  })

  it('pPrDefault widowControl off reaches unstyled paragraphs (docdefaults-widowcontrol)', async () => {
    const styles = wrap(
      '<w:docDefaults><w:rPrDefault/><w:pPrDefault><w:pPr><w:widowControl w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults>',
    )
    const { paras } = await load('<w:p><w:r><w:t>Body</w:t></w:r></w:p>', styles)
    expect(paras[0].widowControl).toBe(false)
    expect(paras[0].source.widowControl).toBe('docDefaults')
  })

  it('docDefaults spacing under the Normal style under direct (link-styles-normal-template)', async () => {
    const styles = wrap(
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:before="120" w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
        '<w:pPr><w:spacing w:after="400"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>',
    )
    const { paras } = await load(
      '<w:p><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:spacing w:before="0" w:line="480" w:lineRule="exact"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>',
      styles,
    )
    expect(paras[0].spacing).toEqual({
      before: 120,
      after: 400,
      beforeAuto: false,
      afterAuto: false,
      line: { rule: 'auto', raw: 276 },
    })
    expect(paras[0].source).toMatchObject({
      spaceBefore: 'docDefaults',
      spaceAfter: 'paraStyle',
      line: 'docDefaults',
    })
    expect(paras[1].spacing).toMatchObject({
      before: 0,
      after: 400,
      line: { rule: 'exact', raw: 480 },
    })
    expect(paras[1].source.spaceBefore).toBe('direct')
  })

  it('compat flags travel on the context (run-size-clamp-numbering-tab-compat)', async () => {
    const { ctx } = await load('<w:p><w:r><w:t>x</w:t></w:r></w:p>', undefined, {
      settings:
        '<w:compat><w:doNotUseHTMLParagraphAutoSpacing/><w:doNotUseIndentAsNumberingTabStop/>' +
        '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>',
    })
    expect(ctx.compat).toEqual({
      mode: 15,
      additiveParaSpacing: true,
      indentNotNumberingTabStop: true,
    })
  })

  it('table-style paragraph spacing reaches cell paragraphs through ctx.table (table-style)', async () => {
    const styles = wrap(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
        '<w:pPr><w:spacing w:after="200"/></w:pPr></w:style>' +
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
        '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
        '<w:rPr><w:sz w:val="18"/></w:rPr>' +
        '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr></w:tblStylePr></w:style>',
    )
    const { ctx } = await load('<w:p><w:r><w:t>x</w:t></w:r></w:p>', styles)
    const cellCtx: ResolveCtx = { ...ctx, table: { styleId: 'TableGrid', cond: ['firstRow'] } }
    const cell = resolveParagraphPropsXml('<w:p><w:r><w:t>c</w:t></w:r></w:p>', cellCtx)
    expect(cell.spacing).toMatchObject({ after: 200, line: { rule: 'auto', raw: 240 } })
    // the paragraph style still outranks the table style for what it declares
    expect(cell.source).toMatchObject({ spaceAfter: 'paraStyle', line: 'tableStyle' })
    const run = resolveRunProps(undefined, cell, cellCtx)
    expect(run).toMatchObject({ bold: true, color: 'FFFFFF', sizeHalfPoints: 18 })
    expect(run.source).toMatchObject({ bold: 'tableStyle', sizeHalfPoints: 'tableStyle' })
    const body = resolveRunProps(undefined, cell, {
      ...ctx,
      table: { styleId: 'TableGrid', cond: [] },
    })
    expect(body.bold).toBe(false)
  })
})

describe('resolveRunProps', () => {
  const RTL_STYLES = wrap(
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
      '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/>' +
      '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/>' +
      '<w:rPr><w:bCs/><w:szCs w:val="28"/></w:rPr></w:style>',
  )

  it('rtl runs read the Cs twins up the whole chain (rtl-runs)', async () => {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:bidi/></w:pPr>' +
      '<w:r><w:rPr><w:rtl/></w:rPr><w:t>a</w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="Strong"/><w:rtl/></w:rPr><w:t>b</w:t></w:r>' +
      '<w:r><w:t>Latin</w:t></w:r></w:p>'
    const { paras, ctx } = await load(xml, RTL_STYLES)
    const [plain, strong, latin] = runs(xml, paras[0], ctx)
    expect(plain).toMatchObject({ cs: true, sizeHalfPoints: 22, bold: false })
    expect(plain.source.sizeHalfPoints).toBe('docDefaults')
    expect(strong).toMatchObject({ cs: true, sizeHalfPoints: 28, bold: true })
    expect(strong.source).toMatchObject({ sizeHalfPoints: 'charStyle', bold: 'charStyle' })
    expect(latin).toMatchObject({ cs: false, sizeHalfPoints: 32, bold: true })
    expect(latin.source.bold).toBe('paraStyle')
  })

  it('an rtl paragraph mark sizes by szCs, falling back to the chain (rtl-runs)', async () => {
    const { paras } = await load(
      '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr><w:b/><w:sz w:val="36"/><w:rtl/></w:rPr></w:pPr></w:p>' +
        '<w:p><w:pPr><w:rPr><w:sz w:val="36"/><w:szCs w:val="40"/><w:rtl/></w:rPr></w:pPr></w:p>' +
        '<w:p><w:pPr><w:rPr><w:sz w:val="36"/><w:szCs w:val="40"/></w:rPr></w:pPr></w:p>',
      RTL_STYLES,
    )
    expect(paras.map((e) => e.mark.sizeHalfPoints)).toEqual([24, 40, 36])
    expect(paras[0].mark.bold).toBe(false)
  })

  it('w:sz=0 clamps to one half-point (run-size-clamp)', async () => {
    const xml =
      '<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>Q</w:t></w:r>' +
      '<w:r><w:rPr><w:sz w:val="0"/></w:rPr><w:t>p</w:t></w:r></w:p>'
    const { paras, ctx } = await load(xml)
    expect(runs(xml, paras[0], ctx).map((r) => r.sizeHalfPoints)).toEqual([21, 1])
  })

  it('vanish inherits from paragraph / character / default styles, specVanish and explicit off stay visible (run-vanish)', async () => {
    const styles = wrap(
      '<w:style w:type="paragraph" w:styleId="HiddenPara"><w:name w:val="HiddenPara"/><w:rPr><w:vanish/></w:rPr></w:style>' +
        '<w:style w:type="character" w:styleId="HiddenChar"><w:name w:val="HiddenChar"/><w:rPr><w:vanish/></w:rPr></w:style>',
    )
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="HiddenPara"/></w:pPr>' +
      '<w:r><w:t>inherited</w:t></w:r>' +
      '<w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>unhidden</w:t></w:r>' +
      '<w:r><w:rPr><w:vanish/><w:specVanish/></w:rPr><w:t>sep</w:t></w:r></w:p>'
    const { paras, ctx } = await load(xml, styles)
    expect(runs(xml, paras[0], ctx).map((r) => r.vanish)).toEqual([true, false, true])
    expect(paras[0].hidden).toBe(true)
    expect(paras[0].mark.source.vanish).toBe('paraStyle')

    const charXml =
      '<w:p><w:r><w:rPr><w:rStyle w:val="HiddenChar"/></w:rPr><w:t>h</w:t></w:r><w:r><w:t>v</w:t></w:r></w:p>'
    const c = await load(charXml, styles)
    expect(runs(charXml, c.paras[0], c.ctx).map((r) => r.vanish)).toEqual([true, false])
    expect(c.paras[0].hidden).toBe(false)

    const defaults = wrap(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normlny"><w:name w:val="Normlny"/><w:rPr><w:vanish/></w:rPr></w:style>',
    )
    const d = await load('<w:p><w:r><w:t>form</w:t></w:r></w:p>', defaults)
    expect(d.paras[0].hidden).toBe(true)
    expect(d.doc.extras.resolveDiff).toEqual([])
  })

  it('toggle properties XOR across the style layers, direct wins (character-styles)', async () => {
    const styles = wrap(
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="BoldPara"><w:name w:val="BoldPara"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>' +
        '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr></w:style>',
    )
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="BoldPara"/></w:pPr>' +
      '<w:r><w:t>para bold</w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>xor regular</w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="Strong"/><w:b/></w:rPr><w:t>direct bold</w:t></w:r></w:p>'
    const { paras, ctx } = await load(xml, styles)
    const [a, b, c] = runs(xml, paras[0], ctx)
    expect(a.bold).toBe(true)
    expect(b).toMatchObject({ bold: false, color: 'C00000' })
    expect(b.source.bold).toBe('charStyle')
    expect(c.bold).toBe(true)
    expect(c.source.bold).toBe('direct')
  })
})

describe('shadow mode', () => {
  it('records no disagreement on a document the legacy path already resolves', async () => {
    const styles = wrap(
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
        '<w:pPr><w:tabs><w:tab w:val="center" w:pos="4000"/></w:tabs><w:autoSpaceDN w:val="0"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="H"><w:name w:val="H"/><w:basedOn w:val="Normal"/>' +
        '<w:pPr><w:keepNext/><w:spacing w:before="240"/><w:ind w:left="300"/></w:pPr><w:rPr><w:vanish/></w:rPr></w:style>',
    )
    const { doc } = await load(
      '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="100" w:after="0"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="H"/><w:tabs><w:tab w:val="clear" w:pos="4000"/><w:tab w:val="right" w:pos="9000"/></w:tabs></w:pPr><w:r><w:t>b</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:pPr></w:p>' +
        '<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/><w:keepLines w:val="0"/></w:pPr><w:r><w:t>d</w:t></w:r></w:p>',
      styles,
    )
    expect(doc.extras.resolveDiff).toEqual([])
  })
})
