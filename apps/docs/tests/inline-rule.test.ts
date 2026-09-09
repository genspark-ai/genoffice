import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import { docStyleCss } from '../src/renderer/doc-style-css'
import { runSpanSpecs } from '../src/renderer/editor/protected-render'
import { makeGapHfEl } from '../src/renderer/editor/hf-dom'
import { inlineRuleDecls, inlineRuleStyle } from '../src/renderer/editor/inline-rule'
import { parseDocx, saveDocx, type HeaderFooter } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { executeOps } from '../src/renderer/ai/ops'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import type { ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const HR_XML =
  '<w:pict><v:rect style="width:0;height:1.5pt" o:hr="t" fillcolor="#a0a0a0"/></w:pict>'
const RULE = { colorHex: 'A0A0A0', thicknessPx: 2 }

describe('VML horizontal rule sharing a text paragraph', () => {
  it('renders the rule run as a full-width inline rule, not an image', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'Title' },
              { type: 'docInlineImage', attrs: { dataUrl: '', xml: HR_XML, rule: RULE } },
            ],
          },
        ],
      },
    })
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.querySelector('img[data-inline-image]')).toBeNull()
    const rule = p.querySelector('span.doc-inline-rule') as HTMLElement
    expect(rule).not.toBeNull()
    expect(rule.style.getPropertyValue('--doc-rule-color')).toBe('#A0A0A0')
    expect(rule.style.getPropertyValue('--doc-rule-h')).toBe('2px')
    expect(rule.style.fontSize).toBe('')
    editor.destroy()
  })

  it('renders a rule-only paragraph as one rule line inside the paragraph (Word: one line of the run font + paragraph spacing)', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0, spaceAfter: 160 },
            content: [{ type: 'docInlineImage', attrs: { dataUrl: '', xml: HR_XML, rule: RULE } }],
          },
        ],
      },
    })
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(editor.view.dom.querySelector('.doc-protected')).toBeNull()
    expect(p.querySelector('span.doc-inline-rule')).not.toBeNull()
    expect(p.style.marginBottom).not.toBe('')
    editor.destroy()
  })

  it('flags only a rule that opens a paragraph continuing with text as the lead rule', () => {
    const hr = { text: '', image: { dataUrl: '', xml: HR_XML, rule: RULE } }
    const lead = runsToInline([hr, { text: 'Title', bold: true }, hr])
    expect(lead[0].attrs?.leadRule).toBe(true)
    expect(lead[2].attrs?.leadRule).toBeUndefined()
    expect(runsToInline([hr])[0].attrs?.leadRule).toBeUndefined()
    expect(runsToInline([{ text: 'Title' }, hr])[1].attrs?.leadRule).toBeUndefined()
    expect(runsToInline([hr, hr])[0].attrs?.leadRule).toBeUndefined()

    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', attrs: { docxIndex: 0 }, content: lead }],
      },
    })
    const rules = editor.view.dom.querySelectorAll<HTMLElement>('span.doc-inline-rule')
    expect(rules).toHaveLength(2)
    expect(rules[0].hasAttribute('data-lead-rule')).toBe(true)
    expect(rules[1].hasAttribute('data-lead-rule')).toBe(false)
    const html = editor.getHTML()
    editor.commands.setContent(html)
    const pasted = editor.getJSON().content?.[0].content as Array<{
      attrs?: Record<string, unknown>
    }>
    expect(pasted[0].attrs?.leadRule).toBe(true)
    expect(pasted[2].attrs?.leadRule).toBe(false)
    editor.destroy()
  })

  it('gives the lead rule a compact line (Word: rule height + 5.5pt) except under fixed line heights', () => {
    const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
    const lead = /\.doc-inline-rule\[data-lead-rule\] \{([^}]*)\}/.exec(css)?.[1] ?? ''
    expect(lead).toContain('padding-top: 5.5pt')
    expect(lead).toContain('padding-bottom: 0;')
    expect(lead).not.toMatch(/margin-(top|bottom):|height:/)
    const fixed =
      /\.doc-lh-fixed > \.doc-inline-rule\[data-lead-rule\] \{([^}]*)\}/.exec(css)?.[1] ?? ''
    expect(fixed).toContain('padding-top: calc(1lh - 0.22em - var(--doc-rule-h, 1px))')
    expect(fixed).toContain('padding-bottom: 0.22em')
  })

  it('sizes and aligns a rule that declares its own width', () => {
    const rule = { ...RULE, widthPx: 576, align: 'center' as const }
    expect(inlineRuleDecls(rule)).toEqual([
      '--doc-rule-color:#A0A0A0',
      '--doc-rule-h:2px',
      'width:576px;max-width:100%',
      'margin-left:auto',
      'margin-right:auto',
    ])
    expect(inlineRuleStyle({ ...RULE, widthPx: 576, align: 'right' })).toMatchObject({
      width: '576px',
      maxWidth: '100%',
      marginLeft: 'auto',
    })
    expect(inlineRuleDecls({ ...RULE, align: 'center' })).toEqual([
      '--doc-rule-color:#A0A0A0',
      '--doc-rule-h:2px',
    ])
  })

  it('sizes the rule line by the rule run font, not the paragraph strut', () => {
    const nodes = runsToInline([
      { text: 'Title', bold: true, sizeHalfPoints: 28 },
      { text: '', sizeHalfPoints: 20, image: { dataUrl: '', xml: HR_XML, rule: RULE } },
    ])
    expect(nodes[1]).toMatchObject({
      type: 'docInlineImage',
      attrs: { rule: { ...RULE, sizeHalfPoints: 20 } },
    })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', attrs: { docxIndex: 0 }, content: nodes }],
      },
    })
    const rule = editor.view.dom.querySelector('span.doc-inline-rule') as HTMLElement
    expect(rule.style.fontSize).toBe('10pt')
    editor.destroy()
  })

  it('publishes the document base font size for rule runs without w:sz', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('Normal', {
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: { sizeHalfPoints: 22 },
    } as StyleInfo)
    styles.set('Big', {
      styleId: 'Big',
      name: 'Big',
      type: 'paragraph',
      display: { sizeHalfPoints: 28 },
    } as StyleInfo)
    const css = docStyleCss({ styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull)
    expect(css).toContain('.page-wrap, .doc-page, .pv-page { --doc-base-fs:11pt }')
    expect(css).toContain('[data-style="Big"] { font-size:14pt;--doc-base-fs:14pt')
  })

  it('maps the rule node back to a run carrying the pict fragment', () => {
    const runs = inlineToRuns([
      { type: 'text', text: 'Title' },
      { type: 'docInlineImage', attrs: { dataUrl: '', xml: HR_XML, rule: RULE } },
    ])
    expect(runs).toHaveLength(2)
    expect(runs[1].image).toEqual({ dataUrl: '', xml: HR_XML, rule: RULE })
  })

  it('still drops an image node without media or rule', () => {
    expect(inlineToRuns([{ type: 'docInlineImage', attrs: { dataUrl: '', xml: HR_XML } }])).toEqual(
      [],
    )
  })

  it('follows a fixed (exact/atLeast) paragraph line height for the rule line', () => {
    const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
    const rule = /\.doc-inline-rule \{([^}]*)\}/.exec(css)?.[1] ?? ''
    expect(rule).toContain('padding-top: calc(1lh - 0.22em - var(--doc-rule-h, 1px))')
    expect(rule).not.toMatch(/margin-(top|bottom):/)
    // the global border-box sizing would swallow the padded height and paint nothing
    expect(rule).toContain('box-sizing: content-box')
    expect(rule).toContain('background-clip: content-box')
    expect(rule.replace(/\s+/g, ' ')).toContain(
      'line-height: calc( var(--doc-line-factor-latin, var(--doc-line-factor, 1.2)) * var(--doc-line-mult, 1) * 1em );',
    )
    expect(css).toMatch(/\.doc-lh-fixed > \.doc-inline-rule \{\s*line-height: inherit;\s*\}/)
  })

  it('round-trips the rule node through HTML (copy/paste keeps the pict fragment)', () => {
    const source = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'Title' },
              { type: 'docInlineImage', attrs: { dataUrl: '', xml: HR_XML, rule: RULE } },
            ],
          },
        ],
      },
    })
    const html = source.view.dom.innerHTML
    source.destroy()
    const pasted = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: html,
    })
    const para = pasted.state.doc.firstChild!
    const nodes: PmNode[] = []
    para.forEach((n) => nodes.push(n))
    const node = nodes.find((n) => n.type.name === 'docInlineImage')
    expect(node?.attrs).toMatchObject({ dataUrl: '', xml: HR_XML, rule: RULE })
    pasted.destroy()
  })

  it('renders the rule in textbox/table-cell run specs instead of a broken image', () => {
    const specs = runSpanSpecs({
      text: '',
      sizeHalfPoints: 20,
      image: { dataUrl: '', xml: HR_XML, rule: RULE },
    })
    expect(specs).toEqual([
      [
        'span',
        {
          class: 'doc-inline-rule',
          style: '--doc-rule-color:#A0A0A0;--doc-rule-h:2px;font-size:10pt',
        },
      ],
    ])
  })

  it('renders the rule in paginated header/footer cells instead of a broken image', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        {
          runs: [],
          row: { heightTwips: 0, heightRule: 'auto', indentTwips: 0, tabOverflow: 'wrap' },
          cells: [
            {
              paras: [
                [
                  { text: 'Title', bold: true },
                  { text: '', sizeHalfPoints: 20, image: { dataUrl: '', xml: HR_XML, rule: RULE } },
                ],
              ],
              widthTwips: 5000,
            },
          ],
        },
      ],
    } as unknown as HeaderFooter
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 2, pageTotal: 3 })
    expect(el.querySelector('img')).toBeNull()
    const rule = el.querySelector<HTMLElement>('span.doc-inline-rule')!
    expect(rule.style.getPropertyValue('--doc-rule-color')).toBe('#A0A0A0')
    expect(rule.style.fontSize).toBe('10pt')
  })

  it('leaves an untouched rule paragraph byte-identical and keeps rPr on an identity edit', async () => {
    const theme =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
      '<a:themeElements><a:clrScheme name="Office">' +
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
      '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
      '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
      '</a:clrScheme></a:themeElements></a:theme>'
    const hr = (id: string) =>
      '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:pict>' +
      `<v:rect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="${id}" ` +
      'style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0" stroked="f"/>' +
      '</w:pict></w:r>'
    const color = '<w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>'
    const rPr = `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/>${color}<w:sz w:val="28"/></w:rPr>`
    const heading =
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${hr('_x0000_i1025')}` +
      `<w:r>${rPr}<w:t>Séquence</w:t></w:r><w:r>${rPr}<w:t xml:space="preserve"> 1 : La couleur</w:t></w:r>` +
      `${hr('_x0000_i1026')}</w:p>`
    const bytes = await buildDocx({
      bodyXml: heading + '<w:p><w:r><w:t>Objectif</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'word/theme/theme1.xml',
          xml: theme,
          contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        },
      ],
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)

    const noop = pmDocToSavePlan(editor.state.doc.toJSON() as never, parsed.blocks)
    expect(noop.saveBlocks.map((b) => b.kind)).toEqual(['original', 'original'])
    expect(await saveDocx(parsed, noop.saveBlocks)).toBe(bytes)

    // an identity replace leaves the paragraph untouched (still 'original')
    const same = executeOps(editor, [
      { op: 'findReplace', find: 'Séquence', replace: 'Séquence', matchCase: true },
    ] as never)
    expect(same.ok).toBe(true)
    expect(
      pmDocToSavePlan(editor.state.doc.toJSON() as never, parsed.blocks).saveBlocks.map(
        (b) => b.kind,
      ),
    ).toEqual(['original', 'original'])

    // a real edit rebuilds the paragraph: text run rPr verbatim, rule runs whole
    const edit = executeOps(editor, [
      { op: 'findReplace', find: 'Séquence', replace: 'Chapitre', matchCase: true },
    ] as never)
    expect(edit.ok).toBe(true)
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as never, parsed.blocks)
    expect(plan.saveBlocks.map((b) => b.kind)).toEqual(['generated', 'original'])
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    const para = /<w:p[ >][\s\S]*?<\/w:p>/.exec(xml)![0]
    expect(para).toContain(hr('_x0000_i1025'))
    expect(para).toContain(hr('_x0000_i1026'))
    expect(para).toContain(`${rPr}<w:t xml:space="preserve">Chapitre 1 : La couleur</w:t>`)
    expect(para).not.toMatch(/w:color w:val="(?!365F91)/)
    editor.destroy()
  })
})
