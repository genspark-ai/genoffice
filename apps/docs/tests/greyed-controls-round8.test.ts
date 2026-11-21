import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import {
  addSwitch,
  fieldResultKind,
  formatFieldNumber,
  formatFieldText,
  formatNumericPicture,
} from '../src/renderer/editor/field-switches'
import {
  detectLanguages,
  mergedLangSlots,
  setDetectLanguageEnabled,
} from '../src/renderer/editor/language-detect'
import { insertHorizontalLine } from '../src/renderer/editor/horizontal-line'
import { unstyleDocument } from '../src/renderer/editor/style-delete'
import {
  createLabelSheet,
  hasLabelSheet,
  joinLabelSheets,
  labelTableIndex,
  updateLabels,
} from '../src/renderer/mail-merge/label-sheet'
import { customGridLayout } from '../src/renderer/mail-merge/layout'
import { mergeDocuments } from '../src/renderer/mail-merge/template'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const editorOf = (content: PmNode[]) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
const para = (text: string, attrs?: Record<string, unknown>): PmNode => ({
  type: 'docParagraph',
  ...(attrs ? { attrs } : {}),
  content: text ? [{ type: 'text', text }] : [],
})

describe('field switches', () => {
  it('formats a number by the \\* switch', () => {
    const page = (name: string, n: number) => formatFieldNumber(`PAGE \\* ${name}`, n)
    expect(page('roman', 14)).toBe('xiv')
    expect(page('ROMAN', 2024)).toBe('MMXXIV')
    expect(page('alphabetic', 28)).toBe('bb')
    expect(page('ALPHABETIC', 3)).toBe('C')
    expect([1, 2, 3, 11, 22, 113].map((n) => page('Ordinal', n))).toEqual([
      '1st',
      '2nd',
      '3rd',
      '11th',
      '22nd',
      '113th',
    ])
    expect(page('CardText', 121)).toBe('one hundred twenty-one')
    expect(page('OrdText', 21)).toBe('twenty-first')
    expect(page('OrdText', 40)).toBe('fortieth')
    expect(page('Hex', 255)).toBe('FF')
    expect(page('DollarText', 12)).toBe('twelve and 00/100')
    expect(page('Arabic', 7)).toBe('7')
    expect(formatFieldNumber('PAGE \\* MERGEFORMAT', 7)).toBe('7')
  })

  it('fills a \\# picture', () => {
    expect(formatNumericPicture(5, '00')).toBe('05')
    expect(formatNumericPicture(1234.5, '#,##0.00')).toBe('1,234.50')
    expect(formatNumericPicture(-3, '$#,##0.00;($#,##0.00)')).toBe('($3.00)')
    expect(formatNumericPicture(12, '0%')).toBe('12%')
    expect(formatFieldNumber('NUMPAGES \\# "0.00"', 3)).toBe('3.00')
  })

  it('capitalizes a text result', () => {
    expect(formatFieldText('FILENAME \\* Upper', 'my report.docx')).toBe('MY REPORT.DOCX')
    expect(formatFieldText('FILENAME \\* Caps', 'my report')).toBe('My Report')
    expect(formatFieldText('FILENAME \\* FirstCap', 'my REPORT')).toBe('My report')
    expect(formatFieldText('FILENAME \\* MERGEFORMAT', 'a.docx')).toBe('a.docx')
  })

  it('adds a switch ahead of MERGEFORMAT and replaces one of its kind', () => {
    expect(addSwitch('PAGE \\* MERGEFORMAT', '\\* roman')).toBe('PAGE \\* roman \\* MERGEFORMAT')
    expect(addSwitch('PAGE \\* roman \\* MERGEFORMAT', '\\* Ordinal')).toBe(
      'PAGE \\* Ordinal \\* MERGEFORMAT',
    )
    expect(addSwitch('DATE \\@ "d MMMM yyyy"', '\\@ "yyyy-MM-dd"')).toBe('DATE \\@ "yyyy-MM-dd"')
    expect(addSwitch('FILENAME', '\\p')).toBe('FILENAME \\p')
  })

  it('offers options only for fields whose result is worked out here', () => {
    expect(fieldResultKind('page')).toBe('number')
    expect(fieldResultKind('DATE \\@ "yyyy"')).toBe('date')
    expect(fieldResultKind('MERGEFIELD Name')).toBe('text')
    expect(fieldResultKind('TOC \\o "1-3"')).toBeNull()
    expect(fieldResultKind('')).toBeNull()
  })
})

describe('language detection', () => {
  it('tells the common languages apart, one per w:lang slot', () => {
    expect(detectLanguages('The quick brown fox was not in the house that they had')).toEqual({
      val: 'en-US',
    })
    expect(
      detectLanguages('Le chat est dans la maison et il ne veut pas sortir avec nous'),
    ).toEqual({ val: 'fr-FR' })
    expect(detectLanguages('Der Hund ist nicht in dem Haus und er wird auch nicht kommen')).toEqual(
      {
        val: 'de-DE',
      },
    )
    expect(detectLanguages('\u3053\u308c\u306f\u65e5\u672c\u8a9e\u3067\u3059')).toEqual({
      eastAsia: 'ja-JP',
    })
    expect(detectLanguages('\ud55c\uad6d\uc5b4 \ubb38\uc7a5')).toEqual({ eastAsia: 'ko-KR' })
    expect(
      detectLanguages('\u8fd9\u662f\u4e2d\u6587 and the text was in the house that they had'),
    ).toEqual({ eastAsia: 'zh-CN', val: 'en-US' })
    expect(detectLanguages('\u041f\u0440\u0438\u0432\u0435\u0442 \u043c\u0438\u0440')).toEqual({
      val: 'ru-RU',
    })
    expect(detectLanguages('ok')).toEqual({})
  })

  it('keeps a variant the text already carries and reports no change', () => {
    expect(mergedLangSlots({ val: 'en-GB' }, { val: 'en-US' })).toBeNull()
    expect(mergedLangSlots({ val: 'en-GB' }, { val: 'fr-FR' })).toEqual({ val: 'fr-FR' })
    expect(mergedLangSlots(null, { eastAsia: 'ja-JP' })).toEqual({ eastAsia: 'ja-JP' })
  })

  it('marks the paragraph being typed only while the option is on', () => {
    const sentence = 'Le chat est dans la maison et il ne veut pas sortir avec nous'
    const langOf = (editor: Editor) =>
      editor.state.doc.firstChild?.firstChild?.marks.find((m) => m.type.name === 'docTextStyle')
        ?.attrs.lang ?? null
    setDetectLanguageEnabled(false)
    const off = editorOf([para('')])
    off.commands.insertContent(sentence)
    expect(langOf(off)).toBeNull()
    setDetectLanguageEnabled(true)
    const on = editorOf([para('')])
    on.commands.insertContent(sentence)
    expect(JSON.parse(String(langOf(on)))).toEqual({ val: 'fr-FR' })
    setDetectLanguageEnabled(false)
  })
})

describe('deleting a style from the text', () => {
  it('paragraphs go back to the default style and a styled heading to body text', () => {
    const editor = editorOf([
      para('a', { styleId: 'MyQuote' }),
      {
        type: 'docHeading',
        attrs: { level: 1, styleId: 'MyQuote' },
        content: [{ type: 'text', text: 'h' }],
      },
      para('b', { styleId: 'Other' }),
    ])
    expect(unstyleDocument(editor, 'MyQuote', 'paragraph')).toBe(true)
    const blocks = editor.getJSON().content!
    expect(blocks.map((b) => [b.type, b.attrs?.styleId ?? null])).toEqual([
      ['docParagraph', null],
      ['docParagraph', null],
      ['docParagraph', 'Other'],
    ])
    expect(unstyleDocument(editor, 'MyQuote', 'paragraph')).toBe(false)
  })

  it('a character style comes off the runs that carry it', () => {
    const editor = editorOf([
      {
        type: 'docParagraph',
        content: [
          { type: 'text', text: 'x', marks: [{ type: 'docTextStyle', attrs: { styleId: 'Em' } }] },
          {
            type: 'text',
            text: 'y',
            marks: [{ type: 'docTextStyle', attrs: { styleId: 'Keep' } }],
          },
        ],
      },
    ])
    expect(unstyleDocument(editor, 'Em', 'character')).toBe(true)
    const ids: unknown[] = []
    editor.state.doc.descendants((n) => {
      if (n.isText)
        ids.push(n.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.styleId ?? null)
    })
    expect(ids).toEqual([null, 'Keep'])
  })
})

describe('Horizontal Line', () => {
  it('goes into an empty paragraph, or a new one below, and is saved as Word writes it', async () => {
    const editor = editorOf([para('before'), para('')])
    editor.commands.setTextSelection(3)
    expect(insertHorizontalLine(editor)).toBe(true)
    let blocks = editor.getJSON().content!
    expect(blocks.map((b) => b.content?.[0]?.type ?? null)).toEqual([
      'text',
      'docInlineImage',
      null,
    ])
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertHorizontalLine(editor)
    blocks = editor.getJSON().content!
    expect(blocks).toHaveLength(3)
    expect(blocks[2].content?.[0]?.type).toBe('docInlineImage')

    const parsed = await parseDocx(await buildBlankDocx())
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml.match(/o:hr="t"/g)).toHaveLength(2)
    expect(xml).toContain('o:hrstd="t"')
    const reopened = await parseDocx(saved)
    const rules = blocksToPmDoc(reopened.blocks).content!.filter(
      (b) => (b.content?.[0]?.attrs?.rule as { thicknessPx?: number } | null)?.thicknessPx === 2,
    )
    expect(rules).toHaveLength(2)
  })
})

describe('labels main document', () => {
  const field = (name: string): PmNode => ({
    type: 'text',
    text: `\u00ab${name}\u00bb`,
    marks: [{ type: 'instrField', attrs: { instr: `MERGEFIELD ${name}` } }],
  })
  const layout = customGridLayout('letter', 2, 2)
  const labelTexts = (doc: PmNode): string[] => {
    const table = doc.content![0]
    const text = (n: PmNode): string => n.text ?? (n.content ?? []).map(text).join('')
    return table
      .content!.flatMap((row) => row.content ?? [])
      .map((cell) => text(cell).trim())
      .filter((t) => t !== '')
  }

  it('Create Label Sheet lays the document out as labels that follow the first', () => {
    const editor = editorOf([{ type: 'docParagraph', content: [field('Name')] }])
    expect(hasLabelSheet(editor.state.doc)).toBe(false)
    expect(createLabelSheet(editor, layout)).toBe(true)
    expect(hasLabelSheet(editor.state.doc)).toBe(true)
    const doc = editor.getJSON() as PmNode
    expect(labelTableIndex(doc)).toBe(0)
    expect(labelTexts(doc)).toHaveLength(4)
    const ids = new Set<unknown>()
    editor.state.doc.descendants((n) => {
      const m = n.marks.find((x) => x.type.name === 'instrField' && x.attrs.instr === 'NEXT')
      if (m) ids.add(m.attrs.fieldId)
    })
    expect(ids.size).toBe(3)
    // made again on another grid, the sheet keeps the first label
    expect(createLabelSheet(editor, customGridLayout('letter', 3, 2))).toBe(true)
    const again = labelTexts(editor.getJSON() as PmNode)
    expect(again).toHaveLength(6)
    expect(new Set(again).size).toBe(1)
  })

  it('Update Labels copies the first label into the others, each after its NEXT field', () => {
    const editor = editorOf([{ type: 'docParagraph', content: [field('Name')] }])
    createLabelSheet(editor, layout)
    // the first label gains a second field
    let firstCellEnd = 0
    editor.state.doc.descendants((node, pos) => {
      if (!firstCellEnd && node.type.name === 'docParagraph') firstCellEnd = pos + node.nodeSize - 1
    })
    editor
      .chain()
      .setTextSelection(firstCellEnd)
      .insertContent([{ type: 'text', text: ' ' }, field('City')])
      .run()
    expect(updateLabels(editor)).toBe(true)
    const texts = labelTexts(editor.getJSON() as PmNode)
    expect(texts).toHaveLength(4)
    expect(new Set(texts)).toEqual(new Set(['\u00abName\u00bb \u00abCity\u00bb']))
    // a second update changes nothing more
    const before = JSON.stringify(labelTexts(editor.getJSON() as PmNode))
    updateLabels(editor)
    expect(JSON.stringify(labelTexts(editor.getJSON() as PmNode))).toBe(before)
    expect(updateLabels(editorOf([para('plain')]))).toBe(false)
  })

  it('the merge fills each label with the next record and adds a sheet when one is full', () => {
    const editor = editorOf([{ type: 'docParagraph', content: [field('Name')] }])
    createLabelSheet(editor, layout)
    const snapshot = editor.getJSON() as PmNode
    const records = ['A', 'B', 'C', 'D', 'E', 'F'].map((Name) => ({ Name }))
    const merged = mergeDocuments(snapshot, records).documents.map((d) => d.doc)
    expect(merged).toHaveLength(2)
    const sheet = joinLabelSheets(merged, labelTableIndex(snapshot))!
    expect(labelTexts(sheet)).toEqual(['A', 'B', 'C', 'D', 'E', 'F'])
    const cells = sheet.content![0].content!.flatMap((row) => row.content ?? [])
    expect(cells.every((cell) => (cell.content ?? []).length > 0)).toBe(true)
  })
})
