import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  formValuesOf,
  upsertFromForm,
  type StyleFormValues,
} from '../src/renderer/components/StyleDialog'

// word/styles.xml of a document whose one Body style carries the given rPr XML
const parseWithStyle = async (rPr: string) =>
  parseDocx(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>\u4e2d\u6587 English</w:t></w:r></w:p>',
      stylesXml:
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body Text"/>' +
        `<w:rPr>${rPr}</w:rPr></w:style></w:styles>`,
    }),
  )

const formValues = (over: Partial<StyleFormValues> = {}): StyleFormValues => ({
  name: 'Body Text',
  type: 'paragraph',
  basedOn: null,
  font: 'Times New Roman',
  eastAsiaFont: 'SimSun',
  sizePt: null,
  bold: false,
  italic: false,
  color: '',
  align: null,
  spaceBeforePt: null,
  spaceAfterPt: null,
  lineSpacing: null,
  ...over,
})

describe('style dialog East Asian font', () => {
  it('a style with East Asian and Latin fonts round-trips docx save and reopen', async () => {
    const parsed = await parseWithStyle(
      '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="SimSun"/>',
    )
    // open parses the pair into the dialog
    const values = formValuesOf(parsed.styles.get('Body')!)
    expect(values.font).toBe('Times New Roman')
    expect(values.eastAsiaFont).toBe('SimSun')
    // editing both pickers persists the pair into styles.xml
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      styleUpserts: [
        upsertFromForm('Body', { ...values, font: 'Arial', eastAsiaFont: 'KaiTi' }, 'all', false),
      ],
    })
    const stylesXml = await (await JSZip.loadAsync(saved)).file('word/styles.xml')!.async('string')
    expect(stylesXml).toContain('w:ascii="Arial"')
    expect(stylesXml).toContain('w:eastAsia="KaiTi"')
    const reopened = await parseDocx(saved)
    expect(reopened.styles.get('Body')!.display).toMatchObject({
      fontAscii: 'Arial',
      eastAsiaFont: 'KaiTi',
    })
  })

  it('the dialog writes both font slots into the style upsert', () => {
    const up = upsertFromForm('Body', formValues(), 'all', true)
    expect(up.rPr).toMatchObject({ font: 'Times New Roman', eastAsiaFont: 'SimSun' })
    // only the field the user touched is patched
    const touched = new Set<keyof StyleFormValues>(['eastAsiaFont'])
    expect(upsertFromForm('Body', formValues(), touched, false).rPr).toEqual({
      eastAsiaFont: 'SimSun',
    })
    // clearing the East Asian picker clears the slot
    expect(upsertFromForm('Body', formValues({ eastAsiaFont: '' }), touched, false).rPr).toEqual({
      eastAsiaFont: null,
    })
  })

  it('a style without an East Asian font opens with the picker empty', async () => {
    const parsed = await parseWithStyle(
      '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>',
    )
    const values = formValuesOf(parsed.styles.get('Body')!)
    expect(values.font).toBe('Times New Roman')
    expect(values.eastAsiaFont).toBe('')
  })
})
