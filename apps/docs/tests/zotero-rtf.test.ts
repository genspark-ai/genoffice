import { describe, expect, it } from 'vitest'
import {
  zoteroCodeFromInstruction,
  zoteroInstructionFromCode,
} from '../src/renderer/zotero/controller'
import { parseZoteroRtf, zoteroRtfToText } from '../src/renderer/zotero/rtf'

describe('Zotero RTF conversion', () => {
  it('decodes Unicode, escaped punctuation, tabs, and paragraph breaks', () => {
    const rtf =
      "{\\rtf1\\ansi\\uc1{\\fonttbl{\\f0 Arial;}}Smith \\'96 Jones\\tab 2024\\par \\u20013?\\u25991?}"
    expect(zoteroRtfToText(rtf)).toBe('Smith – Jones\t2024\n中文')
  })

  it('splits bibliography entries joined with Zotero RTF line continuations', () => {
    const rtf =
      '{\\rtf1\\ansi [1]\\tab First entry\\\r\n' +
      '[2]\\tab Second entry\\\r\n' +
      '[3]\\tab Third entry}'

    expect(
      parseZoteroRtf(rtf).paragraphs.map((paragraph) =>
        paragraph.runs.map((run) => run.text).join(''),
      ),
    ).toEqual(['[1]\tFirst entry', '[2]\tSecond entry', '[3]\tThird entry'])
  })

  it('leaves Zotero plain-text output unchanged', () => {
    expect(zoteroRtfToText('(Smith, 2024)')).toBe('(Smith, 2024)')
  })

  it('preserves character styles and bibliography hanging indents', () => {
    const rtf =
      '{\\rtf1\\ansi\\pard\\li720\\fi-360 Alpha {\\i Journal} {\\super 2}\\par ' +
      '\\pard\\li720\\fi-360 {\\b Beta} {\\sub n}\\par}'

    expect(parseZoteroRtf(rtf)).toEqual({
      paragraphs: [
        {
          indentLeft: 720,
          indentFirstLine: -360,
          runs: [
            { text: 'Alpha ' },
            { text: 'Journal', italic: true },
            { text: ' ' },
            { text: '2', vertAlign: 'superscript' },
          ],
        },
        {
          indentLeft: 720,
          indentFirstLine: -360,
          runs: [
            { text: 'Beta', bold: true },
            { text: ' ' },
            { text: 'n', vertAlign: 'subscript' },
          ],
        },
      ],
    })
  })

  it('handles Zotero style resets, sizing, spacing, and alignment', () => {
    const parsed = parseZoteroRtf(
      '{\\rtf1\\pard\\qr\\ri120\\sb40\\sa80\\fs20 ' +
        '{\\ul Under} {\\strike Strike} {\\scaps Small} plain\\par}',
    )

    expect(parsed.paragraphs[0]).toMatchObject({
      align: 'right',
      indentRight: 120,
      spaceBefore: 40,
      spaceAfter: 80,
    })
    expect(parsed.paragraphs[0].runs).toEqual([
      { text: 'Under', underline: true, sizeHalfPoints: 20 },
      { text: ' ', sizeHalfPoints: 20 },
      { text: 'Strike', strike: true, sizeHalfPoints: 20 },
      { text: ' ', sizeHalfPoints: 20 },
      { text: 'Small', caps: 'small', sizeHalfPoints: 20 },
      { text: ' plain', sizeHalfPoints: 20 },
    ])
  })
})

describe('Zotero field code compatibility', () => {
  it('maps between the wire code and Word ADDIN field instruction', () => {
    const code = 'ITEM CSL_CITATION {"citationID":"abc"}'
    expect(zoteroInstructionFromCode(code)).toBe(`ADDIN ZOTERO_${code}`)
    expect(zoteroCodeFromInstruction(` ADDIN ZOTERO_${code} `)).toBe(code)
  })

  it('supports older ZOTERO_ and CSL_ prefixes', () => {
    expect(zoteroCodeFromInstruction('ZOTERO_BIBL {} CSL_BIBLIOGRAPHY')).toBe(
      'BIBL {} CSL_BIBLIOGRAPHY',
    )
    expect(zoteroCodeFromInstruction('CSL_TEMP')).toBe('TEMP')
  })
})
