import { describe, expect, it } from 'vitest'

import {
  applyTint,
  describeStyleColor,
  echoStyleColor,
  normalizeStyleColor,
  resolveStyleColor,
} from '@genoffice/xlsx-gateway/domain/style-color'
import { StylesheetEditor } from '@genoffice/xlsx-gateway/gateway/xlsx-styles'
import { parseStylesheetFormats } from '@genoffice/xlsx-gateway/gateway/xlsx-style-read'
import { formatOpLabel } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { fromNeutralStyle } from '../src/renderer/edit-journal'

const STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>`

describe('style colors', () => {
  it('parses theme shorthand, names and objects into theme index + tint', () => {
    expect(normalizeStyleColor('accent1')).toEqual({ theme: 4 })
    expect(normalizeStyleColor('accent1+40%')).toEqual({ theme: 4, tint: 0.4 })
    expect(normalizeStyleColor('dk2 - 25%')).toEqual({ theme: 3, tint: -0.25 })
    expect(normalizeStyleColor('tx1')).toEqual({ theme: 1 })
    expect(normalizeStyleColor({ theme: 'accent6', tint: 0.5 })).toEqual({ theme: 9, tint: 0.5 })
    expect(normalizeStyleColor({ theme: 10 })).toEqual({ theme: 10 })
    expect(normalizeStyleColor('#ff00aa')).toBe('#FF00AA')
    expect(() => normalizeStyleColor('accent7')).toThrow(/Unknown color/)
    expect(() => normalizeStyleColor({ theme: 12 })).toThrow(/Unknown theme/)
  })

  it('applies tints on the HLS luminance channel within one step of the Office swatches', () => {
    const swatches: [string, number, string][] = [
      ['#4472C4', 0.4, '#8EA9DB'],
      ['#4472C4', 0.8, '#D9E2F3'],
      ['#4472C4', -0.25, '#2F5597'],
      ['#4472C4', -0.5, '#1F3864'],
      ['#ED7D31', 0.4, '#F4B183'],
      ['#FFFFFF', -0.5, '#7F7F7F'],
      ['#000000', 0.5, '#7F7F7F'],
      ['#70AD47', 0.6, '#C5E0B4'],
      ['#44546A', 0.8, '#D6DCE4'],
    ]
    for (const [base, tint, expected] of swatches) {
      const got = applyTint(base, tint)
      for (let i = 1; i < 7; i += 2) {
        const diff = Math.abs(
          parseInt(got.slice(i, i + 2), 16) - parseInt(expected.slice(i, i + 2), 16),
        )
        expect(diff, `${base} ${tint} → ${got} vs ${expected}`).toBeLessThanOrEqual(1)
      }
    }
    expect(applyTint('#4472C4', 0)).toBe('#4472C4')
    expect(applyTint('#000000', 1)).toBe('#FFFFFF')
    expect(applyTint('#FFFFFF', -1)).toBe('#000000')
  })

  it('resolves against the workbook palette and echoes the slot', () => {
    const palette = Array.from(
      { length: 12 },
      (_, i) => `#${(i * 17).toString(16).padStart(2, '0').repeat(3)}`,
    )
    expect(resolveStyleColor({ theme: 4 }, palette)).toBe('#444444')
    expect(resolveStyleColor('#abcdef')).toBe('#ABCDEF')
    expect(echoStyleColor({ theme: 5, tint: 0.4 })).toMatchObject({ theme: 'accent2', tint: 0.4 })
    expect(echoStyleColor({ theme: 5, tint: 0.4 }).rgb).toMatch(/^#F4B[01]8[234]$/)
    expect(describeStyleColor({ theme: 3, tint: -0.25 })).toBe('dk2-25%')
    expect(describeStyleColor('#FF0000')).toBe('#FF0000')
  })
})

describe('StylesheetEditor theme colors and fills', () => {
  it('writes theme + tint on fonts, fills and borders instead of a baked rgb', () => {
    const editor = new StylesheetEditor(STYLES)
    editor.resolveStyle(0, {
      fontColor: { theme: 1 },
      fillColor: { theme: 4, tint: 0.4 },
      borderBottom: { style: 'thin', color: { theme: 3, tint: -0.25 } },
    })
    const xml = editor.serialize()
    expect(xml).toContain('<font><color theme="1"/><sz val="11"/><name val="Calibri"/></font>')
    expect(xml).toContain(
      '<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/><bgColor indexed="64"/></patternFill></fill>',
    )
    expect(xml).toContain('<bottom style="thin"><color theme="3" tint="-0.25"/></bottom>')
    expect(xml).not.toContain('rgb="FF8EA9DB"')
  })

  it('writes pattern and gradient fills and dedupes them', () => {
    const editor = new StylesheetEditor(STYLES)
    const a = editor.resolveStyle(0, {
      fill: { pattern: 'lightGray', fg: '#FF0000', bg: { theme: 0 } },
    })
    const b = editor.resolveStyle(0, {
      fill: { pattern: 'lightGray', fg: '#FF0000', bg: { theme: 0 } },
    })
    expect(b).toBe(a)
    editor.resolveStyle(0, {
      fill: {
        gradient: {
          angle: 90,
          stops: [
            { position: 0, color: { theme: 4 } },
            { position: 1, color: '#FFFFFF' },
          ],
        },
      },
    })
    editor.resolveStyle(0, {
      fill: {
        gradient: {
          type: 'path',
          left: 0.5,
          right: 0.5,
          top: 0.5,
          bottom: 0.5,
          stops: [
            { position: 0, color: '#FFFFFF' },
            { position: 1, color: '#000000' },
          ],
        },
      },
    })
    const xml = editor.serialize()
    expect(xml).toContain(
      '<fill><patternFill patternType="lightGray"><fgColor rgb="FFFF0000"/><bgColor theme="0"/></patternFill></fill>',
    )
    expect(xml).toContain(
      '<fill><gradientFill degree="90"><stop position="0"><color theme="4"/></stop><stop position="1"><color rgb="FFFFFFFF"/></stop></gradientFill></fill>',
    )
    expect(xml).toContain(
      '<gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5"><stop position="0">',
    )
    expect(xml).toContain('<fills count="5">')
  })

  it('fill wins over fillColor and null clears back to none', () => {
    const editor = new StylesheetEditor(STYLES)
    const index = editor.resolveStyle(0, { fillColor: '#00FF00', fill: null })
    expect(editor.serialize()).toContain(`<cellXfs count="${index === 0 ? 1 : 2}">`)
    expect(editor.serialize()).not.toContain('00FF00')
  })

  it('reads back what it wrote: theme slots, patterns and gradients', () => {
    const editor = new StylesheetEditor(STYLES)
    const themed = editor.resolveStyle(0, {
      fontColor: { theme: 4, tint: 0.4 },
      fillColor: { theme: 5 },
    })
    const patterned = editor.resolveStyle(0, {
      fill: { pattern: 'darkTrellis', fg: '#123456', bg: '#654321' },
    })
    const graded = editor.resolveStyle(0, {
      fill: {
        gradient: {
          angle: 45,
          stops: [
            { position: 0, color: { theme: 4, tint: 0.4 } },
            { position: 1, color: '#FFFFFF' },
          ],
        },
      },
    })
    const stored = parseStylesheetFormats(editor.serialize())
    expect(stored.fontColors[stored.xfs[themed]!.fontId]).toEqual({ theme: 4, tint: 0.4 })
    expect(stored.fills[stored.xfs[themed]!.fillId]).toEqual({ pattern: 'solid', fg: { theme: 5 } })
    expect(stored.fills[stored.xfs[patterned]!.fillId]).toEqual({
      pattern: 'darkTrellis',
      fg: '#123456',
      bg: '#654321',
    })
    expect(stored.fills[stored.xfs[graded]!.fillId]).toEqual({
      gradient: {
        angle: 45,
        stops: [
          { position: 0, color: { theme: 4, tint: 0.4 } },
          { position: 1, color: '#FFFFFF' },
        ],
      },
    })
    expect(stored.fills[0]).toBeNull()
    // fill 1 (gray125) carries no foreground: Excel's automatic color, nothing to echo
    expect(stored.fills[1]).toBeNull()
    expect(stored.fontColors[0]).toEqual({ theme: 1 })
  })
})

describe('app display of theme colors', () => {
  it('maps gateway theme colors and fills to a single rgb for Univer', () => {
    expect(
      fromNeutralStyle({ fontColor: { theme: 1 }, fillColor: { theme: 4, tint: 0.4 } }),
    ).toEqual({
      cl: { rgb: '#000000' },
      bg: { rgb: resolveStyleColor({ theme: 4, tint: 0.4 }) },
    })
    expect(
      fromNeutralStyle({
        fill: {
          gradient: {
            stops: [
              { position: 0, color: '#112233' },
              { position: 1, color: '#FFFFFF' },
            ],
          },
        },
      }),
    ).toMatchObject({ bg: { rgb: '#112233' } })
    expect(fromNeutralStyle({ fill: { pattern: 'lightGray', fg: { theme: 5 } } })).toMatchObject({
      bg: { rgb: '#ED7D31' },
    })
    expect(fromNeutralStyle({ fill: null })).toMatchObject({ bg: { rgb: '' } })
  })

  it('labels theme colors and fills readably', () => {
    expect(
      formatOpLabel({
        op: 'format_range',
        sheetId: 's',
        range: 'A1:B2',
        format: {
          fontColor: 'accent1+40%',
          fill: { pattern: 'lightGray', fg: { theme: 'dk2', tint: -0.25 } },
          border: { type: 'all', color: 'accent2' },
        },
      }),
    ).toBe('A1:B2: font color accent1+40%, fill lightGray dk2-25%, border all accent2')
  })
})
