import { describe, expect, it } from 'vitest'

import {
  applyCfRules,
  buildDxfXml,
  cfRuleUnsaveableReason,
  iconSetSaveable,
} from '../src/gateway/xlsx-cf'

const SHEET =
  '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  '<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>' +
  '<pageMargins left="0.7"/></worksheet>'

class FakeDxfs {
  readonly entries: string[] = []
  internDxf(dxfXml: string): number {
    const existing = this.entries.indexOf(dxfXml)
    if (existing !== -1) return existing
    this.entries.push(dxfXml)
    return this.entries.length - 1
  }
}

const range = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }

describe('applyCfRules', () => {
  it('replaces every existing section with the snapshot', () => {
    const dxfs = new FakeDxfs()
    const xml = applyCfRules(
      SHEET,
      [
        {
          ranges: [range],
          stopIfTrue: false,
          rule: {
            type: 'highlightCell',
            subType: 'number',
            operator: 'greaterThan',
            value: 5,
            style: { bg: { rgb: '#FFF2CC' } },
          },
        },
      ],
      dxfs,
    )
    expect(xml).not.toContain('operator="lessThan"')
    expect(xml).toContain(
      '<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting><pageMargins',
    )
    expect(dxfs.entries[0]).toBe(
      '<dxf><fill><patternFill><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf>',
    )
  })

  it('removes all sections when the snapshot is empty', () => {
    const xml = applyCfRules(SHEET, [], new FakeDxfs())
    expect(xml).not.toContain('conditionalFormatting')
  })

  it('serializes text, dedup, rank, average, and expression rules', () => {
    const dxfs = new FakeDxfs()
    const style = { cl: { rgb: '#C00000' }, bl: 1 }
    const xml = applyCfRules(
      SHEET,
      [
        {
          ranges: [range],
          stopIfTrue: true,
          rule: {
            type: 'highlightCell',
            subType: 'text',
            operator: 'containsText',
            value: 'a"b',
            style,
          },
        },
        {
          ranges: [range],
          stopIfTrue: false,
          rule: { type: 'highlightCell', subType: 'duplicateValues', style },
        },
        {
          ranges: [range],
          stopIfTrue: false,
          rule: {
            type: 'highlightCell',
            subType: 'rank',
            isBottom: true,
            isPercent: false,
            value: 10,
            style,
          },
        },
        {
          ranges: [range],
          stopIfTrue: false,
          rule: { type: 'highlightCell', subType: 'average', operator: 'lessThanOrEqual', style },
        },
        {
          ranges: [range],
          stopIfTrue: false,
          rule: { type: 'highlightCell', subType: 'formula', value: '=$A1>5', style },
        },
      ],
      dxfs,
    )
    expect(xml).toContain(
      'type="containsText" dxfId="0" priority="1" stopIfTrue="1" operator="containsText" text="a&quot;b"',
    )
    expect(xml).toContain('<formula>NOT(ISERROR(SEARCH("a""b",A1)))</formula>')
    expect(xml).toContain('<cfRule type="duplicateValues" dxfId="0" priority="2"/>')
    expect(xml).toContain('type="top10" dxfId="0" priority="3" bottom="1" rank="10"')
    expect(xml).toContain(
      'type="aboveAverage" dxfId="0" priority="4" aboveAverage="0" equalAverage="1"',
    )
    expect(xml).toContain(
      '<cfRule type="expression" dxfId="0" priority="5"><formula>$A1&gt;5</formula></cfRule>',
    )
    // One shared style interns once.
    expect(dxfs.entries).toHaveLength(1)
    expect(dxfs.entries[0]).toBe('<dxf><font><b/><color rgb="FFC00000"/></font></dxf>')
  })

  it('serializes color scales, data bars, and icon sets', () => {
    const dxfs = new FakeDxfs()
    const xml = applyCfRules(
      SHEET,
      [
        {
          ranges: [range],
          stopIfTrue: false,
          rule: {
            type: 'colorScale',
            config: [
              { index: 0, color: '#FF0000', value: { type: 'min' } },
              { index: 1, color: '#FFFF00', value: { type: 'percentile', value: 50 } },
              { index: 2, color: '#00FF00', value: { type: 'max' } },
            ],
          },
        },
        {
          ranges: [range],
          stopIfTrue: false,
          rule: {
            type: 'dataBar',
            isShowValue: false,
            config: {
              min: { type: 'min' },
              max: { type: 'max' },
              isGradient: true,
              positiveColor: '#638EC6',
              nativeColor: '#FF0000',
            },
          },
        },
        {
          ranges: [range],
          stopIfTrue: false,
          rule: {
            type: 'iconSet',
            isShowValue: true,
            config: [
              {
                iconType: '3TrafficLights1',
                iconId: '0',
                operator: 'greaterThanOrEqual',
                value: { type: 'percent', value: 67 },
              },
              {
                iconType: '3TrafficLights1',
                iconId: '1',
                operator: 'greaterThanOrEqual',
                value: { type: 'percent', value: 33 },
              },
              {
                iconType: '3TrafficLights1',
                iconId: '2',
                operator: 'greaterThanOrEqual',
                value: { type: 'min' },
              },
            ],
          },
        },
      ],
      dxfs,
    )
    expect(xml).toContain(
      '<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
        '<color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale>',
    )
    expect(xml).toContain(
      '<dataBar showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>',
    )
    expect(xml).toContain(
      '<iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>',
    )
  })

  it('maps the worst-first rating sets and strict thresholds', () => {
    const xml = applyCfRules(
      SHEET,
      [
        {
          ranges: [range],
          stopIfTrue: false,
          rule: {
            type: 'iconSet',
            isShowValue: true,
            config: [
              {
                iconType: '4Rating',
                iconId: '3',
                operator: 'greaterThanOrEqual',
                value: { type: 'percent', value: 75 },
              },
              {
                iconType: '4Rating',
                iconId: '2',
                operator: 'greaterThanOrEqual',
                value: { type: 'percent', value: 50 },
              },
              {
                iconType: '4Rating',
                iconId: '1',
                operator: 'greaterThan',
                value: { type: 'percent', value: 25 },
              },
              {
                iconType: '4Rating',
                iconId: '0',
                operator: 'greaterThanOrEqual',
                value: { type: 'min' },
              },
            ],
          },
        },
      ],
      new FakeDxfs(),
    )
    expect(xml).toContain(
      '<iconSet iconSet="4Rating"><cfvo type="percent" val="0"/><cfvo type="percent" val="25" gte="0"/>' +
        '<cfvo type="percent" val="50"/><cfvo type="percent" val="75"/></iconSet>',
    )
  })

  it('flags icon configs the base schema cannot hold', () => {
    const entry = (iconType: string, iconId: string) => ({ iconType, iconId })
    expect(
      iconSetSaveable([entry('3Arrows', '0'), entry('3Arrows', '1'), entry('3Arrows', '2')]),
    ).toBe(true)
    expect(
      iconSetSaveable([
        entry('4Rating', '3'),
        entry('4Rating', '2'),
        entry('4Rating', '1'),
        entry('4Rating', '0'),
      ]),
    ).toBe(true)
    expect(
      iconSetSaveable([entry('3Stars', '0'), entry('3Stars', '1'), entry('3Stars', '2')]),
    ).toBe(false)
    expect(
      iconSetSaveable([entry('3Arrows', '0'), entry('3Flags', '1'), entry('3Arrows', '2')]),
    ).toBe(false)
    expect(
      iconSetSaveable([entry('3Arrows', '0'), entry('3Arrows', '2'), entry('3Arrows', '1')]),
    ).toBe(false)
  })

  it('fails closed on Univer-only icon sets and time periods', () => {
    expect(() =>
      applyCfRules(
        SHEET,
        [
          {
            ranges: [range],
            stopIfTrue: false,
            rule: {
              type: 'iconSet',
              config: [
                { iconType: '3Stars', iconId: '0', value: { type: 'min' } },
                { iconType: '3Stars', iconId: '1', value: { type: 'max' } },
              ],
            },
          },
        ],
        new FakeDxfs(),
      ),
    ).toThrow(/icon set/)
    expect(() =>
      applyCfRules(
        SHEET,
        [
          {
            ranges: [range],
            stopIfTrue: false,
            rule: { type: 'highlightCell', subType: 'timePeriod', operator: 'today', style: {} },
          },
        ],
        new FakeDxfs(),
      ),
    ).toThrow(/Date-occurring/)
  })
})

const X14_EXT =
  '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"' +
  ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:conditionalFormattings>' +
  '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
  '<x14:cfRule type="iconSet" priority="1" id="{X14-A}"><x14:iconSet iconSet="3Stars"/></x14:cfRule>' +
  '<x14:cfRule type="iconSet" priority="3" id="{X14-B}"><x14:iconSet iconSet="5Boxes"/></x14:cfRule>' +
  '<xm:sqref>B1:B5</xm:sqref>' +
  '</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>'

const LINKED_BAR =
  '<conditionalFormatting sqref="C1:C5">' +
  '<cfRule type="dataBar" priority="2">' +
  '<dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' +
  '<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"' +
  ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:id>{DATA-BAR-GUID}</x14:id></ext></extLst>' +
  '</cfRule></conditionalFormatting>'

const X14_SHEET = SHEET.replace('</worksheet>', `${X14_EXT}</worksheet>`)
const barRange = { startRow: 0, endRow: 4, startColumn: 2, endColumn: 2 }
const barRule = {
  ranges: [barRange],
  stopIfTrue: false,
  rule: {
    type: 'dataBar',
    config: { min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6' },
  },
}
const highlight = (ranges: (typeof range)[]) => ({
  ranges,
  stopIfTrue: false,
  rule: {
    type: 'highlightCell',
    subType: 'number',
    operator: 'greaterThan',
    value: 5,
    style: { bg: { rgb: '#FFF2CC' } },
  },
})

describe('applyCfRules with x14 extensions', () => {
  it('edits plain rules and keeps the x14 extLst byte-identical', () => {
    const xml = applyCfRules(X14_SHEET, [highlight([range])], new FakeDxfs())
    expect(xml).toContain(X14_EXT)
    expect(xml).not.toContain('operator="lessThan"')
    expect(xml).toContain('<cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan">')
    const removedAll = applyCfRules(X14_SHEET, [], new FakeDxfs())
    expect(removedAll).toContain(X14_EXT)
    expect(removedAll).not.toContain('<conditionalFormatting ')
  })

  it('assigns priorities around the values x14 rules occupy', () => {
    const xml = applyCfRules(
      X14_SHEET,
      [highlight([range]), highlight([{ ...range, startColumn: 3, endColumn: 3 }])],
      new FakeDxfs(),
    )
    expect(xml).toContain('priority="2" operator="greaterThan"')
    expect(xml).toContain('priority="4" operator="greaterThan"')
    expect(xml).not.toMatch(/<cfRule type="cellIs"[^>]* priority="[13]"/)
  })

  it('passes an unchanged x14-linked data bar through verbatim and avoids its priority', () => {
    const sheet = X14_SHEET.replace('<pageMargins', `${LINKED_BAR}<pageMargins`)
    const dxfs = new FakeDxfs()
    const xml = applyCfRules(sheet, [barRule, highlight([range])], dxfs)
    expect(xml).toContain(LINKED_BAR)
    expect(xml).toContain(X14_EXT)
    // 1 and 3 are x14, 2 is the preserved bar — the new rule gets 4.
    expect(xml).toContain('<cfRule type="cellIs" dxfId="0" priority="4" operator="greaterThan">')
    expect(xml.split('<conditionalFormatting sqref="C1:C5">')).toHaveLength(2)
  })

  it('rejects modifying, retyping, or deleting the x14-linked data bar rule', () => {
    const sheet = X14_SHEET.replace('<pageMargins', `${LINKED_BAR}<pageMargins`)
    const recolored = {
      ...barRule,
      rule: { ...barRule.rule, config: { ...barRule.rule.config, positiveColor: '#FF0000' } },
    }
    expect(() => applyCfRules(sheet, [recolored], new FakeDxfs())).toThrow(
      /data-bar extension format/,
    )
    expect(() => applyCfRules(sheet, [highlight([barRange])], new FakeDxfs())).toThrow(
      /data-bar extension format/,
    )
    expect(() => applyCfRules(sheet, [highlight([range])], new FakeDxfs())).toThrow(
      /data-bar extension format/,
    )
    expect(() => applyCfRules(sheet, [], new FakeDxfs())).toThrow(/data-bar extension format/)
  })

  it('appends new dxfs without disturbing entries x14 rules may reference', () => {
    const dxfs = new FakeDxfs()
    dxfs.entries.push('<dxf><font><b/></font></dxf>', '<dxf><fill/></dxf>')
    const xml = applyCfRules(X14_SHEET, [highlight([range])], dxfs)
    expect(dxfs.entries.slice(0, 2)).toEqual(['<dxf><font><b/></font></dxf>', '<dxf><fill/></dxf>'])
    expect(xml).toContain('dxfId="2"')
  })
})

describe('buildDxfXml', () => {
  it('maps the Univer highlight style onto dxf font and fill', () => {
    expect(
      buildDxfXml({
        bl: 1,
        it: 1,
        ul: { s: 1 },
        st: { s: 1 },
        cl: { rgb: '#112233' },
        bg: { rgb: 'AABBCC' },
      }),
    ).toBe(
      '<dxf><font><b/><i/><strike/><u/><color rgb="FF112233"/></font>' +
        '<fill><patternFill><bgColor rgb="FFAABBCC"/></patternFill></fill></dxf>',
    )
    expect(buildDxfXml(undefined)).toBe('<dxf></dxf>')
  })
})

describe('cfRuleUnsaveableReason', () => {
  it('accepts every rule shape the save path serializes', () => {
    const saveable = [
      {
        type: 'highlightCell',
        subType: 'number',
        operator: 'greaterThan',
        value: 5,
        style: { bg: { rgb: '#FFF2CC' } },
      },
      {
        type: 'highlightCell',
        subType: 'text',
        operator: 'containsText',
        value: 'x',
        style: {},
      },
      { type: 'highlightCell', subType: 'average', operator: 'greaterThan', style: {} },
      { type: 'highlightCell', subType: 'formula', value: '=A1>0', style: {} },
      {
        type: 'colorScale',
        config: [
          { index: 0, color: '#FFFFFF', value: { type: 'min' } },
          { index: 1, color: '#63BE7B', value: { type: 'max' } },
        ],
      },
    ]
    for (const rule of saveable) {
      expect(cfRuleUnsaveableReason(rule)).toBeNull()
    }
  })

  it('reports the save-side reason for rules that would fail the save', () => {
    expect(
      cfRuleUnsaveableReason({ type: 'highlightCell', subType: 'timePeriod', style: {} }),
    ).toMatch(/Date-occurring/)
    expect(
      cfRuleUnsaveableReason({
        type: 'highlightCell',
        subType: 'average',
        operator: 'equal',
        style: {},
      }),
    ).toMatch(/cannot be saved/)
    expect(cfRuleUnsaveableReason({ type: 'somethingElse' })).toMatch(/Unsupported/)
  })
})
