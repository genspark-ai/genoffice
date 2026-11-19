import { describe, expect, it } from 'vitest'

import { applyTabColor } from '@genoffice/xlsx-gateway/gateway/xlsx-tab-color'

const OPEN = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
const BARE = `${OPEN}<dimension ref="A1"/><sheetData/></worksheet>`

describe('applyTabColor', () => {
  it('creates sheetPr ahead of dimension when the part has none', () => {
    expect(applyTabColor(BARE, '#92D050')).toBe(
      `${OPEN}<sheetPr><tabColor rgb="FF92D050"/></sheetPr><dimension ref="A1"/><sheetData/></worksheet>`,
    )
  })

  it('accepts lowercase and bare hex', () => {
    expect(applyTabColor(BARE, 'ff0000')).toContain('<tabColor rgb="FFFF0000"/>')
    expect(() => applyTabColor(BARE, 'red')).toThrow(/Unsupported tab color/)
  })

  it('keeps existing sheetPr attributes and children, tabColor first', () => {
    const xml = `${OPEN}<sheetPr codeName="Sheet1"><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>`
    expect(applyTabColor(xml, '#FF0000')).toBe(
      `${OPEN}<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>`,
    )
  })

  it('expands a self-closing sheetPr', () => {
    const xml = `${OPEN}<sheetPr filterMode="1"/><sheetData/></worksheet>`
    expect(applyTabColor(xml, '#00FF00')).toBe(
      `${OPEN}<sheetPr filterMode="1"><tabColor rgb="FF00FF00"/></sheetPr><sheetData/></worksheet>`,
    )
  })

  it('replaces an existing tabColor whatever its color model', () => {
    const xml = `${OPEN}<sheetPr><tabColor theme="4" tint="0.4"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>`
    expect(applyTabColor(xml, '#0000FF')).toBe(
      `${OPEN}<sheetPr><tabColor rgb="FF0000FF"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>`,
    )
  })

  it('clearing removes tabColor but keeps the other sheetPr children', () => {
    const xml = `${OPEN}<sheetPr><tabColor rgb="FF92D050"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>`
    expect(applyTabColor(xml, null)).toBe(
      `${OPEN}<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>`,
    )
  })

  it('clearing drops a sheetPr left empty and is a no-op without one', () => {
    const xml = `${OPEN}<sheetPr><tabColor rgb="FF92D050"/></sheetPr><sheetData/></worksheet>`
    expect(applyTabColor(xml, null)).toBe(BARE.replace('<dimension ref="A1"/>', ''))
    expect(applyTabColor(BARE, null)).toBe(BARE)
    const attributed = `${OPEN}<sheetPr codeName="S"><tabColor rgb="FF92D050"/></sheetPr><sheetData/></worksheet>`
    expect(applyTabColor(attributed, null)).toBe(
      `${OPEN}<sheetPr codeName="S"></sheetPr><sheetData/></worksheet>`,
    )
  })
})
