import { describe, expect, it } from 'vitest'

import { claimsSelection, verifySheetsResponse } from '../src/renderer/ai/response-verify'

const SELECT_OK = [{ name: 'select_range', ok: true }]
const SELECT_FAILED = [{ name: 'select_range', ok: false }]

describe('claimsSelection', () => {
  it('detects Chinese selection claims', () => {
    expect(claimsSelection('已把选区定位到 Kira 那一行方便你查看。')).toBe(true)
    expect(claimsSelection('已选中 Kira 所在的整行 A15:C15。')).toBe(true)
    expect(claimsSelection('已经为你高亮了汇总行。')).toBe(true)
    expect(claimsSelection('已為您選中 A15:C15。')).toBe(true)
    expect(claimsSelection('已跳转到 Sheet2 的 B5。')).toBe(true)
  })

  it('detects English selection claims', () => {
    expect(claimsSelection('I have selected row 15 for you.')).toBe(true)
    expect(claimsSelection("I've now highlighted the total row.")).toBe(true)
    expect(claimsSelection('The selection has been moved to A15:C15.')).toBe(true)
  })

  it('ignores attributive references to the existing selection', () => {
    // the attributive "already-selected range" form refers to the user's own
    // selection, not a claimed action
    expect(claimsSelection('已选中的区域将被加粗。')).toBe(false)
    expect(claimsSelection('对已选中的单元格应用了货币格式。')).toBe(false)
  })

  it('ignores ordinary answers without selection claims', () => {
    expect(claimsSelection('Kira 排名第一，共有 7 个产品。')).toBe(false)
    expect(claimsSelection('You can select the row with ⌘G if you like.')).toBe(false)
    expect(claimsSelection('已将 B2 的值改为 42。')).toBe(false)
  })
})

describe('verifySheetsResponse', () => {
  it('returns a correction when a claim has no successful select_range behind it', () => {
    const correction = verifySheetsResponse('已选中 Kira 所在的整行 A15:C15。', [])
    expect(correction).toContain('select_range')
    expect(correction).toContain('NOT moved')
  })

  it('treats a failed select_range like no call at all', () => {
    expect(verifySheetsResponse('已定位到 A15:C15。', SELECT_FAILED)).not.toBeNull()
  })

  it('accepts the claim when select_range succeeded during the run', () => {
    expect(verifySheetsResponse('已选中 Kira 所在的整行 A15:C15。', SELECT_OK)).toBeNull()
  })

  it('accepts replies without any selection claim regardless of calls', () => {
    expect(verifySheetsResponse('Kira 排名第一。', [])).toBeNull()
  })
})
