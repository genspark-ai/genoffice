import { describe, it, expect } from 'vitest'
import {
  extractJsonObject,
  escapeBrokenStringLiterals,
  parseOutlineJson,
} from '../src/renderer/ai/outline-json'

describe('parseOutlineJson', () => {
  it('parses valid JSON', () => {
    const obj = parseOutlineJson(
      '{"core_hook":"hook","pages":[{"title":"Cover","brief":"b","layout":"cover","image_queries":[]}]}',
    )
    expect(obj?.core_hook).toBe('hook')
    expect(Array.isArray(obj?.pages)).toBe(true)
  })

  it('strips markdown fences', () => {
    const obj = parseOutlineJson('```json\n{"core_hook":"h","pages":[]}\n```')
    expect(obj?.core_hook).toBe('h')
  })

  it('repairs unescaped Chinese quotes inside brief (reproduces position-671-style errors)', () => {
    // The model often writes text with unescaped inner double quotes straight into the JSON string
    const broken =
      '{"core_hook":"钩子","pages":[{"title":"封面","type":"cover","brief":"主标题：王虹。副标：她说"数学之美"属于每一个提问的人。","layout":"cover_dark_minimal","image_queries":[]}]}'
    expect(() => JSON.parse(broken)).toThrow()
    const obj = parseOutlineJson(broken)
    expect(obj).not.toBeNull()
    expect((obj!.pages as Array<{ brief: string }>)[0]!.brief).toContain('数学之美')
  })

  it('strips trailing commas', () => {
    const obj = parseOutlineJson(
      '{"core_hook":"h","pages":[{"title":"t","brief":"b","layout":"cover","image_queries":[],}],}',
    )
    expect(obj?.core_hook).toBe('h')
  })

  it('normalizes curly quotes', () => {
    const obj = parseOutlineJson('{"core_hook":“钩子”,"pages":[]}')
    expect(obj?.core_hook).toBe('钩子')
  })

  it('returns null when completely broken', () => {
    expect(parseOutlineJson('not json at all')).toBeNull()
    expect(parseOutlineJson('{')).toBeNull()
  })
})

describe('extractJsonObject / escapeBrokenStringLiterals', () => {
  it('slices from the first to the last brace', () => {
    expect(extractJsonObject('prefix {"a":1} suffix')).toBe('{"a":1}')
  })

  it('quotes inside values are escaped so it can be parsed', () => {
    const fixed = escapeBrokenStringLiterals('{"a":"说"你好"吧"}')
    expect(JSON.parse(fixed)).toEqual({ a: '说"你好"吧' })
  })
})
