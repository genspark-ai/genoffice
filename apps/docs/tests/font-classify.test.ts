import { describe, expect, it } from 'vitest'
import { isEastAsianFontName } from '../src/renderer/font-list'

describe('isEastAsianFontName', () => {
  it('classifies CJK families across naming conventions', () => {
    for (const f of [
      'KaiTi',
      'SimSun',
      '宋体',
      '楷体_GB2312',
      '仿宋',
      '微软雅黑',
      '等线',
      '方正小标宋简体',
      'PingFang SC',
      'Microsoft YaHei',
      'Yu Mincho',
      'MS Gothic',
      'Malgun Gothic',
      'Batang',
      'Microsoft JhengHei',
      'PMingLiU',
      'Noto Sans CJK SC',
      'Source Han Serif',
    ]) {
      expect(isEastAsianFontName(f), f).toBe(true)
    }
  })

  it('classifies Latin families as non-East-Asian', () => {
    for (const f of [
      'Times New Roman',
      'Arial',
      'Calibri',
      'Cambria',
      'Georgia',
      'Courier New',
      'Segoe UI',
      'Consolas',
      'Garamond',
    ]) {
      expect(isEastAsianFontName(f), f).toBe(false)
    }
  })
})
