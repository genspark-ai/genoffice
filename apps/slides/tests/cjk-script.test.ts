import { describe, it, expect } from 'vitest'
import { classifyCjkScript, classifyCjkScriptByNameScript } from '../src/shared/cjk-script'

describe('classifyCjkScript', () => {
  it('classifies Korean vendor faces by name keywords', () => {
    expect(classifyCjkScript('SamsungOneKorean 300')).toBe('ko')
    expect(classifyCjkScript('Adobe Korean Std')).toBe('ko')
    expect(classifyCjkScript('NanumGothic')).toBe('ko')
    expect(classifyCjkScript('KoPub바탕체 Bold')).toBe('ko')
  })

  it('leaves Latin and other scripts alone', () => {
    expect(classifyCjkScript('Arial')).toBeNull()
    expect(classifyCjkScript('Meiryo')).toBe('ja')
    expect(classifyCjkScript('Microsoft JhengHei')).toBe('tc')
  })
})

describe('classifyCjkScriptByNameScript', () => {
  it('reads only the script of the letters in the name (PowerPoint Latin-text substitution)', () => {
    expect(classifyCjkScriptByNameScript('함초롬돋움')).toBe('ko')
    expect(classifyCjkScriptByNameScript('LG스마트체 Regular')).toBe('ko')
    expect(classifyCjkScriptByNameScript('游ゴシック')).toBe('ja')
    expect(classifyCjkScriptByNameScript('微软雅黑')).toBe('sc')
    expect(classifyCjkScriptByNameScript('微軟正黑體')).toBe('tc')
    // romanized keywords do not count: prod_026's digits in this face set in Calibri
    expect(classifyCjkScriptByNameScript('NanumSquareExtraBold')).toBeNull()
    expect(classifyCjkScriptByNameScript('Malgun Gothic')).toBeNull()
  })
})
