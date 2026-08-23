import { describe, it, expect } from 'vitest'
import { classifyCjkScript } from '../src/shared/cjk-script'

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
