import { describe, expect, it } from 'vitest'
import { aiStrings } from '../src/renderer/i18n/strings-ai'

describe('clarify arrow labels', () => {
  it('every locale labels the prev/next question arrows', () => {
    for (const [locale, strings] of Object.entries(aiStrings)) {
      expect(strings.aiClarifyPrev, `${locale}.aiClarifyPrev`).toBeTruthy()
      expect(strings.aiClarifyNext, `${locale}.aiClarifyNext`).toBeTruthy()
    }
  })
})
