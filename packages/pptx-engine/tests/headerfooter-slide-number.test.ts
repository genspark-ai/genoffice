import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { applyHeaderFooter, openPptx } from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('applyHeaderFooter cached slide numbers', () => {
  it('caches the real deck position when only a subset is targeted', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const idx = opened.deck.slides.length - 1
    expect(idx).toBeGreaterThan(0)
    applyHeaderFooter(opened, { slideNum: true }, [idx])
    const sp = opened.deck.slides[idx]!.elements.find((el) =>
      el.anchor.originalXml.includes('slidenum'),
    )
    expect(sp).toBeDefined()
    const cached = /<a:t>(\d+)<\/a:t>/.exec(sp!.anchor.originalXml)
    expect(cached?.[1]).toBe(String(idx + 1))
  })

  it('numbers every slide 1..n when the whole deck is targeted', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    applyHeaderFooter(opened, { slideNum: true })
    opened.deck.slides.forEach((slide, i) => {
      const sp = slide.elements.find((el) => el.anchor.originalXml.includes('slidenum'))
      expect(sp).toBeDefined()
      const cached = /<a:t>(\d+)<\/a:t>/.exec(sp!.anchor.originalXml)
      expect(cached?.[1]).toBe(String(i + 1))
    })
  })
})
