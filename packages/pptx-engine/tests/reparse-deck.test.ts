import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openPptx,
  savePptx,
  commitSaved,
  reparseDeck,
  addElement,
  applyThemeToArchive,
  remapDeckColors,
  type ThemeSpec,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

const SPEC: ThemeSpec = {
  name: 'Test Theme',
  colors: {
    dk1: '111111',
    lt1: 'FEFEFE',
    dk2: '222222',
    lt2: 'EEEEEE',
    accent1: 'AA0000',
    accent2: '00AA00',
    accent3: '0000AA',
    accent4: 'AAAA00',
    accent5: '00AAAA',
    accent6: 'AA00AA',
    hlink: '0563C1',
    folHlink: '954F72',
  },
}

describe('reparseDeck', () => {
  it('matches the openPptx(savePptx()) roundtrip', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const roundtrip = await openPptx(await savePptx(opened))

    const reparsed = reparseDeck(opened)
    expect(reparsed.deck.slides.length).toBe(roundtrip.deck.slides.length)
    expect(reparsed.deck.size).toEqual(roundtrip.deck.size)
    for (let i = 0; i < reparsed.deck.slides.length; i++) {
      expect(reparsed.deck.slides[i]!.originalXml).toBe(roundtrip.deck.slides[i]!.originalXml)
      expect(reparsed.deck.slides[i]!.elements.length).toBe(
        roundtrip.deck.slides[i]!.elements.length,
      )
    }
  })

  it('apply-theme sequence: commitSaved bakes edits, surgery lands, reparse resolves new colors', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    addElement(opened.deck.slides[0]!, {
      kind: 'textbox',
      offset: { x: 914400, y: 914400, cx: 6096000, cy: 914400 },
      paragraphs: [{ runs: [{ text: 'pre-theme edit', fontSize: 24 }] }],
    })

    commitSaved(opened)
    expect(applyThemeToArchive(opened, SPEC)).toBeGreaterThan(0)
    remapDeckColors(opened, SPEC)
    const reparsed = reparseDeck(opened)

    // The pre-surgery edit survives, and the new theme is what a save→reopen would see
    expect(reparsed.deck.slides[0]!.originalXml).toContain('pre-theme edit')
    const viaZip = await openPptx(await savePptx(reparsed))
    expect(viaZip.archive.readText('ppt/theme/theme1.xml')).toContain('Test Theme')
    expect(reparsed.archive.readText('ppt/theme/theme1.xml')).toContain('AA0000')
  })
})
