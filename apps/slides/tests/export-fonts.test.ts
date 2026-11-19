import { afterEach, describe, expect, it, vi } from 'vitest'

import { collectExportFontCss, familyNames } from '../src/renderer/export-fonts'

const BYTES = new Uint8Array([0x00, 0x01, 0x00, 0x00, 0x41]).buffer
const B64 = 'AAEAAEE='

/** jsdom's CSSOM drops @font-face src descriptors, so the sheet is faked at the CSSOM level. */
function installSheet(faces: Array<Record<string, string>>): void {
  const rules = faces.map((decl) => ({
    cssText: `@font-face { ${Object.entries(decl)
      .map(([k, v]) => `${k}: ${v};`)
      .join(' ')} }`,
    style: { getPropertyValue: (k: string) => decl[k] ?? '' },
  }))
  Object.defineProperty(document, 'styleSheets', {
    configurable: true,
    value: [{ cssRules: rules }],
  })
}

afterEach(() => {
  vi.unstubAllGlobals()
  delete (document as { styleSheets?: unknown }).styleSheets
  delete (window as { slidesApi?: unknown }).slidesApi
})

describe('export fonts', () => {
  it('familyNames: splits stacks, strips quotes, drops generic families, normalizes case', () => {
    const names = familyNames([
      "'Noto Sans SC', 'PingFang SC', sans-serif",
      'Calibri, Carlito, Arial, sans-serif',
      "'Carlito GO'",
    ])
    expect([...names].sort()).toEqual([
      'arial',
      'calibri',
      'carlito',
      'carlito go',
      'noto sans sc',
      'pingfang sc',
    ])
  })

  it('inlines bundled @font-face rules the page uses and skips the rest', async () => {
    installSheet([
      {
        'font-family': "'Carlito'",
        src: "url('/assets/Carlito-Bold.ttf')",
        'font-weight': 'bold',
        'font-style': 'normal',
      },
      { 'font-family': "'Other'", src: "url('/assets/Other.ttf')" },
    ])
    const fetchMock = vi.fn(async (url: string) => ({
      ok: true,
      arrayBuffer: async () => BYTES,
      url,
    }))
    vi.stubGlobal('fetch', fetchMock)
    ;(window as { slidesApi?: unknown }).slidesApi = { privateFontFaces: async () => [] }

    const css = await collectExportFontCss(['Calibri, Carlito, Arial, sans-serif'])
    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetchMock.mock.calls[0]![0]).toContain('Carlito-Bold.ttf')
    expect(css).toBe(
      `@font-face{font-family:"Carlito";src:url(data:font/ttf;base64,${B64});font-weight:bold;font-style:normal;font-display:block}`,
    )
  })

  it('inlines Office-private faces by family with their weight and style', async () => {
    ;(window as { slidesApi?: unknown }).slidesApi = {
      privateFontFaces: async () => [
        { id: 'f1', family: 'Meiryo UI', bold: true, italic: false },
        { id: 'f2', family: 'Unused', bold: false, italic: false },
      ],
      privateFontData: async (id: string) => (id === 'f1' ? BYTES : null),
    }
    const css = await collectExportFontCss(["'Meiryo UI', Meiryo, sans-serif"])
    expect(css).toBe(
      `@font-face{font-family:"Meiryo UI";src:url(data:font/ttf;base64,${B64});font-weight:700;font-style:normal;font-display:block}`,
    )
  })

  it('returns nothing for generic-only stacks and survives a missing bridge', async () => {
    expect(await collectExportFontCss(['sans-serif'])).toBe('')
    expect(await collectExportFontCss(['Calibri'])).toBe('')
  })
})
