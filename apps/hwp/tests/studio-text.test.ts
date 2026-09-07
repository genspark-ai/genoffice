import { describe, expect, it, vi } from 'vitest'
import {
  clipPlainText,
  createStudioFacade,
  fileNameOf,
  getPlainText,
  getSelectionText,
  hasSelection,
  PLAIN_TEXT_MAX_CHARS,
  PLAIN_TEXT_UNAVAILABLE,
  stripHmlToPlainText,
  type StudioTextSource,
} from '../src/renderer/studio-text'

function studio(
  partial: Partial<StudioTextSource> & { textFile?: unknown; selected?: unknown } = {},
): StudioTextSource {
  return {
    pageCount: partial.pageCount ?? (async () => 3),
    exportHml:
      partial.exportHml ??
      (async () => new TextEncoder().encode('<HML><P>fallback body</P></HML>')),
    getHmlSaveState: partial.getHmlSaveState,
    getSelectionContext:
      partial.getSelectionContext ?? (async () => ({ collapsed: true, selectedTextSha256: null })),
    hwpctrl: {
      call:
        partial.hwpctrl?.call ??
        (async (method, args) => {
          if (method !== 'GetTextFile') return null
          if (args?.[1] === 'saveblock') return partial.selected ?? null
          return partial.textFile ?? null
        }),
    },
  }
}

describe('stripHmlToPlainText', () => {
  it('drops tags and decodes entities', () => {
    expect(stripHmlToPlainText('<P>안녕 &amp; hello&nbsp;world</P>')).toBe('안녕 & hello world')
  })

  it('skips invalid numeric entities', () => {
    expect(stripHmlToPlainText('ok&#999999999;end')).toBe('okend')
    expect(stripHmlToPlainText('ok&#x110000;end')).toBe('okend')
  })
})

describe('clipPlainText', () => {
  it('leaves short text unchanged', () => {
    expect(clipPlainText('ok')).toBe('ok')
  })

  it('truncates long text', () => {
    const out = clipPlainText('x'.repeat(PLAIN_TEXT_MAX_CHARS + 10))
    expect(out.endsWith('[truncated]')).toBe(true)
    expect(out.length).toBeLessThan(PLAIN_TEXT_MAX_CHARS + 20)
  })
})

describe('getPlainText', () => {
  it('prefers hwpctrl GetTextFile TEXT', async () => {
    const call = vi.fn(async (_method: string, args?: unknown[]) => {
      expect(args).toEqual(['TEXT', ''])
      return 'from ctrl'
    })
    const text = await getPlainText(studio({ hwpctrl: { call } }))
    expect(text).toBe('from ctrl')
    expect(call).toHaveBeenCalledOnce()
  })

  it('falls back to HML when GetTextFile returns nothing', async () => {
    const text = await getPlainText(studio({ textFile: '' }))
    expect(text).toBe('fallback body')
  })

  it('does not call exportHml when HML is not savable', async () => {
    const exportHml = vi.fn(async () => new Uint8Array())
    await expect(
      getPlainText(
        studio({
          textFile: '',
          exportHml,
          getHmlSaveState: async () => ({ hmlSavable: false }),
        }),
      ),
    ).rejects.toThrow(PLAIN_TEXT_UNAVAILABLE)
    expect(exportHml).not.toHaveBeenCalled()
  })

  it('reports unavailable when exportHml fails', async () => {
    await expect(
      getPlainText(
        studio({
          textFile: '',
          exportHml: async () => {
            throw new Error('export failed')
          },
        }),
      ),
    ).rejects.toThrow(PLAIN_TEXT_UNAVAILABLE)
  })
})

describe('getSelectionText / hasSelection', () => {
  it('reads selection with GetTextFile saveblock and clips it', async () => {
    const call = vi.fn(async (_method: string, args?: unknown[]) => {
      expect(args).toEqual(['TEXT', 'saveblock'])
      return 'y'.repeat(PLAIN_TEXT_MAX_CHARS + 5)
    })
    const text = await getSelectionText(studio({ hwpctrl: { call } }))
    expect(text?.endsWith('[truncated]')).toBe(true)
    expect(call).toHaveBeenCalledOnce()
  })

  it('reports a selection from context when raw text is unavailable', async () => {
    const src = studio({
      selected: null,
      getSelectionContext: async () => ({ collapsed: false, selectedTextSha256: 'abc' }),
    })
    expect(await getSelectionText(src)).toBeNull()
    expect(await hasSelection(src)).toBe(true)
  })

  it('reports no selection when collapsed', async () => {
    expect(await hasSelection(studio({ selected: 'picked' }))).toBe(false)
  })
})

describe('fileNameOf', () => {
  it('uses the last path segment and untitled fallback', () => {
    expect(fileNameOf('/tmp/a/memo.hwp')).toBe('memo.hwp')
    expect(fileNameOf(null)).toBe('untitled.hwp')
  })
})

describe('createStudioFacade', () => {
  it('exposes pageCount and text helpers', async () => {
    const facade = createStudioFacade(
      studio({
        textFile: 'body',
        selected: 'sel',
        getSelectionContext: async () => ({ collapsed: false, selectedTextSha256: 'sha' }),
      }),
    )
    expect(await facade.pageCount()).toBe(3)
    expect(await facade.getPlainText()).toBe('body')
    expect(await facade.getSelectionText()).toBe('sel')
    expect(await facade.hasSelection()).toBe(true)
  })
})
