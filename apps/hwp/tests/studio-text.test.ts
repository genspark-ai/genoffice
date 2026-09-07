import { describe, expect, it, vi } from 'vitest'
import {
  clipPlainText,
  createStudioFacade,
  currentPage,
  fileNameOf,
  normalizeReplacement,
  PARAGRAPH_MAX_CODE_POINTS,
  readSelectionState,
  replaceCurrentParagraph,
  replaceCurrentSelection,
  replaceParagraphAt,
  listBodyParagraphs,
  listDocumentFields,
  setDocumentField,
  listDocumentTables,
  replaceTableCell,
  spliceParagraphText,
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
    getDocumentState: partial.getDocumentState,
    applyTextCommand: partial.applyTextCommand,
    focusTarget: partial.focusTarget,
    _request: partial._request,
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

describe('readSelectionState', () => {
  it('reads page and selection from one getSelectionContext call', async () => {
    const getSelectionContext = vi.fn(async () => ({
      collapsed: false,
      selectedTextSha256: 'sha',
      page: 2,
    }))
    expect(await readSelectionState(studio({ getSelectionContext }))).toEqual({
      page: 2,
      hasSelection: true,
    })
    expect(getSelectionContext).toHaveBeenCalledOnce()
  })
})

describe('currentPage', () => {
  it('returns a positive page from selection context', async () => {
    expect(
      await currentPage(
        studio({
          getSelectionContext: async () => ({
            collapsed: true,
            selectedTextSha256: null,
            page: 3,
          }),
        }),
      ),
    ).toBe(3)
  })

  it('returns null when page is missing or not positive', async () => {
    expect(await currentPage(studio())).toBeNull()
    expect(
      await currentPage(
        studio({
          getSelectionContext: async () => ({
            collapsed: true,
            selectedTextSha256: null,
            page: 0,
          }),
        }),
      ),
    ).toBeNull()
  })
})

describe('fileNameOf', () => {
  it('uses the last path segment and untitled fallback', () => {
    expect(fileNameOf('/tmp/a/memo.hwp')).toBe('memo.hwp')
    expect(fileNameOf(null)).toBe('untitled.hwp')
  })
})

describe('replaceCurrentParagraph', () => {
  it('rejects control characters and overlong replacements', () => {
    expect(() => normalizeReplacement('a\nb')).toThrow('control characters')
    expect(() => normalizeReplacement('x'.repeat(PARAGRAPH_MAX_CODE_POINTS + 1))).toThrow(
      'at most',
    )
  })

  it('applies a prepared fence through applyTextCommand', async () => {
    const applyTextCommand = vi.fn(async (command: { replacement: string }) => {
      expect(command.replacement).toBe('새 문장')
      expect(command.expectedBeforeSha256).toBe('aa'.repeat(32))
      return {
        target: { kind: 'body_paragraph' as const, section: 0, paragraph: 0, charOffset: 0 as const, length: 3 },
      }
    })
    const result = await replaceCurrentParagraph(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method) => {
          expect(method).toBe('prepareTextCommand')
          return {
            editable: true,
            reason: null,
            target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 4 },
            text: '이전',
            textSha256: 'aa'.repeat(32),
            formatSha256: 'cc'.repeat(32),
            adjacentContextSha256: 'dd'.repeat(32),
          }
        },
      }),
      '새 문장',
    )
    expect(result).toEqual({ before: '이전', after: '새 문장' })
    expect(applyTextCommand).toHaveBeenCalledOnce()
  })

  it('splices a selection inside the current paragraph', async () => {
    const applyTextCommand = vi.fn(async (command: { replacement: string }) => {
      expect(command.replacement).toBe('안녕 세상')
      return {
        target: { kind: 'body_paragraph' as const, section: 0, paragraph: 0, charOffset: 0 as const, length: 5 },
      }
    })
    const result = await replaceCurrentSelection(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async () => ({
          editable: true,
          reason: null,
          target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 4 },
          text: '안녕 세계',
          textSha256: 'aa'.repeat(32),
          formatSha256: 'cc'.repeat(32),
          adjacentContextSha256: 'dd'.repeat(32),
          selectionStart: 3,
          selectionEnd: 5,
        }),
      }),
      '세상',
    )
    expect(result).toEqual({ before: '세계', after: '세상' })
    expect(applyTextCommand).toHaveBeenCalledOnce()
  })

  it('replaces a paragraph by index from a fresh list', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 1, charOffset: 0 as const, length: 2 },
    }))
    const result = await replaceParagraphAt(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method) => {
          expect(method).toBe('listBodyParagraphs')
          return [
            {
              editable: true,
              reason: null,
              target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 1 },
              text: '하나',
              textSha256: 'aa'.repeat(32),
              formatSha256: 'cc'.repeat(32),
              adjacentContextSha256: 'dd'.repeat(32),
            },
            {
              editable: true,
              reason: null,
              target: { kind: 'body_paragraph', section: 0, paragraph: 1, charOffset: 0, length: 1 },
              text: '둘',
              textSha256: 'ee'.repeat(32),
              formatSha256: 'ff'.repeat(32),
              adjacentContextSha256: '11'.repeat(32),
            },
          ]
        },
      }),
      1,
      '둘둘',
    )
    expect(result).toEqual({ before: '둘', after: '둘둘' })
  })

  it('lists body paragraphs', async () => {
    const items = await listBodyParagraphs(
      studio({
        _request: async () => [
          {
            editable: false,
            reason: 'table',
            target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
            text: '',
            textSha256: null,
            formatSha256: null,
            adjacentContextSha256: null,
          },
        ],
      }),
    )
    expect(items).toEqual([
      { index: 0, editable: false, reason: 'table', section: 0, paragraph: 0, text: '' },
    ])
  })

  it('sets a field through prepare and falls back to hwpctrl', async () => {
    const fields = await listDocumentFields(
      studio({
        _request: async () => [{ name: '기안자', value: '', type: 'clickhere' }],
      }),
    )
    expect(fields[0]?.name).toBe('기안자')
    const result = await setDocumentField(
      studio({
        _request: async (method, params) => {
          if (method === 'listFields') return [{ name: '기안자', value: '', type: 'clickhere' }]
          expect(method).toBe('setField')
          expect(params).toEqual({ name: '기안자', value: '홍길동' })
          return { ok: true }
        },
      }),
      '기안자',
      '홍길동',
    )
    expect(result).toEqual({ name: '기안자', before: '', after: '홍길동' })
  })

  it('replaces a table cell by row and column', async () => {
    const result = await replaceTableCell(
      studio({
        _request: async (method, params) => {
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 2,
                control: 0,
                rows: 1,
                cols: 2,
                cells: [
                  { index: 0, row: 0, col: 0, text: 'A' },
                  { index: 1, row: 0, col: 1, text: 'B' },
                ],
              },
            ]
          }
          expect(method).toBe('replaceCell')
          expect(params).toEqual({
            section: 0,
            paragraph: 2,
            control: 0,
            cellIndex: 1,
            text: 'C',
          })
          return { ok: true }
        },
      }),
      0,
      0,
      1,
      'C',
    )
    expect(result).toEqual({ before: 'B', after: 'C' })
    await expect(listDocumentTables(studio({ _request: async () => [] }))).resolves.toEqual([])
  })

  it('splices UTF-16 and code-point ranges', () => {
    expect(spliceParagraphText('안녕세계', 2, 4, '세상')).toBe('안녕세상')
    expect(spliceParagraphText('ab', 0, 1, 'z')).toBe('zb')
  })

  it('does not apply when the paragraph is not editable', async () => {
    await expect(
      replaceCurrentParagraph(
        studio({
          getDocumentState: async () => ({
            documentEpoch: 1,
            changeSeq: 0,
            documentSha256: 'bb'.repeat(32),
          }),
          applyTextCommand: async () => {
            throw new Error('should not apply')
          },
          _request: async () => ({
            editable: false,
            reason: 'not_editable',
            target: null,
            text: null,
            textSha256: null,
            formatSha256: null,
            adjacentContextSha256: null,
          }),
        }),
        '새 문장',
      ),
    ).rejects.toThrow('not_editable')
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
    expect(await facade.currentPage()).toBeNull()
    expect(await facade.getPlainText()).toBe('body')
    expect(await facade.getSelectionText()).toBe('sel')
    expect(await facade.hasSelection()).toBe(true)
  })
})
