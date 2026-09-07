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
  insertContent,
  splitInsertParagraphs,
  INSERT_CONTENT_MAX_PARAS,
  listBodyParagraphs,
  listDocumentFields,
  setDocumentField,
  listDocumentTables,
  replaceTableCell,
  insertDocumentTable,
  applyParagraphFormat,
  applyTableCellFormat,
  editDocumentTable,
  styleDocumentTable,
  setDocumentPage,
  FORMAT_EMPTY_RANGE,
  TABLE_NOT_FOUND,
  spliceParagraphText,
  getPlainText,
  getSelectionText,
  hasSelection,
  PLAIN_TEXT_MAX_CHARS,
  PLAIN_TEXT_UNAVAILABLE,
  PARAGRAPH_PREPARE_UNAVAILABLE,
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

  it('keeps a successful apply when focusTarget still has the old length', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 0, charOffset: 0 as const, length: 8 },
    }))
    const focusTarget = vi.fn(async (target: { length: number }) => {
      if (target.length !== 5) {
        throw new Error('exact body paragraph target을 찾을 수 없습니다.')
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
        focusTarget,
        _request: async () => ({
          editable: true,
          reason: null,
          target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 8 },
          text: '8. 기대 효과',
          textSha256: 'aa'.repeat(32),
          formatSha256: 'cc'.repeat(32),
          adjacentContextSha256: 'dd'.repeat(32),
        }),
      }),
      '8. 효과',
    )
    expect(result).toEqual({ before: '8. 기대 효과', after: '8. 효과' })
    expect(focusTarget).toHaveBeenCalledWith(
      expect.objectContaining({ section: 0, paragraph: 0, length: 5 }),
    )
  })

  it('does not fail the replace when focusTarget throws after apply', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 0, charOffset: 0 as const, length: 8 },
    }))
    const result = await replaceCurrentParagraph(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        focusTarget: async () => {
          throw new Error('exact body paragraph target을 찾을 수 없습니다.')
        },
        _request: async () => ({
          editable: true,
          reason: null,
          target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 8 },
          text: '8. 기대 효과',
          textSha256: 'aa'.repeat(32),
          formatSha256: 'cc'.repeat(32),
          adjacentContextSha256: 'dd'.repeat(32),
        }),
      }),
      '8. 효과',
    )
    expect(result).toEqual({ before: '8. 기대 효과', after: '8. 효과' })
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

  it('splits insert text on newlines and rejects a trailing-only empty string', () => {
    expect(splitInsertParagraphs('안녕\n세상\n')).toEqual(['안녕', '세상'])
    expect(() => splitInsertParagraphs('')).toThrow(/must not be empty/)
    expect(() => splitInsertParagraphs(`${'x'.repeat(8)}\n`.repeat(INSERT_CONTENT_MAX_PARAS + 1))).toThrow(
      /at most/,
    )
  })

  it('fills an empty caret paragraph then inserts the rest', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const applyTextCommand = vi.fn(async (command: { replacement: string }) => ({
      target: {
        kind: 'body_paragraph' as const,
        section: 0,
        paragraph: command.replacement === '안녕' ? 0 : 1,
        charOffset: 0 as const,
        length: command.replacement.length,
      },
    }))
    const listed = [
      {
        editable: true,
        reason: null,
        target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
        text: '',
        textSha256: 'aa'.repeat(32),
        formatSha256: 'cc'.repeat(32),
        adjacentContextSha256: 'dd'.repeat(32),
      },
    ]
    const result = await insertContent(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'prepareTextCommand') return listed[0]
          if (method === 'insertFilledParagraphs') {
            return { section: 0, index: params?.index, count: (params?.texts as string[]).length }
          }
          if (method === 'listBodyParagraphs') return listed.map((item) => ({ ...item }))
          throw new Error(`unexpected ${method}`)
        },
      }),
      '안녕\n세상',
    )
    expect(result).toEqual({ count: 2, start: 0 })
    expect(calls.some((call) => call.method === 'insertFilledParagraphs')).toBe(true)
    expect(calls.find((call) => call.method === 'insertFilledParagraphs')?.params?.texts).toEqual(['세상'])
    expect(applyTextCommand.mock.calls.map((call) => call[0].replacement)).toEqual(['안녕'])
  })

  it('inserts after a numbered paragraph without rewriting it', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 1, charOffset: 0 as const, length: 2 },
    }))
    const listed = [
      {
        editable: true,
        reason: null,
        target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 1 },
        text: '기존',
        textSha256: 'aa'.repeat(32),
        formatSha256: 'cc'.repeat(32),
        adjacentContextSha256: 'dd'.repeat(32),
      },
    ]
    const result = await insertContent(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method, params) => {
          if (method === 'insertFilledParagraphs') {
            return { section: 0, index: params?.index, count: (params?.texts as string[]).length }
          }
          if (method === 'listBodyParagraphs') return listed.map((item) => ({ ...item }))
          throw new Error(`unexpected ${method}`)
        },
      }),
      '추가',
      0,
    )
    expect(result).toEqual({ count: 1, start: 1 })
    expect(applyTextCommand).not.toHaveBeenCalled()
  })

  it('inserts at the start when afterIndex is -1 and the first paragraph has text', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 0, charOffset: 0 as const, length: 2 },
    }))
    const listed = [
      {
        editable: true,
        reason: null,
        target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 1 },
        text: '기존',
        textSha256: 'aa'.repeat(32),
        formatSha256: 'cc'.repeat(32),
        adjacentContextSha256: 'dd'.repeat(32),
      },
    ]
    let insertAt: number | undefined
    const result = await insertContent(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method, params) => {
          if (method === 'insertFilledParagraphs') {
            insertAt = params?.index as number
            return { section: 0, index: insertAt, count: (params?.texts as string[]).length }
          }
          if (method === 'listBodyParagraphs') return listed.map((item) => ({ ...item }))
          throw new Error(`unexpected ${method}`)
        },
      }),
      '앞',
      -1,
    )
    expect(result).toEqual({ count: 1, start: 0 })
    expect(insertAt).toBe(0)
    expect(applyTextCommand).not.toHaveBeenCalled()
  })

  it('skips a locked empty caret paragraph and inserts after it', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 1, charOffset: 0 as const, length: 2 },
    }))
    const listed = [
      {
        editable: false,
        reason: 'control',
        target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
        text: '',
        textSha256: null,
        formatSha256: null,
        adjacentContextSha256: null,
      },
    ]
    let insertAt: number | undefined
    const result = await insertContent(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method, params) => {
          if (method === 'prepareTextCommand') {
            return {
              editable: false,
              reason: 'control',
              target: listed[0]!.target,
              text: '',
              textSha256: null,
              formatSha256: null,
              adjacentContextSha256: null,
            }
          }
          if (method === 'insertFilledParagraphs') {
            insertAt = params?.index as number
            return { section: 0, index: insertAt, count: (params?.texts as string[]).length }
          }
          if (method === 'listBodyParagraphs') return listed.map((item) => ({ ...item }))
          throw new Error(`unexpected ${method}`)
        },
      }),
      '계획서',
    )
    expect(result).toEqual({ count: 1, start: 1 })
    expect(insertAt).toBe(1)
    expect(applyTextCommand).not.toHaveBeenCalled()
  })

  it('inserts after the last paragraph when the caret has no body target', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 1, charOffset: 0 as const, length: 2 },
    }))
    const listed = [
      {
        editable: false,
        reason: 'not_editable',
        target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
        text: '',
        textSha256: null,
        formatSha256: null,
        adjacentContextSha256: null,
      },
    ]
    const result = await insertContent(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method) => {
          if (method === 'prepareTextCommand') {
            return {
              editable: false,
              reason: 'not_editable',
              target: null,
              text: null,
              textSha256: null,
              formatSha256: null,
              adjacentContextSha256: null,
            }
          }
          if (method === 'insertFilledParagraphs') {
            return { section: 0, index: 1, count: 1 }
          }
          if (method === 'listBodyParagraphs') return listed.map((item) => ({ ...item }))
          throw new Error(`unexpected ${method}`)
        },
      }),
      '초안',
    )
    expect(result).toEqual({ count: 1, start: 1 })
    expect(applyTextCommand).not.toHaveBeenCalled()
  })

  it('sends a locked-first draft in one insertFilledParagraphs call', async () => {
    const applyTextCommand = vi.fn(async () => ({
      target: { kind: 'body_paragraph' as const, section: 0, paragraph: 2, charOffset: 0 as const, length: 2 },
    }))
    const listed = [
      {
        editable: false,
        reason: 'control',
        target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
        text: '',
        textSha256: null,
        formatSha256: null,
        adjacentContextSha256: null,
      },
    ]
    const result = await insertContent(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method, params) => {
          if (method === 'prepareTextCommand') {
            return {
              editable: false,
              reason: 'control',
              target: listed[0]!.target,
              text: '',
              textSha256: null,
              formatSha256: null,
              adjacentContextSha256: null,
            }
          }
          if (method === 'insertFilledParagraphs') {
            return { section: 0, index: params?.index, count: 1 }
          }
          if (method === 'listBodyParagraphs') return listed.map((item) => ({ ...item }))
          throw new Error(`unexpected ${method}`)
        },
      }),
      '본문',
    )
    expect(result.start).toBe(1)
    expect(applyTextCommand).not.toHaveBeenCalled()
  })

  it('writes every line of a multi-paragraph insert after a locked first paragraph', async () => {
    const applyTextCommand = vi.fn(async (command: { replacement: string }) => ({
      target: {
        kind: 'body_paragraph' as const,
        section: 0,
        paragraph: 1,
        charOffset: 0 as const,
        length: command.replacement.length,
      },
    }))
    const listed = [
      {
        editable: false,
        reason: 'control',
        target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
        text: '',
        textSha256: null,
        formatSha256: null,
        adjacentContextSha256: null,
      },
    ]
    let filled: string[] | undefined
    const result = await insertContent(
      studio({
        getDocumentState: async () => ({
          documentEpoch: 1,
          changeSeq: 0,
          documentSha256: 'bb'.repeat(32),
        }),
        applyTextCommand,
        _request: async (method, params) => {
          if (method === 'prepareTextCommand') {
            return {
              editable: false,
              reason: 'control',
              target: listed[0]!.target,
              text: '',
              textSha256: null,
              formatSha256: null,
              adjacentContextSha256: null,
            }
          }
          if (method === 'insertFilledParagraphs') {
            filled = params?.texts as string[]
            return { section: 0, index: params?.index, count: filled.length }
          }
          if (method === 'listBodyParagraphs') return listed.map((item) => ({ ...item }))
          throw new Error(`unexpected ${method}`)
        },
      }),
      '제목\n1. 개요\n본문',
    )
    expect(result).toEqual({ count: 3, start: 1 })
    expect(filled).toEqual(['제목', '1. 개요', '본문'])
    expect(applyTextCommand).not.toHaveBeenCalled()
  })

  it('rejects an out-of-range afterIndex', async () => {
    await expect(
      insertContent(
        studio({
          _request: async (method) => {
            if (method === 'listBodyParagraphs') {
              return [
                {
                  editable: true,
                  reason: null,
                  target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 1 },
                  text: '기존',
                  textSha256: 'aa'.repeat(32),
                  formatSha256: 'cc'.repeat(32),
                  adjacentContextSha256: 'dd'.repeat(32),
                },
              ]
            }
            throw new Error(`unexpected ${method}`)
          },
        }),
        '추가',
        3,
      ),
    ).rejects.toThrow(/out of range/)
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

  it('sets a field through prepareTextCommand setField', async () => {
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

  it('falls back to hwpctrl PutFieldText when setField is unavailable', async () => {
    const call = vi.fn(async (method: string, args?: unknown[]) => {
      expect(method).toBe('PutFieldText')
      expect(args).toEqual(['기안자', '홍길동'])
      return null
    })
    const result = await setDocumentField(
      studio({
        hwpctrl: { call },
        _request: async (method) => {
          if (method === 'listFields') return [{ name: '기안자', value: '', type: 'clickhere' }]
          throw new Error('setField missing')
        },
      }),
      '기안자',
      '홍길동',
    )
    expect(result).toEqual({ name: '기안자', before: '', after: '홍길동' })
    expect(call).toHaveBeenCalledOnce()
  })

  it('does not hide listFields or listTables failures', async () => {
    await expect(
      listDocumentFields(
        studio({
          _request: async () => {
            throw new Error('listFields missing')
          },
        }),
      ),
    ).rejects.toThrow('listFields missing')
    await expect(
      listDocumentTables(
        studio({
          _request: async () => {
            throw new Error('listTables missing')
          },
        }),
      ),
    ).rejects.toThrow('listTables missing')
    await expect(listDocumentFields(studio({ _request: async () => ({}) }))).rejects.toThrow(
      PARAGRAPH_PREPARE_UNAVAILABLE,
    )
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

  it('inserts a table after the last paragraph and fills cells', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await insertDocumentTable(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 2 },
                text: '제목',
              },
            ]
          }
          if (method === 'insertBodyParagraphs') return { section: 0, index: 1, count: 1 }
          if (method === 'insertTable') return { section: 0, paragraph: 1, control: 0, rows: 2, cols: 2 }
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 1,
                control: 0,
                rows: 2,
                cols: 2,
                cells: [
                  { index: 0, row: 0, col: 0, text: '' },
                  { index: 1, row: 0, col: 1, text: '' },
                  { index: 2, row: 1, col: 0, text: '' },
                  { index: 3, row: 1, col: 1, text: '' },
                ],
              },
            ]
          }
          if (method === 'replaceCell') return { ok: true }
          throw new Error(`unexpected ${method}`)
        },
      }),
      2,
      2,
      [['이름', '역할']],
    )
    expect(result).toEqual({ table: 0, rows: 2, cols: 2, unfilled: [] })
    expect(calls.some((call) => call.method === 'insertBodyParagraphs')).toBe(true)
    expect(calls.find((call) => call.method === 'insertTable')?.params).toEqual({
      section: 0,
      index: 1,
      rows: 2,
      cols: 2,
    })
    expect(calls.filter((call) => call.method === 'replaceCell')).toHaveLength(2)
  })

  it('throws when every provided cell write fails', async () => {
    await expect(
      insertDocumentTable(
        studio({
          _request: async (method) => {
            if (method === 'listBodyParagraphs') {
              return [
                {
                  editable: true,
                  reason: null,
                  target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                  text: '',
                },
              ]
            }
            if (method === 'insertTable') return { section: 0, paragraph: 0, control: 0, rows: 1, cols: 1 }
            if (method === 'listTables') {
              return [
                {
                  section: 0,
                  paragraph: 0,
                  control: 0,
                  rows: 1,
                  cols: 1,
                  cells: [{ index: 0, row: 0, col: 0, text: '' }],
                },
              ]
            }
            if (method === 'replaceCell') throw new Error('cell locked')
            return {}
          },
        }),
        1,
        1,
        [['값']],
      ),
    ).rejects.toThrow(/cell writes failed/)
  })

  it('keeps a created table if only some cell writes fail', async () => {
    let writes = 0
    const result = await insertDocumentTable(
      studio({
        _request: async (method) => {
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                text: '',
              },
            ]
          }
          if (method === 'insertTable') return { section: 0, paragraph: 0, control: 0, rows: 1, cols: 2 }
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 0,
                control: 0,
                rows: 1,
                cols: 2,
                cells: [
                  { index: 0, row: 0, col: 0, text: '' },
                  { index: 1, row: 0, col: 1, text: '' },
                ],
              },
            ]
          }
          if (method === 'replaceCell') {
            writes += 1
            if (writes === 2) throw new Error('cell locked')
            return { ok: true }
          }
          return {}
        },
      }),
      1,
      2,
      [['A', 'B']],
    )
    expect(result).toEqual({ table: 0, rows: 1, cols: 2, unfilled: ['r0c1'] })
  })

  it('fills the inserted table when a later table already exists', async () => {
    const cells: Array<Record<string, unknown>> = []
    const result = await insertDocumentTable(
      studio({
        _request: async (method, params) => {
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 2 },
                text: '제목',
              },
            ]
          }
          if (method === 'insertBodyParagraphs') return { section: 0, index: 1, count: 1 }
          if (method === 'insertTable') return { section: 0, paragraph: 1, control: 0, rows: 1, cols: 1 }
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 1,
                control: 0,
                rows: 1,
                cols: 1,
                cells: [{ index: 0, row: 0, col: 0, text: '' }],
              },
              {
                section: 0,
                paragraph: 8,
                control: 0,
                rows: 2,
                cols: 2,
                cells: [{ index: 0, row: 0, col: 0, text: '끝 표' }],
              },
            ]
          }
          if (method === 'replaceCell') {
            cells.push(params ?? {})
            return { ok: true }
          }
          throw new Error(`unexpected ${method}`)
        },
      }),
      1,
      1,
      [['새 칸']],
    )
    expect(result).toEqual({ table: 0, rows: 1, cols: 1, unfilled: [] })
    expect(cells).toEqual([
      { section: 0, paragraph: 1, control: 0, cellIndex: 0, text: '새 칸' },
    ])
  })

  it('does not fall back to the last table when insertTable coords are missing', async () => {
    let filled = false
    await expect(
      insertDocumentTable(
        studio({
          _request: async (method) => {
            if (method === 'listBodyParagraphs') {
              return [
                {
                  editable: true,
                  reason: null,
                  target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                  text: '',
                },
              ]
            }
            if (method === 'insertTable') return { section: 0, paragraph: 1, control: 0, rows: 1, cols: 1 }
            if (method === 'listTables') {
              return [
                {
                  section: 0,
                  paragraph: 8,
                  control: 0,
                  rows: 1,
                  cols: 1,
                  cells: [{ index: 0, row: 0, col: 0, text: '끝 표' }],
                },
              ]
            }
            if (method === 'replaceCell') {
              filled = true
              return { ok: true }
            }
            return {}
          },
        }),
        1,
        1,
        [['값']],
      ),
    ).rejects.toThrow('table cell not found')
    expect(filled).toBe(false)
  })

  it('applies char and para format to a paragraph', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await applyParagraphFormat(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 4 },
                text: '제목',
                selectionStart: null,
                selectionEnd: null,
              },
            ]
          }
          return { ok: true }
        },
      }),
      { bold: true, fontSize: 16, align: 'center' },
      0,
    )
    expect(result).toEqual({ indexes: [0], applied: ['bold=true', 'fontSize=16', 'align=center'] })
    expect(calls.find((call) => call.method === 'applyBodyCharFormat')?.params).toEqual({
      section: 0,
      paragraph: 0,
      start: 0,
      end: 4,
      format: { bold: true, fontSize: 1600 },
    })
    expect(calls.find((call) => call.method === 'applyBodyParaFormat')?.params).toEqual({
      section: 0,
      paragraph: 0,
      format: { alignment: 'center' },
    })
  })

  it('maps color, font, indent, and line spacing onto studio format payloads', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    await applyParagraphFormat(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 2 },
                text: '제목',
              },
            ]
          }
          return { ok: true }
        },
      }),
      { color: 'F00', font: 'Pretendard', lineSpacing: 1.5, indentLeft: 10, indentFirstLine: 20 },
      0,
    )
    expect(calls.find((call) => call.method === 'applyBodyCharFormat')?.params).toMatchObject({
      format: { textColor: '#ff0000', fontName: 'Pretendard' },
    })
    expect(calls.find((call) => call.method === 'applyBodyParaFormat')?.params).toMatchObject({
      format: { lineSpacing: 150, lineSpacingType: 'Percent', marginLeft: 2000, indent: 4000 },
    })
  })

  it('formats only the caret selection when no paragraph index is given', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await applyParagraphFormat(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 6 },
                text: '안녕세계',
              },
            ]
          }
          if (method === 'prepareTextCommand') {
            return {
              editable: true,
              reason: null,
              target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 6 },
              text: '안녕세계',
              selectionStart: 2,
              selectionEnd: 6,
            }
          }
          return { ok: true }
        },
      }),
      { bold: true },
    )
    expect(result.indexes).toEqual([0])
    expect(calls.find((call) => call.method === 'applyBodyCharFormat')?.params).toEqual({
      section: 0,
      paragraph: 0,
      start: 2,
      end: 6,
      format: { bold: true },
    })
  })

  it('skips locked rows when formatting several paragraphs', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await applyParagraphFormat(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: false,
                reason: 'control',
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                text: '',
              },
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 1, charOffset: 0, length: 4 },
                text: '제목',
              },
            ]
          }
          return { ok: true }
        },
      }),
      { bold: true },
      undefined,
      [0, 1],
    )
    expect(result.indexes).toEqual([1])
    expect(calls.find((call) => call.method === 'applyBodyCharFormat')?.params).toMatchObject({
      paragraph: 1,
    })
  })

  it('rejects format on a locked table paragraph', async () => {
    let called = false
    await expect(
      applyParagraphFormat(
        studio({
          _request: async (method) => {
            if (method === 'listBodyParagraphs') {
              return [
                {
                  editable: false,
                  reason: 'table',
                  target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                  text: '',
                },
              ]
            }
            called = true
            return { ok: true }
          },
        }),
        { bold: true },
        0,
      ),
    ).rejects.toThrow('table')
    expect(called).toBe(false)
  })

  it('allows format on a mixed-formatting paragraph', async () => {
    const calls: string[] = []
    await applyParagraphFormat(
      studio({
        _request: async (method) => {
          calls.push(method)
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: false,
                reason: 'mixed formatting',
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 4 },
                text: '제목',
              },
            ]
          }
          return { ok: true }
        },
      }),
      { bold: true },
      0,
    )
    expect(calls).toContain('applyBodyCharFormat')
  })

  it('skips character format on an empty paragraph but still aligns', async () => {
    const calls: string[] = []
    await applyParagraphFormat(
      studio({
        _request: async (method) => {
          calls.push(method)
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                text: '',
              },
            ]
          }
          return { ok: true }
        },
      }),
      { bold: true, align: 'center' },
      0,
    )
    expect(calls).not.toContain('applyBodyCharFormat')
    expect(calls).toContain('applyBodyParaFormat')
  })

  it('skips empty paragraphs in a format batch and continues', async () => {
    const formatted: number[] = []
    const result = await applyParagraphFormat(
      studio({
        _request: async (method, params) => {
          if (method === 'listBodyParagraphs') {
            return [
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                text: '',
              },
              {
                editable: true,
                reason: null,
                target: { kind: 'body_paragraph', section: 0, paragraph: 1, charOffset: 0, length: 2 },
                text: '제목',
              },
            ]
          }
          if (method === 'applyBodyCharFormat') formatted.push(params?.paragraph as number)
          return { ok: true }
        },
      }),
      { bold: true },
      undefined,
      [0, 1],
    )
    expect(result.indexes).toEqual([1])
    expect(formatted).toEqual([1])
  })

  it('rejects character-only format on an empty paragraph', async () => {
    await expect(
      applyParagraphFormat(
        studio({
          _request: async (method) => {
            if (method === 'listBodyParagraphs') {
              return [
                {
                  editable: true,
                  reason: null,
                  target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                  text: '',
                },
              ]
            }
            return { ok: true }
          },
        }),
        { bold: true },
        0,
      ),
    ).rejects.toThrow(FORMAT_EMPTY_RANGE)
  })

  it('applies char format to every cell in a table row', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await applyTableCellFormat(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 2,
                control: 0,
                rows: 2,
                cols: 2,
                cells: [
                  { index: 0, row: 0, col: 0, text: '메뉴' },
                  { index: 1, row: 0, col: 1, text: '가격' },
                  { index: 2, row: 1, col: 0, text: '아메리카노' },
                  { index: 3, row: 1, col: 1, text: '4500' },
                ],
              },
            ]
          }
          return { ok: true }
        },
      }),
      { bold: true, fontSize: 13 },
      { table: 0, row: 0 },
    )
    expect(result).toEqual({
      table: 0,
      row: 0,
      cols: [0, 1],
      applied: ['bold=true', 'fontSize=13'],
    })
    const charCalls = calls.filter((call) => call.method === 'applyCellCharFormat')
    expect(charCalls).toHaveLength(2)
    expect(charCalls[0]?.params).toEqual({
      section: 0,
      paragraph: 2,
      control: 0,
      cellIndex: 0,
      cellPara: 0,
      start: 0,
      end: 2,
      format: { bold: true, fontSize: 1300 },
    })
    expect(charCalls[1]?.params).toMatchObject({ cellIndex: 1, end: 2 })
    expect(calls.some((call) => call.method === 'applyBodyCharFormat')).toBe(false)
  })

  it('formats one table cell when col is set', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await applyTableCellFormat(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 1,
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
          return { ok: true }
        },
      }),
      { align: 'center' },
      { table: 0, row: 0, col: 1 },
    )
    expect(result.cols).toEqual([1])
    expect(calls.filter((call) => call.method === 'applyCellParaFormat')).toHaveLength(1)
    expect(calls.find((call) => call.method === 'applyCellParaFormat')?.params).toMatchObject({
      cellIndex: 1,
      format: { alignment: 'center' },
    })
  })

  it('uses UTF-16 units for cell char ranges', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    await applyTableCellFormat(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 1,
                control: 0,
                rows: 1,
                cols: 1,
                cells: [{ index: 0, row: 0, col: 0, text: '😀메뉴' }],
              },
            ]
          }
          return { ok: true }
        },
      }),
      { bold: true },
      { table: 0, row: 0 },
    )
    expect(calls.find((call) => call.method === 'applyCellCharFormat')?.params).toMatchObject({
      start: 0,
      end: 4,
    })
  })

  it('rejects bold-only format on an empty table row', async () => {
    await expect(
      applyTableCellFormat(
        studio({
          _request: async (method) => {
            if (method === 'listTables') {
              return [
                {
                  section: 0,
                  paragraph: 1,
                  control: 0,
                  rows: 1,
                  cols: 2,
                  cells: [
                    { index: 0, row: 0, col: 0, text: '' },
                    { index: 1, row: 0, col: 1, text: '' },
                  ],
                },
              ]
            }
            return { ok: true }
          },
        }),
        { bold: true },
        { table: 0, row: 0 },
      ),
    ).rejects.toThrow(FORMAT_EMPTY_RANGE)
  })

  it('rejects a missing table index', async () => {
    await expect(
      applyTableCellFormat(
        studio({
          _request: async (method) => {
            if (method === 'listTables') {
              return [
                {
                  section: 0,
                  paragraph: 1,
                  control: 0,
                  rows: 1,
                  cols: 1,
                  cells: [{ index: 0, row: 0, col: 0, text: '칸' }],
                },
              ]
            }
            return { ok: true }
          },
        }),
        { bold: true },
        { table: 1, row: 0 },
      ),
    ).rejects.toThrow(TABLE_NOT_FOUND)
  })

  it('inserts a table row after the given index', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await editDocumentTable(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 2,
                control: 0,
                rows: 2,
                cols: 2,
                cells: [
                  { index: 0, row: 0, col: 0, text: 'a' },
                  { index: 1, row: 0, col: 1, text: 'b' },
                ],
              },
            ]
          }
          return { ok: true }
        },
      }),
      { action: 'insert_row', table: 0, row: 0 },
    )
    expect(result).toEqual({ table: 0, action: 'insert_row', detail: 'row after 0' })
    expect(calls.find((call) => call.method === 'insertTableRow')?.params).toEqual({
      section: 0,
      paragraph: 2,
      control: 0,
      row: 0,
      after: true,
    })
  })

  it('merges a header range and fills the first row', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    await editDocumentTable(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 1,
                control: 0,
                rows: 2,
                cols: 3,
                cells: [
                  { index: 0, row: 0, col: 0, text: 'h' },
                  { index: 1, row: 0, col: 1, text: '' },
                  { index: 2, row: 0, col: 2, text: '' },
                ],
              },
            ]
          }
          return { ok: true }
        },
      }),
      { action: 'merge', table: 0, row: 0, col: 0, endRow: 0, endCol: 2 },
    )
    expect(calls.find((call) => call.method === 'mergeTableCells')?.params).toMatchObject({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
    })
    const styled = await styleDocumentTable(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'listTables') {
            return [
              {
                section: 0,
                paragraph: 1,
                control: 0,
                rows: 2,
                cols: 3,
                cells: [
                  { index: 0, row: 0, col: 0, text: 'h' },
                  { index: 1, row: 0, col: 1, text: '' },
                  { index: 2, row: 0, col: 2, text: '' },
                ],
              },
            ]
          }
          return { ok: true }
        },
      }),
      { table: 0, row: 0, fill: 'EEE', valign: 'center' },
    )
    expect(styled.applied).toEqual(['fill=#eeeeee', 'valign=center'])
    expect(calls.filter((call) => call.method === 'setCellProperties')).toHaveLength(3)
    expect(calls.find((call) => call.method === 'setCellProperties')?.params).toMatchObject({
      cellIndex: 0,
      props: { fillType: 'solid', fillColor: '#eeeeee', verticalAlign: 'Center' },
    })
  })

  it('sets landscape A4 and two columns', async () => {
    const calls: Array<{ method: string; params?: Record<string, unknown> }> = []
    const result = await setDocumentPage(
      studio({
        _request: async (method, params) => {
          calls.push({ method, params })
          if (method === 'getPageDef') return { width: 59528, height: 84188, landscape: false }
          if (method === 'getColumnDef') return { columnCount: 1, columnType: 0, sameWidth: true, spacing: 0 }
          return { ok: true }
        },
      }),
      { paper: 'A4', orientation: 'landscape', columns: 2 },
    )
    expect(result.applied).toEqual(['paper=A4', 'orientation=landscape', 'columns=2'])
    expect(calls.find((call) => call.method === 'setPageDef')?.params).toMatchObject({
      section: 0,
      props: { width: 84188, height: 59528, landscape: true },
    })
    expect(calls.find((call) => call.method === 'setColumnDef')?.params).toMatchObject({
      section: 0,
      count: 2,
      sameWidth: true,
    })
  })

  it('rejects list format on table cells', async () => {
    await expect(
      applyTableCellFormat(
        studio({
          _request: async () => {
            throw new Error('should not list tables')
          },
        }),
        { list: 'bullet' },
        { table: 0, row: 0 },
      ),
    ).rejects.toThrow('list is not supported in table cells')
  })

  it('does not treat a generic format error as mixed formatting', async () => {
    await expect(
      applyParagraphFormat(
        studio({
          _request: async (method) => {
            if (method === 'listBodyParagraphs') {
              return [
                {
                  editable: false,
                  reason: 'cannot apply format to control',
                  target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 0 },
                  text: '',
                },
              ]
            }
            return { ok: true }
          },
        }),
        { bold: true },
        0,
      ),
    ).rejects.toThrow('cannot apply format to control')
  })

  it('rejects out-of-range font size instead of clamping', async () => {
    await expect(
      applyParagraphFormat(
        studio({
          _request: async (method) => {
            if (method === 'listBodyParagraphs') {
              return [
                {
                  editable: true,
                  reason: null,
                  target: { kind: 'body_paragraph', section: 0, paragraph: 0, charOffset: 0, length: 2 },
                  text: '제목',
                },
              ]
            }
            return { ok: true }
          },
        }),
        { fontSize: 80 },
        0,
      ),
    ).rejects.toThrow(/between 8 and 72/)
  })

  it('splices UTF-16 ranges and rejects offsets past the string', () => {
    expect(spliceParagraphText('안녕세계', 2, 4, '세상')).toBe('안녕세상')
    expect(spliceParagraphText('ab', 0, 1, 'z')).toBe('zb')
    expect(() => spliceParagraphText('ab', 0, 3, 'z')).toThrow('nothing is selected')
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
