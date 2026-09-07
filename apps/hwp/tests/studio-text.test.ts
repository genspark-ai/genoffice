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
