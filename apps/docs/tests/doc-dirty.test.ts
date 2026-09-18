import { describe, expect, it, vi } from 'vitest'
import {
  isDocDirty,
  resetCrossDocEditState,
  type CrossDocEditStateSink,
  type DocDirtyState,
} from '../src/renderer/doc-dirty'
import { openedFileStartsDirty } from '../src/renderer/doc-state'

function cleanState(): DocDirtyState {
  return {
    dirtyRef: { current: false },
    sectionDirty: false,
    sectionsDirty: [],
    trailingStartType: null,
    pageColorDirty: false,
    headerDirty: false,
    footerDirty: false,
    hfVariantsDirty: [],
    sectionHfEdits: {},
    pgNumEdit: null,
    pgNumDirtySections: [],
    numberingDirty: false,
    styleUpserts: {},
    titlePgDirty: false,
    evenOddHfDirty: false,
    watermarkDirty: false,
    inksDirty: false,
    notesDirty: false,
    sourcesDirty: false,
    zoteroDocumentDataDirty: false,
    themeFontsDirty: false,
    themeColorsDirty: false,
    commentsDirty: false,
    protectionDirty: false,
    writeProtectionDirty: false,
    removePersonalInfoDirty: false,
  }
}

describe('isDocDirty', () => {
  it('is false for a fully clean state', () => {
    expect(isDocDirty(cleanState())).toBe(false)
  })

  it('is true when only the body changed', () => {
    expect(isDocDirty({ ...cleanState(), dirtyRef: { current: true } })).toBe(true)
  })

  it('treats restored recovery content as unsaved body content', () => {
    const state = cleanState()
    state.dirtyRef.current = openedFileStartsDirty({ recovered: true })
    expect(isDocDirty(state)).toBe(true)
    expect(openedFileStartsDirty({})).toBe(false)
  })

  // the regression: autosave ignored comment/protection edits; the recovery push ignored everything but the body
  it.each([
    ['commentsDirty', { commentsDirty: true }],
    ['protectionDirty', { protectionDirty: true }],
    ['writeProtectionDirty', { writeProtectionDirty: true }],
    ['removePersonalInfoDirty', { removePersonalInfoDirty: true }],
    ['sectionDirty', { sectionDirty: true }],
    ['headerDirty', { headerDirty: true }],
    ['footerDirty', { footerDirty: true }],
    ['pageColorDirty', { pageColorDirty: true }],
    ['titlePgDirty', { titlePgDirty: true }],
    ['evenOddHfDirty', { evenOddHfDirty: true }],
    ['watermarkDirty', { watermarkDirty: true }],
    ['inksDirty', { inksDirty: true }],
    ['notesDirty', { notesDirty: true }],
    ['sourcesDirty', { sourcesDirty: true }],
    ['zoteroDocumentDataDirty', { zoteroDocumentDataDirty: true }],
    ['themeFontsDirty', { themeFontsDirty: true }],
    ['themeColorsDirty', { themeColorsDirty: true }],
    ['numberingDirty', { numberingDirty: true }],
    ['sectionsDirty', { sectionsDirty: [1] }],
    ['hfVariantsDirty', { hfVariantsDirty: ['headerFirst'] }],
    ['pgNumDirtySections', { pgNumDirtySections: [0] }],
    ['sectionHfEdits', { sectionHfEdits: { '3:header': {} } }],
    ['defaultFonts', { defaultFonts: { eastAsiaFont: 'SimSun' } }],
    ['styleUpserts', { styleUpserts: { Heading1: {} } }],
    ['pgNumEdit', { pgNumEdit: { fmt: 'decimal' } }],
    ['trailingStartType', { trailingStartType: 'nextPage' }],
  ] as const)('is true when only %s changed', (_name, patch) => {
    expect(isDocDirty({ ...cleanState(), ...patch })).toBe(true)
  })
})

describe('resetCrossDocEditState', () => {
  function mockSink(): CrossDocEditStateSink & Record<string, ReturnType<typeof vi.fn>> {
    return {
      setSectionsDirty: vi.fn(),
      setTrailingStartType: vi.fn(),
      setSectionHfEdits: vi.fn(),
      setPgNumEdit: vi.fn(),
      setPgNumDirtySections: vi.fn(),
      setPendingNumbering: vi.fn(),
      setStyleUpserts: vi.fn(),
      setDefaultFonts: vi.fn(),
    }
  }

  it('clears every section/numbering/style edit state a swap must not inherit', () => {
    const sink = mockSink()
    resetCrossDocEditState(sink)
    expect(sink.setSectionsDirty).toHaveBeenCalledWith([])
    expect(sink.setTrailingStartType).toHaveBeenCalledWith(null)
    expect(sink.setSectionHfEdits).toHaveBeenCalledWith({})
    expect(sink.setPgNumEdit).toHaveBeenCalledWith(null)
    expect(sink.setPgNumDirtySections).toHaveBeenCalledWith([])
    expect(sink.setPendingNumbering).toHaveBeenCalledWith({ newDefs: [], restartNums: [] })
    expect(sink.setStyleUpserts).toHaveBeenCalledWith({})
    expect(sink.setDefaultFonts).toHaveBeenCalledWith(undefined)
  })

  it('leaves a state carrying those edits clean afterwards', () => {
    // the cross-document leak: doc A set all of these, doc B opens pristine
    const leaked: DocDirtyState = {
      ...cleanState(),
      sectionsDirty: [0],
      trailingStartType: 'nextPage',
      sectionHfEdits: { '0:header': {} },
      pgNumEdit: { fmt: 'decimal' },
      pgNumDirtySections: [1],
      styleUpserts: { Heading1: {} },
    }
    expect(isDocDirty(leaked)).toBe(true)
    const cleared: DocDirtyState = {
      ...leaked,
      sectionsDirty: [],
      trailingStartType: null,
      sectionHfEdits: {},
      pgNumEdit: null,
      pgNumDirtySections: [],
      styleUpserts: {},
    }
    // numberingDirty is the live flag; pendingNumbering itself is not polled,
    // but buildDocBytes applies its indices — hence the reset
    expect(isDocDirty({ ...cleared, numberingDirty: false })).toBe(false)
  })
})
