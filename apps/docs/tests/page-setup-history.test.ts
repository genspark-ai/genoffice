import { Editor } from '@tiptap/core'
import type { Transaction } from '@tiptap/pm/state'
import {
  applySectionSettings,
  sectionSettingsFromXml,
  type SectionInfo,
  type SectionSettings,
} from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  pageSetupChange,
  patchSections,
  recordPageSetupChange,
  replayedPageSetup,
  watermarkStateOf,
  type PageSetupPatch,
  type PageSetupState,
} from '../src/renderer/editor/page-setup-history'

const para = (text: string) => ({
  type: 'docParagraph',
  attrs: {},
  content: [{ type: 'text', text }],
})

const LETTER = {
  pageWidth: 12240,
  pageHeight: 15840,
  orientation: 'portrait',
  marginTop: 1440,
  marginRight: 1440,
  marginBottom: 1440,
  marginLeft: 1440,
  columns: 1,
} as unknown as SectionSettings

const info = (settings: SectionSettings, extra: Partial<SectionInfo> = {}): SectionInfo => ({
  settings,
  startType: 'nextPage',
  firstBlockIndex: 0,
  lastBlockIndex: 0,
  sectPrXml: '<w:sectPr/>',
  titlePg: false,
  headerRefs: {},
  footerRefs: {},
  ...extra,
})

const state = (settings: SectionSettings, extra: Partial<PageSetupState> = {}): PageSetupState => ({
  section: settings,
  sections: [info(settings)],
  trailingStartType: null,
  mirrorMargins: false,
  gutterAtTop: false,
  titlePg: false,
  evenOddHf: false,
  notes: { footnote: {}, endnote: {} },
  sectionNotes: [
    {
      footnote: { ...extra.notes?.footnote, ...settings.footnotePr },
      endnote: { ...extra.notes?.endnote, ...settings.endnotePr },
    },
  ],
  autoHyphenation: false,
  hyphenation: { caps: true, limit: 0, zoneTwips: 360 },
  pageColor: null,
  watermark: { text: null, style: null, picture: null, dirty: false },
  typography: { spacingControl: 'none', strictKinsoku: false, kinsokuCustom: null },
  ...extra,
})

function harness() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para('alpha')] } as never,
  })
  let live = state(LETTER)
  const replayed: PageSetupPatch[] = []
  editor.on('transaction', ({ transaction }: { transaction: Transaction }) => {
    for (const patch of replayedPageSetup(transaction)) {
      replayed.push(patch)
      live = {
        ...live,
        ...patch,
        section: live.section && { ...live.section, ...patch.section },
        sections: patchSections(live.sections, patch),
        notes: {
          footnote: { ...live.notes.footnote, ...patch.notes?.footnote },
          endnote: { ...live.notes.endnote, ...patch.notes?.endnote },
        },
        hyphenation: { ...live.hyphenation, ...patch.hyphenation },
        watermark: patch.watermark?.state ?? live.watermark,
        typography: { ...live.typography, ...patch.typography },
      }
    }
  })
  const change = (next: PageSetupState, joinPrevious = false) => {
    const diff = pageSetupChange(live, next)
    live = next
    if (diff) recordPageSetupChange(editor, diff, joinPrevious)
  }
  return { editor, change, replayed, live: () => live }
}

describe('page setup change', () => {
  it('is nothing when the two states hold the same values', () => {
    expect(pageSetupChange(state(LETTER), state({ ...LETTER }))).toBeNull()
  })

  it('carries only the settings that differ, a removed one as undefined', () => {
    const before = state({ ...LETTER, vAlign: 'center' } as SectionSettings)
    const after = state({ ...LETTER, marginLeft: 2880 })
    const change = pageSetupChange(before, after)!
    expect(change.undo.section).toEqual({ marginLeft: 1440, vAlign: 'center' })
    expect(change.redo.section).toEqual({ marginLeft: 2880, vAlign: undefined })
    expect(change.redo.sections).toEqual({ 0: { settings: { marginLeft: 2880 } } })
    expect('mirrorMargins' in change.undo).toBe(false)
  })

  it('leaves alone what another change made to the same section since', () => {
    const wide = { ...LETTER, marginLeft: 2880 }
    const { undo } = pageSetupChange(state(LETTER), state(wide))!
    const since = [info({ ...wide, columns: 2 }, { titlePg: true, pageNumberStart: 5 })]
    const [back] = patchSections(since, undo)
    expect(back.settings.marginLeft).toBe(1440)
    expect(back.settings.columns).toBe(2)
    expect(back.titlePg).toBe(true)
    expect(back.pageNumberStart).toBe(5)
  })

  it('puts back a gutter and equal columns the file did not state, after a save wrote the new ones', () => {
    const fileXml =
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '<w:cols w:space="720"/></w:sectPr>'
    const before = sectionSettingsFromXml(fileXml)
    expect(before.gutter).toBeUndefined()
    const after: SectionSettings = {
      ...before,
      gutter: 500,
      marginLeft: 1940,
      columns: 2,
      colEqualWidth: false,
      colWidths: [3000, 5000],
      colSpaces: [720],
    }
    const savedXml = applySectionSettings(fileXml, after)
    expect(savedXml).toContain('w:gutter="500"')
    expect(savedXml).toContain('w:equalWidth="0"')
    const { undo } = pageSetupChange(state(before), state(after))!
    const [back] = patchSections([info(sectionSettingsFromXml(savedXml))], undo)
    const undoneXml = applySectionSettings(savedXml, back.settings)
    expect(sectionSettingsFromXml(undoneXml)).toEqual(before)
    expect(undoneXml).toContain('w:gutter="0"')
    expect(undoneXml).toContain('w:left="1440"')
    expect(undoneXml).not.toContain('<w:col ')
  })

  it.each<[string, Partial<SectionSettings>]>([
    ['header and footer distance', { headerDist: 300, footerDist: 400 }],
    ['text direction', { textDirection: 'tbRl' }],
    ['vertical alignment', { vAlign: 'center' }],
    ['a rule between two columns', { columns: 2, colSep: true }],
    ['line numbers', { lineNumbers: { countBy: 1, start: 1, restart: 'continuous' } }],
    ['a line grid', { docGrid: { type: 'lines', linePitch: 400 } }],
    ['a page border', { pageBorder: true }],
    ['landscape', { orientation: 'landscape', pageWidth: 15840, pageHeight: 12240 }],
  ])('takes %s out of the file again when undone after a save', (_name, set) => {
    const fileXml =
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '<w:cols w:space="720"/></w:sectPr>'
    const before = sectionSettingsFromXml(fileXml)
    const after = { ...before, ...set } as SectionSettings
    const savedXml = applySectionSettings(fileXml, after)
    expect(sectionSettingsFromXml(savedXml)).toMatchObject(set)
    const { undo } = pageSetupChange(state(before), state(after))!
    const [back] = patchSections([info(sectionSettingsFromXml(savedXml))], undo)
    const undoneXml = applySectionSettings(savedXml, back.settings)
    // a grid goes back to no grid by losing its type: the element and its pitch stay, and snap nothing
    const { docGrid, ...written } = sectionSettingsFromXml(undoneXml)
    expect(written).toEqual(before)
    expect(docGrid?.type ?? 'default').toBe('default')
  })

  it('states the note properties a section did not have with the values that were in effect', () => {
    const base = { footnote: { numFmt: 'upperRoman' }, endnote: {} }
    const before = state(LETTER, { notes: base })
    const after = state(
      { ...LETTER, footnotePr: { numFmt: 'lowerLetter', numStart: 3 } } as never,
      {
        notes: { footnote: { numFmt: 'lowerLetter', numStart: 3 }, endnote: {} },
      },
    )
    const { undo, redo } = pageSetupChange(before, after)!
    expect(undo.notes).toEqual({ footnote: { numFmt: 'upperRoman', numStart: 1 } })
    expect(redo.notes).toEqual({ footnote: { numFmt: 'lowerLetter', numStart: 3 } })
    expect(undo.section).toEqual({ footnotePr: { numFmt: 'upperRoman', numStart: 1 } })
    expect(undo.sections).toEqual({
      0: { settings: { footnotePr: { numFmt: 'upperRoman', numStart: 1 } } },
    })
    expect('endnote' in undo.notes!).toBe(false)
  })

  it('writes the earlier note numbering of a section when undone after a save', () => {
    const fileXml =
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '<w:cols w:space="720"/></w:sectPr>'
    const before = sectionSettingsFromXml(fileXml)
    const after: SectionSettings = {
      ...before,
      endnotePr: { numFmt: 'decimal', numRestart: 'eachSect' },
    }
    const savedXml = applySectionSettings(fileXml, after)
    expect(savedXml).toContain('<w:numFmt w:val="decimal"/>')
    expect(savedXml).toContain('<w:numRestart w:val="eachSect"/>')
    const { undo } = pageSetupChange(state(before), state(after))!
    const [back] = patchSections([info(sectionSettingsFromXml(savedXml))], undo)
    const undoneXml = applySectionSettings(savedXml, back.settings)
    expect(undoneXml).toContain('<w:numFmt w:val="lowerRoman"/>')
    expect(undoneXml).toContain('<w:numRestart w:val="continuous"/>')
    expect(undoneXml).not.toContain('eachSect')
  })

  it('carries hyphenation: the switch, and only the rules that differ', () => {
    const after = state(LETTER, {
      autoHyphenation: true,
      hyphenation: { caps: false, limit: 2, zoneTwips: 360 },
    })
    const { undo, redo } = pageSetupChange(state(LETTER), after)!
    expect(undo).toEqual({ autoHyphenation: false, hyphenation: { caps: true, limit: 0 } })
    expect(redo).toEqual({ autoHyphenation: true, hyphenation: { caps: false, limit: 2 } })
  })

  it('does not describe a change that added a section', () => {
    const after = state(LETTER)
    after.sections = [info(LETTER, { pendingBreak: true }), info(LETTER)]
    expect(pageSetupChange(state(LETTER), after)).toBeNull()
  })

  it('carries the page number format and a first-page choice of a section', () => {
    const before = state(LETTER)
    const after = state(LETTER)
    after.sections = [
      info(LETTER, {
        pageNumberFmt: 'lowerRoman',
        pageNumberStart: 3,
        pageNumberChapter: { style: 1, sep: 'hyphen' },
        titlePg: true,
      }),
    ]
    const { undo, redo } = pageSetupChange(before, after)!
    expect(undo.sections).toEqual({ 0: { pgNum: { chapter: null }, titlePg: false } })
    expect(redo.sections).toEqual({
      0: {
        pgNum: { fmt: 'lowerRoman', start: 3, chapter: { style: 1, sep: 'hyphen' } },
        titlePg: true,
      },
    })
    const back = patchSections(after.sections, undo)[0]
    expect([
      back.pageNumberFmt,
      back.pageNumberStart,
      back.pageNumberChapter,
      back.titlePg,
    ]).toEqual([undefined, undefined, undefined, false])
    expect(patchSections(before.sections, redo)).toEqual(after.sections)
  })

  it('carries the page colour, the watermark as it stood, and only the typography options that differ', () => {
    const before = state(LETTER, { pageColor: null })
    const draft = watermarkStateOf({ text: 'DRAFT', colorHex: 'FF0000' })
    const after = state(LETTER, {
      pageColor: 'FFF2CC',
      watermark: draft,
      typography: { spacingControl: 'kana', strictKinsoku: false, kinsokuCustom: null },
    })
    const { undo, redo } = pageSetupChange(before, after)!
    expect(undo).toEqual({
      pageColor: null,
      // the file's own watermark: not one the save writes
      watermark: { state: { text: null, style: null, picture: null, dirty: false } },
      typography: { spacingControl: 'none' },
    })
    expect(redo).toEqual({
      pageColor: 'FFF2CC',
      watermark: { state: draft },
      typography: { spacingControl: 'kana' },
    })
    expect(draft).toEqual({
      text: 'DRAFT',
      style: { colorHex: 'FF0000' },
      picture: null,
      dirty: true,
    })
    expect(watermarkStateOf(null)).toEqual({ text: null, style: null, picture: null, dirty: true })
    expect(watermarkStateOf('COPY').style).toBeNull()
  })

  it('carries the sectPr of a section behind an unsaved break', () => {
    const before = state(LETTER)
    before.sections = [info(LETTER, { pendingBreak: true, sectPrXml: '<a/>' }), info(LETTER)]
    const after = state(LETTER)
    after.sections = [
      info({ ...LETTER, columns: 2 }, { pendingBreak: true, sectPrXml: '<b/>' }),
      info(LETTER),
    ]
    const { undo } = pageSetupChange(before, after)!
    expect(undo.sections).toEqual({ 0: { settings: { columns: 1 }, sectPrXml: '<a/>' } })
    expect(patchSections(after.sections, undo)[0].sectPrXml).toBe('<a/>')
    // once saved the section is an ordinary one, and its sectPr is the file's
    const saved = [info(after.sections[0].settings, { sectPrXml: '<saved/>' }), info(LETTER)]
    expect(patchSections(saved, undo)[0].sectPrXml).toBe('<saved/>')
  })
})

describe('page setup in the undo history', () => {
  it('is one step, undone and redone in order with text edits', () => {
    const { editor, change, replayed, live } = harness()
    editor.commands.insertContentAt(6, ' one')
    change(state({ ...LETTER, marginLeft: 2880, columns: 2 }, { mirrorMargins: true }))
    editor.commands.insertContentAt(10, ' two')
    expect(editor.state.doc.textContent).toBe('alpha one two')
    expect(replayed).toEqual([])

    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('alpha one')
    expect(replayed).toEqual([])

    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('alpha one')
    expect(replayed).toHaveLength(1)
    expect(live().section).toMatchObject({ marginLeft: 1440, columns: 1 })
    expect(live().sections[0].settings).toMatchObject({ marginLeft: 1440, columns: 1 })
    expect(live().mirrorMargins).toBe(false)

    editor.commands.redo()
    expect(replayed).toHaveLength(2)
    expect(live().section).toMatchObject({ marginLeft: 2880, columns: 2 })
    expect(live().mirrorMargins).toBe(true)

    editor.commands.undo()
    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('alpha')
    expect(editor.can().undo()).toBe(false)
    editor.destroy()
  })

  it('takes note and hyphenation settings as one step each, in order with text edits', () => {
    const { editor, change, replayed, live } = harness()
    change(state(LETTER, { notes: { footnote: { numFmt: 'upperRoman' }, endnote: {} } }))
    editor.commands.insertContentAt(6, ' one')
    change({
      ...live(),
      autoHyphenation: true,
      hyphenation: { caps: false, limit: 3, zoneTwips: 360 },
    })

    editor.commands.undo()
    expect(replayed).toHaveLength(1)
    expect(live().autoHyphenation).toBe(false)
    expect(live().hyphenation).toEqual({ caps: true, limit: 0, zoneTwips: 360 })
    expect(live().notes.footnote.numFmt).toBe('upperRoman')

    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('alpha')
    expect(replayed).toHaveLength(1)

    editor.commands.undo()
    expect(replayed).toHaveLength(2)
    expect(live().notes.footnote.numFmt).toBe('decimal')
    expect(editor.can().undo()).toBe(false)

    editor.commands.redo()
    editor.commands.redo()
    editor.commands.redo()
    expect(live().notes.footnote.numFmt).toBe('upperRoman')
    expect(live().autoHyphenation).toBe(true)
    expect(live().hyphenation.limit).toBe(3)
    expect(editor.state.doc.textContent).toBe('alpha one')
    editor.destroy()
  })

  it('takes a page colour and a typography option as one step each, in order with text edits', () => {
    const { editor, change, replayed, live } = harness()
    change(state(LETTER, { pageColor: 'FFF2CC' }))
    editor.commands.insertContentAt(6, ' one')
    change(
      state(LETTER, {
        pageColor: 'FFF2CC',
        typography: { spacingControl: 'punct', strictKinsoku: true, kinsokuCustom: null },
      }),
    )
    editor.commands.undo()
    expect(replayed).toEqual([{ typography: { spacingControl: 'none', strictKinsoku: false } }])
    expect(live().pageColor).toBe('FFF2CC')
    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('alpha')
    expect(replayed).toHaveLength(1)
    editor.commands.undo()
    expect(live().pageColor).toBeNull()
    editor.commands.redo()
    editor.commands.redo()
    editor.commands.redo()
    expect(live().pageColor).toBe('FFF2CC')
    expect(live().typography).toEqual({
      spacingControl: 'punct',
      strictKinsoku: true,
      kinsokuCustom: null,
    })
    expect(editor.state.doc.textContent).toBe('alpha one')
    editor.destroy()
  })

  it('loses its redo to an edit made after the undo', () => {
    const { editor, change, replayed, live } = harness()
    change(state({ ...LETTER, marginLeft: 2880 }))
    editor.commands.undo()
    expect(editor.can().redo()).toBe(true)
    editor.commands.insertContentAt(6, '!')
    expect(editor.can().redo()).toBe(false)
    editor.commands.redo()
    expect(replayed).toHaveLength(1)
    expect(live().section?.marginLeft).toBe(1440)
    editor.destroy()
  })

  it('makes the document a new one without touching its content', () => {
    const { editor, change } = harness()
    const before = editor.state.doc
    change(state({ ...LETTER, marginLeft: 2880 }))
    expect(editor.state.doc).not.toBe(before)
    expect(editor.state.doc.eq(before)).toBe(true)
    expect(editor.state.doc.firstChild).toBe(before.firstChild)
    editor.destroy()
  })

  it('undoes together with the document edit the same command made', () => {
    const { editor, change, replayed } = harness()
    editor.commands.insertContentAt(6, ' one')
    change(state({ ...LETTER, columns: 2 }), true)
    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('alpha')
    expect(replayed).toHaveLength(1)
    expect(editor.can().undo()).toBe(false)
    editor.destroy()
  })
})
