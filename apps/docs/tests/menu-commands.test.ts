/**
 * The application menu's commands run through one handler table. These pin
 * which host action each command reaches and the guards in front of it.
 */
import { insertCaptionBlock } from '../src/renderer/components/ReferenceDialogs'
import { describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { COMMANDS, type ActionCommandId, type CommandDef } from '../src/shared/commands'
import {
  MENU_COMMAND_HANDLERS,
  NONE,
  runMenuCommand,
  canRepeatCommand,
  type CommandHost,
} from '../src/renderer/commands/menu-commands'
import { setActiveSubEditor } from '../src/renderer/editor/active-editor'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { grammarState, setGrammarFindings } from '../src/renderer/editor/grammar-marks'
import { ignoredOnce } from '../src/renderer/editor/spell-menu'
import { autocorrectLists } from '../src/renderer/autocorrect-pref'
import { hfHasPageField } from '../src/renderer/editor/hf-dom'
import { mountHfEditor } from '../src/renderer/editor/hf-editor'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'plain paragraph' }] }],
    },
  })
}

function makeHost(over: Partial<CommandHost> = {}): CommandHost {
  return {
    editor: null,
    hasDoc: true,
    newFile: vi.fn(),
    openFile: vi.fn(),
    openRecent: vi.fn(),
    save: vi.fn(),
    nameWithAi: vi.fn(),
    exportPdf: vi.fn(),
    exportHtml: vi.fn(),
    exportMarkdown: vi.fn(),
    exportImages: vi.fn(),
    printDoc: vi.fn(),
    stepZoom: vi.fn(),
    setZoom: vi.fn(),
    zoomFit: vi.fn(),
    toggleAi: vi.fn(),
    toggleDark: vi.fn(),
    toggleMarks: vi.fn(),
    toggleStylesPane: vi.fn(),
    stepFontSize: vi.fn(),
    clearFormatting: vi.fn(),
    listNumId: vi.fn(() => '7'),
    createListDef: vi.fn(() => '9'),
    listStyleNum: vi.fn(() => ({ numId: '11', kind: 'ordered' as const })),
    openListDialog: vi.fn(),
    addListStyle: vi.fn(),
    modifyListStyle: vi.fn(),
    deleteListStyle: vi.fn(),
    applyStyle: vi.fn(),
    aiPreset: vi.fn(),
    openDialog: vi.fn(),
    openChartDialog: vi.fn(),
    openTableDialog: vi.fn(),
    findStep: vi.fn(),
    pageLayout: vi.fn(() => null),
    openMailMerge: vi.fn(),
    labelMainDocument: false,
    pagesOf: vi.fn(() => null),
    checkSpelling: vi.fn(),
    thesaurus: vi.fn(),
    applyHeadingLevel: vi.fn(),
    translate: vi.fn(),
    pickMergeDataSource: vi.fn(),
    toggleMergePreview: vi.fn(),
    hasMergeSource: false,
    openRuleDialog: vi.fn(),
    onDisk: true,
    moveFile: vi.fn(),
    toggleReadOnlyRecommended: vi.fn(),
    editNote: vi.fn(),
    convertNote: vi.fn(),
    noNotes: vi.fn(),
    openTabsDialog: vi.fn(),
    openPhoneticGuide: vi.fn(),
    insertFile: vi.fn(),
    openTextEffects: vi.fn(),
    openEqLayout: vi.fn(),
    openFitText: vi.fn(),
    openMarkCitation: vi.fn(),
    openMarkIndexEntry: vi.fn(),
    openMarkTocEntry: vi.fn(),
    nextMisspelling: vi.fn(),
    showMergeFields: vi.fn(),
    copyTextPrompt: vi.fn(),
    autoMarkIndex: vi.fn(),
    setIndexSection: vi.fn(),
    openAutoText: vi.fn(),
    newFromTemplate: vi.fn(),
    saveAsTemplate: vi.fn(),
    openLanguage: vi.fn(),
    toggleSelectionPane: vi.fn(),
    toggleCitationsPane: vi.fn(),
    toggleDrawing: vi.fn(),
    setDesignAsDefault: vi.fn(),
    nothingToCrop: vi.fn(),
    selectChartData: vi.fn(),
    showObjectFormatTab: vi.fn(),
    citationHeading: vi.fn((n: number) => `Category ${n}`),
    openAsianLayout: vi.fn(),
    openPasteSpecial: vi.fn(),
    openPictureDialog: vi.fn(),
    pickImage: vi.fn(async () => null),
    openFind: vi.fn(),
    showLayoutTab: vi.fn(),
    showTableDesignTab: vi.fn(),
    defineTableStyle: vi.fn(),
    insertField: vi.fn(),
    insertNote: vi.fn(),
    blocks: [],
    sources: [],
    citationStyle: 'apa',
    setCitationStyle: vi.fn(),
    headingPages: vi.fn(() => null),
    section: null,
    headerFooter: {
      header: null,
      footer: null,
      setHeader: vi.fn(),
      setFooter: vi.fn(),
      setFacingPageNumber: vi.fn(),
      edit: vi.fn(),
      openPageNumberFormat: vi.fn(),
    },
    headerFooterAction: vi.fn(),
    setInkTool: vi.fn(),
    clearInk: vi.fn(),
    view: {
      mode: 'print',
      setMode: vi.fn(),
      readMode: false,
      setReadMode: vi.fn(),
      openPagePreview: vi.fn(),
      openZoomDialog: vi.fn(),
      toggleRuler: vi.fn(),
      toggleGrid: vi.fn(),
      toggleNavigation: vi.fn(),
      toggleThumbnails: vi.fn(),
      toggleSplit: vi.fn(),
      toggleRibbon: vi.fn(),
      toggleFocus: vi.fn(),
      newWindow: vi.fn(),
      arrangeAll: vi.fn(),
      mergeWindows: vi.fn(),
      switchWindow: vi.fn(),
      closePane: vi.fn(),
    },
    design: {
      setTheme: vi.fn(),
      setPageColor: vi.fn(),
      setWatermark: vi.fn(),
      setParagraphSpacing: vi.fn(),
      applyStyleSet: vi.fn(),
      resetStyleSet: vi.fn(),
      saveStyleSet: vi.fn(),
      deleteStyleSet: vi.fn(),
      savedTheme: vi.fn(),
      openedTheme: vi.fn(),
      browseTheme: vi.fn(),
      saveTheme: vi.fn(),
      deleteTheme: vi.fn(),
    },
    review: {
      revisionDisplay: 'all',
      toggleSpellcheck: vi.fn(),
      resetIgnoredWords: vi.fn(),
      toggleGrammarMarks: vi.fn(),
      deleteComment: vi.fn(),
      openHighlightChanges: vi.fn(),
      resolveComment: vi.fn(),
      reopenComment: vi.fn(),
      replyToComment: vi.fn(),
      editComment: vi.fn(),
      replaceMisspelling: vi.fn(),
      ignoreSpelling: vi.fn(),
      learnSpelling: vi.fn(),
      respell: vi.fn(),
      deleteAllComments: vi.fn(),
      gotoComment: vi.fn(),
      toggleComments: vi.fn(),
      toggleTrackChanges: vi.fn(),
      setRevisionDisplay: vi.fn(),
      toggleMarkup: vi.fn(),
      setBalloonMode: vi.fn(),
      openTrackChangesOptions: vi.fn(),
      resolveRevisions: vi.fn(),
      gotoRevision: vi.fn(),
      toggleReviewPane: vi.fn(),
      compare: vi.fn(),
      combine: vi.fn(),
      protect: vi.fn(),
    },
    gridBasePt: 10.5,
    setSection: vi.fn(),
    setSectionsIn: vi.fn(),
    setMirrorMargins: vi.fn(),
    setGutterAtTop: vi.fn(),
    setSectionStart: vi.fn(),
    insertSectionBreak: vi.fn(),
    insertImage: vi.fn(),
    openLinks: vi.fn(),
    insertPageBreak: vi.fn(),
    insertBlankPage: vi.fn(),
    newComment: vi.fn(),
    wordCount: vi.fn(),
    aiProofread: vi.fn(),
    toggleTableGridlines: vi.fn(),
    tableSectionWidthPx: 600,
    notInTable: vi.fn(),
    ...over,
  }
}

/** command, optional payload → the host action it reaches and its arguments */
const ROUTES: [ActionCommandId, string | undefined, keyof CommandHost, unknown[]][] = [
  ['new', undefined, 'newFile', []],
  ['open', undefined, 'openFile', []],
  ['open-path', '/tmp/a.docx', 'openRecent', ['/tmp/a.docx']],
  ['save', undefined, 'save', [false]],
  ['save-as', undefined, 'save', [true]],
  ['name-with-ai', undefined, 'nameWithAi', []],
  ['zoom-in', undefined, 'stepZoom', [1]],
  ['zoom-out', undefined, 'stepZoom', [-1]],
  ['zoom-100', undefined, 'setZoom', [100]],
  ['zoom-set', '150', 'setZoom', [150]],
  ['zoom-page-width', undefined, 'zoomFit', ['width']],
  ['zoom-whole-page', undefined, 'zoomFit', ['page']],
  ['zoom-multi-page', undefined, 'zoomFit', ['multi']],
  ['toggle-ai', undefined, 'toggleAi', []],
  ['toggle-dark', undefined, 'toggleDark', []],
  ['show-marks', undefined, 'toggleMarks', []],
  ['list-dialog', 'number', 'openListDialog', ['number', undefined]],
  ['list-dialog', 'bullet#2', 'openListDialog', ['bullet', 2]],
  ['list-dialog', 'style', 'openListDialog', ['style', undefined]],
  ['list-dialog', 'styleModify#1ai', 'openListDialog', ['styleModify', undefined, '1ai']],
  ['ai-preset', 'Summarize this', 'aiPreset', ['Summarize this']],
  ['styles-pane', undefined, 'toggleStylesPane', []],
  ['font-dialog', undefined, 'openDialog', ['font']],
  ['paragraph-dialog', undefined, 'openDialog', ['paragraph']],
  ['autocorrect-options', undefined, 'openDialog', ['autocorrect']],
  ['preferences', undefined, 'openDialog', ['preferences']],
  ['shortcuts', undefined, 'openDialog', ['shortcuts']],
  ['page-setup', undefined, 'showLayoutTab', []],
  ['find', undefined, 'openFind', ['find']],
  ['replace', undefined, 'openFind', ['replace']],
  ['goto', undefined, 'openFind', ['goto']],
  ['export-pdf', undefined, 'exportPdf', []],
  ['export-html', undefined, 'exportHtml', []],
  ['export-markdown', undefined, 'exportMarkdown', []],
  ['export-images', undefined, 'exportImages', []],
  ['print', undefined, 'printDoc', []],
  ['word-count', undefined, 'wordCount', []],
  ['ai-proofread', undefined, 'aiProofread', []],
  ['table-gridlines', undefined, 'toggleTableGridlines', []],
]

/** commands that also need an editor */
const EDITOR_ROUTES: [ActionCommandId, keyof CommandHost, unknown[]][] = [
  ['insert-table', 'openDialog', ['table']],
  ['insert-link', 'openDialog', ['link']],
  ['insert-equation', 'openDialog', ['equation']],
  ['insert-image', 'insertImage', []],
  ['edit-links', 'openLinks', []],
  ['insert-page-break', 'insertPageBreak', []],
  ['insert-blank-page', 'insertBlankPage', []],
  ['insert-chart', 'openChartDialog', [undefined]],
  ['insert-bookmark', 'openDialog', ['bookmark']],
  ['insert-caption', 'openDialog', ['caption']],
  ['format-document', 'openDialog', ['document']],
  ['format-equation-options', 'openDialog', ['equationOptions']],
  ['format-change-case', 'openDialog', ['changeCase']],
  ['insert-field-dialog', 'openDialog', ['field']],
  ['insert-date-time', 'openDialog', ['dateTime']],
  ['format-drop-cap', 'openDialog', ['dropCap']],
  ['insert-watermark', 'openDialog', ['watermark']],
  ['insert-cross-reference', 'openDialog', ['crossReference']],
  ['insert-comment', 'newComment', []],
]

/** commands that do nothing without an open document */
const NEEDS_DOC: ActionCommandId[] = [
  'font-dialog',
  'paragraph-dialog',
  'find',
  'replace',
  'goto',
  'print',
  'word-count',
  'ai-proofread',
]

describe('menu command handlers', () => {
  it('has a handler for every command that runs, and none for native roles or placeholders', () => {
    const sent = Object.entries<CommandDef>(COMMANDS)
      .filter(([, def]) => !def.role && !def.planned)
      .map(([id]) => id)
    expect(Object.keys(MENU_COMMAND_HANDLERS).sort()).toEqual(sent.sort())
  })

  it.each(ROUTES)('%s reaches the host', (command, payload, action, args) => {
    const host = makeHost()
    runMenuCommand(host, command, payload)
    expect(host[action]).toHaveBeenCalledTimes(1)
    expect(host[action]).toHaveBeenCalledWith(...args)
  })

  it.each(EDITOR_ROUTES)('%s needs an editor and a document', (command, action, args) => {
    const noEditor = makeHost()
    runMenuCommand(noEditor, command)
    expect(noEditor[action]).not.toHaveBeenCalled()

    const editor = makeEditor()
    const noDoc = makeHost({ editor, hasDoc: false })
    runMenuCommand(noDoc, command)
    expect(noDoc[action]).not.toHaveBeenCalled()

    const host = makeHost({ editor })
    runMenuCommand(host, command)
    expect(host[action]).toHaveBeenCalledWith(...args)
    editor.destroy()
  })

  it.each(NEEDS_DOC)('%s does nothing without a document', (command) => {
    const host = makeHost({ hasDoc: false })
    runMenuCommand(host, command)
    for (const value of Object.values(host)) {
      if (typeof value === 'function') expect(value).not.toHaveBeenCalled()
    }
  })

  it('a payload command without a payload does nothing', () => {
    const host = makeHost()
    runMenuCommand(host, 'open-path')
    runMenuCommand(host, 'zoom-set')
    expect(host.openRecent).not.toHaveBeenCalled()
    expect(host.setZoom).not.toHaveBeenCalled()
  })

  it('bold, italic and underline toggle the mark on the editor', () => {
    const editor = makeEditor()
    editor.commands.selectAll()
    const host = makeHost({ editor })
    for (const mark of ['bold', 'italic', 'underline'] as const) {
      runMenuCommand(host, mark)
      expect(editor.isActive(mark)).toBe(true)
    }
    editor.destroy()
  })

  it('the grammar commands of the right-click menu replace, ignore and open the dialog at the finding', () => {
    const editor = makeEditor()
    const finding = {
      id: 'g-test',
      from: 1,
      to: 6,
      quote: 'plain',
      message: 'wrong',
      suggestions: ['simple'],
      key: 'k',
    }
    setGrammarFindings(editor, [finding])
    const host = makeHost({ editor })
    runMenuCommand(host, 'grammar-dialog', 'g-test')
    expect(host.checkSpelling).toHaveBeenLastCalledWith('g-test')
    runMenuCommand(host, 'grammar-replace', JSON.stringify({ id: 'g-test', text: 'simple' }))
    expect(editor.getText()).toBe('simple paragraph')
    expect(grammarState(editor.state)!.findings).toHaveLength(0)
    editor.commands.undo()
    expect(editor.getText()).toBe('plain paragraph')

    setGrammarFindings(editor, [finding])
    runMenuCommand(host, 'grammar-ignore', 'g-test')
    expect(grammarState(editor.state)!.findings).toHaveLength(0)
    expect(grammarState(editor.state)!.ignored.has('k')).toBe(true)
    expect(editor.getText()).toBe('plain paragraph')
    editor.destroy()
  })

  it('a synonym pick replaces the selected word and keeps its formatting', () => {
    const editor = makeEditor()
    editor.chain().setTextSelection({ from: 1, to: 6 }).setMark('bold').run()
    runMenuCommand(makeHost({ editor }), 'synonym-replace', 'simple')
    expect(editor.getText()).toBe('simple paragraph')
    expect(editor.state.doc.nodeAt(1)?.marks.map((m) => m.type.name)).toContain('bold')
    expect(editor.state.doc.nodeAt(8)?.marks.map((m) => m.type.name)).not.toContain('bold')
    // a bare caret stands for the word it is in
    editor.commands.setTextSelection(3)
    runMenuCommand(makeHost({ editor }), 'synonym-replace', 'plain')
    expect(editor.getText()).toBe('plain paragraph')
    expect(editor.state.doc.nodeAt(1)?.marks.map((m) => m.type.name)).toContain('bold')
    // Repeat puts the same word on the next selection
    editor.commands.setTextSelection({ from: 7, to: 16 })
    runMenuCommand(makeHost({ editor }), 'repeat')
    expect(editor.getText()).toBe('plain plain')
    editor.destroy()
  })

  it('the character effects Word for Mac has keys for toggle on the selection', () => {
    const editor = makeEditor()
    editor.commands.selectAll()
    const host = makeHost({ editor })
    const style = () => editor.getAttributes('docTextStyle')
    runMenuCommand(host, 'all-caps')
    expect(style().caps).toBe('all')
    runMenuCommand(host, 'small-caps')
    expect(style().caps).toBe('small')
    runMenuCommand(host, 'small-caps')
    expect(style().caps).toBe('none')
    runMenuCommand(host, 'hidden-text')
    expect(style()).toMatchObject({ vanish: true, vanishOwn: true })
    runMenuCommand(host, 'hidden-text')
    expect(style()).toMatchObject({ vanish: false, vanishOwn: false })
    runMenuCommand(host, 'double-underline')
    expect(editor.getAttributes('underline').line).toBe('double')
    runMenuCommand(host, 'word-underline')
    expect(editor.getAttributes('underline').line).toBe('words')
    runMenuCommand(host, 'word-underline')
    expect(editor.isActive('underline')).toBe(false)
    editor.destroy()
  })

  it('the bullet key applies the list and never takes it off', () => {
    const editor = makeEditor()
    const host = makeHost({ editor, listNumId: () => '7' })
    runMenuCommand(host, 'bullets', 'on')
    expect(editor.isActive('docListItem', { kind: 'bullet' })).toBe(true)
    runMenuCommand(host, 'bullets', 'on')
    expect(editor.isActive('docListItem', { kind: 'bullet' })).toBe(true)
    runMenuCommand(host, 'bullets')
    expect(editor.isActive('docListItem')).toBe(false)
    editor.destroy()
  })

  it('strikethrough, subscript and superscript toggle on the selection', () => {
    const editor = makeEditor()
    editor.commands.selectAll()
    const host = makeHost({ editor })
    runMenuCommand(host, 'strikethrough')
    expect(editor.isActive('strike')).toBe(true)
    runMenuCommand(host, 'subscript')
    expect(editor.getAttributes('docTextStyle').vertAlign).toBe('subscript')
    runMenuCommand(host, 'superscript')
    expect(editor.getAttributes('docTextStyle').vertAlign).toBe('superscript')
    runMenuCommand(host, 'superscript')
    expect(editor.getAttributes('docTextStyle').vertAlign ?? null).toBeNull()
    editor.destroy()
  })

  it('a grid pick inserts a table of that size and brings up Table Design', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-table-grid', 'wide')
    expect(host.showTableDesignTab).not.toHaveBeenCalled()
    runMenuCommand(host, 'insert-table-grid', '2,3')
    expect(host.showTableDesignTab).toHaveBeenCalledTimes(1)
    let rows = 0
    let cells = 0
    editor.state.doc.descendants((node) => {
      if (node.type.spec.tableRole === 'row') rows++
      if (node.type.spec.tableRole === 'cell') cells++
    })
    expect([rows, cells]).toEqual([2, 6])
    editor.destroy()
  })

  it('symbol, field and drop cap act on the body of an editable document only', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-symbol', '\u00a9')
    expect(editor.getText()).toContain('\u00a9')
    runMenuCommand(host, 'insert-field', 'DATE')
    expect(host.insertField).toHaveBeenCalledWith('DATE')
    runMenuCommand(host, 'drop-cap', 'drop')
    expect(JSON.parse(String(editor.getAttributes('docParagraph').dropCap))).toEqual({
      type: 'drop',
      lines: 3,
    })
    // the Drop Cap dialog sends the lines to drop along; out of range is pulled back in
    runMenuCommand(host, 'drop-cap', 'margin,5')
    expect(JSON.parse(String(editor.getAttributes('docParagraph').dropCap))).toEqual({
      type: 'margin',
      lines: 5,
    })
    runMenuCommand(host, 'drop-cap', 'drop,99')
    expect(JSON.parse(String(editor.getAttributes('docParagraph').dropCap)).lines).toBe(10)
    runMenuCommand(host, 'drop-cap', NONE)
    expect(editor.getAttributes('docParagraph').dropCap ?? null).toBeNull()
    editor.setEditable(false)
    runMenuCommand(host, 'insert-field', 'TIME')
    expect(host.insertField).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('Compare Documents and Combine Documents open their dialog, on a read-only document too', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'tools-combine-documents')
    runMenuCommand(host, 'compare-documents')
    expect(host.review.combine).toHaveBeenCalledTimes(1)
    expect(host.review.compare).toHaveBeenCalledTimes(1)
    editor.setEditable(false)
    runMenuCommand(host, 'tools-combine-documents')
    expect(host.review.combine).toHaveBeenCalledTimes(2)
    editor.destroy()
  })

  it('page setup picks change the section the caret is in', () => {
    const editor = makeEditor()
    const section = {
      pageWidth: 12240,
      pageHeight: 15840,
      orientation: 'portrait',
      marginTop: 1440,
      marginRight: 1440,
      marginBottom: 1440,
      marginLeft: 1440,
      columns: 1,
    } as NonNullable<CommandHost['section']>
    const host = makeHost({ editor, section })
    const last = () => vi.mocked(host.setSection).mock.lastCall?.[0]

    runMenuCommand(host, 'page-orientation', 'portrait')
    expect(host.setSection).not.toHaveBeenCalled()
    runMenuCommand(host, 'page-orientation', 'landscape')
    expect(last()).toMatchObject({ orientation: 'landscape', pageWidth: 15840, pageHeight: 12240 })

    runMenuCommand(host, 'page-size', '11906,16838')
    expect(last()).toMatchObject({ pageWidth: 11906, pageHeight: 16838 })
    runMenuCommand(host, 'page-columns', '2')
    expect(last()).toMatchObject({ columns: 2 })

    const narrow = { top: 720, right: 720, bottom: 720, left: 720, mirror: true }
    runMenuCommand(host, 'page-margins', JSON.stringify(narrow))
    expect(last()).toMatchObject({ marginTop: 720, marginLeft: 720 })
    expect(host.setMirrorMargins).toHaveBeenCalledWith(true)

    // the Document dialog's changes land in one write: none of them overwrites another
    const before = vi.mocked(host.setSection).mock.calls.length
    const fixed = { ...section, marginTopFixed: true } as typeof section
    const fixedHost = makeHost({ editor, section: fixed })
    runMenuCommand(
      fixedHost,
      'document-settings',
      JSON.stringify({
        margins: { top: 2880, right: 1440, bottom: 1440, left: 1440 },
        mirror: false,
        headerDist: 360,
        footerDist: 1080,
      }),
    )
    expect(fixedHost.setSection).toHaveBeenCalledTimes(1)
    expect(vi.mocked(fixedHost.setSection).mock.lastCall?.[0]).toMatchObject({
      marginTop: 2880,
      marginTopFixed: undefined,
      headerDist: 360,
      footerDist: 1080,
    })
    // a distance alone leaves the margins, and a fixed margin, as they were read
    runMenuCommand(fixedHost, 'document-settings', JSON.stringify({ footerDist: 200 }))
    expect(vi.mocked(fixedHost.setSection).mock.lastCall?.[0]).toMatchObject({
      marginTop: 1440,
      marginTopFixed: true,
      footerDist: 200,
    })
    runMenuCommand(fixedHost, 'document-settings', JSON.stringify({}))
    expect(fixedHost.setSection).toHaveBeenCalledTimes(2)
    // the gutter rides with the margins that hold it; alignment and columns stand alone
    runMenuCommand(
      fixedHost,
      'document-settings',
      JSON.stringify({
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
        gutter: 360,
        vAlign: 'center',
        columns: 2,
      }),
    )
    expect(vi.mocked(fixedHost.setSection).mock.lastCall?.[0]).toMatchObject({
      marginLeft: 1800,
      gutter: 360,
      vAlign: 'center',
      columns: 2,
    })
    runMenuCommand(fixedHost, 'document-settings', JSON.stringify({ vAlign: 'top' }))
    expect(vi.mocked(fixedHost.setSection).mock.lastCall?.[0].vAlign).toBeUndefined()
    runMenuCommand(fixedHost, 'document-settings', JSON.stringify({ columns: 0 }))
    expect(fixedHost.setSection).toHaveBeenCalledTimes(4)
    // the section start goes to the host on its own; an unknown one goes nowhere
    runMenuCommand(fixedHost, 'document-settings', JSON.stringify({ startType: 'continuous' }))
    expect(fixedHost.setSectionStart).toHaveBeenCalledWith('continuous')
    runMenuCommand(fixedHost, 'document-settings', JSON.stringify({ startType: 'sideways' }))
    expect(fixedHost.setSectionStart).toHaveBeenCalledTimes(1)
    expect(fixedHost.setSection).toHaveBeenCalledTimes(4)
    // the gutter position is the document's; the caret's section takes its gutter to that side
    runMenuCommand(
      fixedHost,
      'document-settings',
      JSON.stringify({
        margins: { top: 2160, right: 1440, bottom: 1440, left: 1440 },
        gutter: 720,
        gutterAtTop: true,
      }),
    )
    expect(fixedHost.setGutterAtTop).toHaveBeenLastCalledWith(true)
    expect(vi.mocked(fixedHost.setSection).mock.lastCall?.[0]).toMatchObject({
      marginTop: 2160,
      marginLeft: 1440,
      gutter: 720,
      gutterAtTop: true,
    })
    const bound = makeHost({
      editor,
      section: { ...section, marginLeft: 2160, gutter: 720 },
    })
    runMenuCommand(bound, 'document-settings', JSON.stringify({ gutterAtTop: true }))
    expect(vi.mocked(bound.setSection).mock.lastCall?.[0]).toMatchObject({
      marginTop: 2160,
      marginLeft: 1440,
      gutterAtTop: true,
    })
    runMenuCommand(
      bound,
      'document-settings',
      JSON.stringify({
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        gutter: 0,
        gutterAtTop: true,
      }),
    )
    expect(vi.mocked(bound.setSection).mock.lastCall?.[0].gutterAtTop).toBeUndefined()

    // the text flow rides in the same write; horizontal takes the tag away, a flow the page cannot lay out is not set
    const flowHost = makeHost({ editor, section })
    const flow = () => vi.mocked(flowHost.setSection).mock.lastCall?.[0]
    runMenuCommand(
      flowHost,
      'document-settings',
      JSON.stringify({ textDirection: 'tbRl', vAlign: 'center' }),
    )
    expect(flowHost.setSection).toHaveBeenCalledTimes(1)
    expect(flow()).toMatchObject({ textDirection: 'tbRl', vAlign: 'center' })
    runMenuCommand(flowHost, 'document-settings', JSON.stringify({ textDirection: 'lrTb' }))
    runMenuCommand(flowHost, 'document-settings', JSON.stringify({ textDirection: 'sideways' }))
    expect(flowHost.setSection).toHaveBeenCalledTimes(1)
    const rotated = makeHost({ editor, section: { ...section, textDirection: 'tbRlV' } })
    runMenuCommand(rotated, 'document-settings', JSON.stringify({ textDirection: 'tbRlV' }))
    expect(rotated.setSection).not.toHaveBeenCalled()
    runMenuCommand(rotated, 'document-settings', JSON.stringify({ footerDist: 200 }))
    expect(vi.mocked(rotated.setSection).mock.lastCall?.[0].textDirection).toBe('tbRlV')
    runMenuCommand(rotated, 'document-settings', JSON.stringify({ textDirection: 'lrTb' }))
    expect(vi.mocked(rotated.setSection).mock.lastCall?.[0].textDirection).toBeUndefined()

    // Layout > Text Direction in body text is the same write on the caret's section
    const body = makeHost({ editor, section })
    runMenuCommand(body, 'format-text-direction', 'lrTb')
    runMenuCommand(body, 'format-text-direction', 'btLr')
    expect(body.setSection).not.toHaveBeenCalled()
    runMenuCommand(body, 'format-text-direction', 'tbRl')
    expect(vi.mocked(body.setSection).mock.lastCall?.[0].textDirection).toBe('tbRl')
    expect(body.openDialog).not.toHaveBeenCalled()
    const turned = makeHost({ editor, section: { ...section, textDirection: 'tbRlV' } })
    runMenuCommand(turned, 'format-text-direction', 'tbRl')
    expect(turned.setSection).not.toHaveBeenCalled()
    runMenuCommand(turned, 'format-text-direction', 'lrTb')
    expect(vi.mocked(turned.setSection).mock.lastCall?.[0].textDirection).toBeUndefined()

    // suppress endnotes is set and cleared on the section, and left alone by other changes
    const notesHost = makeHost({ editor, section })
    const notes = () => vi.mocked(notesHost.setSection).mock.lastCall?.[0]
    runMenuCommand(notesHost, 'document-settings', JSON.stringify({ suppressEndnotes: false }))
    expect(notesHost.setSection).not.toHaveBeenCalled()
    runMenuCommand(notesHost, 'document-settings', JSON.stringify({ suppressEndnotes: true }))
    expect(notes()?.noEndnote).toBe(true)
    const suppressing = makeHost({ editor, section: { ...section, noEndnote: true } })
    runMenuCommand(suppressing, 'document-settings', JSON.stringify({ suppressEndnotes: true }))
    expect(suppressing.setSection).not.toHaveBeenCalled()
    runMenuCommand(suppressing, 'document-settings', JSON.stringify({ footerDist: 200 }))
    expect(vi.mocked(suppressing.setSection).mock.lastCall?.[0].noEndnote).toBe(true)
    runMenuCommand(suppressing, 'document-settings', JSON.stringify({ suppressEndnotes: false }))
    expect(vi.mocked(suppressing.setSection).mock.lastCall?.[0].noEndnote).toBeUndefined()

    // Apply to beyond the caret's section: each section takes the dialog's values, a
    // section the margins do not fit is left as it is, and the new section starts on a new page
    const wide = makeHost({ editor, section })
    runMenuCommand(
      wide,
      'document-settings',
      JSON.stringify({
        margins: { top: 2000, right: 1440, bottom: 1440, left: 1440 },
        vAlign: 'center',
        scope: 'forward',
      }),
    )
    expect(wide.setSection).not.toHaveBeenCalled()
    const [docScope, docChange, docBreak] = vi.mocked(wide.setSectionsIn).mock.lastCall!
    expect([docScope, docBreak]).toEqual(['forward', 'nextPage'])
    expect(docChange(section)).toMatchObject({ marginTop: 2000, vAlign: 'center' })
    const tiny = { ...section, pageHeight: 3000 }
    expect(docChange(tiny)).toBe(tiny)
    runMenuCommand(
      wide,
      'document-settings',
      JSON.stringify({ columns: 2, startType: 'oddPage', scope: 'forward' }),
    )
    expect(vi.mocked(wide.setSectionsIn).mock.lastCall![2]).toBe('oddPage')
    expect(wide.setSectionStart).not.toHaveBeenCalled()
    // another number of columns drops the widths made for the old number
    const toTwo = vi.mocked(wide.setSectionsIn).mock.lastCall![1]
    const three = { ...section, columns: 3, colWidths: [2000, 2000, 3000], colEqualWidth: false }
    const narrowed = toTwo(three)
    expect(narrowed.columns).toBe(2)
    expect(narrowed.colWidths).toBeUndefined()
    expect(narrowed.colEqualWidth).toBe(true)
    // a section that already has that many keeps its own widths
    const two = { ...section, columns: 2, colWidths: [2640, 6000], colEqualWidth: false }
    expect(toTwo(two).colWidths).toEqual([2640, 6000])
    // the document grid: a type the page draws is written with the pitch the section had
    const gridded = makeHost({
      editor,
      section: { ...section, docGrid: { type: 'default', linePitch: 312 } },
    })
    runMenuCommand(gridded, 'document-settings', JSON.stringify({ gridType: 'lines' }))
    expect(vi.mocked(gridded.setSection).mock.lastCall?.[0].docGrid).toEqual({
      type: 'lines',
      linePitch: 312,
    })
    runMenuCommand(gridded, 'document-settings', JSON.stringify({ gridType: 'bogus' }))
    expect(gridded.setSection).toHaveBeenCalledTimes(1)
    // a character pitch goes with a grid that has cells, within what the line can hold
    runMenuCommand(gridded, 'document-settings', JSON.stringify({ charSpace: 532 }))
    expect(gridded.setSection).toHaveBeenCalledTimes(1)
    runMenuCommand(
      gridded,
      'document-settings',
      JSON.stringify({ gridType: 'snapToChars', charSpace: 532 }),
    )
    expect(vi.mocked(gridded.setSection).mock.lastCall?.[0].docGrid).toEqual({
      type: 'snapToChars',
      linePitch: 312,
      charSpace: 532,
    })
    const celled = makeHost({
      editor,
      section: { ...section, docGrid: { type: 'linesAndChars', linePitch: 312, charSpace: 532 } },
    })
    runMenuCommand(celled, 'document-settings', JSON.stringify({ charSpace: -40960 }))
    runMenuCommand(celled, 'document-settings', JSON.stringify({ charSpace: 1.5 }))
    expect(celled.setSection).not.toHaveBeenCalled()
    runMenuCommand(celled, 'document-settings', JSON.stringify({ charSpace: 0 }))
    expect(vi.mocked(celled.setSection).mock.lastCall?.[0].docGrid).toEqual({
      type: 'linesAndChars',
      linePitch: 312,
      charSpace: 0,
    })
    // a line pitch on its own keeps the grid the section has; one below a point is not a pitch
    const lined = makeHost({
      editor,
      section: { ...section, docGrid: { type: 'lines', linePitch: 312, charSpace: -820 } },
    })
    runMenuCommand(lined, 'document-settings', JSON.stringify({ linePitch: 480 }))
    expect(vi.mocked(lined.setSection).mock.lastCall?.[0].docGrid).toEqual({
      type: 'lines',
      linePitch: 480,
      charSpace: -820,
    })
    runMenuCommand(lined, 'document-settings', JSON.stringify({ linePitch: 5 }))
    expect(lined.setSection).toHaveBeenCalledTimes(1)

    // Format > Columns: unequal widths, the line between, and back to equal columns
    const cols = makeHost({ editor, section })
    const lastCols = () => vi.mocked(cols.setSection).mock.lastCall?.[0]
    runMenuCommand(
      cols,
      'columns-settings',
      JSON.stringify({ columns: 2, colWidths: [2640, 6000], lineBetween: true }),
    )
    expect(lastCols()).toMatchObject({
      columns: 2,
      colWidths: [2640, 6000],
      colEqualWidth: false,
      colSep: true,
    })
    const unequal = makeHost({ editor, section: lastCols()! })
    runMenuCommand(unequal, 'columns-settings', JSON.stringify({ columns: 2, lineBetween: false }))
    const evened = vi.mocked(unequal.setSection).mock.lastCall?.[0]
    expect(evened).toMatchObject({ columns: 2, colEqualWidth: true })
    expect(evened?.colWidths).toBeUndefined()
    expect(evened?.colSep).toBeUndefined()
    // a width list of the wrong length is not a layout; one column has no line
    runMenuCommand(
      cols,
      'columns-settings',
      JSON.stringify({ columns: 1, colWidths: [2640, 6000], lineBetween: true }),
    )
    expect(lastCols()?.colWidths).toBeUndefined()
    expect(lastCols()?.colSep).toBeUndefined()
    expect(lastCols()?.colEqualWidth).toBeUndefined()
    const colCalls = vi.mocked(cols.setSection).mock.calls.length
    runMenuCommand(cols, 'columns-settings', JSON.stringify({ columns: 0 }))
    expect(cols.setSection).toHaveBeenCalledTimes(colCalls)
    // Apply to beyond the caret's section: the change is handed over per section, the
    // new columns start on the same page, and unequal widths follow each section's text width
    runMenuCommand(
      cols,
      'columns-settings',
      JSON.stringify({ columns: 2, colWidths: [2640, 6000], scope: 'forward' }),
    )
    expect(cols.setSection).toHaveBeenCalledTimes(colCalls)
    const [scope, change, breakType] = vi.mocked(cols.setSectionsIn).mock.lastCall!
    expect([scope, breakType]).toEqual(['forward', 'continuous'])
    expect(change(section)).toMatchObject({ columns: 2, colWidths: [2640, 6000] })
    const textWidth = section.pageWidth - section.marginLeft - section.marginRight - 720
    const narrower = change({ ...section, marginLeft: section.marginLeft + textWidth / 2 })
    expect(narrower.colWidths).toEqual([1320, 3000])
    // each column's own gap, the gap of equal columns, and the next column as the start
    runMenuCommand(
      cols,
      'columns-settings',
      JSON.stringify({
        columns: 3,
        colWidths: [2000, 3000, 2780],
        colSpaces: [360, 500],
        scope: 'forward',
        startNewColumn: true,
      }),
    )
    const [, gapped, columnBreak] = vi.mocked(cols.setSectionsIn).mock.lastCall!
    expect(columnBreak).toBe('nextColumn')
    expect(gapped(section)).toMatchObject({ colWidths: [2000, 3000, 2780], colSpaces: [360, 500] })
    runMenuCommand(
      cols,
      'columns-settings',
      JSON.stringify({ columns: 2, colSpace: 360, scope: 'whole', startNewColumn: true }),
    )
    const [, evenly, sameArea] = vi.mocked(cols.setSectionsIn).mock.lastCall!
    expect(sameArea).toBe('continuous')
    const even = evenly({ ...section, columns: 3, colWidths: [1, 2, 3], colSpaces: [4, 5] })
    expect(even).toMatchObject({ columns: 2, colSpace: 360 })
    expect(even.colSpaces).toBeUndefined()
    runMenuCommand(cols, 'columns-settings', JSON.stringify({ columns: 3, scope: 'whole' }))
    expect(vi.mocked(cols.setSectionsIn).mock.lastCall![0]).toBe('whole')
    expect(vi.mocked(cols.setSectionsIn).mock.lastCall![1](section)).toMatchObject({ columns: 3 })
    runMenuCommand(cols, 'columns-settings', JSON.stringify({ columns: 3, scope: 'section' }))
    expect(cols.setSection).toHaveBeenCalledTimes(colCalls + 1)
    // the ribbon's Two on a section with its own widths evens them out
    runMenuCommand(unequal, 'page-columns', '2')
    expect(vi.mocked(unequal.setSection).mock.lastCall?.[0]).toMatchObject({
      columns: 2,
      colEqualWidth: true,
    })
    expect(host.setSection).toHaveBeenCalledTimes(before)

    const calls = vi.mocked(host.setSection).mock.calls.length
    const tooWide = { top: 720, right: 9000, bottom: 720, left: 9000 }
    runMenuCommand(host, 'page-margins', JSON.stringify(tooWide))
    expect(host.setSection).toHaveBeenCalledTimes(calls)

    runMenuCommand(host, 'insert-section-break', 'continuous')
    expect(host.insertSectionBreak).toHaveBeenCalledWith('continuous')
    runMenuCommand(host, 'insert-section-break', 'sideways')
    expect(host.insertSectionBreak).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('comment commands act on the thread they name, or on the one at the caret', () => {
    const host = makeHost()
    const { review } = host
    runMenuCommand(host, 'delete-comment', 'c7')
    expect(review.deleteComment).toHaveBeenLastCalledWith('c7')
    runMenuCommand(host, 'delete-comment')
    expect(review.deleteComment).toHaveBeenLastCalledWith(undefined)
    runMenuCommand(host, 'comment-resolve', 'c7')
    expect(review.resolveComment).toHaveBeenLastCalledWith('c7')
    runMenuCommand(host, 'comment-resolve')
    expect(review.resolveComment).toHaveBeenLastCalledWith(undefined)
    runMenuCommand(host, 'comment-reply', 'c7')
    expect(review.replyToComment).toHaveBeenCalledWith('c7')
    // a reply needs a thread to answer
    runMenuCommand(host, 'comment-reply')
    expect(review.replyToComment).toHaveBeenCalledTimes(1)
  })

  it('Thumbnail reaches the view, Advanced Symbol opens its dialog, Next Footnote selects the next reference', () => {
    const note = (id: string) => ({ type: 'docNoteRef', attrs: { kind: 'footnote', id, num: 1 } })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'one' }, note('1')] },
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'two' },
              { type: 'docNoteRef', attrs: { kind: 'endnote', id: '9', num: 1 } },
              note('2'),
            ],
          },
        ],
      },
    })
    const host = makeHost({ editor })
    runMenuCommand(host, 'view-thumbnails')
    expect(host.view.toggleThumbnails).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'insert-advanced-symbol')
    expect(host.openDialog).toHaveBeenCalledWith('symbol')
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'insert-advanced-symbol')
    expect(closed.openDialog).not.toHaveBeenCalled()

    const selectedNote = () => {
      const { selection } = editor.state
      return selection instanceof NodeSelection ? String(selection.node.attrs.id) : null
    }
    editor.commands.setTextSelection(1)
    runMenuCommand(host, 'footnote-next')
    expect(selectedNote()).toBe('1')
    // the endnote between the two is passed over
    runMenuCommand(host, 'footnote-next')
    expect(selectedNote()).toBe('2')
    // past the last one the selection stays where it is
    runMenuCommand(host, 'footnote-next')
    expect(selectedNote()).toBe('2')
    editor.destroy()
  })

  it('Start Mail Merge, Envelopes and Labels open the merge panel on their kind of output', () => {
    const host = makeHost()
    runMenuCommand(host, 'mailings-start')
    expect(host.openMailMerge).toHaveBeenLastCalledWith('perRecord')
    runMenuCommand(host, 'tools-envelopes')
    expect(host.openMailMerge).toHaveBeenLastCalledWith('envelopes')
    runMenuCommand(host, 'tools-labels')
    expect(host.openMailMerge).toHaveBeenLastCalledWith('labels')
    runMenuCommand(host, 'mailings-finish-merge')
    expect(host.openMailMerge).toHaveBeenLastCalledWith('single')
    // a file that says it is a labels main document merges as labels
    const labels = makeHost({ labelMainDocument: true })
    runMenuCommand(labels, 'mailings-start')
    expect(labels.openMailMerge).toHaveBeenLastCalledWith('labels')
    runMenuCommand(labels, 'mailings-finish-merge')
    expect(labels.openMailMerge).toHaveBeenLastCalledWith('labels')
    const closed = makeHost({ hasDoc: false })
    runMenuCommand(closed, 'tools-labels')
    expect(closed.openMailMerge).not.toHaveBeenCalled()
  })

  it('Markup hides the markup that is showing and brings all of it back; Page Numbers opens its dialog', () => {
    const shown = makeHost()
    runMenuCommand(shown, 'view-markup')
    expect(shown.review.setRevisionDisplay).toHaveBeenLastCalledWith('none')
    for (const mode of ['none', 'original'] as const) {
      const hidden = makeHost()
      hidden.review.revisionDisplay = mode
      runMenuCommand(hidden, 'view-markup')
      expect(hidden.review.setRevisionDisplay).toHaveBeenLastCalledWith('all')
    }
    const simple = makeHost()
    simple.review.revisionDisplay = 'simple'
    runMenuCommand(simple, 'view-markup')
    expect(simple.review.setRevisionDisplay).toHaveBeenLastCalledWith('none')

    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-page-numbers')
    expect(host.openDialog).toHaveBeenCalledWith('pageNumbers')
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'insert-page-numbers')
    expect(closed.openDialog).not.toHaveBeenCalled()
    runMenuCommand(host, 'format-tabs')
    expect(host.openTabsDialog).toHaveBeenCalledWith(editor)
    runMenuCommand(host, 'format-phonetic-guide')
    expect(host.openPhoneticGuide).toHaveBeenCalledWith(editor)
    runMenuCommand(host, 'format-fit-text')
    expect(host.openFitText).toHaveBeenCalledWith(editor)
    runMenuCommand(host, 'format-two-lines-in-one')
    expect(host.openAsianLayout).toHaveBeenLastCalledWith(editor, 'combine')
    runMenuCommand(host, 'format-horizontal-in-vertical')
    expect(host.openAsianLayout).toHaveBeenLastCalledWith(editor, 'vert')
    runMenuCommand(closed, 'format-fit-text')
    runMenuCommand(closed, 'format-two-lines-in-one')
    expect(closed.openFitText).not.toHaveBeenCalled()
    expect(closed.openAsianLayout).not.toHaveBeenCalled()
    runMenuCommand(closed, 'format-phonetic-guide')
    expect(closed.openPhoneticGuide).not.toHaveBeenCalled()
    runMenuCommand(host, 'format-enclose-characters')
    expect(host.openEqLayout).toHaveBeenLastCalledWith(editor, 'enclose')
    runMenuCommand(host, 'format-combine-characters')
    expect(host.openEqLayout).toHaveBeenLastCalledWith(editor, 'combine')
    runMenuCommand(closed, 'format-enclose-characters')
    runMenuCommand(closed, 'format-combine-characters')
    expect(closed.openEqLayout).not.toHaveBeenCalled()
    runMenuCommand(host, 'format-bullets-numbering')
    expect(host.openDialog).toHaveBeenLastCalledWith('bulletsNumbering')
    runMenuCommand(closed, 'format-bullets-numbering')
    expect(closed.openDialog).not.toHaveBeenCalled()
    runMenuCommand(closed, 'format-tabs')
    expect(closed.openTabsDialog).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('Distributed, Line Numbers, Clear Contents and Advanced Find act on the document', () => {
    const editor = makeEditor()
    const section = { columns: 1 } as NonNullable<CommandHost['section']>
    const host = makeHost({ editor, section })
    runMenuCommand(host, 'align-distributed')
    expect(editor.getAttributes('docParagraph').align).toBe('distribute')

    runMenuCommand(host, 'layout-line-numbers', 'continuous')
    expect(host.setSection).toHaveBeenLastCalledWith({
      columns: 1,
      lineNumbers: { countBy: 1, start: 1, restart: 'continuous' },
    })
    // what the section already counts by is kept when only the restart changes
    const counting = makeHost({
      editor,
      section: { ...section, lineNumbers: { countBy: 5, start: 1, restart: 'newPage' } },
    })
    runMenuCommand(counting, 'layout-line-numbers', 'newSection')
    expect(counting.setSection).toHaveBeenLastCalledWith({
      columns: 1,
      lineNumbers: { countBy: 5, start: 1, restart: 'newSection' },
    })
    runMenuCommand(counting, 'layout-line-numbers', 'none')
    expect(counting.setSection).toHaveBeenLastCalledWith({ columns: 1 })
    runMenuCommand(counting, 'layout-line-numbers', 'sideways')
    expect(counting.setSection).toHaveBeenCalledTimes(2)
    // the Line Numbers dialog writes the whole numbering; one it cannot use is dropped
    const typed = { countBy: 2, start: 5, distance: 360, restart: 'continuous' }
    runMenuCommand(counting, 'layout-line-numbers', JSON.stringify(typed))
    expect(counting.setSection).toHaveBeenLastCalledWith({ columns: 1, lineNumbers: typed })
    runMenuCommand(counting, 'layout-line-numbers', JSON.stringify({ ...typed, countBy: 0 }))
    runMenuCommand(counting, 'layout-line-numbers', JSON.stringify({ ...typed, restart: 'never' }))
    expect(counting.setSection).toHaveBeenCalledTimes(3)
    runMenuCommand(counting, 'layout-line-numbering-options')
    expect(counting.openDialog).toHaveBeenLastCalledWith('lineNumbers')

    editor.commands.setTextSelection({ from: 1, to: 6 })
    runMenuCommand(host, 'edit-clear-contents')
    expect(editor.getText()).toBe(' paragraph')

    runMenuCommand(host, 'edit-advanced-find')
    expect(host.openFind).toHaveBeenLastCalledWith('find')
    editor.destroy()
  })

  it('the Tools, Mailings and References placeholders that now run reach the app', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'tools-thesaurus')
    expect(host.thesaurus).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'tools-translate-selection')
    expect(host.translate).toHaveBeenLastCalledWith('selection')
    runMenuCommand(host, 'tools-translate-document')
    expect(host.translate).toHaveBeenLastCalledWith('document')
    runMenuCommand(host, 'tools-translation-language')
    expect(host.openDialog).toHaveBeenLastCalledWith('translationLanguage')
    runMenuCommand(host, 'mailings-select-recipients')
    expect(host.pickMergeDataSource).toHaveBeenCalledTimes(1)
    // Preview Results asks for a list when there is none, and toggles once there is one
    runMenuCommand(host, 'mailings-preview-results')
    expect(host.pickMergeDataSource).toHaveBeenCalledTimes(2)
    expect(host.toggleMergePreview).not.toHaveBeenCalled()
    const listed = makeHost({ editor, hasMergeSource: true })
    runMenuCommand(listed, 'mailings-preview-results')
    expect(listed.toggleMergePreview).toHaveBeenCalledTimes(1)
    expect(listed.pickMergeDataSource).not.toHaveBeenCalled()
    runMenuCommand(host, 'mailings-finish-merge')
    expect(host.openMailMerge).toHaveBeenLastCalledWith('single')

    // Add Text: a heading level puts the paragraph in the table of contents, 0 takes it out
    runMenuCommand(host, 'toc-add-text', '2')
    expect(host.applyHeadingLevel).toHaveBeenLastCalledWith(2)
    runMenuCommand(host, 'toc-add-text', '0')
    expect(host.applyHeadingLevel).toHaveBeenLastCalledWith(0)
    runMenuCommand(host, 'toc-add-text', '7')
    expect(host.applyHeadingLevel).toHaveBeenCalledTimes(2)

    // none of them runs without a document, except Customize Keyboard
    const closed = makeHost({ editor, hasDoc: false })
    for (const id of [
      'tools-thesaurus',
      'tools-translate-selection',
      'tools-translate-document',
      'tools-translation-language',
      'mailings-select-recipients',
      'mailings-finish-merge',
    ] as const)
      runMenuCommand(closed, id)
    expect(closed.thesaurus).not.toHaveBeenCalled()
    expect(closed.translate).not.toHaveBeenCalled()
    expect(closed.pickMergeDataSource).not.toHaveBeenCalled()
    expect(closed.openMailMerge).not.toHaveBeenCalled()
    expect(closed.openDialog).not.toHaveBeenCalled()
    runMenuCommand(closed, 'tools-customize-keyboard')
    expect(closed.openDialog).toHaveBeenLastCalledWith('customizeKeyboard')
    editor.destroy()
  })

  it('the View, Tools, Mailings and File placeholders of the eleventh batch reach the app', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'view-ribbon')
    expect(host.view.toggleRibbon).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'tools-reset-ignored')
    expect(host.review.resetIgnoredWords).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'tools-hide-grammar')
    expect(host.review.toggleGrammarMarks).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'file-properties')
    expect(host.openDialog).toHaveBeenLastCalledWith('properties')
    runMenuCommand(host, 'file-always-read-only')
    expect(host.toggleReadOnlyRecommended).toHaveBeenCalledTimes(1)

    // the recipient dialogs need a list: without one the list is asked for
    for (const id of [
      'mailings-edit-recipients',
      'mailings-filter-recipients',
      'mailings-find-recipient',
    ] as const)
      runMenuCommand(host, id)
    expect(host.pickMergeDataSource).toHaveBeenCalledTimes(3)
    const listed = makeHost({ editor, hasMergeSource: true })
    runMenuCommand(listed, 'mailings-edit-recipients')
    expect(listed.openDialog).toHaveBeenLastCalledWith('recipients')
    runMenuCommand(listed, 'mailings-filter-recipients')
    expect(listed.openDialog).toHaveBeenLastCalledWith('filterRecipients')
    runMenuCommand(listed, 'mailings-find-recipient')
    expect(listed.openDialog).toHaveBeenLastCalledWith('findRecipient')
    expect(listed.pickMergeDataSource).not.toHaveBeenCalled()

    // the ribbon folds with no document open; the rest wait for one
    const closed = makeHost({ editor, hasDoc: false, hasMergeSource: true })
    for (const id of [
      'tools-reset-ignored',
      'file-properties',
      'file-always-read-only',
      'mailings-edit-recipients',
      'view-footnotes',
      'update-index',
    ] as const)
      runMenuCommand(closed, id)
    expect(closed.review.resetIgnoredWords).not.toHaveBeenCalled()
    expect(closed.openDialog).not.toHaveBeenCalled()
    expect(closed.toggleReadOnlyRecommended).not.toHaveBeenCalled()
    expect(closed.noNotes).not.toHaveBeenCalled()
    runMenuCommand(closed, 'view-ribbon')
    expect(closed.view.toggleRibbon).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('Rename, Move, Number, Sort, Text Direction and Focus reach the app', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'file-rename')
    expect(host.openDialog).toHaveBeenLastCalledWith('rename')
    runMenuCommand(host, 'file-move')
    expect(host.moveFile).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'insert-number')
    expect(host.openDialog).toHaveBeenLastCalledWith('number')
    runMenuCommand(host, 'sort')
    expect(host.openDialog).toHaveBeenLastCalledWith('sort')
    runMenuCommand(host, 'format-text-direction')
    expect(host.openDialog).toHaveBeenLastCalledWith('textDirection')
    runMenuCommand(host, 'view-focus')
    expect(host.view.toggleFocus).toHaveBeenCalledTimes(1)

    // a document never saved has no file to rename or move
    const unsaved = makeHost({ editor, onDisk: false })
    runMenuCommand(unsaved, 'file-rename')
    runMenuCommand(unsaved, 'file-move')
    expect(unsaved.openDialog).not.toHaveBeenCalled()
    expect(unsaved.moveFile).not.toHaveBeenCalled()

    // a read-only document takes no number, sort or text direction
    editor.setEditable(false)
    const locked = makeHost({ editor })
    for (const id of ['insert-number', 'sort', 'format-text-direction'] as const)
      runMenuCommand(locked, id)
    expect(locked.openDialog).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('the chart types, Borders and Shading, Font Substitution and Rules reach the app', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-chart-bar')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('barH')
    runMenuCommand(host, 'insert-chart-area')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('area')
    runMenuCommand(host, 'insert-chart-radar')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('radar')
    runMenuCommand(host, 'insert-chart-scatter')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('scatter')
    runMenuCommand(host, 'insert-chart', 'radar')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('radar')
    runMenuCommand(host, 'insert-chart', 'nonsense')
    expect(host.openChartDialog).toHaveBeenLastCalledWith(undefined)
    runMenuCommand(host, 'format-borders-shading')
    expect(host.openDialog).toHaveBeenLastCalledWith('bordersShading')
    runMenuCommand(host, 'format-font-substitution')
    expect(host.openDialog).toHaveBeenLastCalledWith('fontSubstitution')

    // Next Record goes straight in; the two conditional rules need a list for its fields
    runMenuCommand(host, 'mailings-rules', 'next')
    expect(host.insertField).toHaveBeenLastCalledWith('NEXT')
    runMenuCommand(host, 'mailings-rules', 'skipIf')
    expect(host.pickMergeDataSource).toHaveBeenCalledTimes(1)
    expect(host.openRuleDialog).not.toHaveBeenCalled()
    const listed = makeHost({ editor, hasMergeSource: true })
    runMenuCommand(listed, 'mailings-rules', 'if')
    expect(listed.openRuleDialog).toHaveBeenLastCalledWith('if')
    runMenuCommand(listed, 'mailings-rules', 'other')
    expect(listed.openRuleDialog).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('Borders and Shading writes the sides and takes them off', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const borders = () => editor.state.doc.firstChild?.attrs.borders ?? null
    runMenuCommand(host, 'paragraph-borders', 'tb')
    expect(borders()).toBe('tb')
    runMenuCommand(host, 'paragraph-borders', 'none')
    expect(borders()).toBeNull()
    // the dialog's payload: each side with its line
    const lines = { t: { color: 'C00000', szPt: 1.5, val: 'double' }, b: { szPt: 0.75 } }
    runMenuCommand(host, 'paragraph-borders', JSON.stringify(lines))
    expect(borders()).toBe('tb')
    expect(editor.state.doc.firstChild?.attrs.borderLines).toBe(JSON.stringify(lines))
    // the ribbon's sides keep the look the lines have
    runMenuCommand(host, 'paragraph-borders', 't')
    expect(editor.state.doc.firstChild?.attrs.borderLines).toBe(JSON.stringify(lines))
    editor.destroy()
  })

  it('the right-click spelling items replace, ignore, learn and pair the word they were opened on', () => {
    const editor = makeEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'a mispeled word' }] }],
    })
    const host = makeHost({ editor })
    const pick = JSON.stringify({ pos: 3, word: 'mispeled', to: 'misspelled', marked: false })
    runMenuCommand(host, 'spell-replace', pick)
    expect(host.review.replaceMisspelling).toHaveBeenLastCalledWith(
      3,
      'mispeled',
      'misspelled',
      false,
    )

    runMenuCommand(host, 'spell-ignore-all', 'mispeled')
    expect(host.review.ignoreSpelling).toHaveBeenLastCalledWith('mispeled')
    runMenuCommand(host, 'spell-add-word', 'mispeled')
    expect(host.review.learnSpelling).toHaveBeenLastCalledWith('mispeled')

    // Ignore Once: this occurrence only, and the text is looked at again
    runMenuCommand(host, 'spell-ignore-once', JSON.stringify({ pos: 3, word: 'mispeled' }))
    expect(ignoredOnce(editor.state, 3, 11)).toBe(true)
    expect(host.review.respell).toHaveBeenCalledTimes(1)
    // a word that is no longer there is left alone
    runMenuCommand(host, 'spell-ignore-once', JSON.stringify({ pos: 13, word: 'mispeled' }))
    expect(host.review.respell).toHaveBeenCalledTimes(1)

    // AutoCorrect: the pair is in the replace table, and the word is corrected
    localStorage.clear()
    runMenuCommand(host, 'spell-autocorrect', pick)
    expect(autocorrectLists().replace).toContainEqual(['mispeled', 'misspelled'])
    expect(host.review.replaceMisspelling).toHaveBeenCalledTimes(2)
    localStorage.clear()

    // Spelling… starts the check at the word
    runMenuCommand(host, 'tools-spelling', '3')
    expect(editor.state.selection.from).toBe(3)
    expect(host.checkSpelling).toHaveBeenCalledTimes(1)

    // nothing that changes the text runs on a read-only document, and a broken payload is dropped
    editor.setEditable(false)
    runMenuCommand(host, 'spell-replace', pick)
    runMenuCommand(host, 'spell-autocorrect', pick)
    editor.setEditable(true)
    runMenuCommand(host, 'spell-replace', 'not json')
    expect(host.review.replaceMisspelling).toHaveBeenCalledTimes(2)
    expect(autocorrectLists().replace).not.toContainEqual(['mispeled', 'misspelled'])
    editor.destroy()
  })

  it('the right-click comment and note items reach the thread or note they name', () => {
    const editor = makeEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'one' },
            { type: 'docNoteRef', attrs: { kind: 'footnote', id: '2', num: 1 } },
            { type: 'text', text: ' two' },
            { type: 'docNoteRef', attrs: { kind: 'endnote', id: '2', num: 1 } },
          ],
        },
      ],
    })
    const host = makeHost({ editor })
    runMenuCommand(host, 'comment-edit', 'c7')
    expect(host.review.editComment).toHaveBeenLastCalledWith('c7')
    runMenuCommand(host, 'comment-reopen', 'c7')
    expect(host.review.reopenComment).toHaveBeenLastCalledWith('c7')
    runMenuCommand(host, 'comment-reopen')
    expect(host.review.reopenComment).toHaveBeenCalledTimes(1)

    // from the note's text to its mark in the body: the mark is selected
    runMenuCommand(host, 'note-goto', 'endnote:ref:2')
    const { selection } = editor.state
    expect(selection instanceof NodeSelection && selection.node.attrs.kind).toBe('endnote')
    expect(selection.from).toBe(9)
    // from the mark to the note's text
    runMenuCommand(host, 'note-goto', 'footnote:text:2')
    expect(host.editNote).toHaveBeenLastCalledWith('footnote', '2')

    runMenuCommand(host, 'note-convert', 'footnote:2')
    expect(host.convertNote).toHaveBeenLastCalledWith('footnote', '2')
    runMenuCommand(host, 'note-convert', 'sidenote:2')
    editor.setEditable(false)
    runMenuCommand(host, 'note-convert', 'endnote:2')
    expect(host.convertNote).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('View > Footnotes opens the note after the caret, else the first, and says when there is none', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'view-footnotes')
    expect(host.noNotes).toHaveBeenCalledTimes(1)
    expect(host.editNote).not.toHaveBeenCalled()

    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'one' },
            { type: 'docNoteRef', attrs: { kind: 'footnote', id: '2', num: 1 } },
            { type: 'text', text: ' two' },
            { type: 'docNoteRef', attrs: { kind: 'endnote', id: '5', num: 1 } },
            { type: 'text', text: ' end' },
          ],
        },
      ],
    })
    editor.commands.setTextSelection(1)
    runMenuCommand(host, 'view-footnotes')
    expect(host.editNote).toHaveBeenLastCalledWith('footnote', '2')
    // between the two references
    editor.commands.setTextSelection(7)
    runMenuCommand(host, 'view-footnotes')
    expect(host.editNote).toHaveBeenLastCalledWith('endnote', '5')
    // past the last one: back to the first
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    runMenuCommand(host, 'view-footnotes')
    expect(host.editNote).toHaveBeenLastCalledWith('footnote', '2')
    editor.destroy()
  })

  it('Update Index rebuilds the index in place from the entries marked now', () => {
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => {})
    const editor = makeEditor()
    const host = makeHost({ editor })
    const mark = (term: string) => ({ type: 'docXeMark', attrs: { term } })
    // no index yet: nothing to update
    runMenuCommand(host, 'update-index')
    expect(alert).toHaveBeenCalledTimes(1)

    editor.commands.setContent({
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'pear' }, mark('pear')] },
        { type: 'docParagraph', content: [{ type: 'text', text: 'tail' }] },
      ],
    })
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    runMenuCommand(host, 'insert-index')
    const indexLines = () => {
      const lines: string[] = []
      editor.state.doc.forEach((node) => {
        if (node.type.name === 'docProtected') lines.push(node.attrs.fieldDisplay.left)
      })
      return lines
    }
    expect(indexLines()).toEqual(['pear'])

    // a new entry is marked: the index picks it up where it stands
    editor.commands.insertContentAt(1, [{ type: 'text', text: 'apple' }, mark('apple')])
    runMenuCommand(host, 'update-index')
    expect(indexLines()).toEqual(['apple', 'pear'])
    expect(editor.state.doc.firstChild?.textContent).toContain('apple')
    expect(alert).toHaveBeenCalledTimes(1)
    alert.mockRestore()
    editor.destroy()
  })

  it('an index gets a section for its columns, from the ribbon as from the dialog; one column and an unchanged count do not', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const mark = (term: string) => ({ type: 'docXeMark', attrs: { term } })
    const reset = () => {
      editor.commands.setContent({
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'pear' }, mark('pear')] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'tail' }] },
        ],
      })
      editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    }
    const spec = (columns: string) => JSON.stringify({ instr: ` INDEX ${columns}`, leader: 'dot' })
    const section = { columns: 1, colEqualWidth: false } as never

    reset()
    runMenuCommand(host, 'insert-index', spec(''))
    expect(host.setIndexSection).not.toHaveBeenCalled()

    // the ribbon's button is the dialog's OK on its defaults: two columns in a section of their own
    reset()
    runMenuCommand(host, 'insert-index')
    expect(host.setIndexSection).toHaveBeenCalledTimes(1)
    let [nth, change, wrap] = vi.mocked(host.setIndexSection).mock.calls[0]!
    expect([nth, wrap]).toEqual([0, true])
    expect(change(section)).toMatchObject({ columns: 2, colEqualWidth: true })
    const instrs: string[] = []
    editor.state.doc.forEach((node) => {
      const m = /<w:instrText[^>]*>([^<]*INDEX[^<]*)</.exec(String(node.attrs.genXml ?? ''))
      if (m) instrs.push(m[1])
    })
    expect(instrs).toEqual([' INDEX \\c "2" '])
    vi.mocked(host.setIndexSection).mockClear()

    reset()
    runMenuCommand(host, 'insert-index', spec('\\c "3" '))
    expect(host.setIndexSection).toHaveBeenCalledTimes(1)
    ;[nth, change, wrap] = vi.mocked(host.setIndexSection).mock.calls[0]!
    expect([nth, wrap]).toEqual([0, true])
    expect(change(section)).toMatchObject({ columns: 3, colEqualWidth: true })

    // the selection is in that index now: the same count again leaves its section alone
    let at = 0
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'docProtected') at = offset
    })
    editor.commands.setNodeSelection(at)
    runMenuCommand(host, 'insert-index', spec('\\c "3" '))
    expect(host.setIndexSection).toHaveBeenCalledTimes(1)
    // one column: the section it has takes it, and none is made for it
    runMenuCommand(host, 'insert-index', spec(''))
    expect(host.setIndexSection).toHaveBeenCalledTimes(2)
    ;[nth, change, wrap] = vi.mocked(host.setIndexSection).mock.calls[1]!
    expect([nth, wrap]).toEqual([0, false])
    expect(change(section)).toMatchObject({ columns: 1 })

    // a second index after the one already there: it is the one that gets the section
    editor.commands.insertContentAt(editor.state.doc.content.size, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'after' }],
    })
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    runMenuCommand(host, 'insert-index', spec('\\c "2" '))
    ;[nth, , wrap] = vi.mocked(host.setIndexSection).mock.calls[2]!
    expect([nth, wrap]).toEqual([1, true])

    // the host moved the caret out of the index, as inserting section breaks does: the same index is rebuilt
    const indexCount = () => {
      let count = 0
      editor.state.doc.forEach((node) => {
        if (node.type.name === 'docProtected') count++
      })
      return count
    }
    const had = indexCount()
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'docProtected') at = offset
    })
    editor.commands.setNodeSelection(at)
    vi.mocked(host.setIndexSection).mockImplementationOnce(() => {
      editor.commands.setTextSelection(1)
    })
    runMenuCommand(host, 'insert-index', spec('\\c "4" '))
    expect(host.setIndexSection).toHaveBeenCalledTimes(4)
    expect(vi.mocked(host.setIndexSection).mock.calls[3]![0]).toBe(1)
    expect(indexCount()).toBe(had)

    // the entries are gone: the rebuild fails, so the section is left as it is
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => {})
    const marks: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docXeMark') marks.push(pos)
    })
    const tr = editor.state.tr
    for (const pos of marks.reverse()) tr.delete(pos, pos + 1)
    editor.view.dispatch(tr)
    editor.state.doc.forEach((node, offset) => {
      if (node.type.name === 'docProtected') at = offset
    })
    editor.commands.setNodeSelection(at)
    runMenuCommand(host, 'insert-index', spec('\\c "3" '))
    expect(alert).toHaveBeenCalledTimes(1)
    expect(host.setIndexSection).toHaveBeenCalledTimes(4)
    alert.mockRestore()
    editor.destroy()
  })

  it('Character Shading toggles the grey; Index and Tables and Hyphenation open their dialogs', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    editor.commands.setTextSelection({ from: 1, to: 6 })
    runMenuCommand(host, 'character-shading')
    expect(editor.getAttributes('docTextStyle').shading).toBe('D9D9D9')
    runMenuCommand(host, 'character-shading')
    expect(editor.getAttributes('docTextStyle').shading ?? null).toBeNull()

    // Character Border: the box and the run's own border go on together, and come off as an explicit none
    runMenuCommand(host, 'character-border')
    expect(editor.getAttributes('docTextStyle')).toMatchObject({
      bdr: '{"val":"single","sz":4}',
      bdrOwn: '{"val":"single","sz":4}',
    })
    runMenuCommand(host, 'character-border')
    expect(editor.getAttributes('docTextStyle')).toMatchObject({ bdr: null, bdrOwn: 'none' })

    // Borders and Shading on the text: the line and the fill it picked, then off
    const dashed = JSON.stringify({ val: 'dashed', sz: 12, color: 'C00000' })
    runMenuCommand(host, 'character-border', dashed)
    expect(editor.getAttributes('docTextStyle')).toMatchObject({ bdr: dashed, bdrOwn: dashed })
    runMenuCommand(host, 'character-border', 'not a border')
    expect(editor.getAttributes('docTextStyle').bdr).toBe(dashed)
    runMenuCommand(host, 'character-border', NONE)
    expect(editor.getAttributes('docTextStyle')).toMatchObject({ bdr: null, bdrOwn: 'none' })
    runMenuCommand(host, 'character-shading', 'FFF2CC')
    expect(editor.getAttributes('docTextStyle').shading).toBe('FFF2CC')
    runMenuCommand(host, 'character-shading', NONE)
    expect(editor.getAttributes('docTextStyle').shading ?? null).toBeNull()

    runMenuCommand(host, 'insert-index-tables')
    expect(host.openDialog).toHaveBeenLastCalledWith('indexTables')
    runMenuCommand(host, 'tools-hyphenation')
    expect(host.openDialog).toHaveBeenLastCalledWith('hyphenation')
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'insert-index-tables')
    runMenuCommand(closed, 'tools-hyphenation')
    expect(closed.openDialog).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('Paragraph Spacing hands the picked built-in to the design host', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'design-paragraph-spacing', 'relaxed')
    expect(host.design.setParagraphSpacing).toHaveBeenLastCalledWith({
      id: 'relaxed',
      afterTwips: 120,
      line: 1.5,
    })
    runMenuCommand(host, 'design-paragraph-spacing', 'airy')
    expect(host.design.setParagraphSpacing).toHaveBeenCalledTimes(1)
    // a spacing typed by hand arrives whole; a negative one is dropped
    const custom = { beforeTwips: 120, afterTwips: 360, line: 1.5 }
    runMenuCommand(host, 'design-paragraph-spacing', JSON.stringify(custom))
    expect(host.design.setParagraphSpacing).toHaveBeenLastCalledWith(custom)
    runMenuCommand(host, 'design-paragraph-spacing', JSON.stringify({ ...custom, afterTwips: -1 }))
    expect(host.design.setParagraphSpacing).toHaveBeenCalledTimes(2)
    runMenuCommand(host, 'design-custom-paragraph-spacing')
    expect(host.openDialog).toHaveBeenLastCalledWith('paragraphSpacing')
    // the style is the document's: the caret may be in a header, footer or text box
    const strip = makeEditor()
    setActiveSubEditor(strip)
    try {
      runMenuCommand(host, 'design-paragraph-spacing', 'double')
      expect(host.design.setParagraphSpacing).toHaveBeenCalledTimes(3)
    } finally {
      setActiveSubEditor(null)
      strip.destroy()
    }
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'design-paragraph-spacing', 'double')
    expect(closed.design.setParagraphSpacing).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('a style set pick hands a known set to the design host', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'design-style-set', 'elegant')
    expect(host.design.applyStyleSet).toHaveBeenLastCalledWith('elegant')
    runMenuCommand(host, 'design-style-set', 'ornate')
    expect(host.design.applyStyleSet).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'design-style-set', 'custom:set3')
    expect(host.design.applyStyleSet).toHaveBeenLastCalledWith({ custom: 'set3' })
    runMenuCommand(host, 'design-style-set-reset')
    expect(host.design.resetStyleSet).toHaveBeenCalledTimes(1)
    // Save asks for a name first; the dialog runs the command again with it
    runMenuCommand(host, 'design-style-set-save')
    expect(host.openDialog).toHaveBeenLastCalledWith('saveStyleSet')
    expect(host.design.saveStyleSet).not.toHaveBeenCalled()
    runMenuCommand(host, 'design-style-set-save', ' Mine ')
    expect(host.design.saveStyleSet).toHaveBeenLastCalledWith('Mine')
    runMenuCommand(host, 'design-style-set-delete', 'set3')
    expect(host.design.deleteStyleSet).toHaveBeenLastCalledWith('set3')
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'design-style-set', 'elegant')
    expect(closed.design.applyStyleSet).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('Mark Entry opens its dialog on an open, editable document', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'mark-index-entry')
    expect(host.openMarkIndexEntry).toHaveBeenCalledTimes(1)
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'mark-index-entry')
    expect(closed.openMarkIndexEntry).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('Mark Citation opens its dialog; the table of authorities lists a category and updates', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'toa-mark-citation')
    expect(host.openMarkCitation).toHaveBeenCalledTimes(1)
    editor.commands.insertContentAt(6, {
      type: 'docXeMark',
      attrs: { term: '', ta: '\\l "Baldwin v. Alberti" \\c 2' },
    } as never)
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    runMenuCommand(host, 'toa-insert', '2')
    const lines = () => {
      const out: string[] = []
      editor.state.doc.forEach((node) => {
        if (/TOAHeading|TableofAuthorities/.test(String(node.attrs.genXml ?? '')))
          out.push((node.attrs.fieldDisplay as { left: string }).left)
      })
      return out
    }
    expect(host.citationHeading).toHaveBeenLastCalledWith(2)
    expect(lines()).toEqual(['Category 2', 'Baldwin v. Alberti'])
    runMenuCommand(host, 'toa-update')
    expect(lines()).toEqual(['Category 2', 'Baldwin v. Alberti'])
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'toa-mark-citation')
    expect(closed.openMarkCitation).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('New from Template runs with or without a document; Save as Template needs one', () => {
    const editor = makeEditor()
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'file-new-from-template')
    runMenuCommand(closed, 'file-save-as-template')
    expect(closed.newFromTemplate).toHaveBeenCalledTimes(1)
    expect(closed.saveAsTemplate).not.toHaveBeenCalled()
    const host = makeHost({ editor })
    runMenuCommand(host, 'file-save-as-template')
    expect(host.saveAsTemplate).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('AutoText, Reveal Formatting, Selection Pane and Language open with a document only', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-autotext')
    runMenuCommand(host, 'insert-autotext-new')
    expect(vi.mocked(host.openAutoText).mock.calls).toEqual([['insert'], ['new']])
    runMenuCommand(host, 'view-reveal-formatting')
    expect(host.openDialog).toHaveBeenLastCalledWith('revealFormatting')
    runMenuCommand(host, 'format-selection-pane')
    expect(host.toggleSelectionPane).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'citations-pane')
    expect(host.toggleCitationsPane).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'edit-select-data')
    expect(host.selectChartData).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'format-object')
    expect(host.showObjectFormatTab).toHaveBeenCalledTimes(1)
    editor.commands.selectAll()
    runMenuCommand(host, 'tools-language')
    expect(host.openLanguage).toHaveBeenCalledWith(editor)
    const closed = makeHost({ editor, hasDoc: false })
    runMenuCommand(closed, 'insert-autotext')
    runMenuCommand(closed, 'view-reveal-formatting')
    runMenuCommand(closed, 'format-selection-pane')
    runMenuCommand(closed, 'tools-language')
    runMenuCommand(closed, 'format-object')
    expect(closed.showObjectFormatTab).not.toHaveBeenCalled()
    runMenuCommand(closed, 'citations-pane')
    runMenuCommand(closed, 'edit-select-data')
    expect(closed.selectChartData).not.toHaveBeenCalled()
    expect(closed.toggleCitationsPane).not.toHaveBeenCalled()
    expect(closed.toggleSelectionPane).not.toHaveBeenCalled()
    expect(closed.openLanguage).not.toHaveBeenCalled()
    expect(closed.openAutoText).not.toHaveBeenCalled()
    expect(closed.openDialog).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('Find Next and Find Previous step the search, and Columns needs a section to act on', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'edit-find-next')
    expect(host.findStep).toHaveBeenLastCalledWith(1)
    runMenuCommand(host, 'edit-find-previous')
    expect(host.findStep).toHaveBeenLastCalledWith(-1)
    runMenuCommand(makeHost({ editor, hasDoc: false }), 'edit-find-next')
    expect(host.findStep).toHaveBeenCalledTimes(2)

    runMenuCommand(host, 'format-columns')
    expect(host.openDialog).not.toHaveBeenCalled()
    const withSection = makeHost({
      editor,
      section: { columns: 1 } as NonNullable<CommandHost['section']>,
    })
    runMenuCommand(withSection, 'format-columns')
    expect(withSection.openDialog).toHaveBeenCalledWith('columns')
    editor.destroy()
  })

  it('review commands reach the review host with their payload decoded', () => {
    const host = makeHost()
    const { review } = host
    runMenuCommand(host, 'delete-comments', 'resolved')
    expect(review.deleteAllComments).toHaveBeenCalledWith(true)
    runMenuCommand(host, 'comment-previous')
    expect(review.gotoComment).toHaveBeenCalledWith(-1)
    runMenuCommand(host, 'revision-display', 'original')
    expect(review.setRevisionDisplay).toHaveBeenCalledWith('original')
    runMenuCommand(host, 'revision-display', 'sideways')
    expect(review.setRevisionDisplay).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'markup-filter', 'format')
    expect(review.toggleMarkup).toHaveBeenLastCalledWith('format')
    runMenuCommand(host, 'markup-filter', 'author:Ada')
    expect(review.toggleMarkup).toHaveBeenLastCalledWith({ author: 'Ada' })
    runMenuCommand(host, 'accept-revision', 'shown')
    expect(review.resolveRevisions).toHaveBeenLastCalledWith('accept', 'shown')
    runMenuCommand(host, 'reject-revision', JSON.stringify({ author: 'Ada' }))
    expect(review.resolveRevisions).toHaveBeenLastCalledWith('reject', { author: 'Ada' })
    runMenuCommand(host, 'revision-next')
    expect(review.gotoRevision).toHaveBeenCalledWith(1)
  })

  it('review commands do nothing without a document', () => {
    const host = makeHost({ hasDoc: false })
    runMenuCommand(host, 'track-changes')
    runMenuCommand(host, 'protect-document')
    expect(host.review.toggleTrackChanges).not.toHaveBeenCalled()
    expect(host.review.protect).not.toHaveBeenCalled()
  })

  it('footnote, endnote and citation reach the host and the sources list', () => {
    const editor = makeEditor()
    const source = { tag: 'Knuth1997', title: 'The Art of Computer Programming' }
    const host = makeHost({ editor, sources: [source] as CommandHost['sources'] })
    runMenuCommand(host, 'insert-footnote')
    runMenuCommand(host, 'insert-endnote')
    expect(vi.mocked(host.insertNote).mock.calls).toEqual([['footnote'], ['endnote']])
    const before = editor.getText()
    runMenuCommand(host, 'insert-citation', 'Nobody2000')
    expect(editor.getText()).toBe(before)
    runMenuCommand(host, 'insert-citation', 'Knuth1997')
    expect(editor.getText().length).toBeGreaterThan(before.length)
    editor.destroy()
  })

  it('design picks reach the design host; a theme sets the halves it carries', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const fonts = { major: 'Georgia', minor: 'Verdana' }
    runMenuCommand(host, 'theme-fonts', JSON.stringify({ fonts }))
    expect(host.design.setTheme).toHaveBeenLastCalledWith({ fonts })
    // a pick from Word's lists goes by id: a theme carries all three parts, a scheme its own
    runMenuCommand(host, 'document-theme', 'facet')
    const facet = vi.mocked(host.design.setTheme).mock.lastCall![0]
    expect(facet.fonts?.major).toBe('Trebuchet MS')
    expect(facet.colors?.accent1).toBe('90C226')
    expect(facet.format).toMatchObject({ name: 'Facet' })
    runMenuCommand(host, 'theme-colors', 'blueWarm')
    expect(vi.mocked(host.design.setTheme).mock.lastCall![0]).toEqual({
      colors: expect.objectContaining({ name: 'Blue Warm', accent1: '4A66AC' }),
    })
    runMenuCommand(host, 'theme-fonts', 'georgia')
    expect(vi.mocked(host.design.setTheme).mock.lastCall![0]).toEqual({
      fonts: expect.objectContaining({ major: 'Georgia', minor: 'Georgia' }),
    })
    runMenuCommand(host, 'document-theme', 'no-such-theme')
    expect(host.design.setTheme).toHaveBeenCalledTimes(4)
    // a theme file of the themes folder goes by its file name; one that is gone picks nothing
    const mine = { fonts, colors: { accent1: '112233' }, format: { name: 'Mine', xml: '<a/>' } }
    vi.mocked(host.design.savedTheme).mockImplementation((name) =>
      name === 'Mine.thmx' ? mine : undefined,
    )
    runMenuCommand(host, 'document-theme', 'custom:Mine.thmx')
    expect(host.design.setTheme).toHaveBeenLastCalledWith(mine)
    runMenuCommand(host, 'document-theme', 'custom:Gone.thmx')
    expect(host.design.setTheme).toHaveBeenCalledTimes(5)
    // a theme read from a file comes as its own values, format scheme included
    runMenuCommand(host, 'document-theme', JSON.stringify(mine))
    expect(host.design.setTheme).toHaveBeenLastCalledWith(mine)
    // Reset puts back the theme from when the document was opened, once that is known
    runMenuCommand(host, 'theme-reset')
    expect(host.design.setTheme).toHaveBeenCalledTimes(6)
    vi.mocked(host.design.openedTheme).mockReturnValue({ fonts })
    runMenuCommand(host, 'theme-reset')
    expect(host.design.setTheme).toHaveBeenLastCalledWith({ fonts })
    runMenuCommand(host, 'theme-browse')
    expect(host.design.browseTheme).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'theme-save-current')
    expect(host.design.saveTheme).toHaveBeenCalledTimes(1)
    runMenuCommand(host, 'theme-custom-delete')
    runMenuCommand(host, 'theme-custom-delete', 'Mine.thmx')
    expect(vi.mocked(host.design.deleteTheme).mock.calls).toEqual([['Mine.thmx']])
    runMenuCommand(host, 'page-color', 'FFF2CC')
    runMenuCommand(host, 'page-color', NONE)
    expect(vi.mocked(host.design.setPageColor).mock.calls).toEqual([['FFF2CC'], [null]])
    runMenuCommand(host, 'watermark', 'DRAFT')
    runMenuCommand(host, 'watermark', NONE)
    expect(vi.mocked(host.design.setWatermark).mock.calls).toEqual([['DRAFT'], [null]])
    editor.setEditable(false)
    runMenuCommand(host, 'watermark', 'SAMPLE')
    expect(host.design.setWatermark).toHaveBeenCalledTimes(2)
    editor.destroy()
  })

  it('a layout pick toggles against Print Layout, and Print Layout leaves Read Mode', () => {
    const host = makeHost()
    runMenuCommand(host, 'view-mode', 'outline')
    expect(host.view.setMode).toHaveBeenLastCalledWith('outline')
    host.view.mode = 'outline'
    runMenuCommand(host, 'view-mode', 'outline')
    expect(host.view.setMode).toHaveBeenLastCalledWith('print')
    runMenuCommand(host, 'view-mode', 'print')
    expect(host.view.setReadMode).toHaveBeenCalledWith(false)
    runMenuCommand(host, 'view-mode', 'sideways')
    expect(host.view.setMode).toHaveBeenCalledTimes(3)
    runMenuCommand(host, 'read-mode')
    expect(host.view.setReadMode).toHaveBeenLastCalledWith(true)
  })

  it('Draft is a layout like the others, from the menu and from the ribbon', () => {
    const host = makeHost()
    runMenuCommand(host, 'view-draft')
    expect(host.view.setMode).toHaveBeenLastCalledWith('draft')
    host.view.mode = 'draft'
    runMenuCommand(host, 'view-mode', 'draft')
    expect(host.view.setMode).toHaveBeenLastCalledWith('print')
  })

  it('Arrange All asks the host to tile the windows, once a document is open', () => {
    const closed = makeHost({ hasDoc: false })
    runMenuCommand(closed, 'view-arrange-all')
    expect(closed.view.arrangeAll).not.toHaveBeenCalled()
    const host = makeHost()
    runMenuCommand(host, 'view-arrange-all')
    expect(host.view.arrangeAll).toHaveBeenCalledTimes(1)
  })

  it('Merge All Windows asks the host to bring the windows back, once a document is open', () => {
    const closed = makeHost({ hasDoc: false })
    runMenuCommand(closed, 'view-merge-windows')
    expect(closed.view.mergeWindows).not.toHaveBeenCalled()
    const host = makeHost()
    runMenuCommand(host, 'view-merge-windows')
    expect(host.view.mergeWindows).toHaveBeenCalledTimes(1)
  })

  it('a new window needs no document; the other view commands do', () => {
    const host = makeHost({ hasDoc: false })
    runMenuCommand(host, 'toggle-ruler')
    runMenuCommand(host, 'new-window')
    expect(host.view.toggleRuler).not.toHaveBeenCalled()
    expect(host.view.newWindow).toHaveBeenCalledTimes(1)
  })

  it('a header and footer action and an ink tool arrive decoded', () => {
    const host = makeHost()
    const action = { type: 'pageNumber', kind: 'footer', align: 'center' }
    runMenuCommand(host, 'header-footer', JSON.stringify(action))
    expect(host.headerFooterAction).toHaveBeenCalledWith(action)
    runMenuCommand(host, 'ink-tool', 'highlighter')
    runMenuCommand(host, 'ink-tool', 'crayon')
    expect(vi.mocked(host.setInkTool).mock.calls).toEqual([['highlighter']])
    runMenuCommand(host, 'ink-clear')
    expect(host.clearInk).toHaveBeenCalledTimes(1)
  })

  it('Toggle Drawing and Set as Default need a document', () => {
    const closed = makeHost({ hasDoc: false })
    runMenuCommand(closed, 'edit-toggle-drawing')
    runMenuCommand(closed, 'design-set-default')
    expect(closed.toggleDrawing).not.toHaveBeenCalled()
    expect(closed.setDesignAsDefault).not.toHaveBeenCalled()
    const host = makeHost()
    runMenuCommand(host, 'edit-toggle-drawing')
    runMenuCommand(host, 'design-set-default')
    expect(host.toggleDrawing).toHaveBeenCalledTimes(1)
    expect(host.setDesignAsDefault).toHaveBeenCalledTimes(1)
  })

  it('the picture dialogs open only on a picture', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'picture-alt-text')
    runMenuCommand(host, 'picture-size-position')
    expect(host.openDialog).not.toHaveBeenCalled()
  })

  it('the list buttons turn the paragraph into a list item and back', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'bullets')
    expect(editor.getAttributes('docListItem')).toMatchObject({ kind: 'bullet', numId: '7' })
    expect(host.listNumId).toHaveBeenCalledWith('bullet', true)
    runMenuCommand(host, 'numbering')
    expect(editor.getAttributes('docListItem').kind).toBe('ordered')
    runMenuCommand(host, 'numbering')
    expect(editor.isActive('docListItem')).toBe(false)
    editor.destroy()
  })

  it('a gallery card defines its levels, applies them and joins the recents', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const levels = [{ numFmt: 'decimal', lvlText: '%1.' }]
    runMenuCommand(host, 'list-preset', JSON.stringify({ levels, recent: 'numbers' }))
    expect(host.createListDef).toHaveBeenCalledWith(levels)
    expect(editor.getAttributes('docListItem')).toMatchObject({ kind: 'ordered', numId: '9' })
    runMenuCommand(host, 'list-level', '2')
    expect(editor.getAttributes('docListItem').ilvl).toBe(2)
    runMenuCommand(host, 'list-level', 'deep')
    expect(editor.getAttributes('docListItem').ilvl).toBe(2)
    runMenuCommand(host, 'list-none')
    expect(editor.isActive('docListItem')).toBe(false)
    editor.destroy()
  })

  it('a list style puts its list on the paragraph, keeping the level of a list item', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'list-style', '1ai')
    expect(host.listStyleNum).toHaveBeenCalledWith('1ai')
    expect(editor.getAttributes('docListItem')).toMatchObject({
      kind: 'ordered',
      numId: '11',
      ilvl: 0,
    })
    runMenuCommand(host, 'list-level', '1')
    runMenuCommand(host, 'list-style', '111111')
    expect(editor.getAttributes('docListItem').ilvl).toBe(1)
    editor.destroy()
  })

  it('a list style over a mixed-level selection leaves each item on its own level', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const item = (text: string, ilvl: number) => ({
      type: 'docListItem',
      attrs: { kind: 'bullet', numId: '3', ilvl },
      content: [{ type: 'text', text }],
    })
    editor.commands.setContent({
      type: 'doc',
      content: [
        item('one', 0),
        item('two', 2),
        { type: 'docParagraph', content: [{ type: 'text', text: 'plain' }] },
      ],
    })
    editor.commands.selectAll()
    runMenuCommand(host, 'list-style', '1ai')
    const items: unknown[] = []
    editor.state.doc.forEach((node) =>
      items.push([node.type.name, node.attrs.numId, node.attrs.ilvl]),
    )
    expect(items).toEqual([
      ['docListItem', '11', 0],
      ['docListItem', '11', 2],
      ['docListItem', '11', 0],
    ])
    editor.destroy()
  })

  it('an unknown list style leaves the paragraph alone', () => {
    const editor = makeEditor()
    const host = makeHost({ editor, listStyleNum: vi.fn(() => null) })
    runMenuCommand(host, 'list-style', 'Missing')
    expect(editor.isActive('docListItem')).toBe(false)
    editor.destroy()
  })

  it('a style pick reaches the host only on an editable document', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'apply-style', 'Heading1')
    expect(host.applyStyle).toHaveBeenCalledWith('Heading1')
    editor.setEditable(false)
    runMenuCommand(host, 'apply-style', 'Heading2')
    expect(host.applyStyle).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('change case takes its mode from the payload and ignores an unknown one', () => {
    const editor = makeEditor()
    editor.commands.selectAll()
    const host = makeHost({ editor })
    runMenuCommand(host, 'change-case', 'bogus')
    expect(editor.getText()).toBe('plain paragraph')
    runMenuCommand(host, 'change-case', 'upper')
    expect(editor.getText()).toBe('PLAIN PARAGRAPH')
    editor.destroy()
  })

  it('font size steps and clear formatting reach the host only on an editable document', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'font-grow')
    runMenuCommand(host, 'font-shrink')
    runMenuCommand(host, 'clear-formatting')
    expect(host.stepFontSize).toHaveBeenNthCalledWith(1, 1)
    expect(host.stepFontSize).toHaveBeenNthCalledWith(2, -1)
    expect(host.clearFormatting).toHaveBeenCalledTimes(1)
    editor.setEditable(false)
    runMenuCommand(host, 'font-grow')
    runMenuCommand(host, 'clear-formatting')
    expect(host.stepFontSize).toHaveBeenCalledTimes(2)
    expect(host.clearFormatting).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('font, size, colour and highlight take their value from the payload', () => {
    const editor = makeEditor()
    editor.commands.selectAll()
    const host = makeHost({ editor })
    const style = () => editor.getAttributes('docTextStyle')
    runMenuCommand(host, 'font-family', 'Georgia')
    expect(style().fontAscii).toBe('Georgia')
    runMenuCommand(host, 'font-size', '13.5')
    expect(style().sizeHalfPoints).toBe(27)
    runMenuCommand(host, 'font-size', 'big')
    runMenuCommand(host, 'font-size')
    expect(style().sizeHalfPoints).toBe(27)
    runMenuCommand(host, 'font-color', 'C00000')
    expect(style().color).toBe('C00000')
    runMenuCommand(host, 'font-color', NONE)
    expect(style().color ?? null).toBeNull()
    runMenuCommand(host, 'highlight', 'yellow')
    expect(style().highlight).toBe('yellow')
    runMenuCommand(host, 'highlight', NONE)
    expect(style().highlight ?? null).toBeNull()
    editor.destroy()
  })

  it('line spacing, paragraph spacing and shading set the paragraph', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const para = () => editor.getAttributes('docParagraph')
    runMenuCommand(host, 'line-spacing', '1.5')
    expect(para().lineSpacing).toBe(1.5)
    runMenuCommand(host, 'space-before')
    expect(para().spaceBefore).toBe(240)
    runMenuCommand(host, 'space-before')
    expect(para().spaceBefore).toBe(0)
    runMenuCommand(host, 'space-after')
    expect(para().spaceAfter).toBe(240)
    runMenuCommand(host, 'paragraph-shading', 'FFF2CC')
    expect(para().shadingFill).toBe('FFF2CC')
    runMenuCommand(
      host,
      'paragraph-shading',
      JSON.stringify({ fill: 'FFF2CC', pattern: 'pct25', color: 'C00000' }),
    )
    expect(para().shadingPattern).toBe(JSON.stringify({ val: 'pct25', color: 'C00000' }))
    expect(para().shadingDisplay).toBe('EFB699')
    // the ribbon's fill takes the pattern off, as Word's Shading button does
    runMenuCommand(host, 'paragraph-shading', 'DDEBF7')
    expect(para().shadingFill).toBe('DDEBF7')
    expect(para().shadingPattern ?? null).toBeNull()
    expect(para().shadingDisplay ?? null).toBeNull()
    runMenuCommand(host, 'paragraph-shading', NONE)
    expect(para().shadingFill ?? null).toBeNull()
    editor.destroy()
  })

  it('the indent commands step the paragraph indent and stop at the margin', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const indentLeft = () => Number(editor.getAttributes('docParagraph').indentLeft ?? 0)
    runMenuCommand(host, 'indent-increase')
    const stepped = indentLeft()
    expect(stepped).toBeGreaterThan(0)
    runMenuCommand(host, 'indent-decrease')
    runMenuCommand(host, 'indent-decrease')
    expect(indentLeft()).toBe(0)
    editor.destroy()
  })

  it('formatting commands leave a read-only document alone', () => {
    const editor = makeEditor()
    editor.commands.selectAll()
    editor.setEditable(false)
    const host = makeHost({ editor })
    const before = editor.getJSON()
    for (const command of [
      'bold',
      'strikethrough',
      'all-caps',
      'hidden-text',
      'double-underline',
      'superscript',
      'align-center',
      'indent-increase',
      'bullets',
      'direction-rtl',
      'space-before',
    ] as const)
      runMenuCommand(host, command)
    expect(editor.getJSON()).toEqual(before)
    editor.destroy()
  })

  it('on a table selected as a whole, a table command acts on every cell', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-table-grid', '2,3')
    const rowTexts = () => {
      const rows: string[] = []
      editor.state.doc.descendants((node) => {
        if (node.type.spec.tableRole === 'row') rows.push(node.textContent)
      })
      return rows
    }
    // mark the last row, so where a new row lands can be told apart
    let lastCell = -1
    editor.state.doc.descendants((node, at) => {
      if (node.type.spec.tableRole === 'cell' || node.type.spec.tableRole === 'header_cell')
        lastCell = at
    })
    editor
      .chain()
      .setTextSelection(lastCell + 2)
      .insertContent('tail')
      .run()
    expect(rowTexts()).toEqual(['', 'tail'])

    let tablePos = -1
    editor.state.doc.descendants((node, at) => {
      if (node.type.name === 'docTable' && tablePos < 0) tablePos = at
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, tablePos)),
    )
    runMenuCommand(host, 'table-insert-rows-below')
    // below the last row of the selection, not below the first row
    expect(rowTexts()).toEqual(['', 'tail', ''])
    expect(host.notInTable).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('a row inserted from the ribbon or the menu takes the neighbouring row\u2019s cell look', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-table-grid', '1,2')
    const cells = () => {
      const found: { pos: number; align: unknown }[] = []
      editor.state.doc.descendants((node, pos) => {
        const role = node.type.spec.tableRole
        if (role === 'cell' || role === 'header_cell') found.push({ pos, align: node.attrs.align })
      })
      return found
    }
    // give the only row a look a bare new row would not have
    let tr = editor.state.tr
    for (const { pos } of cells())
      tr = tr.setNodeMarkup(pos, undefined, { ...tr.doc.nodeAt(pos)!.attrs, align: 'bottom' })
    editor.view.dispatch(tr)
    editor.commands.setTextSelection(cells()[0].pos + 2)

    runMenuCommand(host, 'table-insert-rows-below')
    expect(cells().map((cell) => cell.align)).toEqual(['bottom', 'bottom', 'bottom', 'bottom'])
    editor.destroy()
  })

  it('a new comment opens the comments pane on its way, from the menu and the ribbon alike', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-comment')
    // the host's newComment is the one action that shows the pane and starts the comment
    expect(host.newComment).toHaveBeenCalledTimes(1)
    // commenting is allowed on a body that takes no edits (the comments-only restriction)
    editor.setEditable(false)
    runMenuCommand(host, 'insert-comment')
    expect(host.newComment).toHaveBeenCalledTimes(2)
    editor.destroy()
  })

  it('cell alignment, table look and table style take their pick from the payload', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-table-grid', '2,2')
    let firstCell = -1
    editor.state.doc.descendants((node, pos) => {
      if (firstCell < 0 && node.type.spec.tableRole?.endsWith('cell')) firstCell = pos
    })
    editor.commands.setTextSelection(firstCell + 2)
    const table = () => editor.getAttributes('docTable')
    const cell = () => editor.state.doc.nodeAt(firstCell)!.attrs

    runMenuCommand(host, 'table-cell-align', 'bottom,center')
    expect(cell().vAlign).toBe('bottom')
    runMenuCommand(host, 'table-cell-align', 'sideways,center')
    expect(cell().vAlign).toBe('bottom')

    runMenuCommand(host, 'table-look', 'lastRow=true')
    expect(table().tblLook).toMatchObject({ lastRow: true })
    runMenuCommand(host, 'table-look', 'lastRow=false')
    expect(table().tblLook).toMatchObject({ lastRow: false })

    runMenuCommand(host, 'table-style', 'GridTable4')
    expect(table().tblStyleId).toBe('GridTable4')
    runMenuCommand(host, 'table-style', NONE)
    expect(table().tblStyleId).toBeNull()
    expect(host.notInTable).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('rotate and flip turn the selected picture', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              blockType: 'image',
              imageDataUrl:
                'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
            },
          },
        ],
      },
    })
    const host = makeHost({ editor })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    const picture = () => editor.state.doc.firstChild!.attrs

    runMenuCommand(host, 'picture-rotate', '90')
    expect(picture().imageRotDeg).toBe(90)
    runMenuCommand(host, 'picture-rotate', '-90')
    // back upright is stored as no rotation at all
    expect(picture().imageRotDeg ?? null).toBeNull()
    runMenuCommand(host, 'picture-rotate', '-90')
    expect(picture().imageRotDeg).toBe(270)

    runMenuCommand(host, 'picture-flip', 'h')
    expect(picture().imageFlipH).toBe(true)
    runMenuCommand(host, 'picture-flip', 'h')
    expect(picture().imageFlipH).toBe(false)
    runMenuCommand(host, 'picture-flip', 'v')
    expect(picture().imageFlipV).toBe(true)
    editor.destroy()
  })

  it('bring forward and send backward move the selected picture one rank; an unknown step does nothing', () => {
    const picture = (z: number) => ({
      type: 'docProtected',
      attrs: {
        blockType: 'image',
        imageWrap: 'front',
        imageZOrder: z,
        imageDataUrl:
          'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
      },
    })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [picture(3), picture(9)] },
    })
    const host = makeHost({ editor })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    const rank = () => editor.state.doc.firstChild!.attrs.imageZOrder

    runMenuCommand(host, 'arrange-order', 'forward')
    expect(rank()).toBe(4)
    runMenuCommand(host, 'arrange-order', 'backward')
    runMenuCommand(host, 'arrange-order', 'backward')
    expect(rank()).toBe(2)
    runMenuCommand(host, 'arrange-order', 'front')
    expect(rank()).toBe(10)
    runMenuCommand(host, 'arrange-order', 'back')
    expect(rank()).toBe(8)
    runMenuCommand(host, 'arrange-order', 'constructor')
    runMenuCommand(host, 'arrange-order')
    expect(rank()).toBe(8)
    editor.destroy()
  })

  it('a caption lands below the selected picture instead of replacing it', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: { blockType: 'image', imageDataUrl: 'data:image/png;base64,AAAA' },
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'body' }] },
        ],
      },
    })
    const caption = { docxIndex: null, blockType: 'passthrough', genXml: '<w:p/>' }
    const kinds = () => {
      const out: string[] = []
      editor.state.doc.forEach((n) => out.push(n.attrs.blockType ?? n.type.name))
      return out
    }

    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    insertCaptionBlock(editor, caption)
    expect(kinds()).toEqual(['image', 'passthrough', 'docParagraph'])
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    insertCaptionBlock(editor, caption, 'above')
    expect(kinds()).toEqual(['passthrough', 'image', 'passthrough', 'docParagraph'])

    // with a caret the caption goes where the caret is
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertCaptionBlock(editor, caption)
    expect(kinds().filter((k) => k === 'image')).toHaveLength(1)
    expect(kinds().filter((k) => k === 'passthrough')).toHaveLength(3)
    editor.destroy()
  })

  it('a caption for an in-line picture follows its paragraph, or its table when it sits in a cell', () => {
    const inline = {
      type: 'docInlineImage',
      attrs: { dataUrl: 'data:image/png;base64,AAAA', widthPx: 120, heightPx: 80 },
    }
    const para = (content: object[]) => ({ type: 'docParagraph', content })
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          para([{ type: 'text', text: 'a' }, inline]),
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [{ type: 'docTableCell', content: [para([inline])] }],
              },
              {
                type: 'docTableRow',
                content: [
                  { type: 'docTableCell', content: [para([{ type: 'text', text: 'row 2' }])] },
                ],
              },
            ],
          },
          para([{ type: 'text', text: 'end' }]),
        ],
      },
    })
    const caption = { docxIndex: null, blockType: 'passthrough', genXml: '<w:p/>' }
    const top = () => {
      const out: string[] = []
      editor.state.doc.forEach((n) => out.push(n.type.name))
      return out
    }
    const pictures = () => {
      let count = 0
      editor.state.doc.descendants((n) => {
        if (n.type.name === 'docInlineImage') count++
      })
      return count
    }
    const selectPicture = (nth: number) => {
      const found: number[] = []
      editor.state.doc.descendants((n, pos) => {
        if (n.type.name === 'docInlineImage') found.push(pos)
      })
      editor.view.dispatch(
        editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, found[nth])),
      )
    }

    selectPicture(0)
    insertCaptionBlock(editor, caption)
    expect(top()).toEqual(['docParagraph', 'docProtected', 'docTable', 'docParagraph'])

    // a cell takes no caption block: it goes below the table
    selectPicture(1)
    insertCaptionBlock(editor, caption)
    expect(top()).toEqual([
      'docParagraph',
      'docProtected',
      'docTable',
      'docProtected',
      'docParagraph',
    ])
    expect(pictures()).toBe(2)
    // the table stays whole: both rows are still in it
    expect(editor.state.doc.child(2).childCount).toBe(2)
    editor.destroy()
  })

  it('crop and remove background open on the selected picture; change picture asks for a file', async () => {
    const dataUrl =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docProtected', attrs: { blockType: 'image', imageDataUrl: dataUrl } },
          { type: 'docParagraph', content: [{ type: 'text', text: 'body' }] },
        ],
      },
    })
    const host = makeHost({ editor })

    // the caret is in the text: there is no picture to act on
    editor.commands.setTextSelection(editor.state.doc.content.size - 2)
    runMenuCommand(host, 'picture-edit', 'crop')
    runMenuCommand(host, 'picture-replace')
    expect(host.openPictureDialog).not.toHaveBeenCalled()
    expect(host.pickImage).not.toHaveBeenCalled()

    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    runMenuCommand(host, 'picture-edit', 'crop')
    expect(host.openPictureDialog).toHaveBeenLastCalledWith('crop', dataUrl)
    runMenuCommand(host, 'picture-edit', 'cutout')
    expect(host.openPictureDialog).toHaveBeenLastCalledWith('cutout', dataUrl)
    runMenuCommand(host, 'picture-edit', 'sideways')
    expect(host.openPictureDialog).toHaveBeenCalledTimes(2)
    runMenuCommand(host, 'picture-replace')
    expect(host.pickImage).toHaveBeenCalledTimes(1)
    editor.destroy()
  })

  it('change picture from the icons and from the clipboard put the new bytes on the selected picture', async () => {
    const png = (pixel: string) => `data:image/png;base64,${pixel}`
    const old = png('AAAA')
    const pasted = png(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: { blockType: 'image', imageDataUrl: old, imageWidthPx: 200, imageHeightPx: 80 },
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'body' }] },
        ],
      },
    })
    const host = makeHost({ editor })
    const picture = () => editor.state.doc.child(0).attrs
    // jsdom loads no images: every picture reads as 40 x 20
    const RealImage = globalThis.Image
    class LoadedImage {
      naturalWidth = 40
      naturalHeight = 20
      onload: (() => void) | null = null
      set src(_url: string) {
        queueMicrotask(() => this.onload?.())
      }
    }
    vi.stubGlobal('Image', LoadedImage)
    const settle = () => new Promise((resolve) => setTimeout(resolve, 20))

    // no picture selected: neither asks for anything
    editor.commands.setTextSelection(editor.state.doc.content.size - 2)
    runMenuCommand(host, 'picture-replace-icon')
    expect(host.openDialog).not.toHaveBeenCalledWith('changePictureIcon')

    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    runMenuCommand(host, 'picture-replace-icon')
    expect(host.openDialog).toHaveBeenLastCalledWith('changePictureIcon')

    // text on the clipboard changes nothing
    Object.assign(navigator, {
      clipboard: {
        read: () =>
          Promise.resolve([{ types: ['text/plain'], getType: async () => new Blob(['x']) }]),
      },
    })
    runMenuCommand(host, 'picture-replace-clipboard')
    await settle()
    expect(picture().imageDataUrl).toBe(old)

    const bytes = Uint8Array.from(atob(pasted.slice(pasted.indexOf(',') + 1)), (ch) =>
      ch.charCodeAt(0),
    )
    const blob = new Blob([bytes], { type: 'image/png' })
    Object.assign(navigator, {
      clipboard: {
        // the bitmap is taken, whatever other picture format is listed before it
        read: () =>
          Promise.resolve([
            {
              types: ['image/svg+xml', 'image/png'],
              getType: async (type: string) =>
                type === 'image/png' ? blob : new Blob(['<svg/>'], { type }),
            },
          ]),
      },
    })
    runMenuCommand(host, 'picture-replace-clipboard')
    await settle()
    // the width stays, the height follows the new picture
    expect(picture()).toMatchObject({ imageDataUrl: pasted, imageWidthPx: 200, imageHeightPx: 100 })
    expect(picture().genImage).toMatchObject({ mime: 'image/png', widthPx: 200, heightPx: 100 })
    vi.stubGlobal('Image', RealImage)
    editor.destroy()
  })

  it('wrap reaches a picture in the header, footer or text box that has the focus', () => {
    const picture = {
      type: 'docProtected',
      attrs: {
        blockType: 'image',
        imageDataUrl:
          'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
      },
    }
    const body = makeEditor()
    const strip = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [picture] },
    })
    strip.view.dispatch(strip.state.tr.setSelection(NodeSelection.create(strip.state.doc, 0)))
    const host = makeHost({ editor: body })
    const bodyBefore = body.getJSON()
    setActiveSubEditor(strip)
    try {
      runMenuCommand(host, 'picture-wrap', 'square')
      expect(strip.state.doc.firstChild!.attrs.imageWrap).toBe('square')
      expect(body.getJSON()).toEqual(bodyBefore)
      // a drop cap belongs to a body paragraph: with the caret elsewhere it leaves the body alone
      runMenuCommand(host, 'format-drop-cap')
      expect(host.openDialog).not.toHaveBeenCalled()
      runMenuCommand(host, 'drop-cap', 'drop,4')
      expect(body.getJSON()).toEqual(bodyBefore)
    } finally {
      setActiveSubEditor(null)
      strip.destroy()
      body.destroy()
    }
  })

  it('rotate and flip act on a selected picture and on nothing else', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const before = editor.getJSON()
    // the caret is in plain text: there is no picture to turn
    runMenuCommand(host, 'picture-rotate', '90')
    runMenuCommand(host, 'picture-flip', 'h')
    runMenuCommand(host, 'picture-flip', 'diagonal')
    expect(editor.getJSON()).toEqual(before)
    editor.destroy()
  })

  it('the Insert tab header, footer and page number picks rewrite the strip through the host', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    const hf = host.headerFooter
    runMenuCommand(host, 'edit-header-footer', 'footer')
    expect(hf.edit).toHaveBeenCalledWith('footer')
    runMenuCommand(host, 'remove-header-footer', 'header')
    expect(hf.setHeader).toHaveBeenLastCalledWith({ text: '' })
    runMenuCommand(host, 'page-number', 'footer,center')
    expect(hfHasPageField(vi.mocked(hf.setFooter).mock.lastCall?.[0])).toBe(true)
    runMenuCommand(host, 'page-number', 'margin,center')
    expect(hf.setFooter).toHaveBeenCalledTimes(1)
    // Inside / Outside follow facing pages: handed over as such, not as one strip
    runMenuCommand(host, 'page-number', 'footer,outside')
    runMenuCommand(host, 'page-number', 'header,inside')
    expect(vi.mocked(hf.setFacingPageNumber).mock.calls).toEqual([
      ['footer', 'outside'],
      ['header', 'inside'],
    ])
    expect(hf.setFooter).toHaveBeenCalledTimes(1)

    // only the strip that carries a page number is rewritten when they are removed
    hf.footer = vi.mocked(hf.setFooter).mock.lastCall![0]
    hf.header = { text: 'Title' }
    vi.mocked(hf.setHeader).mockClear()
    runMenuCommand(host, 'page-number-remove')
    expect(hf.setFooter).toHaveBeenCalledTimes(2)
    expect(hfHasPageField(vi.mocked(hf.setFooter).mock.lastCall?.[0])).toBe(false)
    expect(hf.setHeader).not.toHaveBeenCalled()
    editor.destroy()
  })

  it('a chart type picked in the menu opens the dialog on that type', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'insert-chart', 'pie')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('pie')
    runMenuCommand(host, 'insert-chart', 'sunburst')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('sunburst')
    // a type there is no such chart for falls back to the dialog's own default
    runMenuCommand(host, 'insert-chart', 'bubble3d')
    expect(host.openChartDialog).toHaveBeenLastCalledWith(undefined)
    runMenuCommand(host, 'insert-chart-box-whisker')
    expect(host.openChartDialog).toHaveBeenLastCalledWith('boxWhisker')
    editor.destroy()
  })

  it('a command sent before the editor has read a click finds the selection the page shows', async () => {
    const element = document.createElement('div')
    document.body.appendChild(element)
    const editor = new Editor({
      element,
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'first' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'second' }] },
        ],
      },
    })
    await vi.waitFor(() => expect(editor.isInitialized).toBe(true))
    editor.commands.setTextSelection(2)
    editor.view.focus()
    const second = editor.view.domAtPos(10)
    document.getSelection()?.collapse(second.node, second.offset)
    expect(editor.state.selection.from).toBe(2)

    let seen = -1
    const host = makeHost({ editor })
    vi.mocked(host.openDialog).mockImplementation(() => {
      seen = editor.state.selection.from
    })
    runMenuCommand(host, 'shortcuts')
    expect(seen).toBe(10)
    editor.destroy()
    element.remove()
  })

  it('a table command outside a table tells the user and opens nothing', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    for (const command of [
      'table-insert-rows-below',
      'table-select-row',
      'table-merge-cells',
      'table-properties',
      'table-insert-cells',
    ] as const) {
      runMenuCommand(host, command)
    }
    expect(host.notInTable).toHaveBeenCalledTimes(5)
    expect(host.openTableDialog).not.toHaveBeenCalled()
    expect(editor.getText()).toBe('plain paragraph')
    editor.destroy()
  })

  it('Formula opens in body text as in a table, and not on a block selected as a whole', () => {
    const editor = makeEditor()
    const host = makeHost({ editor })
    runMenuCommand(host, 'table-formula')
    expect(host.openTableDialog).toHaveBeenLastCalledWith('formula')
    expect(host.notInTable).not.toHaveBeenCalled()

    editor.commands.setNodeSelection(0)
    runMenuCommand(host, 'table-formula')
    expect(host.openTableDialog).toHaveBeenCalledTimes(1)
    expect(host.notInTable).toHaveBeenCalledTimes(1)

    editor.commands.setTextSelection(1)
    runMenuCommand(host, 'insert-table-grid', '2,2')
    runMenuCommand(host, 'table-formula')
    expect(host.openTableDialog).toHaveBeenCalledTimes(2)
    editor.destroy()
  })

  describe('Edit > Repeat', () => {
    const twoParagraphs = (): Editor => {
      const editor = makeEditor()
      editor.commands.setContent({
        type: 'doc',
        content: ['first paragraph', 'second paragraph'].map((text) => ({
          type: 'docParagraph',
          content: [{ type: 'text', text }],
        })),
      })
      return editor
    }
    const paragraphs = (editor: Editor) => editor.getJSON().content ?? []
    const selectSecond = (editor: Editor) => {
      const from = editor.state.doc.child(0).nodeSize + 1
      editor.commands.setTextSelection({ from, to: from + 6 })
    }

    it('runs the last command, payload and all, on the new selection', () => {
      const editor = twoParagraphs()
      const host = makeHost({ editor })
      editor.commands.setTextSelection({ from: 1, to: 6 })
      expect(canRepeatCommand(host)).toBe(false)
      runMenuCommand(host, 'align-center')
      expect(canRepeatCommand(host)).toBe(true)
      selectSecond(editor)
      expect(canRepeatCommand(host)).toBe(true)
      runMenuCommand(host, 'repeat')
      expect(paragraphs(editor)[1].attrs).toEqual(paragraphs(editor)[0].attrs)
      expect(paragraphs(editor)[1].attrs?.align).toBe('center')

      editor.commands.setTextSelection({ from: 1, to: 6 })
      runMenuCommand(host, 'font-color', 'C00000')
      selectSecond(editor)
      runMenuCommand(host, 'repeat')
      const marks = paragraphs(editor)[1].content?.[0].marks ?? []
      expect(marks.some((mark) => mark.attrs?.color === 'C00000')).toBe(true)
      // a repeated command stays the last one
      expect(canRepeatCommand(host)).toBe(true)
      editor.destroy()
    })

    it('ends when the document moves on without a command: typing is not repeated', () => {
      const editor = twoParagraphs()
      const host = makeHost({ editor })
      editor.commands.setTextSelection({ from: 1, to: 6 })
      runMenuCommand(host, 'align-center')
      selectSecond(editor)
      editor.commands.insertContent('typed')
      expect(canRepeatCommand(host)).toBe(false)
      const before = editor.getJSON()
      runMenuCommand(host, 'repeat')
      expect(editor.getJSON()).toEqual(before)
      editor.destroy()
    })

    it('does not remember a command that left the document as it was', () => {
      const editor = twoParagraphs()
      const host = makeHost({ editor })
      runMenuCommand(host, 'align-center')
      editor.commands.insertContent('typed')
      runMenuCommand(host, 'word-count')
      runMenuCommand(host, 'font-dialog')
      expect(canRepeatCommand(host)).toBe(false)
      editor.destroy()
    })

    it('remembers edits only: a command that swaps the document is not run again', () => {
      const editor = twoParagraphs()
      const host = makeHost({ editor, newFile: () => editor.commands.insertContent('new') })
      runMenuCommand(host, 'new')
      expect(editor.getText()).toContain('new')
      expect(canRepeatCommand(host)).toBe(false)
      editor.destroy()
    })

    it('carries into a text box entered after the command, until that box is edited', () => {
      const editor = twoParagraphs()
      const box = makeEditor()
      const host = makeHost({ editor })
      editor.commands.setTextSelection({ from: 1, to: 6 })
      runMenuCommand(host, 'align-center')
      setActiveSubEditor(box)
      try {
        expect(canRepeatCommand(host)).toBe(true)
        box.commands.selectAll()
        runMenuCommand(host, 'repeat')
        expect(box.getJSON().content?.[0].attrs?.align).toBe('center')
        box.commands.insertContent('typed')
        expect(canRepeatCommand(host)).toBe(false)
      } finally {
        setActiveSubEditor(null)
        box.destroy()
        editor.destroy()
      }
    })

    it('outlives a header or footer that was visited and closed without an edit', () => {
      const editor = twoParagraphs()
      const strip = makeEditor()
      const host = makeHost({ editor })
      editor.commands.setTextSelection({ from: 1, to: 6 })
      runMenuCommand(host, 'align-center')
      setActiveSubEditor(strip)
      expect(canRepeatCommand(host)).toBe(true)
      setActiveSubEditor(null)
      strip.destroy()
      expect(canRepeatCommand(host)).toBe(true)
      editor.destroy()
    })

    const openStrip = () => {
      const el = document.createElement('div')
      document.body.append(el)
      const commits: unknown[] = []
      const handle = mountHfEditor(el, {
        value: { text: '', paras: [{ align: 'left', runs: [{ text: 'Header' }] }] },
        pageNo: '1',
        pageTotal: '1',
        spellcheck: false,
        onCommit: (next) => commits.push(next),
        onExit: () => el.remove(),
      })
      setActiveSubEditor(handle.editor)
      return { handle, commits }
    }
    const centredBody = () => {
      const editor = twoParagraphs()
      const host = makeHost({ editor })
      editor.commands.setTextSelection({ from: 1, to: 6 })
      runMenuCommand(host, 'align-center')
      return { editor, host }
    }

    it('outlives a real header strip that closes untouched', () => {
      const { editor, host } = centredBody()
      const { handle, commits } = openStrip()
      expect(canRepeatCommand(host)).toBe(true)
      handle.exit()
      expect(commits).toHaveLength(0)
      expect(canRepeatCommand(host)).toBe(true)
      editor.destroy()
    })

    it('ends when the header was typed in, and stays ended once the strip has closed', () => {
      const { editor, host } = centredBody()
      const { handle, commits } = openStrip()
      expect(canRepeatCommand(host)).toBe(true)
      handle.editor.commands.insertContent('typed')
      expect(canRepeatCommand(host)).toBe(false)
      handle.exit()
      expect(commits).toHaveLength(1)
      expect(canRepeatCommand(host)).toBe(false)
      runMenuCommand(host, 'repeat')
      expect(paragraphs(editor)[1].attrs?.align).not.toBe('center')
      editor.destroy()
    })

    it('ends on a header edit that closed before Repeat was looked at again', () => {
      const { editor, host } = centredBody()
      const { handle } = openStrip()
      handle.editor.commands.insertContent('typed')
      handle.exit()
      expect(canRepeatCommand(host)).toBe(false)
      editor.destroy()
    })

    it('ends on a header edit seen only after it was typed', () => {
      const { editor, host } = centredBody()
      const { handle } = openStrip()
      handle.editor.commands.insertContent('typed')
      canRepeatCommand(host)
      handle.exit()
      expect(canRepeatCommand(host)).toBe(false)
      editor.destroy()
    })

    it('ends when an edited header strip is discarded by its owner', () => {
      const { editor, host } = centredBody()
      const { handle } = openStrip()
      expect(canRepeatCommand(host)).toBe(true)
      handle.editor.commands.insertContent('typed')
      handle.discard()
      expect(canRepeatCommand(host)).toBe(false)
      editor.destroy()
    })

    it('stays after the command was repeated inside a header that then closes', () => {
      const { editor, host } = centredBody()
      const { handle, commits } = openStrip()
      handle.editor.commands.selectAll()
      runMenuCommand(host, 'repeat')
      handle.exit()
      expect(commits).toHaveLength(1)
      expect(canRepeatCommand(host)).toBe(true)
      selectSecond(editor)
      runMenuCommand(host, 'repeat')
      expect(paragraphs(editor)[1].attrs?.align).toBe('center')
      editor.destroy()
    })

    it('is Redo while there is something to redo', () => {
      const editor = twoParagraphs()
      const host = makeHost({ editor })
      editor.commands.setTextSelection({ from: 1, to: 6 })
      runMenuCommand(host, 'align-center')
      const centered = editor.getJSON()
      runMenuCommand(host, 'undo')
      expect(editor.getJSON()).not.toEqual(centered)
      expect(canRepeatCommand(host)).toBe(false)
      runMenuCommand(host, 'repeat')
      expect(editor.getJSON()).toEqual(centered)
      editor.destroy()
    })

    it('repeats Text Direction on the kind of target it turned: cells, not the section', () => {
      const editor = twoParagraphs()
      const section = { columns: 1 } as NonNullable<CommandHost['section']>
      const host = makeHost({ editor, section })
      editor.commands.setTextSelection(editor.state.doc.content.size - 1)
      runMenuCommand(host, 'insert-table-grid', '2,2')
      const directions = () => {
        const found: unknown[] = []
        editor.state.doc.descendants((node) => {
          if (node.type.spec.tableRole === 'cell') found.push(node.attrs.textDirection)
        })
        return found
      }
      runMenuCommand(host, 'format-text-direction', 'tbRl')
      expect(directions().filter((d) => d === 'tbRl')).toHaveLength(1)
      expect(canRepeatCommand(host)).toBe(true)

      editor.commands.setTextSelection(1)
      runMenuCommand(host, 'repeat')
      expect(host.setSection).not.toHaveBeenCalled()
      expect(host.notInTable).toHaveBeenCalledTimes(1)
      expect(directions().filter((d) => d === 'tbRl')).toHaveLength(1)
      editor.destroy()
    })
  })
})
