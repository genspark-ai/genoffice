/**
 * A ribbon control that stands for a command takes its greyed-out and pressed
 * looks from the command, by the rules the menu bar follows.
 */
import { describe, expect, it } from 'vitest'
import { offByDocument } from '../src/renderer/commands/menu-state'
import { RIBBON_COMMAND_IDS, ribbonCommandPressed } from '../src/renderer/commands/ribbon-controls'
import { COMMANDS, type CommandDef } from '../src/shared/commands'
import { EMPTY_FORMAT_STATE } from '../src/renderer/components/ribbon-format-state'

const VIEW = {
  showAi: false,
  showMarks: false,
  showStylesPane: false,
  readOnlyRecommended: false,
  mergePreview: false,
  hasComments: false,
  hasMergeSource: false,
  windowsDetached: false,
}
const open = {
  hasDoc: true,
  editable: true,
  inTable: false,
  inTextBox: false,
  cellsSelected: false,
  labelSheet: false,
  hasSelection: false,
  pictureSelected: false,
  hasNotes: false,
  hasChart: false,
  canSplitTable: false,
  canComment: true,
  inEmptyParagraph: false,
  hasComments: false,
  hasTableOfFigures: false,
  hasIndex: false,
  hasTableOfAuthorities: false,
  hasMergeSource: false,
  windowsDetached: false,
}

describe('ribbon command state', () => {
  it('a read-only document greys the formatting commands, not the view ones', () => {
    const readOnly = { ...open, editable: false }
    for (const id of ['bold', 'strikethrough', 'superscript', 'indent-increase'] as const)
      expect(offByDocument(id, readOnly), id).toBe(true)
    for (const id of ['find', 'replace', 'show-marks', 'styles-pane', 'toggle-ai'] as const)
      expect(offByDocument(id, readOnly), id).toBe(false)
  })

  it('with no document only the AI entry stays on', () => {
    const none = { ...open, hasDoc: false, editable: false }
    expect(offByDocument('toggle-ai', none)).toBe(false)
    for (const id of ['find', 'show-marks', 'styles-pane', 'bold'] as const)
      expect(offByDocument(id, none), id).toBe(true)
  })

  it('paragraph direction is off inside a text box', () => {
    const inTextBox = { ...open, inTextBox: true }
    for (const id of ['direction-ltr', 'direction-rtl'] as const) {
      expect(offByDocument(id, open), id).toBe(false)
      expect(offByDocument(id, inTextBox), id).toBe(true)
    }
    expect(offByDocument('bullets', inTextBox)).toBe(false)
  })

  it('controls on the tabs follow the selection and the document as their menu items do', () => {
    // References > Show Notes
    expect(offByDocument('view-footnotes', open)).toBe(true)
    expect(offByDocument('view-footnotes', { ...open, hasNotes: true })).toBe(false)
    // Insert > Date & Time, Review > Thesaurus, Home > Phonetic Guide and Enclose Characters
    const picture = { ...open, hasSelection: true, pictureSelected: true }
    const cells = { ...open, inTable: true, hasSelection: true, cellsSelected: true }
    for (const id of [
      'insert-date-time',
      'tools-thesaurus',
      'format-phonetic-guide',
      'format-enclose-characters',
    ] as const) {
      expect(offByDocument(id, open), id).toBe(false)
      expect(offByDocument(id, picture), id).toBe(true)
      expect(offByDocument(id, cells), id).toBe(true)
    }
  })

  it('controls that need something the document does not have yet are greyed until it does', () => {
    const needs = [
      ['update-table-of-figures', 'hasTableOfFigures'],
      ['update-index', 'hasIndex'],
      ['toa-update', 'hasTableOfAuthorities'],
      ['show-comments', 'hasComments'],
      ['view-merge-windows', 'windowsDetached'],
      ['mailings-edit-recipients', 'hasMergeSource'],
      ['mailings-filter-recipients', 'hasMergeSource'],
      ['mailings-find-recipient', 'hasMergeSource'],
      ['mailings-rules', 'hasMergeSource'],
      ['mailings-preview-results', 'hasMergeSource'],
      ['mailings-finish-merge', 'hasMergeSource'],
    ] as const
    for (const [id, input] of needs) {
      expect(offByDocument(id, open), id).toBe(true)
      expect(offByDocument(id, { ...open, [input]: true }), id).toBe(false)
    }
    // what starts a merge or attaches the list stays live without one
    for (const id of ['mailings-start', 'mailings-select-recipients', 'update-toc'] as const)
      expect(offByDocument(id, open), id).toBe(false)
  })

  it('New Comment needs a place to hold on to, and Drop Cap a paragraph with text', () => {
    expect(offByDocument('insert-comment', open)).toBe(false)
    expect(offByDocument('insert-comment', { ...open, canComment: false })).toBe(true)
    for (const id of ['drop-cap', 'format-drop-cap'] as const) {
      expect(offByDocument(id, open), id).toBe(false)
      expect(offByDocument(id, { ...open, inEmptyParagraph: true }), id).toBe(true)
    }
  })

  it('the list and direction buttons show the paragraph at the caret', () => {
    const list = { ...EMPTY_FORMAT_STATE, listOrdered: true, bidi: true }
    expect(ribbonCommandPressed('numbering', list, VIEW)).toBe(true)
    expect(ribbonCommandPressed('bullets', list, VIEW)).toBe(false)
    expect(ribbonCommandPressed('direction-rtl', list, VIEW)).toBe(true)
    expect(ribbonCommandPressed('direction-ltr', list, VIEW)).toBe(false)
  })

  it('Always Open Read-Only shows pressed while the document asks for it', () => {
    expect(ribbonCommandPressed('file-always-read-only', EMPTY_FORMAT_STATE, VIEW)).toBe(false)
    expect(
      ribbonCommandPressed('file-always-read-only', EMPTY_FORMAT_STATE, {
        ...VIEW,
        readOnlyRecommended: true,
      }),
    ).toBe(true)
  })

  it('Preview Results shows pressed while the merge fields show a recipient', () => {
    expect(ribbonCommandPressed('mailings-preview-results', EMPTY_FORMAT_STATE, VIEW)).toBe(false)
    expect(
      ribbonCommandPressed('mailings-preview-results', EMPTY_FORMAT_STATE, {
        ...VIEW,
        mergePreview: true,
      }),
    ).toBe(true)
  })

  it('a control Word has and we have not built yet is greyed out whatever the document', () => {
    const placeholders = RIBBON_COMMAND_IDS.filter((id) => (COMMANDS[id] as CommandDef).planned)
    expect(placeholders).not.toContain('view-arrange-all')
    expect(COMMANDS['view-draft']).not.toHaveProperty('planned')
    for (const id of placeholders) {
      expect(offByDocument(id, open), id).toBe(true)
      expect(ribbonCommandPressed(id, EMPTY_FORMAT_STATE, VIEW), id).toBe(false)
    }
  })

  it('Word Count on the Review tab is the menu command, live once a document is open', () => {
    expect(COMMANDS['word-count']).not.toHaveProperty('planned')
    expect(offByDocument('word-count', open)).toBe(false)
    expect(offByDocument('word-count', { ...open, hasDoc: false })).toBe(true)
  })

  it('pressed follows the formatting at the caret', () => {
    const format = { ...EMPTY_FORMAT_STATE, bold: true, strike: true, vertAlign: 'subscript' }
    expect(ribbonCommandPressed('bold', format, VIEW)).toBe(true)
    expect(ribbonCommandPressed('italic', format, VIEW)).toBe(false)
    expect(ribbonCommandPressed('strikethrough', format, VIEW)).toBe(true)
    expect(ribbonCommandPressed('subscript', format, VIEW)).toBe(true)
    expect(ribbonCommandPressed('superscript', format, VIEW)).toBe(false)
  })

  it('a paragraph with no alignment of its own shows its reading direction', () => {
    expect(ribbonCommandPressed('align-left', EMPTY_FORMAT_STATE, VIEW)).toBe(true)
    const rtl = { ...EMPTY_FORMAT_STATE, bidi: true }
    expect(ribbonCommandPressed('align-right', rtl, VIEW)).toBe(true)
    expect(ribbonCommandPressed('align-left', rtl, VIEW)).toBe(false)
    const centered = { ...rtl, align: 'center' }
    expect(ribbonCommandPressed('align-center', centered, VIEW)).toBe(true)
    expect(ribbonCommandPressed('align-right', centered, VIEW)).toBe(false)
  })

  it('pressed follows the panes the app shows', () => {
    const view = { ...VIEW, showAi: true, showMarks: true }
    expect(ribbonCommandPressed('toggle-ai', EMPTY_FORMAT_STATE, view)).toBe(true)
    expect(ribbonCommandPressed('show-marks', EMPTY_FORMAT_STATE, view)).toBe(true)
    expect(ribbonCommandPressed('styles-pane', EMPTY_FORMAT_STATE, view)).toBe(false)
  })
})
