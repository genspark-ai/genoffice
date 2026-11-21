import {
  COMMANDS,
  type CommandDef,
  type CommandId,
  type CommandMenuState,
} from '../../shared/commands'

/** the app state the menu bar's check marks and greyed-out items are derived from */
export interface MenuStateInput {
  aiSidebar: boolean
  darkPage: boolean
  hasDoc: boolean
  editable: boolean
  inTable: boolean
  /** the caret is in a text box or a header / footer strip */
  inTextBox: boolean
  /** a selection of several table cells */
  cellsSelected: boolean
  /** something is selected where the caret is: text, cells or an object */
  hasSelection: boolean
  /** the selection is a picture */
  pictureSelected: boolean
  /** the document has a footnote or an endnote */
  hasNotes: boolean
  hasChart: boolean
  /** Split Table has somewhere to cut: false in a table that does not sit in the body */
  canSplitTable: boolean
  /** a new comment has something to hold on to at the caret (see canCommentAt) */
  canComment: boolean
  /** the caret's paragraph has nothing in it */
  inEmptyParagraph: boolean
  hasComments: boolean
  hasTableOfFigures: boolean
  hasIndex: boolean
  hasTableOfAuthorities: boolean
  /** a mail-merge recipient list is attached */
  hasMergeSource: boolean
  /** a Docs window stands outside the tab strip */
  windowsDetached: boolean
  /** the document is a sheet of mail-merge labels */
  labelSheet: boolean
  canUndo: boolean
  canRedo: boolean
  /** the last command can run again on the selection (see repeat.ts) */
  canRepeat: boolean
  /** check-as-you-type spelling is on */
  spellcheck: boolean
  /** the document has a file of its own */
  onDisk: boolean
  /** the document asks to be opened read-only */
  readOnlyRecommended: boolean
  /** how the document is shown, for the View menu's check marks */
  view: {
    mode: 'print' | 'web' | 'outline' | 'draft' | 'story'
    readMode: boolean
    ruler: boolean
    grid: boolean
    navigation: boolean
    /** the navigation pane is open on its page thumbnails */
    thumbnails: boolean
    /** tracked changes are shown as markup */
    markup: boolean
    /** everything but the page is hidden */
    focus: boolean
    /** the ribbon's command band is showing */
    ribbon: boolean
    reviewPane: boolean
  }
  /** what Cut, Copy and Paste would act on (see edit-target.ts) */
  target: { hasSelection: boolean; editable: boolean; field: boolean }
}

const ALL = Object.keys(COMMANDS) as CommandId[]

/** still meaningful with no document open */
const WITHOUT_DOC: ReadonlySet<CommandId> = new Set([
  'new',
  'open',
  'open-path',
  'file-new-from-template',
  'toggle-ai',
  'toggle-dark',
  'new-window',
  'ink-tool',
  'autocorrect-options',
  'preferences',
  'shortcuts',
  'tools-customize-keyboard',
  'view-ribbon',
  'table-gridlines',
])

/** act on the table around the caret */
const IN_TABLE: ReadonlySet<CommandId> = new Set(
  ALL.filter((id) => id.startsWith('table-') && id !== 'table-gridlines' && id !== 'table-formula'),
)

/** act on the document's file, which a document never saved does not have */
const ON_DISK: ReadonlySet<CommandId> = new Set(['file-rename', 'file-move'])

/** change the document, so unavailable while it is read-only; selecting is not a change */
const EDITS: ReadonlySet<CommandId> = new Set([
  'file-reduce-size',
  'insert-table',
  'insert-image',
  'insert-page-break',
  'insert-blank-page',
  'insert-chart',
  'insert-chart-bar',
  'insert-chart-area',
  'insert-chart-radar',
  'insert-chart-scatter',
  'insert-chart-stock',
  'insert-chart-treemap',
  'insert-chart-sunburst',
  'insert-chart-histogram',
  'insert-chart-pareto',
  'insert-chart-box-whisker',
  'insert-chart-waterfall',
  'insert-chart-funnel',
  'insert-chart-surface',
  'edit-header-footer',
  'view-header-footer',
  'remove-header-footer',
  'page-number',
  'page-number-format',
  'page-number-remove',
  'insert-footnote',
  'insert-footnote-dialog',
  'insert-endnote',
  'insert-toc',
  'update-toc',
  'insert-citation',
  'insert-bibliography',
  'citation-style',
  'insert-index',
  'insert-cover-page',
  'remove-cover-page',
  'insert-table-grid',
  'insert-shape',
  'insert-text-box',
  'insert-vertical-text-box',
  'insert-wordart',
  'drop-cap',
  'insert-field',
  'insert-date-field',
  'insert-field-dialog',
  'insert-date-time',
  'insert-watermark',
  'format-drop-cap',
  'insert-symbol',
  'insert-font-symbol',
  'insert-advanced-symbol',
  'insert-page-numbers',
  'mailings-start',
  'mailings-select-recipients',
  'mailings-edit-recipients',
  'insert-number',
  'sort',
  'format-text-direction',
  'mailings-filter-recipients',
  'mailings-find-recipient',
  'view-footnotes',
  'mark-index-entry',
  'index-automark',
  'update-index',
  'file-always-read-only',
  'mailings-finish-merge',
  'toc-add-text',
  'tools-translate-selection',
  'tools-translate-document',
  'tools-envelopes',
  'tools-labels',
  'insert-bookmark',
  'insert-caption',
  'insert-cross-reference',
  'insert-link',
  'insert-equation',
  'bold',
  'italic',
  'underline',
  'strikethrough',
  'subscript',
  'superscript',
  'indent-decrease',
  'indent-increase',
  'font-family',
  'font-size',
  'font-color',
  'highlight',
  'character-shading',
  'character-border',
  'insert-index-tables',
  'insert-table-of-figures',
  'update-table-of-figures',
  'toa-mark-citation',
  'toa-insert',
  'tools-language',
  'insert-autotext',
  'insert-file',
  'convert-text-to-table',
  'insert-icons',
  'insert-icon',
  'insert-horizontal-line',
  'mailings-update-labels',
  'insert-autotext-new',
  'toa-update',
  'tools-hyphenation',
  'design-paragraph-spacing',
  'design-custom-paragraph-spacing',
  'line-spacing',
  'space-before',
  'space-after',
  'paragraph-shading',
  'paragraph-borders',
  'format-borders-shading',
  'mailings-rules',
  'font-grow',
  'font-shrink',
  'change-case',
  'format-change-case',
  'clear-formatting',
  'bullets',
  'list-none',
  'list-level',
  'list-preset',
  'list-style',
  'list-style-add',
  'list-style-delete',
  'list-dialog',
  'apply-style',
  'numbering',
  'direction-ltr',
  'direction-rtl',
  'paste-as',
  'align-left',
  'align-center',
  'align-right',
  'align-justify',
  'align-distributed',
  'edit-clear-contents',
  'layout-line-numbers',
  'layout-suppress-line-numbers',
  'layout-line-numbering-options',
  'font-dialog',
  'paragraph-dialog',
  'page-setup',
  'format-document',
  'format-equation-options',
  'document-theme',
  'theme-fonts',
  'theme-colors',
  'page-color',
  'watermark',
  'page-border',
  'page-margins',
  'format-columns',
  'format-tabs',
  'format-phonetic-guide',
  'format-text-effects',
  'text-effects-preset',
  'format-enclose-characters',
  'format-combine-characters',
  'format-fit-text',
  'format-two-lines-in-one',
  'format-horizontal-in-vertical',
  'format-bullets-numbering',
  'document-settings',
  'page-orientation',
  'page-size',
  'page-columns',
  'columns-settings',
  'insert-column-break',
  'insert-text-wrapping-break',
  'insert-section-break',
  'picture-wrap',
  'picture-rotate',
  'picture-flip',
  'picture-replace',
  'picture-edit',
  'picture-position',
  'picture-graphics-fill',
  'shape-text-direction',
  'arrange-order',
  'arrange-align',
  'arrange-group',
  'insert-comment',
  'ai-proofread',
  'table-formula',
  ...ALL.filter((id) => IN_TABLE.has(id) && !id.startsWith('table-select-')),
])

/** an edit of the document's content, which is what Edit > Repeat can run again */
export const editsDocument = (id: CommandId): boolean => EDITS.has(id)

/**
 * Undo, redo and the clipboard act on whatever has the focus, so a text field
 * decides for itself; the document's own rules apply only when it has the focus.
 * Undo and Redo run on the document, so they step aside for a focused field:
 * greyed out, their accelerators fall through to the field's own undo.
 */
/** Cut and Copy need something selected in whatever has the focus; the ribbon greys its buttons by this too */
export function clipboardOff(
  id: 'cut' | 'copy',
  { hasDoc, target }: Pick<MenuStateInput, 'hasDoc' | 'target'>,
): boolean {
  if (!hasDoc && !target.field) return true
  return !target.hasSelection || (id === 'cut' && !target.editable)
}

function editCommandOff(id: CommandId, input: MenuStateInput): boolean | null {
  const { target } = input
  const noTarget = !input.hasDoc && !target.field
  switch (id) {
    case 'undo':
      return target.field || !input.hasDoc || !input.canUndo
    case 'redo':
      return target.field || !input.hasDoc || !input.canRedo
    case 'repeat':
      return target.field || !input.hasDoc || !(input.canRedo || input.canRepeat)
    case 'copy':
    case 'cut':
      return clipboardOff(id, input)
    case 'paste':
    case 'paste-match-style':
      return noTarget || !target.editable
    // opens on the document, so a focused text field or no document leaves it nothing to do
    case 'edit-paste-special':
      return !input.hasDoc || target.field || !target.editable
    case 'select-all':
      return noTarget
    default:
      return null
  }
}

/** Word greys these out while several cells are selected */
const NOT_ON_CELLS: ReadonlySet<CommandId> = new Set([
  'insert-table',
  'insert-link',
  'insert-date-time',
  'insert-field-dialog',
  'insert-advanced-symbol',
  'format-phonetic-guide',
  'format-enclose-characters',
  'format-horizontal-in-vertical',
  'format-combine-characters',
  'format-two-lines-in-one',
  'tools-thesaurus',
  'tools-translate-selection',
])

/** work on text or put some at the caret, which a selected picture has none of */
const NOT_ON_PICTURE: ReadonlySet<CommandId> = new Set([
  'insert-date-time',
  'insert-field-dialog',
  'insert-advanced-symbol',
  'format-drop-cap',
  'format-fit-text',
  'format-phonetic-guide',
  'format-enclose-characters',
  'format-horizontal-in-vertical',
  'format-combine-characters',
  'format-two-lines-in-one',
  'tools-thesaurus',
  'tools-translate-selection',
])

/** act on what is selected, so a bare caret gives them nothing */
const NEEDS_SELECTION: ReadonlySet<CommandId> = new Set(['insert-autotext-new', 'format-fit-text'])

/** act on body text, which a table's cells are not */
const NOT_IN_TABLE: ReadonlySet<CommandId> = new Set(['convert-text-to-table', 'format-drop-cap'])

/** set on the body paragraph, which a text box does not have */
const NOT_IN_TEXT_BOX: ReadonlySet<CommandId> = new Set([
  'direction-ltr',
  'direction-rtl',
  'convert-text-to-table',
  'table-formula',
])

/** work on the recipient list, so Word greys them until one is attached */
const NEEDS_MERGE_SOURCE: ReadonlySet<CommandId> = new Set([
  'mailings-edit-recipients',
  'mailings-filter-recipients',
  'mailings-find-recipient',
  'mailings-rules',
  'mailings-preview-results',
  'mailings-finish-merge',
])

/** hold a place in the menu and are not built yet */
const PLANNED: ReadonlySet<CommandId> = new Set(
  ALL.filter((id) => (COMMANDS[id] as CommandDef).planned),
)

/** what the document and the caret alone rule out; the ribbon greys its controls by this too */
export type DocumentStateInput = Pick<
  MenuStateInput,
  | 'hasDoc'
  | 'editable'
  | 'inTable'
  | 'inTextBox'
  | 'cellsSelected'
  | 'labelSheet'
  | 'hasSelection'
  | 'pictureSelected'
  | 'hasNotes'
  | 'hasChart'
  | 'canSplitTable'
  | 'canComment'
  | 'inEmptyParagraph'
  | 'hasComments'
  | 'hasTableOfFigures'
  | 'hasIndex'
  | 'hasTableOfAuthorities'
  | 'hasMergeSource'
  | 'windowsDetached'
>

export function offByDocument(id: CommandId, input: DocumentStateInput): boolean {
  return (
    PLANNED.has(id) ||
    (!input.hasDoc && !WITHOUT_DOC.has(id)) ||
    (!input.editable && EDITS.has(id)) ||
    (!input.inTable && IN_TABLE.has(id)) ||
    (input.inTable && NOT_IN_TABLE.has(id)) ||
    (input.inTextBox && NOT_IN_TEXT_BOX.has(id)) ||
    (!input.cellsSelected && id === 'table-merge-cells') ||
    (!input.labelSheet && id === 'mailings-update-labels') ||
    (input.cellsSelected && NOT_ON_CELLS.has(id)) ||
    (input.pictureSelected && NOT_ON_PICTURE.has(id)) ||
    (!input.hasSelection && NEEDS_SELECTION.has(id)) ||
    (!input.hasNotes && id === 'view-footnotes') ||
    (!input.hasChart && id === 'edit-select-data') ||
    (!input.canSplitTable && id === 'table-split-table') ||
    (!input.canComment && id === 'insert-comment') ||
    (input.inEmptyParagraph && (id === 'format-drop-cap' || id === 'drop-cap')) ||
    (!input.hasComments && id === 'show-comments') ||
    (!input.hasTableOfFigures && id === 'update-table-of-figures') ||
    (!input.hasIndex && id === 'update-index') ||
    (!input.hasTableOfAuthorities && id === 'toa-update') ||
    (!input.hasMergeSource && NEEDS_MERGE_SOURCE.has(id)) ||
    (!input.windowsDetached && id === 'view-merge-windows') ||
    // in body text a formula goes in at the caret, which a selected picture is not
    (!input.inTable && input.pictureSelected && id === 'table-formula')
  )
}

export function menuCommandState(input: MenuStateInput): CommandMenuState {
  const checked: CommandId[] = []
  if (input.aiSidebar) checked.push('toggle-ai')
  if (input.darkPage) checked.push('toggle-dark')
  const { view } = input
  // Read Mode is its own view: no layout is the current one while it is on
  if (!view.readMode) {
    if (view.mode === 'print') checked.push('view-print-layout')
    if (view.mode === 'web') checked.push('view-web-layout')
    if (view.mode === 'outline') checked.push('view-outline')
    if (view.mode === 'draft') checked.push('view-draft')
  }
  if (view.ruler) checked.push('toggle-ruler')
  if (view.grid) checked.push('toggle-grid')
  if (view.navigation) checked.push('toggle-navigation')
  if (view.thumbnails) checked.push('view-thumbnails')
  if (view.markup) checked.push('view-markup')
  if (view.ribbon) checked.push('view-ribbon')
  if (view.focus) checked.push('view-focus')
  if (input.readOnlyRecommended) checked.push('file-always-read-only')
  if (view.reviewPane) checked.push('review-pane')
  if (!input.spellcheck) checked.push('tools-hide-spelling')
  const disabled = ALL.filter(
    (id) =>
      editCommandOff(id, input) ?? (offByDocument(id, input) || (!input.onDisk && ON_DISK.has(id))),
  )
  const repeat = disabled.includes('repeat') ? null : input.canRedo ? 'redo' : 'repeat'
  return { checked, disabled, ...(repeat ? { repeat } : {}) }
}
