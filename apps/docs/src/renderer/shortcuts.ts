/**
 * The single list of user-facing Docs keyboard shortcuts.
 *
 * Bindings themselves still live where they must fire — the main-process menu
 * accelerators, App.tsx's window handler and the editor keymap — but every one
 * of them is described here once, so the Keyboard Shortcuts sheet, and the test
 * that guards it against drift, read from one place.
 *
 * Chords are written in Mac notation and rewritten to Ctrl/Alt/Shift form on
 * Windows/Linux, exactly like the ribbon's screen tips. `win` is only set where
 * the two platforms genuinely use different keys (macOS eats ⌘H / ⌘M / ⌘Q for
 * Hide / Minimize / Quit, and Word itself differs on a few chords).
 */
import { macShortcutsToWin, platformShortcuts } from '@genoffice/i18n'
import type { ActionCommandId } from '../shared/commands'
import type { StringKey } from './i18n/locale'

export type ShortcutGroupId =
  'file' | 'edit' | 'nav' | 'text' | 'para' | 'insert' | 'review' | 'view'

export interface ShortcutDef {
  id: string
  /** the command the chord runs; without one the binding is built in and cannot be reassigned */
  cmd?: ActionCommandId
  payload?: string
  group: ShortcutGroupId
  /** row label; reuses the command's existing translation wherever one exists */
  labelKey: StringKey
  /** appended verbatim to the label (line-spacing values and the like) */
  labelSuffix?: string
  /** Mac notation, e.g. '⇧⌘S'; alternates are listed as 'A / B' */
  keys: string
  /** Windows/Linux chord when it differs from a notation rewrite of `keys` */
  win?: string
  /** the chords of `keys` the document key table runs on the Mac; the others have handlers of their own */
  docKeys?: string
}

/** in `keys` or `win`: the platform has no chord for the row, which is then not listed there */
const NO_CHORD = ''

/** a chord of Word for Mac that Word for Windows does not share */
const mac = (
  id: string,
  cmd: ActionCommandId,
  group: ShortcutGroupId,
  labelKey: StringKey,
  keys: string,
): ShortcutDef => ({ id, cmd, group, labelKey, keys, win: NO_CHORD })

/** a chord of Word for Mac the document key table runs */
const word = (
  id: string,
  cmd: ActionCommandId,
  group: ShortcutGroupId,
  labelKey: StringKey,
  keys: string,
  payload?: string,
): ShortcutDef => ({ id, cmd, payload, group, labelKey, keys, win: NO_CHORD, docKeys: keys })

/** a caret or selection move: the chords Chromium already answers, then the ones Word adds */
const move = (
  id: string,
  labelKey: StringKey,
  native: string,
  added: string,
  payload: string,
): ShortcutDef => ({
  id,
  cmd: 'caret-move',
  payload,
  group: 'nav',
  labelKey,
  keys: native ? `${native} / ${added}` : added,
  win: NO_CHORD,
  docKeys: added,
})

export const SHORTCUT_GROUPS: readonly { id: ShortcutGroupId; labelKey: StringKey }[] = [
  { id: 'file', labelKey: 'appScGroupFile' },
  { id: 'edit', labelKey: 'appScGroupEdit' },
  { id: 'nav', labelKey: 'appScGroupNav' },
  { id: 'text', labelKey: 'appScGroupText' },
  { id: 'para', labelKey: 'appScGroupPara' },
  { id: 'insert', labelKey: 'appScGroupInsert' },
  { id: 'review', labelKey: 'appScGroupReview' },
  { id: 'view', labelKey: 'appScGroupView' },
]

export const SHORTCUTS: readonly ShortcutDef[] = [
  // ---- File ----
  {
    id: 'new',
    cmd: 'new',
    group: 'file',
    labelKey: 'appNew',
    keys: '⌘N / ⌃N',
    win: 'Ctrl+N',
    docKeys: '⌃N',
  },
  // Word for Mac has no chord for a new window, and ⇧⌘N is its Normal style
  {
    id: 'new-window',
    group: 'file',
    labelKey: 'appScNewWindow',
    keys: NO_CHORD,
    win: 'Ctrl+Shift+N',
  },
  { id: 'open', cmd: 'open', group: 'file', labelKey: 'ribbonOpen', keys: '⌘O / ⇧⌘O' },
  { id: 'save', cmd: 'save', group: 'file', labelKey: 'appSave', keys: '⌘S / ⌃S', win: 'Ctrl+S' },
  { id: 'save-as', cmd: 'save-as', group: 'file', labelKey: 'ribbonSaveAs', keys: '⇧⌘S' },
  {
    id: 'print',
    cmd: 'print',
    group: 'file',
    labelKey: 'appPrintTitle',
    keys: '⌘P / ⌃P',
    win: 'Ctrl+P',
    docKeys: '⌃P',
  },
  mac('new-from-template', 'file-new-from-template', 'file', 'appTemplateGalleryTitle', '⇧⌘P'),
  { id: 'close', group: 'file', labelKey: 'appClose', keys: '⌘W' },

  // ---- Edit ----
  {
    id: 'undo',
    cmd: 'undo',
    group: 'edit',
    labelKey: 'appUndo',
    keys: '⌘Z / ⌃Z',
    win: 'Ctrl+Z',
    docKeys: '⌃Z',
  },
  { id: 'redo', cmd: 'redo', group: 'edit', labelKey: 'appRedo', keys: '⇧⌘Z' },
  {
    id: 'repeat',
    cmd: 'repeat',
    group: 'edit',
    labelKey: 'appScRepeat',
    keys: '⌘Y / ⌃Y / ⌥⏎',
    win: 'Ctrl+Y',
    docKeys: '⌃Y / ⌥⏎',
  },
  {
    id: 'cut',
    cmd: 'edit-text',
    payload: 'cut',
    group: 'edit',
    labelKey: 'appCut',
    keys: '⌘X / ⌃X / F2',
    win: 'Ctrl+X',
    docKeys: '⌃X / F2',
  },
  {
    id: 'copy',
    cmd: 'edit-text',
    payload: 'copy',
    group: 'edit',
    labelKey: 'appCopy',
    keys: '⌘C / ⌃C / F3',
    win: 'Ctrl+C',
    docKeys: '⌃C / F3',
  },
  {
    id: 'paste',
    cmd: 'edit-text',
    payload: 'paste',
    group: 'edit',
    labelKey: 'appPaste',
    keys: '⌘V / ⌃V / F4',
    win: 'Ctrl+V',
    docKeys: '⌃V / F4',
  },
  // Electron's pasteAndMatchStyle role: ⌥⇧⌘V on macOS, Ctrl+Shift+V elsewhere
  {
    id: 'paste-plain',
    group: 'edit',
    labelKey: 'appPastePlain',
    keys: '⌥⇧⌘V',
    win: 'Ctrl+Shift+V',
  },
  word('paste-text', 'paste-as', 'edit', 'appPasteTextOnly', '⇧⌘V', 'text'),
  mac('paste-special', 'edit-paste-special', 'edit', 'ribbonPasteSpecialTitle', '⌃⌘V'),
  {
    id: 'select-all',
    cmd: 'edit-text',
    payload: 'select-all',
    group: 'edit',
    labelKey: 'appScSelectAll',
    keys: '⌘A / ⌃A',
    win: 'Ctrl+A',
    docKeys: '⌃A',
  },
  {
    id: 'find',
    cmd: 'find',
    group: 'edit',
    labelKey: 'appFindPlaceholder',
    keys: '⌘F / ⌃F',
    win: 'Ctrl+F',
  },
  mac('find-next', 'edit-find-next', 'edit', 'ribbonFindNext', '⌘G / ⇧F4 / ⇧⌘F4'),
  mac('find-previous', 'edit-find-previous', 'edit', 'appScFindPrevious', '⇧⌘G'),
  // ⌘H is macOS' Hide role and never reaches the window
  {
    id: 'replace',
    cmd: 'replace',
    group: 'edit',
    labelKey: 'appReplace',
    keys: '⇧⌘H / ⌃H',
    win: 'Ctrl+H',
  },
  // Word for Mac's Go To; Windows Word uses Ctrl+G
  {
    id: 'goto',
    cmd: 'goto',
    group: 'edit',
    labelKey: 'appGoToTitle',
    keys: '⌥⌘G / F5',
    win: 'Ctrl+G',
    docKeys: 'F5',
  },
  { id: 'shortcuts', cmd: 'shortcuts', group: 'edit', labelKey: 'appScTitle', keys: '⌘/' },
  mac('toggle-drawing', 'edit-toggle-drawing', 'edit', 'appScToggleDrawing', '⌃⌘Z'),
  mac('delete-rows', 'table-delete-rows', 'edit', 'ribbonDeleteRow', '⌃⌘X'),
  word('insert-row', 'table-insert-rows-above', 'edit', 'ribbonInsertAbove', '⌃⌘I'),
  {
    id: 'delete-word-back',
    cmd: 'delete-word',
    payload: 'back',
    group: 'edit',
    labelKey: 'appScDeleteWordBack',
    keys: '⌥⌫ / ⌘⌫',
    win: NO_CHORD,
    docKeys: '⌘⌫',
  },
  {
    id: 'delete-word-forward',
    cmd: 'delete-word',
    payload: 'forward',
    group: 'edit',
    labelKey: 'appScDeleteWordForward',
    keys: '⌥⌦ / ⌘⌦',
    win: NO_CHORD,
    docKeys: '⌘⌦',
  },
  word('spike-cut', 'spike', 'edit', 'appScSpikeCut', '⌘F3', 'cut'),
  word('spike-insert', 'spike', 'edit', 'appScSpikeInsert', '⇧⌘F3', 'insert'),
  word('clear-contents', 'edit-clear-contents', 'edit', 'appScClearContents', 'Clear'),
  word('copy-text', 'copy-text', 'edit', 'appScCopyText', '⇧F2'),

  // ---- Navigation and selection ----
  move('line-start', 'appScLineStart', '⌘←', 'Home / ⌃↑', 'line-start'),
  move('line-end', 'appScLineEnd', '⌘→', 'End / ⌃↓', 'line-end'),
  move('para-up', 'appScParaUp', '', '⌘↑ / ⌥↑', 'para-up'),
  move('para-down', 'appScParaDown', '', '⌘↓ / ⌥↓', 'para-down'),
  move('doc-start', 'appScDocStart', '', '⌘Home / ⌥Home', 'doc-start'),
  move('doc-end', 'appScDocEnd', '', '⌘End / ⌥End', 'doc-end'),
  move('select-line-start', 'appScSelLineStart', '⇧⌘←', '⇧Home', 'line-start-extend'),
  move('select-line-end', 'appScSelLineEnd', '⇧⌘→', '⇧End', 'line-end-extend'),
  move('select-para-up', 'appScSelParaUp', '', '⇧⌘↑ / ⌥⇧↑', 'para-up-extend'),
  move('select-para-down', 'appScSelParaDown', '', '⇧⌘↓ / ⌥⇧↓', 'para-down-extend'),
  move('select-doc-start', 'appScSelDocStart', '', '⇧⌘Home', 'doc-start-extend'),
  move('select-doc-end', 'appScSelDocEnd', '', '⇧⌘End', 'doc-end-extend'),
  move('screen-up', 'appScScreenUp', '', 'PageUp', 'screen-up'),
  move('screen-down', 'appScScreenDown', '', 'PageDown', 'screen-down'),
  move('select-screen-up', 'appScSelScreenUp', '', '⇧PageUp', 'screen-up-extend'),
  move('select-screen-down', 'appScSelScreenDown', '', '⇧PageDown', 'screen-down-extend'),
  move('window-start', 'appScWindowStart', '', '⌥⌘PageUp', 'window-start'),
  move('window-end', 'appScWindowEnd', '', '⌥⌘PageDown', 'window-end'),
  move('select-window-start', 'appScSelWindowStart', '', '⌥⇧⌘PageUp', 'window-start-extend'),
  move('select-window-end', 'appScSelWindowEnd', '', '⌥⇧⌘PageDown', 'window-end-extend'),
  word('browse-prev', 'browse', 'nav', 'appScBrowsePrev', '⌘PageUp', 'prev'),
  word('browse-next', 'browse', 'nav', 'appScBrowseNext', '⌘PageDown', 'next'),
  word('section-prev', 'section-go', 'nav', 'appScPrevSection', '⌥PageUp', 'prev'),
  word('section-next', 'section-go', 'nav', 'appScNextSection', '⌥PageDown', 'next'),
  word('extend-selection', 'extend-selection', 'nav', 'appScExtendSelection', 'F8', 'grow'),
  word('shrink-selection', 'extend-selection', 'nav', 'appScShrinkSelection', '⇧F8', 'shrink'),
  word('go-back', 'go-back', 'nav', 'appScGoBack', '⇧F5 / ⌥⌘Z'),
  word('row-start', 'table-cell-go', 'nav', 'appScRowStart', '⌃Home / ⌥⇧Home', 'row-start'),
  word('row-end', 'table-cell-go', 'nav', 'appScRowEnd', '⌃End / ⌥⇧End', 'row-end'),
  word(
    'column-start',
    'table-cell-go',
    'nav',
    'appScColumnStart',
    '⌃PageUp / ⌥⇧PageUp',
    'column-start',
  ),
  word(
    'column-end',
    'table-cell-go',
    'nav',
    'appScColumnEnd',
    '⌃PageDown / ⌥⇧PageDown',
    'column-end',
  ),
  word('select-table', 'table-select-table', 'nav', 'ribbonSelectTable', '⌥Clear'),
  // Word for Mac has no chord for these four: they are there to be given one
  ...(
    [
      ['row-start', 'appScSelRowStart'],
      ['row-end', 'appScSelRowEnd'],
      ['column-start', 'appScSelColumnStart'],
      ['column-end', 'appScSelColumnEnd'],
    ] as const
  ).map(([to, labelKey]): ShortcutDef => ({
    id: `select-${to}`,
    cmd: 'table-cell-go',
    payload: `${to}-extend`,
    group: 'nav',
    labelKey,
    keys: NO_CHORD,
    win: NO_CHORD,
  })),

  // ---- Text formatting ----
  {
    id: 'bold',
    cmd: 'bold',
    group: 'text',
    labelKey: 'appFontBold',
    keys: '⌘B / ⇧⌘B / ⌃B / ⌃⇧B',
    win: 'Ctrl+B',
    docKeys: '⇧⌘B / ⌃B / ⌃⇧B',
  },
  {
    id: 'italic',
    cmd: 'italic',
    group: 'text',
    labelKey: 'appFontItalic',
    keys: '⌘I / ⇧⌘I / ⌃I / ⌃⇧I',
    win: 'Ctrl+I',
    docKeys: '⇧⌘I / ⌃I / ⌃⇧I',
  },
  {
    id: 'underline',
    cmd: 'underline',
    group: 'text',
    labelKey: 'appUnderline',
    keys: '⌘U / ⇧⌘U / ⌃U / ⌃⇧U',
    win: 'Ctrl+U',
    docKeys: '⇧⌘U / ⌃U / ⌃⇧U',
  },
  word('double-underline', 'double-underline', 'text', 'ribbonTcMarkDoubleUnderline', '⇧⌘D'),
  word('word-underline', 'word-underline', 'text', 'appScWordUnderline', '⇧⌘W'),
  word('strikethrough', 'strikethrough', 'text', 'ribbonStrikethrough', '⇧⌘X'),
  word('all-caps', 'all-caps', 'text', 'appAllCaps', '⇧⌘A'),
  word('small-caps', 'small-caps', 'text', 'appSmallCaps', '⇧⌘K'),
  word('hidden-text', 'hidden-text', 'text', 'appHiddenText', '⌃⇧H'),
  {
    id: 'font-dialog',
    cmd: 'font-dialog',
    group: 'text',
    labelKey: 'appFontDialogTitle',
    keys: '⌘D',
  },
  { id: 'grow-font', cmd: 'font-grow', group: 'text', labelKey: 'ribbonGrowFont', keys: '⇧⌘.' },
  {
    id: 'shrink-font',
    cmd: 'font-shrink',
    group: 'text',
    labelKey: 'ribbonShrinkFont',
    keys: '⇧⌘,',
  },
  { id: 'grow-font-1pt', group: 'text', labelKey: 'appScGrowFont1', keys: '⌘]' },
  { id: 'shrink-font-1pt', group: 'text', labelKey: 'appScShrinkFont1', keys: '⌘[' },
  // Word for Mac: ⌘. is Cancel and ⌘, Preferences, so the unshifted pair is Windows/Linux only
  {
    id: 'superscript',
    cmd: 'superscript',
    group: 'text',
    labelKey: 'ribbonSuperscript',
    keys: '⇧⌘=',
    win: 'Ctrl+Shift+= / Ctrl+.',
  },
  {
    id: 'subscript',
    cmd: 'subscript',
    group: 'text',
    labelKey: 'ribbonSubscript',
    keys: '⇧⌘-',
    win: 'Ctrl+,',
  },
  { id: 'change-case', group: 'text', labelKey: 'ribbonChangeCase', keys: '⇧F3' },
  {
    id: 'copy-format',
    group: 'text',
    labelKey: 'ribbonCopyFormat',
    keys: '⌥⌘C / ⇧⌘C',
    win: 'Ctrl+Shift+C',
  },
  // Ctrl+Shift+V is Electron's paste-and-match-style accelerator off the Mac
  {
    id: 'paste-format',
    group: 'text',
    labelKey: 'ribbonPasteFormat',
    keys: '⌥⌘V',
    win: 'Ctrl+Alt+Shift+V',
  },
  { id: 'clear-formatting', group: 'text', labelKey: 'ribbonClearFormatting', keys: '⌃␣' },

  // ---- Paragraph formatting ----
  { id: 'align-left', cmd: 'align-left', group: 'para', labelKey: 'appScAlignLeft', keys: '⌘L' },
  {
    id: 'align-center',
    cmd: 'align-center',
    group: 'para',
    labelKey: 'appScAlignCenter',
    keys: '⌘E',
  },
  { id: 'align-right', cmd: 'align-right', group: 'para', labelKey: 'appScAlignRight', keys: '⌘R' },
  {
    id: 'align-justify',
    cmd: 'align-justify',
    group: 'para',
    labelKey: 'appScJustify',
    keys: '⌘J',
  },
  {
    id: 'spacing-1',
    cmd: 'line-spacing',
    payload: '1',
    group: 'para',
    labelKey: 'ribbonLineSpacing',
    labelSuffix: ' 1.0',
    keys: '⌘1',
  },
  {
    id: 'spacing-15',
    cmd: 'line-spacing',
    payload: '1.5',
    group: 'para',
    labelKey: 'ribbonLineSpacing',
    labelSuffix: ' 1.5',
    keys: '⌘5',
  },
  {
    id: 'spacing-2',
    cmd: 'line-spacing',
    payload: '2',
    group: 'para',
    labelKey: 'ribbonLineSpacing',
    labelSuffix: ' 2.0',
    keys: '⌘2',
  },
  // Word for Mac's pair is ⌃⇧M / ⇧⌘M (⌘M minimizes the window); Word for Windows' is Ctrl+M / Ctrl+Shift+M
  {
    id: 'indent',
    cmd: 'indent-increase',
    group: 'para',
    labelKey: 'ribbonIncreaseIndent',
    keys: '⌃⇧M / ⌃M',
    win: 'Ctrl+M',
  },
  {
    id: 'outdent',
    cmd: 'indent-decrease',
    group: 'para',
    labelKey: 'ribbonDecreaseIndent',
    keys: '⇧⌘M',
    win: 'Ctrl+Shift+M',
  },
  { id: 'hanging-indent', group: 'para', labelKey: 'appScHangingIndent', keys: '⌘T' },
  { id: 'hanging-indent-off', group: 'para', labelKey: 'appScHangingIndentOff', keys: '⇧⌘T' },
  // ⌘Q quits on macOS
  { id: 'clear-para', group: 'para', labelKey: 'appScClearPara', keys: '⌥⌘Q / ⌃Q', win: 'Ctrl+Q' },
  {
    id: 'style-normal',
    cmd: 'style-normal',
    group: 'para',
    labelKey: 'ribbonStyleNormal',
    keys: '⇧⌘N / ⌥⌘0 / ⌥⇧Clear',
    win: 'Ctrl+Alt+0',
    docKeys: '⌥⇧Clear',
  },
  { id: 'style-h1', group: 'para', labelKey: 'ribbonStyleHeading1', keys: '⌥⌘1' },
  { id: 'style-h2', group: 'para', labelKey: 'ribbonStyleHeading2', keys: '⌥⌘2' },
  { id: 'style-h3', group: 'para', labelKey: 'ribbonStyleHeading3', keys: '⌥⌘3' },
  {
    id: 'paragraph-dialog',
    cmd: 'paragraph-dialog',
    group: 'para',
    labelKey: 'appParagraph',
    keys: '⌥⌘M',
  },
  word('list-bullet', 'bullets', 'para', 'ribbonBullets', '⇧⌘L', 'on'),
  // Word for Mac moves a paragraph on ⌃⇧↑ / ⌃⇧↓; ⌥⇧↑ / ⌥⇧↓ extend the selection there
  { id: 'move-up', group: 'para', labelKey: 'appScMoveUp', keys: '⌃⇧↑', win: 'Alt+Shift+Up' },
  { id: 'move-down', group: 'para', labelKey: 'appScMoveDown', keys: '⌃⇧↓', win: 'Alt+Shift+Down' },
  word('outline-promote', 'outline-level', 'para', 'appScPromoteHeading', '⌃⇧←', 'promote'),
  word('outline-demote', 'outline-level', 'para', 'appScDemoteHeading', '⌃⇧→', 'demote'),
  word('outline-expand', 'outline-fold', 'para', 'appScOutlineExpand', '⌃⇧= / ⌃⇧Num+', 'expand'),
  word(
    'outline-collapse',
    'outline-fold',
    'para',
    'appScOutlineCollapse',
    '⌃⇧- / ⌃⇧Num-',
    'collapse',
  ),
  ...Array.from({ length: 9 }, (_, i): ShortcutDef => ({
    ...word(`outline-show-${i + 1}`, 'outline-show', 'para', 'appScOutlineShowLevel', `⌃⇧${i + 1}`),
    payload: String(i + 1),
    labelSuffix: ` ${i + 1}`,
  })),
  word('outline-show-all', 'outline-show', 'para', 'appScOutlineShowAll', '⌃⇧A', 'all'),
  { id: 'list-demote', group: 'para', labelKey: 'appScListDemote', keys: 'Tab' },
  {
    id: 'list-promote',
    group: 'para',
    labelKey: 'appScListPromote',
    keys: '⇧Tab',
    win: 'Shift+Tab',
  },

  // ---- Insert ----
  {
    id: 'page-break',
    cmd: 'insert-page-break',
    group: 'insert',
    labelKey: 'ribbonPageBreak',
    keys: '⌘⏎',
  },
  {
    id: 'column-break',
    cmd: 'insert-column-break',
    group: 'insert',
    labelKey: 'appScColumnBreak',
    keys: '⇧⌘⏎',
  },
  { id: 'line-break', group: 'insert', labelKey: 'appScLineBreak', keys: '⇧⏎' },
  { id: 'link', cmd: 'insert-link', group: 'insert', labelKey: 'ribbonLink', keys: '⌘K' },
  {
    id: 'comment',
    cmd: 'insert-comment',
    group: 'insert',
    labelKey: 'ribbonNewComment',
    keys: '⌥⌘A',
  },
  {
    id: 'footnote',
    cmd: 'insert-footnote',
    group: 'insert',
    labelKey: 'ribbonFootnote',
    keys: '⌥⌘F',
  },
  // Word itself splits here: ⌥⌘E on the Mac (⌥⌘D toggles the Dock), Alt+Ctrl+D on Windows
  {
    id: 'endnote',
    cmd: 'insert-endnote',
    group: 'insert',
    labelKey: 'ribbonEndnote',
    keys: '⌥⌘E',
    win: 'Ctrl+Alt+D',
  },
  {
    id: 'field-date',
    cmd: 'insert-field',
    payload: 'DATE',
    group: 'insert',
    labelKey: 'ribbonFieldDate',
    keys: '⌃⇧D / ⌥⇧D',
    win: 'Alt+Shift+D',
    docKeys: '⌃⇧D',
  },
  {
    id: 'field-time',
    cmd: 'insert-field',
    payload: 'TIME',
    group: 'insert',
    labelKey: 'ribbonFieldTime',
    keys: '⌃⇧T / ⌥⇧T',
    win: 'Alt+Shift+T',
    docKeys: '⌃⇧T',
  },
  word('field-page', 'insert-field', 'insert', 'ribbonFieldPage', '⌃⇧P', 'PAGE'),
  word('field-chars', 'field-insert-chars', 'insert', 'appScFieldChars', '⌘F9'),
  word('hf-link-previous', 'hf-link-previous', 'insert', 'ribbonHfLinkPrev', '⌃⇧R'),
  word('merge-field', 'insert-merge-field', 'insert', 'ribbonMailingsInsertField', '⌃⇧F'),
  word('mark-toc-entry', 'mark-toc-entry', 'insert', 'appScMarkTocEntry', '⌥⇧⌘O'),
  // the keypad's Enter; the main block's is the page break above
  word('section-break', 'insert-section-break', 'insert', 'appScSectionBreak', '⌘⌤', 'nextPage'),
  word('en-dash', 'insert-symbol', 'insert', 'ribbonSpEnDash', '⌘Num-', '\u2013'),
  word('em-dash', 'insert-symbol', 'insert', 'ribbonSpEmDash', '⌥⌘Num-', '\u2014'),
  mac('equation', 'insert-equation', 'insert', 'appScEquation', '⌃='),
  word('bookmark', 'insert-bookmark', 'insert', 'ribbonBookmark', '⇧⌘F5'),
  word('autotext-new', 'insert-autotext-new', 'insert', 'ribbonAutoTextNew', '⌥F3'),
  word('mark-index-entry', 'mark-index-entry', 'insert', 'ribbonMarkEntryTitle', '⌥⇧⌘X'),
  word('mark-citation', 'toa-mark-citation', 'insert', 'ribbonMarkCitation', '⌥⇧⌘I'),
  {
    id: 'nbsp',
    group: 'insert',
    labelKey: 'appScNbsp',
    keys: '⌃⇧␣ / ⇧⌘␣',
    win: 'Ctrl+Shift+Space',
  },
  // Word for Mac's ⇧⌘- is Subscript
  {
    id: 'nb-hyphen',
    group: 'insert',
    labelKey: 'appScNbHyphen',
    keys: '⌥⇧⌘-',
    win: 'Ctrl+Shift+-',
  },
  { id: 'opt-hyphen', group: 'insert', labelKey: 'ribbonSpOptHyphen', keys: '⌥⌘-', win: NO_CHORD },
  { id: 'tab-char', group: 'insert', labelKey: 'appScTabChar', keys: '⌃Tab / ⌥Tab', win: NO_CHORD },

  // ---- Review & tools ----
  {
    id: 'track-changes',
    cmd: 'track-changes',
    group: 'review',
    labelKey: 'ribbonTrackChanges',
    keys: '⇧⌘E',
  },
  // Word for Windows' chord; Word for Mac has none, and ⇧⌘G is its Find Previous
  {
    id: 'word-count',
    cmd: 'word-count',
    group: 'review',
    labelKey: 'appWordCountTitle',
    keys: NO_CHORD,
    win: 'Ctrl+Shift+G',
  },
  mac('thesaurus', 'tools-thesaurus', 'review', 'ribbonThesaurus', '⌃⌥⌘R / ⇧F7'),
  {
    id: 'tools-spelling',
    cmd: 'tools-spelling',
    group: 'review',
    labelKey: 'appSpellingTitle',
    keys: '⌥⌘L / F7',
    win: 'Ctrl+Alt+L',
    docKeys: 'F7',
  },
  word('next-misspelling', 'spell-next', 'review', 'appScNextMisspelling', '⌥F7'),
  // F7 is Word for Mac's Spelling and Grammar
  {
    id: 'proofread',
    cmd: 'ai-proofread',
    group: 'review',
    labelKey: 'appScProofread',
    keys: NO_CHORD,
    win: 'F7',
  },
  word('translate-selection', 'tools-translate-selection', 'review', 'ribbonTranslate', '⌃⌥⌘T'),
  {
    id: 'update-fields',
    group: 'review',
    labelKey: 'appUpdateField',
    keys: 'F9 / ⌥⇧⌘U',
    win: 'F9',
  },
  { id: 'toggle-field-codes', group: 'review', labelKey: 'appToggleFieldCodes', keys: '⌥F9' },
  {
    id: 'field-code-here',
    group: 'review',
    labelKey: 'appScFieldCodeHere',
    keys: '⇧F9',
    win: NO_CHORD,
  },
  word('field-lock', 'field-lock', 'review', 'appScLockFields', '⌘F11 / ⌘3', 'lock'),
  word('field-unlock', 'field-lock', 'review', 'appScUnlockFields', '⇧⌘F11 / ⌘4', 'unlock'),
  word('field-unlink', 'field-unlink', 'review', 'appScUnlinkFields', '⇧⌘F9 / ⌘6'),
  word('field-next', 'field-go', 'review', 'appScNextField', 'F11', 'next'),
  word('field-prev', 'field-go', 'review', 'appScPrevField', '⇧F11', 'prev'),
  word('edit-recipients', 'mailings-edit-recipients', 'review', 'ribbonEditRecipientList', '⌃⇧E'),
  word('finish-merge', 'mailings-finish-merge', 'review', 'ribbonFinishMerge', '⌃⇧N'),

  // ---- View ----
  { id: 'zoom-in', cmd: 'zoom-in', group: 'view', labelKey: 'ribbonZoomIn', keys: '⌘=' },
  { id: 'zoom-out', cmd: 'zoom-out', group: 'view', labelKey: 'ribbonZoomOut', keys: '⌘-' },
  { id: 'zoom-100', cmd: 'zoom-100', group: 'view', labelKey: 'ribbonZoom100Tip', keys: '⌘0' },
  // Word for Mac uses ⌘8 for the ¶ toggle, Windows Ctrl+Shift+8; both work
  {
    id: 'formatting-marks',
    cmd: 'show-marks',
    group: 'view',
    labelKey: 'ribbonShowMarks',
    keys: '⌘8',
    win: 'Ctrl+Shift+8',
  },
  word('print-layout', 'view-print-layout', 'view', 'ribbonPrintLayout', '⌥⌘P'),
  word('outline', 'view-outline', 'view', 'ribbonOutlineView', '⌥⌘O'),
  word('draft', 'view-draft', 'view', 'ribbonDraft', '⌥⌘N'),
  word('split', 'split-view', 'view', 'ribbonSplit', '⌥⌘S'),
  word('close-pane', 'close-pane', 'view', 'appScClosePane', '⌃⇧C'),
  word('window-next', 'window-go', 'view', 'appScNextWindow', '⌘F6 / ⌥F6', 'next'),
  word('window-prev', 'window-go', 'view', 'appScPrevWindow', '⇧⌘F6 / ⌥⇧F6', 'prev'),
  word('styles-pane', 'styles-pane', 'view', 'ribbonStylesPane', '⌥⇧⌘S'),
  mac('focus', 'view-focus', 'view', 'ribbonFocus', '⌃⇧⌘F'),
  mac('ribbon', 'view-ribbon', 'view', 'appScRibbon', '⌥⌘R'),
  mac('format-object', 'format-object', 'view', 'appScFormatObject', '⇧⌘1'),
  mac('selection-pane', 'format-selection-pane', 'view', 'ribbonSelectionPane', '⌥⌘U'),
]

/** platformShortcuts is the identity on macOS and rewrites the notation elsewhere */
const IS_MAC = platformShortcuts('⌘') === '⌘'

/** the chord as this platform's user sees it */
export function shortcutKeys(def: ShortcutDef, isMac = IS_MAC): string {
  if (isMac) return def.keys
  return def.win ?? macShortcutsToWin(def.keys)
}
