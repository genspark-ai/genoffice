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

export type ShortcutGroupId = 'file' | 'edit' | 'text' | 'para' | 'insert' | 'review' | 'view'

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
}

export const SHORTCUT_GROUPS: readonly { id: ShortcutGroupId; labelKey: StringKey }[] = [
  { id: 'file', labelKey: 'appScGroupFile' },
  { id: 'edit', labelKey: 'appScGroupEdit' },
  { id: 'text', labelKey: 'appScGroupText' },
  { id: 'para', labelKey: 'appScGroupPara' },
  { id: 'insert', labelKey: 'appScGroupInsert' },
  { id: 'review', labelKey: 'appScGroupReview' },
  { id: 'view', labelKey: 'appScGroupView' },
]

export const SHORTCUTS: readonly ShortcutDef[] = [
  // ---- File ----
  { id: 'new', cmd: 'new', group: 'file', labelKey: 'appNew', keys: '⌘N' },
  { id: 'new-window', group: 'file', labelKey: 'appScNewWindow', keys: '⇧⌘N' },
  { id: 'open', cmd: 'open', group: 'file', labelKey: 'ribbonOpen', keys: '⌘O / ⇧⌘O' },
  { id: 'save', cmd: 'save', group: 'file', labelKey: 'appSave', keys: '⌘S' },
  { id: 'save-as', cmd: 'save-as', group: 'file', labelKey: 'ribbonSaveAs', keys: '⇧⌘S' },
  { id: 'print', cmd: 'print', group: 'file', labelKey: 'appPrintTitle', keys: '⌘P' },
  { id: 'close', group: 'file', labelKey: 'appClose', keys: '⌘W' },

  // ---- Edit ----
  { id: 'undo', cmd: 'undo', group: 'edit', labelKey: 'appUndo', keys: '⌘Z' },
  { id: 'redo', cmd: 'redo', group: 'edit', labelKey: 'appRedo', keys: '⇧⌘Z / ⌘Y' },
  { id: 'cut', group: 'edit', labelKey: 'appCut', keys: '⌘X' },
  { id: 'copy', group: 'edit', labelKey: 'appCopy', keys: '⌘C' },
  { id: 'paste', group: 'edit', labelKey: 'appPaste', keys: '⌘V' },
  // Electron's pasteAndMatchStyle role: ⌥⇧⌘V on macOS, Ctrl+Shift+V elsewhere
  {
    id: 'paste-plain',
    group: 'edit',
    labelKey: 'appPastePlain',
    keys: '⌥⇧⌘V',
    win: 'Ctrl+Shift+V',
  },
  { id: 'select-all', group: 'edit', labelKey: 'appScSelectAll', keys: '⌘A' },
  { id: 'find', cmd: 'find', group: 'edit', labelKey: 'appFindPlaceholder', keys: '⌘F' },
  // ⌘H is macOS' Hide role and never reaches the window
  {
    id: 'replace',
    cmd: 'replace',
    group: 'edit',
    labelKey: 'appReplace',
    keys: '⌃H',
    win: 'Ctrl+H',
  },
  // Word for Mac's Go To; Windows Word uses Ctrl+G
  { id: 'goto', cmd: 'goto', group: 'edit', labelKey: 'appGoToTitle', keys: '⌥⌘G', win: 'Ctrl+G' },
  { id: 'shortcuts', cmd: 'shortcuts', group: 'edit', labelKey: 'appScTitle', keys: '⌘/' },

  // ---- Text formatting ----
  { id: 'bold', cmd: 'bold', group: 'text', labelKey: 'appFontBold', keys: '⌘B' },
  { id: 'italic', cmd: 'italic', group: 'text', labelKey: 'appFontItalic', keys: '⌘I' },
  { id: 'underline', cmd: 'underline', group: 'text', labelKey: 'appUnderline', keys: '⌘U' },
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
  {
    id: 'superscript',
    cmd: 'superscript',
    group: 'text',
    labelKey: 'ribbonSuperscript',
    keys: '⇧⌘= / ⌘.',
  },
  { id: 'subscript', cmd: 'subscript', group: 'text', labelKey: 'ribbonSubscript', keys: '⌘,' },
  { id: 'change-case', group: 'text', labelKey: 'ribbonChangeCase', keys: '⇧F3' },
  { id: 'copy-format', group: 'text', labelKey: 'ribbonCopyFormat', keys: '⇧⌘C' },
  // Ctrl+Shift+V is Electron's paste-and-match-style accelerator off the Mac
  {
    id: 'paste-format',
    group: 'text',
    labelKey: 'ribbonPasteFormat',
    keys: '⇧⌘V',
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
  // ⌘M is macOS' Minimize accelerator, so the indent chord is Control+M there
  // (Windows/Linux: plain Ctrl+M)
  {
    id: 'indent',
    cmd: 'indent-increase',
    group: 'para',
    labelKey: 'ribbonIncreaseIndent',
    keys: '⌃M',
    win: 'Ctrl+M',
  },
  {
    id: 'outdent',
    cmd: 'indent-decrease',
    group: 'para',
    labelKey: 'ribbonDecreaseIndent',
    keys: '⌃⇧M',
    win: 'Ctrl+Shift+M',
  },
  { id: 'hanging-indent', group: 'para', labelKey: 'appScHangingIndent', keys: '⌘T' },
  { id: 'hanging-indent-off', group: 'para', labelKey: 'appScHangingIndentOff', keys: '⇧⌘T' },
  // ⌘Q quits on macOS
  { id: 'clear-para', group: 'para', labelKey: 'appScClearPara', keys: '⌃Q', win: 'Ctrl+Q' },
  { id: 'style-normal', group: 'para', labelKey: 'ribbonStyleNormal', keys: '⌥⌘0' },
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
  { id: 'move-up', group: 'para', labelKey: 'appScMoveUp', keys: '⌥⇧↑' },
  { id: 'move-down', group: 'para', labelKey: 'appScMoveDown', keys: '⌥⇧↓' },
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
    keys: '⌥⇧D',
  },
  {
    id: 'field-time',
    cmd: 'insert-field',
    payload: 'TIME',
    group: 'insert',
    labelKey: 'ribbonFieldTime',
    keys: '⌥⇧T',
  },
  { id: 'nbsp', group: 'insert', labelKey: 'appScNbsp', keys: '⇧⌘␣' },
  { id: 'nb-hyphen', group: 'insert', labelKey: 'appScNbHyphen', keys: '⇧⌘-' },

  // ---- Review & tools ----
  {
    id: 'track-changes',
    cmd: 'track-changes',
    group: 'review',
    labelKey: 'ribbonTrackChanges',
    keys: '⇧⌘E',
  },
  {
    id: 'word-count',
    cmd: 'word-count',
    group: 'review',
    labelKey: 'appWordCountTitle',
    keys: '⇧⌘G',
  },
  {
    id: 'tools-spelling',
    cmd: 'tools-spelling',
    group: 'review',
    labelKey: 'appSpellingTitle',
    keys: '⌥⌘L',
  },
  { id: 'proofread', cmd: 'ai-proofread', group: 'review', labelKey: 'appScProofread', keys: 'F7' },
  { id: 'update-fields', group: 'review', labelKey: 'appUpdateField', keys: 'F9' },
  { id: 'toggle-field-codes', group: 'review', labelKey: 'appToggleFieldCodes', keys: '⌥F9' },

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
]

/** platformShortcuts is the identity on macOS and rewrites the notation elsewhere */
const IS_MAC = platformShortcuts('⌘') === '⌘'

/** the chord as this platform's user sees it */
export function shortcutKeys(def: ShortcutDef, isMac = IS_MAC): string {
  if (isMac) return def.keys
  return def.win ?? macShortcutsToWin(def.keys)
}
