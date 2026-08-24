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
import type { StringKey } from './i18n/locale'

export type ShortcutGroupId = 'file' | 'edit' | 'text' | 'para' | 'insert' | 'review' | 'view'

export interface ShortcutDef {
  id: string
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
  { id: 'new', group: 'file', labelKey: 'appNew', keys: '⌘N' },
  { id: 'new-window', group: 'file', labelKey: 'appScNewWindow', keys: '⇧⌘N' },
  { id: 'open', group: 'file', labelKey: 'ribbonOpen', keys: '⌘O' },
  { id: 'save', group: 'file', labelKey: 'appSave', keys: '⌘S' },
  { id: 'save-as', group: 'file', labelKey: 'ribbonSaveAs', keys: '⇧⌘S' },
  { id: 'print', group: 'file', labelKey: 'appPrintTitle', keys: '⌘P' },
  { id: 'close', group: 'file', labelKey: 'appClose', keys: '⌘W' },

  // ---- Edit ----
  { id: 'undo', group: 'edit', labelKey: 'appUndo', keys: '⌘Z' },
  { id: 'redo', group: 'edit', labelKey: 'appRedo', keys: '⇧⌘Z / ⌘Y' },
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
  { id: 'find', group: 'edit', labelKey: 'appFindPlaceholder', keys: '⌘F' },
  // ⌘H is macOS' Hide role and never reaches the window
  { id: 'replace', group: 'edit', labelKey: 'appReplace', keys: '⌃H', win: 'Ctrl+H' },
  { id: 'shortcuts', group: 'edit', labelKey: 'appScTitle', keys: '⌘/' },

  // ---- Text formatting ----
  { id: 'bold', group: 'text', labelKey: 'appFontBold', keys: '⌘B' },
  { id: 'italic', group: 'text', labelKey: 'appFontItalic', keys: '⌘I' },
  { id: 'underline', group: 'text', labelKey: 'appUnderline', keys: '⌘U' },
  { id: 'font-dialog', group: 'text', labelKey: 'appFontDialogTitle', keys: '⌘D' },
  { id: 'grow-font', group: 'text', labelKey: 'ribbonGrowFont', keys: '⇧⌘.' },
  { id: 'shrink-font', group: 'text', labelKey: 'ribbonShrinkFont', keys: '⇧⌘,' },
  { id: 'grow-font-1pt', group: 'text', labelKey: 'appScGrowFont1', keys: '⌘]' },
  { id: 'shrink-font-1pt', group: 'text', labelKey: 'appScShrinkFont1', keys: '⌘[' },
  { id: 'superscript', group: 'text', labelKey: 'ribbonSuperscript', keys: '⇧⌘= / ⌘.' },
  { id: 'subscript', group: 'text', labelKey: 'ribbonSubscript', keys: '⌘,' },
  { id: 'change-case', group: 'text', labelKey: 'ribbonChangeCase', keys: '⇧F3' },
  { id: 'clear-formatting', group: 'text', labelKey: 'ribbonClearFormatting', keys: '⌃␣' },

  // ---- Paragraph formatting ----
  { id: 'align-left', group: 'para', labelKey: 'appScAlignLeft', keys: '⌘L' },
  { id: 'align-center', group: 'para', labelKey: 'appScAlignCenter', keys: '⌘E' },
  { id: 'align-right', group: 'para', labelKey: 'appScAlignRight', keys: '⌘R' },
  { id: 'align-justify', group: 'para', labelKey: 'appScJustify', keys: '⌘J' },
  {
    id: 'spacing-1',
    group: 'para',
    labelKey: 'ribbonLineSpacing',
    labelSuffix: ' 1.0',
    keys: '⌘1',
  },
  {
    id: 'spacing-15',
    group: 'para',
    labelKey: 'ribbonLineSpacing',
    labelSuffix: ' 1.5',
    keys: '⌘5',
  },
  {
    id: 'spacing-2',
    group: 'para',
    labelKey: 'ribbonLineSpacing',
    labelSuffix: ' 2.0',
    keys: '⌘2',
  },
  // ⌘M is macOS' Minimize accelerator, so indenting is on ⌃M there too
  { id: 'indent', group: 'para', labelKey: 'ribbonIncreaseIndent', keys: '⌃M', win: 'Ctrl+M' },
  {
    id: 'outdent',
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
  { id: 'paragraph-dialog', group: 'para', labelKey: 'appParagraph', keys: '⌥⌘M' },
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
  { id: 'page-break', group: 'insert', labelKey: 'ribbonPageBreak', keys: '⌘⏎' },
  { id: 'column-break', group: 'insert', labelKey: 'appScColumnBreak', keys: '⇧⌘⏎' },
  { id: 'line-break', group: 'insert', labelKey: 'appScLineBreak', keys: '⇧⏎' },
  { id: 'link', group: 'insert', labelKey: 'ribbonLink', keys: '⌘K' },
  { id: 'comment', group: 'insert', labelKey: 'ribbonNewComment', keys: '⌥⌘A' },
  { id: 'footnote', group: 'insert', labelKey: 'ribbonFootnote', keys: '⌥⌘F' },
  // Word itself splits here: ⌥⌘E on the Mac (⌥⌘D toggles the Dock), Alt+Ctrl+D on Windows
  { id: 'endnote', group: 'insert', labelKey: 'ribbonEndnote', keys: '⌥⌘E', win: 'Ctrl+Alt+D' },
  { id: 'field-date', group: 'insert', labelKey: 'ribbonFieldDate', keys: '⌥⇧D' },
  { id: 'field-time', group: 'insert', labelKey: 'ribbonFieldTime', keys: '⌥⇧T' },
  { id: 'nbsp', group: 'insert', labelKey: 'appScNbsp', keys: '⇧⌘␣' },
  { id: 'nb-hyphen', group: 'insert', labelKey: 'appScNbHyphen', keys: '⇧⌘-' },

  // ---- Review & tools ----
  { id: 'track-changes', group: 'review', labelKey: 'ribbonTrackChanges', keys: '⇧⌘E' },
  { id: 'word-count', group: 'review', labelKey: 'appWordCountTitle', keys: '⇧⌘G' },
  { id: 'proofread', group: 'review', labelKey: 'appScProofread', keys: 'F7' },
  { id: 'update-fields', group: 'review', labelKey: 'appUpdateField', keys: 'F9' },

  // ---- View ----
  { id: 'zoom-in', group: 'view', labelKey: 'ribbonZoomIn', keys: '⌘=' },
  { id: 'zoom-out', group: 'view', labelKey: 'ribbonZoomOut', keys: '⌘-' },
  { id: 'zoom-100', group: 'view', labelKey: 'ribbonZoom100Tip', keys: '⌘0' },
  // Word for Mac uses ⌘8 for the ¶ toggle, Windows Ctrl+Shift+8; both work
  {
    id: 'formatting-marks',
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
