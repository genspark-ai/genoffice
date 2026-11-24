/**
 * Keyboard assignments the open document carries (its Customizations part):
 * which of them this app can run, and how they sit over the machine's own
 * assignments while the document is open.
 */
import {
  KCM_ALT,
  KCM_COMMAND,
  KCM_CONTROL,
  KCM_SHIFT,
  readKeymaps,
  writeKeymaps,
  type KeymapEntry,
} from '@genoffice/docx-engine'
import { toChord, type KeyOverrides } from '../shared/keybindings'
import { IS_MAC, assignChord, bindingCommand, chordOwner, removeChord } from './keybindings'

/** binding id -> the name Word gives the same fixed command ([MS-OE376] 3.1.2.1.2) */
const FCI_NAMES: Record<string, string> = {
  new: 'FileNewDefault',
  open: 'FileOpen',
  save: 'FileSave',
  'save-as': 'FileSaveAs',
  print: 'FilePrint',
  'page-preview': 'FilePrintPreview',
  'page-setup': 'FilePageSetup',
  'file-properties': 'FileProperties',
  undo: 'EditUndo',
  redo: 'EditRedo',
  repeat: 'EditRedoOrRepeat',
  'edit-paste-special': 'EditPasteSpecial',
  find: 'EditFind',
  replace: 'EditReplace',
  goto: 'EditGoTo',
  'edit-links': 'EditLinks',
  'edit-find-next': 'RepeatFind',
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  'double-underline': 'DoubleUnderline',
  'word-underline': 'WordUnderline',
  strikethrough: 'Strikethrough',
  'all-caps': 'AllCaps',
  'small-caps': 'SmallCaps',
  'hidden-text': 'Hidden',
  subscript: 'Subscript',
  superscript: 'Superscript',
  'font-grow': 'GrowFont',
  'font-shrink': 'ShrinkFont',
  'change-case': 'ChangeCase',
  'font-dialog': 'FormatFont',
  'paragraph-dialog': 'FormatParagraph',
  'format-tabs': 'FormatTabs',
  'format-borders-shading': 'FormatBordersAndShading',
  'format-columns': 'FormatColumns',
  'format-drop-cap': 'FormatDropCap',
  'format-bullets-numbering': 'FormatBulletsAndNumbering',
  'format-change-case': 'FormatChangeCase',
  'align-left': 'LeftPara',
  'align-center': 'CenterPara',
  'align-right': 'RightPara',
  'align-justify': 'JustifyPara',
  'align-distributed': 'DistributePara',
  'indent-increase': 'Indent',
  'indent-decrease': 'UnIndent',
  'line-spacing:1': 'SpacePara1',
  'line-spacing:1.5': 'SpacePara15',
  'line-spacing:2': 'SpacePara2',
  'show-marks': 'ShowAll',
  'view-print-layout': 'ViewPage',
  'view-web-layout': 'ViewWeb',
  'view-outline': 'ViewOutline',
  'view-draft': 'ViewNormal',
  'view-header-footer': 'ViewHeader',
  'view-footnotes': 'ViewFootnotes',
  'view-thumbnails': 'ToggleThumbnail',
  'view-ribbon': 'ToggleRibbon',
  'toggle-ruler': 'ViewRuler',
  'toggle-grid': 'ViewGridlines',
  'zoom-dialog': 'ViewZoom',
  'zoom-100': 'ViewZoom100',
  'zoom-page-width': 'ViewZoomPageWidth',
  'zoom-whole-page': 'ViewZoomWholePage',
  'insert-page-break': 'InsertPageBreak',
  'insert-column-break': 'InsertColumnBreak',
  'insert-footnote': 'InsertFootnoteNow',
  'insert-endnote': 'InsertEndnoteNow',
  'insert-footnote-dialog': 'InsertFootnote',
  'insert-bookmark': 'EditBookmark',
  'insert-link': 'InsertHyperlink',
  'insert-cross-reference': 'InsertCrossReference',
  'insert-advanced-symbol': 'InsertSymbol',
  'insert-image': 'InsertPicture',
  'insert-field-dialog': 'InsertField',
  'insert-date-time': 'InsertDateTime',
  'insert-file': 'InsertFile',
  'insert-caption': 'InsertCaption',
  'insert-index-tables': 'InsertIndexAndTables',
  'insert-comment': 'InsertNewComment',
  'insert-equation': 'InsertEquation',
  'insert-horizontal-line': 'InsertHorizontalLine',
  'insert-text-box': 'DrawTextBox',
  'insert-vertical-text-box': 'DrawVerticalTextBox',
  'insert-wordart': 'InsertWordArt',
  'insert-chart': 'InsertChart',
  'insert-citation': 'InsertCitation',
  'insert-bibliography': 'InsertBibliography',
  'mark-index-entry': 'MarkIndexEntry',
  'toa-mark-citation': 'MarkCitation',
  'tools-spelling': 'ToolsProofing',
  'tools-thesaurus': 'ToolsThesaurus',
  'tools-hyphenation': 'ToolsHyphenation',
  'tools-language': 'ToolsLanguage',
  'word-count': 'ToolsWordCount',
  'track-changes': 'ToolsRevisionMarksToggle',
  'compare-documents': 'ToolsCompareVersions',
  'tools-combine-documents': 'ToolsMergeRevisions',
  'protect-document': 'ToolsProtectDocument',
  'tools-envelopes': 'ToolsCreateEnvelope',
  'tools-labels': 'ToolsCreateLabels',
  'tools-customize-keyboard': 'ToolsCustomizeKeyboard',
  'autocorrect-options': 'ToolsAutoCorrect',
  preferences: 'ToolsOptions',
  'accept-revision': 'AcceptChangesSelected',
  'reject-revision': 'RejectChangesSelected',
  'review-pane': 'ReviewingPane',
  'delete-comment': 'DeleteAnnotation',
  'delete-comments': 'DeleteAllCommentsInDoc',
  'insert-table': 'TableInsertTable',
  'table-insert-cells': 'TableInsertCells',
  'table-insert-rows-above': 'TableInsertRowAbove',
  'table-insert-rows-below': 'TableInsertRowBelow',
  'table-insert-cols-left': 'TableInsertColumn',
  'table-insert-cols-right': 'TableInsertColumnRight',
  'table-delete-table': 'TableDeleteTable',
  'table-delete-columns': 'TableDeleteColumn',
  'table-delete-rows': 'TableDeleteRow',
  'table-delete-cells': 'TableDeleteCells',
  'table-select-table': 'TableSelectTable',
  'table-select-column': 'TableSelectColumn',
  'table-select-row': 'TableSelectRow',
  'table-select-cell': 'TableSelectCell',
  'table-merge-cells': 'TableMergeCells',
  'table-split-cells': 'TableSplitCells',
  'table-split-table': 'TableSplit',
  'table-repeat-header': 'TableHeadings',
  'convert-text-to-table': 'TextToTable',
  'table-convert-to-text': 'TableToText',
  'table-formula': 'TableFormula',
  'table-gridlines': 'TableGridlines',
  'table-properties': 'TableProperties',
  sort: 'TableSort',
  'new-window': 'WindowNewWindow',
  'view-arrange-all': 'WindowArrangeAll',
  'split-view': 'DocSplit',
  'close-pane': 'ClosePane',
  'window-go:next': 'NextWindow',
  'window-go:prev': 'PrevWindow',
  'go-back': 'GoBack',
}

const BINDING_BY_FCI = new Map(
  Object.entries(FCI_NAMES).map(([id, name]) => [name.toLowerCase(), id]),
)

/** the name a document stores the binding's command under; undefined when Word has no such command */
export function fciNameOf(id: string): string | undefined {
  return Object.hasOwn(FCI_NAMES, id) ? FCI_NAMES[id] : undefined
}

/** every binding a document can carry a key for */
export const DOCUMENT_BINDINGS: readonly string[] = Object.keys(FCI_NAMES)

// Windows virtual key codes, which a KCM's low byte holds
const VIRTUAL_KEYS: Record<string, number> = {
  Backspace: 0x08,
  Tab: 0x09,
  Clear: 0x0c,
  Enter: 0x0d,
  Space: 0x20,
  PageUp: 0x21,
  PageDown: 0x22,
  End: 0x23,
  Home: 0x24,
  Left: 0x25,
  Up: 0x26,
  Right: 0x27,
  Down: 0x28,
  Delete: 0x2e,
  NumPlus: 0x6b,
  NumMinus: 0x6d,
  ';': 0xba,
  '=': 0xbb,
  ',': 0xbc,
  '-': 0xbd,
  '.': 0xbe,
  '/': 0xbf,
  '`': 0xc0,
  '[': 0xdb,
  '\\': 0xdc,
  ']': 0xdd,
  "'": 0xde,
}
const KEY_BY_VIRTUAL = new Map(Object.entries(VIRTUAL_KEYS).map(([key, code]) => [code, key]))

function virtualKey(key: string): number | null {
  if (/^[A-Z0-9]$/.test(key)) return key.charCodeAt(0)
  const fn = /^F(\d{1,2})$/.exec(key)
  if (fn) return Number(fn[1]) >= 1 && Number(fn[1]) <= 24 ? 0x6f + Number(fn[1]) : null
  return VIRTUAL_KEYS[key] ?? null
}

/** the chord as a KCM (modifier byte, then key code); null for a key a KCM cannot name */
export function chordToKcm(chord: string): number | null {
  const parts = chord.split('+')
  const code = virtualKey(parts.pop() ?? '')
  if (code === null) return null
  const modifiers =
    (parts.includes('Shift') ? KCM_SHIFT : 0) |
    (parts.includes('Ctrl') ? KCM_CONTROL : 0) |
    (parts.includes('Alt') ? KCM_ALT : 0) |
    (parts.includes('Cmd') ? KCM_COMMAND : 0)
  return (modifiers << 8) | code
}

/** null for a key this keyboard model has no chord for (Command off the Mac among them) */
export function kcmToChord(kcm: number, isMac = IS_MAC): string | null {
  const code = kcm & 0xff
  const modifiers = kcm >> 8
  if (modifiers & ~(KCM_SHIFT | KCM_CONTROL | KCM_ALT | KCM_COMMAND)) return null
  const key =
    (code >= 0x30 && code <= 0x39) || (code >= 0x41 && code <= 0x5a)
      ? String.fromCharCode(code)
      : code >= 0x70 && code <= 0x87
        ? `F${code - 0x6f}`
        : KEY_BY_VIRTUAL.get(code)
  if (!key) return null
  const held = [
    ...(modifiers & KCM_CONTROL ? ['Ctrl'] : []),
    ...(modifiers & KCM_ALT ? ['Alt'] : []),
    ...(modifiers & KCM_SHIFT ? ['Shift'] : []),
    ...(modifiers & KCM_COMMAND ? ['Cmd'] : []),
  ]
  return toChord([...held, key].join('+'), isMac)
}

/** one key map of the document */
export interface DocKey {
  entry: KeymapEntry
  /** the chord it is on; null when it is not one key we can spell */
  chord: string | null
  /** the binding it runs; null for a masked key and for a command we do not have */
  id: string | null
}

/** whether the key map takes part in key handling here; the others are only carried */
export const isLiveDocKey = (key: DocKey): boolean =>
  key.chord !== null && (key.entry.mask || key.id !== null)

function docKeyOf(entry: KeymapEntry, isMac: boolean): DocKey {
  const chord =
    entry.kcmPrimary !== null && entry.kcmSecondary === null
      ? kcmToChord(entry.kcmPrimary, isMac)
      : null
  const id = entry.mask ? null : (BINDING_BY_FCI.get(entry.fciName?.toLowerCase() ?? '') ?? null)
  return { entry, chord, id: id && bindingCommand(id) ? id : null }
}

export function docKeysFromXml(xml: string | undefined, isMac = IS_MAC): DocKey[] {
  return xml ? readKeymaps(xml).map((entry) => docKeyOf(entry, isMac)) : []
}

/** the Customizations part holding `keys`; null when there is nothing to write or the part cannot be edited */
export function docKeysXml(source: string | undefined, keys: readonly DocKey[]): string | null {
  return writeKeymaps(
    source ?? null,
    keys.map((key) => key.entry),
  )
}

/** a file must not be able to take Save's key away: the user would believe the document saved */
const holdsSave = (overrides: KeyOverrides, chord: string, isMac: boolean): boolean => {
  const owner = chordOwner(overrides, chord, isMac)
  return owner?.kind === 'binding' && owner.id === 'save'
}

/** the machine's assignments with the document's laid over them, in file order */
export function layerDocKeys(
  overrides: KeyOverrides,
  keys: readonly DocKey[],
  isMac = IS_MAC,
): KeyOverrides {
  let out = overrides
  for (const key of keys) {
    if (!isLiveDocKey(key) || !key.chord) continue
    if (holdsSave(out, key.chord, isMac) && key.id !== 'save') continue
    if (key.id) {
      out = assignChord(out, key.id, key.chord, isMac) ?? out
      continue
    }
    const owner = chordOwner(out, key.chord, isMac)
    if (owner?.kind === 'binding') out = removeChord(out, owner.id, key.chord, isMac)
  }
  return out
}

const withoutChord = (keys: readonly DocKey[], chord: string): DocKey[] =>
  keys.filter((key) => !(isLiveDocKey(key) && key.chord === chord))

const madeKey = (chord: string, id: string | null, kcm: number): DocKey => ({
  entry: { kcmPrimary: kcm, kcmSecondary: null, mask: id === null, fciName: id && FCI_NAMES[id] },
  chord,
  id,
})

/** the document's keys with `chord` given to the binding; null when the document cannot hold that */
export function assignDocKey(
  overrides: KeyOverrides,
  keys: readonly DocKey[],
  id: string,
  chord: string,
  isMac = IS_MAC,
): DocKey[] | null {
  const kcm = chordToKcm(chord)
  if (kcm === null || !fciNameOf(id)) return null
  const rest = withoutChord(keys, chord)
  const under = layerDocKeys(overrides, rest, isMac)
  if (holdsSave(under, chord, isMac) && id !== 'save') return null
  if (!assignChord(under, id, chord, isMac)) return null
  return [...rest, madeKey(chord, id, kcm)]
}

/** the document's keys with `chord` taken off the binding: its own key map goes, a key from elsewhere is masked */
export function removeDocKey(
  overrides: KeyOverrides,
  keys: readonly DocKey[],
  id: string,
  chord: string,
  isMac = IS_MAC,
): DocKey[] | null {
  const rest = withoutChord(keys, chord)
  const owner = chordOwner(layerDocKeys(overrides, rest, isMac), chord, isMac)
  if (owner?.kind !== 'binding' || owner.id !== id) return rest
  if (id === 'save') return null
  const kcm = chordToKcm(chord)
  return kcm === null ? null : [...rest, madeKey(chord, null, kcm)]
}

/** only the key maps this app does not run are left */
export const resetDocKeys = (keys: readonly DocKey[]): DocKey[] =>
  keys.filter((key) => !isLiveDocKey(key))
