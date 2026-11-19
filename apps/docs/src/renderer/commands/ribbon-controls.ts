import type { CommandId } from '../../shared/commands'
import type { RibbonFormatState } from '../components/ribbon-format-state'
import type { StringKey } from '../i18n/locale'

/** how the ribbon names a command: its tooltip, and the caption of a captioned button */
export interface RibbonControl {
  tip: StringKey
  caption?: StringKey
}

const RIBBON_CONTROLS = {
  'toggle-ai': { tip: 'aiOpenAssistant' },
  bold: { tip: 'ribbonBoldTip' },
  italic: { tip: 'ribbonItalicTip' },
  underline: { tip: 'ribbonUnderlineTip' },
  strikethrough: { tip: 'ribbonStrikethrough' },
  subscript: { tip: 'ribbonSubscript' },
  superscript: { tip: 'ribbonSuperscript' },
  'indent-decrease': { tip: 'ribbonDecreaseIndent' },
  'indent-increase': { tip: 'ribbonIncreaseIndent' },
  'show-marks': { tip: 'ribbonShowMarks' },
  'insert-blank-page': { tip: 'ribbonBlankPageTip', caption: 'ribbonBlankPage' },
  'insert-page-break': { tip: 'ribbonPageBreakTip', caption: 'ribbonPageBreak' },
  'insert-image': { tip: 'ribbonPictureTip', caption: 'ribbonPictures' },
  'insert-chart': { tip: 'ribbonChartTip', caption: 'ribbonChart' },
  'insert-link': { tip: 'ribbonLinkTip', caption: 'ribbonLink' },
  'insert-bookmark': { tip: 'ribbonBookmarkTip', caption: 'ribbonBookmark' },
  'insert-cross-reference': { tip: 'ribbonCrossRefTip', caption: 'ribbonCrossRef' },
  'font-grow': { tip: 'ribbonGrowFont' },
  'font-shrink': { tip: 'ribbonShrinkFont' },
  'clear-formatting': { tip: 'ribbonClearFormatting' },
  'format-phonetic-guide': { tip: 'ribbonPhoneticGuide' },
  'character-border': { tip: 'ribbonCharacterBorder' },
  'format-text-effects': { tip: 'ribbonTextEffects' },
  'character-shading': { tip: 'ribbonCharacterShading' },
  'format-enclose-characters': { tip: 'ribbonEncloseCharacters' },
  sort: { tip: 'ribbonSort', caption: 'ribbonSort' },
  'table-convert-to-text': {
    tip: 'ribbonConvertTableToTextTitle',
    caption: 'ribbonConvertTableToTextTitle',
  },
  'table-formula': { tip: 'ribbonFormulaTitle', caption: 'ribbonFormulaTitle' },
  'convert-text-to-table': { tip: 'ribbonConvertTextToTableTitle' },
  'insert-file': { tip: 'ribbonTextFromFile' },
  'align-distributed': { tip: 'ribbonDistributed' },
  'insert-icons': { tip: 'ribbonIcons', caption: 'ribbonIcons' },
  'insert-date-time': { tip: 'ribbonDateTime', caption: 'ribbonDateTime' },
  'insert-number': { tip: 'ribbonNumber', caption: 'ribbonNumber' },
  'tools-thesaurus': { tip: 'ribbonThesaurus', caption: 'ribbonThesaurus' },
  'word-count': { tip: 'ribbonWordCount', caption: 'ribbonWordCount' },
  'tools-language': { tip: 'ribbonProofingLanguage', caption: 'ribbonProofingLanguage' },
  'file-always-read-only': { tip: 'ribbonAlwaysReadOnly', caption: 'ribbonAlwaysReadOnly' },
  'view-focus': { tip: 'ribbonFocus', caption: 'ribbonFocus' },
  'view-arrange-all': { tip: 'ribbonArrangeAll', caption: 'ribbonArrangeAll' },
  'view-merge-windows': { tip: 'ribbonMergeAllWindows', caption: 'ribbonMergeAllWindows' },
  'design-set-default': { tip: 'ribbonSetAsDefault', caption: 'ribbonSetAsDefault' },
  'format-text-direction': { tip: 'ribbonTextDirection', caption: 'ribbonTextDirection' },
  'layout-line-numbers': { tip: 'ribbonLineNumbers', caption: 'ribbonLineNumbers' },
  'tools-hyphenation': { tip: 'ribbonHyphenation', caption: 'ribbonHyphenation' },
  'format-selection-pane': { tip: 'ribbonSelectionPane', caption: 'ribbonSelectionPane' },
  'arrange-align': { tip: 'ribbonAlignObjects', caption: 'ribbonAlignObjects' },
  'arrange-group': { tip: 'ribbonArrangeGroup', caption: 'ribbonArrangeGroup' },
  'toc-add-text': { tip: 'ribbonAddText', caption: 'ribbonAddText' },
  'footnote-next': { tip: 'ribbonNextFootnote', caption: 'ribbonNextFootnote' },
  'view-footnotes': { tip: 'ribbonShowNotes', caption: 'ribbonShowNotes' },
  'citations-pane': { tip: 'ribbonCitationsPane', caption: 'ribbonCitationsPane' },
  'insert-table-of-figures': { tip: 'ribbonInsertTof', caption: 'ribbonInsertTof' },
  'update-table-of-figures': { tip: 'ribbonTocUpdate', caption: 'ribbonTocUpdate' },
  'citation-style': { tip: 'ribbonCitationStyle', caption: 'ribbonCitationStyle' },
  'mark-index-entry': { tip: 'ribbonMarkEntryTitle', caption: 'ribbonMarkEntry' },
  'insert-index': { tip: 'ribbonInsertIndex', caption: 'ribbonInsertIndex' },
  'update-index': { tip: 'ribbonUpdateIndex', caption: 'ribbonUpdateIndex' },
  'toa-mark-citation': { tip: 'ribbonMarkCitation', caption: 'ribbonMarkCitation' },
  'toa-insert': { tip: 'ribbonInsertToa', caption: 'ribbonInsertToa' },
  'toa-update': { tip: 'ribbonTocUpdate', caption: 'ribbonTocUpdate' },
  'tools-envelopes': { tip: 'ribbonEnvelopes', caption: 'ribbonEnvelopes' },
  'tools-labels': { tip: 'ribbonLabels', caption: 'ribbonLabels' },
  'mailings-select-recipients': {
    tip: 'ribbonSelectRecipients',
    caption: 'ribbonSelectRecipients',
  },
  'mailings-edit-recipients': {
    tip: 'ribbonEditRecipientList',
    caption: 'ribbonEditRecipientList',
  },
  'mailings-filter-recipients': {
    tip: 'ribbonFilterRecipients',
    caption: 'ribbonFilterRecipients',
  },
  'mailings-rules': { tip: 'ribbonRules', caption: 'ribbonRules' },
  'mailings-update-labels': { tip: 'ribbonUpdateLabels', caption: 'ribbonUpdateLabels' },
  'mailings-preview-results': { tip: 'ribbonPreviewResults', caption: 'ribbonPreviewResults' },
  'mailings-find-recipient': { tip: 'ribbonFindRecipient', caption: 'ribbonFindRecipient' },
  'mailings-finish-merge': { tip: 'ribbonFinishMerge', caption: 'ribbonFinishMerge' },
  bullets: { tip: 'ribbonBullets' },
  numbering: { tip: 'ribbonNumbering' },
  'direction-ltr': { tip: 'ribbonDirLtrTip' },
  'direction-rtl': { tip: 'ribbonDirRtlTip' },
  'align-left': { tip: 'ribbonAlignLeftTip' },
  'align-center': { tip: 'ribbonAlignCenterTip' },
  'align-right': { tip: 'ribbonAlignRightTip' },
  'align-justify': { tip: 'ribbonJustifyTip' },
  'styles-pane': { tip: 'ribbonStylesPaneTip', caption: 'ribbonStylesPane' },
  find: { tip: 'ribbonFindTip', caption: 'ribbonFind' },
  replace: { tip: 'ribbonReplaceTip', caption: 'ribbonReplace' },
} satisfies Partial<Record<CommandId, RibbonControl>>

export type RibbonCommandId = keyof typeof RIBBON_CONTROLS

export const RIBBON_COMMAND_IDS = Object.keys(RIBBON_CONTROLS) as RibbonCommandId[]

export const ribbonControl = (id: RibbonCommandId): RibbonControl => RIBBON_CONTROLS[id]

/** the app state, beside the selection's formatting, that a ribbon control shows as pressed or greyed */
export interface RibbonViewState {
  showAi: boolean
  showMarks: boolean
  showStylesPane: boolean
  readOnlyRecommended: boolean
  mergePreview: boolean
  hasComments: boolean
  hasMergeSource: boolean
  windowsDetached: boolean
}

export function ribbonCommandPressed(
  id: RibbonCommandId,
  format: RibbonFormatState,
  view: RibbonViewState,
): boolean {
  switch (id) {
    case 'toggle-ai':
      return view.showAi
    case 'show-marks':
      return view.showMarks
    case 'styles-pane':
      return view.showStylesPane
    case 'file-always-read-only':
      return view.readOnlyRecommended
    case 'mailings-preview-results':
      return view.mergePreview
    case 'character-shading':
      return format.charShading
    case 'character-border':
      return format.charBorder
    case 'bold':
      return format.bold
    case 'italic':
      return format.italic
    case 'underline':
      return format.underline
    case 'strikethrough':
      return format.strike
    case 'bullets':
      return format.listBullet
    case 'numbering':
      return format.listOrdered
    case 'direction-ltr':
      return !format.bidi
    case 'direction-rtl':
      return format.bidi
    case 'subscript':
    case 'superscript':
      return format.vertAlign === id
    case 'align-left':
    case 'align-center':
    case 'align-right':
    case 'align-justify':
      // a paragraph with no alignment of its own sits at its reading direction's start
      return (format.align ?? (format.bidi ? 'right' : 'left')) === id.slice('align-'.length)
    case 'align-distributed':
      return format.align === 'distribute'
    default:
      return false
  }
}
