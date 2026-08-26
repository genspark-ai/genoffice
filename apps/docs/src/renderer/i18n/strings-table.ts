import { defineStrings } from '@genoffice/i18n'

const en = {
  ribbonTableStyleOptions: 'Table Style Options',
  ribbonTableFirstRow: 'Header Row',
  ribbonTableLastRow: 'Total Row',
  ribbonTableBandedRows: 'Banded Rows',
  ribbonTableFirstColumn: 'First Column',
  ribbonTableLastColumn: 'Last Column',
  ribbonTableBandedColumns: 'Banded Columns',
  ribbonTablePresetGrid: 'Plain Grid',
  ribbonTablePresetBlueHeader: 'Blue Header',
  ribbonTablePresetBlueBanded: 'Blue Banded',
  ribbonTablePresetGrayBanded: 'Gray Banded',
  ribbonTablePresetGreenHeader: 'Green Header',
  ribbonAutoFit: 'AutoFit',
  ribbonAutoFitContents: 'AutoFit Contents',
  ribbonAutoFitWindow: 'AutoFit Window',
  ribbonFixedColumnWidth: 'Fixed Column Width',
  ribbonRepeatHeaderRows: 'Repeat Header Rows',
  ribbonTableProperties: 'Table Properties',
  ribbonTableData: 'Table',
  ribbonHorizontalPosition: 'Horizontal position',
  ribbonVerticalPosition: 'Vertical position',
  ribbonDistanceFromText: 'Distance from text',
  ribbonCellMargins: 'Default cell margins',
}

/**
 * New table controls deliberately fall back to English until each locale has
 * reviewed terminology. Keeping one complete key set prevents partially
 * translated dialogs and lets language packs override the shard incrementally.
 */
export const tableStrings = defineStrings({
  zh: en,
  en,
  ja: en,
  ko: en,
  fr: en,
  de: en,
  es: en,
  th: en,
  id: en,
  ru: en,
  ar: en,
  pt: en,
  it: en,
  pl: en,
  nl: en,
  ms: en,
  he: en,
  hi: en,
  'zh-TW': en,
})
