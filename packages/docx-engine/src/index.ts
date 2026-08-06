export * from './types'
export { parseDocx, type ParseExtras } from './parse'
export {
  saveDocx,
  findChartWorkbookPath,
  readDocxPartBase64,
  type SaveBlock,
  type SaveOptions,
  type StyleUpsert,
  type ParsedDocFull,
} from './patch'
export {
  TABLE_HEADER_FILL,
  WORDART_PRESETS,
  applyImageWrap,
  buildShapeParagraphXml,
  buildTextboxParagraphXml,
  buildWordArtParagraphXml,
  generateCaptionXml,
  generateIndexFieldXml,
  generateParagraphXml,
  generateTableModelXml,
  generateTableXml,
  generateTocFieldXml,
  mergePPrFormat,
  setPPrChange,
  stripPPrChange,
  patchFieldParagraphXml,
  patchImageParagraphXml,
  patchMathTokens,
  patchTableCellTexts,
  patchTextboxHeights,
  patchTextboxParas,
  patchTextboxSizes,
  patchDrawingExtent,
  buildLineParagraphXml,
  LINE_KINDS,
  type TextboxSizePatch,
  type CellTextsPatch,
  type FieldTextPatch,
  type GenerateContext,
  type ImagePatch,
  type TextboxParaPatch,
  type TableGenOptions,
  type TocEntry,
} from './generate'
export {
  buildChartPartXml,
  buildChartWorkbookXlsxBase64,
  patchChartWorkbookXlsxBase64,
  parseChartPartXml,
  patchChartPartXml,
  CHART_WORKBOOK_REL_TYPE,
  type ChartPatch,
  type ChartSeriesPatch,
} from './chart'
export {
  latexToOmml,
  mathParagraphXml,
  mathTokensOf,
  ommlFragmentsOf,
  ommlToLatex,
  ommlToMathML,
} from './math'
export { scanBody, type BodyElement, type BodyScan } from './scan'
export {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  type BlankDocxOptions,
  type CustomNumberingLevel,
} from './blank'
export {
  DEFAULT_SECTION,
  applySectionSettings,
  applyPageNumType,
  applySectionStartType,
  readPageColor,
  readSections,
  readSectionSettings,
  sectionSettingsFromXml,
} from './section'
export { nextNoteId, parseNotesXml, type NoteKind } from './notes'
export { readWatermarkText } from './watermark'
export {
  INK_NAME_PREFIX,
  anchoredInkRunXml,
  findInkRuns,
  injectInkRunsIntoParagraph,
  stripInkRuns,
} from './ink'
export { bibliographyLine, citationText, parseSourcesXml } from './sources'
export { readThemeColors, readThemeFonts } from './theme'
export { hashProtectionPassword, verifyProtectionPassword } from './protection'
export { decodeSymbolChar, decodeSymbolText, isSymbolFont } from './symbol-fonts'
export { computeListMarkers, formatNumber, type ListItemRef } from './list-markers'
