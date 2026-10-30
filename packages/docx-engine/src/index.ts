// one definition, in the package both carriers and both apps can reach
export { REDACT_EL, REDACT_NS, REDACT_PREFIX } from '@genoffice/agent-core/redact-range'
export * from './types'
export {
  deobfuscateOdttf,
  isSfnt,
  parseFontTable,
  readEmbeddedFonts,
  sfntLineMetrics,
} from './font-table'
export { decodeEntities } from './parse-xml-text'
export { sdtCheckboxGlyphs, sdtCheckboxIsChecked } from './checkbox-control'
export {
  buildContentControlXml,
  isContentControlLock,
  nextContentControlId,
  CONTENT_CONTROL_LOCKS,
  type ContentControlLock,
  type ContentControlShell,
  type ContentControlSpec,
} from './content-control'
export {
  parseDocx,
  reconcileGridColumns,
  paraCsFormat,
  styleRunFormat,
  type ParseExtras,
  type ParseOptions,
} from './parse'
export {
  compareOperands,
  conditionHolds,
  controlEffect,
  ifBranch,
  mergeBlockTree,
  mergeDataFromRows,
  mergeFieldInstruction,
  mergeFieldName,
  mergeFieldSwitches,
  mergeFieldText,
  mergeOutputFileName,
  mergeRecords,
  placeholderText,
  splitInstruction,
  valueOf,
  type ControlEffect,
  type MergeCursorState,
  type MergeData,
  type MergeFieldSwitches,
  type MergeRecord,
} from './mail-merge'
export { assertZipInflatesWithinLimits, assertZipWithinLimits, DOCX_ZIP_LIMITS } from './zip-load'
export { LAZY_MEDIA_SCHEME, isLazyMediaPart, lazyMediaUrl, parseLazyMediaUrl } from './lazy-media'
export { setAltChunkHtmlConverter, type AltChunkHtmlConverter } from './alt-chunk'
export { tocLevelOf } from './parse-fields'
export { mergeStyleBorders } from './parse-props'
export {
  APP_PROPERTIES_PATH,
  APP_PROPERTY_KEYS,
  CORE_PROPERTIES_PATH,
  CORE_PROPERTY_KEYS,
  applyPropertyChanges,
  emptyDocumentProperties,
  parseAppPropertiesXml,
  parseCorePropertiesXml,
  parseCustomPropertiesXml,
  parsePropertyChanges,
  patchAppPropertiesXml,
  patchCorePropertiesXml,
  patchCustomPropertiesXml,
  readDocumentProperties,
  type DocumentProperties,
  type PropertyChanges,
} from './doc-properties'
export {
  saveDocx,
  findChartWorkbookPath,
  missingHeadingLevels,
  readDocxPartBase64,
  type SaveBlock,
  type SaveOptions,
  type StyleUpsert,
  type ParsedDocFull,
} from './patch'
export {
  TABLE_HEADER_FILL,
  applyImageWrap,
  applyImageZOrder,
  applyShapeWrapAt,
  applyShapeZOrderAt,
  type ShapeDrawingLocation,
  buildAnchoredTextboxParagraphXml,
  buildShapeParagraphXml,
  buildTextboxParagraphXml,
  buildWordArtParagraphXml,
  type AnchoredTextboxOptions,
  type TextboxContentParagraph,
  generateCaptionXml,
  generateIndexFieldXml,
  generateParagraphXml,
  generateTableModelXml,
  generateTableXml,
  generateTocFieldXml,
  mergePPrFormat,
  setPPrChange,
  stripPPrChange,
  paraMarkRevisionXml,
  setParaMarkRevision,
  stripParaMarkRevision,
  resolveParaMarkRPrChange,
  paraMarkColor,
  stripParaMarkColor,
  patchFieldParagraphXml,
  patchImageParagraphXml,
  patchMathTokens,
  patchTableCellTexts,
  patchTextboxHeights,
  patchTextboxParas,
  patchTextboxSizes,
  patchShapeStyles,
  type ShapeStylePatch,
  patchDrawingExtent,
  buildLineParagraphXml,
  LINE_KINDS,
  type TextboxSizePatch,
  type CellParaPatch,
  type CellTextsPatch,
  type FieldTextPatch,
  type GenerateContext,
  type ImagePatch,
  type TextboxParaPatch,
  type TextboxParasPatchSet,
  inlinePictureDrawingXml,
  resizeInlinePictureXml,
  NEW_PICTURE_TOKEN_PREFIX,
  type InlinePictureXmlOptions,
  type TableGenOptions,
  type TocEntry,
} from './generate'
export {
  buildChartPartXml,
  buildChartWorkbookXlsxBase64,
  patchChartWorkbookXlsxBase64,
  parseChartPartXml,
  patchChartPartXml,
  lumHex,
  colLetter,
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
export { patchRunPictureXml, type RunPictureGeometry } from './run-picture'
export { scanBody, type BodyElement, type BodyScan } from './scan'
export {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  customLevelFromNumberingLevel,
  customLevelXml,
  mergeLevelXml,
  numberingLevelFromCustom,
  paperSizeForLocale,
  paperSizeForRegion,
  type BlankDocxOptions,
  type BlankPaperSize,
  type CustomNumberingLevel,
} from './blank'
export {
  DEFAULT_SECTION,
  applySectionSettings,
  applyPageNumType,
  applySectionStartType,
  applyTitlePg,
  sectionFromSectPr,
  readPageColor,
  readSections,
  readSectionSettings,
  sectionSettingsFromXml,
  notePropsFromXml,
  xmlFlagOn,
} from './section'
export { nextNoteId, parseNotesXml, type NoteKind } from './notes'
export {
  isPictureWatermark,
  pictureWatermarkPreviewImage,
  readPictureWatermark,
  readWatermarkText,
  type PictureWatermarkInfo,
  type PictureWatermarkSpec,
  type Watermark,
  type WatermarkSpec,
} from './watermark'
export {
  builtInHeadingStyleUpsert,
  mergeStyleXml,
  mergeDefaultFontsXml,
  type DefaultFonts,
  pendingHeadingLevel,
  type StyleHeadingInfo,
  type StyleParaProps,
  type StyleRunProps,
} from './style-upsert'
export {
  INK_NAME_PREFIX,
  anchoredInkRunXml,
  findInkRuns,
  injectInkRunsIntoParagraph,
  stripInkRuns,
} from './ink'
export { bibliographyLine, citationText, parseSourcesXml } from './sources'
export { parseZoteroDocumentDataXml, patchZoteroDocumentDataXml } from './zotero-doc-props'
export { readThemeColors, readThemeFonts } from './theme'
export {
  DEFAULT_SPIN_COUNT,
  MAX_SPIN_COUNT,
  hashProtectionPassword,
  resolveSpinCount,
  verifyProtectionPassword,
} from './protection'
export {
  decodeSymbolChar,
  decodeSymbolText,
  isSymbolFont,
  symbolGlyph,
  symbolPuaChar,
  toSymbolPua,
} from './symbol-fonts'
export {
  type BulletSubstitute,
  substituteBullet,
  markerFallbackFace,
  SEGOE_UI_SYMBOL_RE,
  computeListMarkerInfos,
  computeListMarkers,
  computeListValues,
  customEnumItems,
  formatNumber,
  markerTabAdvance,
  type ListItemRef,
  type ListMarkerInfo,
} from './list-markers'

export { previewFontSettings } from './font-settings'
export {
  applyRevisionResolutions,
  applyTrackedChangesXml,
  countTrackedChangesXml,
  readPropChangeXml,
  resolvePropChangeXml,
  type PropChangeRecord,
  type RevisionResolution,
  type RevisionXmlFilter,
  type RevisionXmlMode,
  type RevisionXmlResult,
  type RevisionXmlType,
} from './revision-xml'
export {
  resolveParagraphProps,
  resolveCtxOf,
  resolveParagraphPropsXml,
  resolveRunProps,
  shadowCompare,
  type EffectivePPr,
  type EffectiveRPr,
  type PropSource,
  type ResolveCompat,
  type ResolveCtx,
  type ResolveDiff,
  type TableCond,
} from './resolve-props'
export {
  compareBlocks,
  compareDocuments,
  DEFAULT_COMPARE_AUTHOR,
  type CompareGranularity,
  type CompareIgnore,
  type CompareOptions,
  type CompareResult,
  type CompareSummary,
} from './compare'
