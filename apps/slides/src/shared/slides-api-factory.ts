/// Transport-agnostic construction of the slides renderer bridge APIs.
///
/// One source of truth for the `window.slidesApi` / `window.desktop` /
/// `window.projectApi` surface: the sandboxed preload builds it on an
/// ipcRenderer transport (Electron), the browser web-bridge builds the exact
/// same object on the HTTP/SSE transport (web version). Channel names, argument
/// shapes, listener wrappers and return coercion are identical — only the
/// transport differs.
import type { RenderSlide } from '@genoffice/pptx-render'
import type { ProjectApi } from '@genoffice/project-store'
import type { IpcTransport } from '@genoffice/ipc-bridge/client'
import type {
  AddChartOp,
  AddElementOp,
  AiRunFailure,
  ApplyEditScriptOp,
  ApplyTxnOp,
  AddImageBytesOp,
  AddInkOp,
  AddMediaBytesOp,
  ReplacePictureBytesOp,
  AddSmartArtOp,
  ApplyThemeOp,
  AddBlankSlideOp,
  AddSlideOp,
  PasteSlideOp,
  RepasteSlideOp,
  AddSlideWithLayoutOp,
  AddTableOp,
  HeaderFooterOp,
  SetLinkOp,
  AiSettings,
  CopyElementsOp,
  PasteElementsOp,
  DuplicateElementsOp,
  EditTableCellOp,
  EditTableStyleOp,
  EditChartOp,
  EditPictureSrcRectOp,
  EditPictureOpacityOp,
  GroupElementsOp,
  UngroupElementOp,
  BatchEditTransformOp,
  SetTableColWidthOp,
  SetTableRowHeightOp,
  SetTableCellAnchorOp,
  TableStructureIpcOp,
  TableMergeIpcOp,
  ReorderElementOp,
  SetAdvanceTimesOp,
  SetAnimationsOp,
  SetSlideHiddenOp,
  SetTransitionOp,
  SectionInfo,
  AddSectionOp,
  RenameSectionOp,
  RemoveSectionOp,
  MoveSectionOp,
  MoveSlideOp,
  AiStreamChunk,
  AiStreamRequest,
  AudienceNavAction,
  ShowInkEvent,
  ShowSyncState,
  DeleteElementOp,
  DesktopFilesApi,
  EditBackgroundOp,
  EditFillOp,
  EditFillImageOp,
  EditStrokeOp,
  FlipElementOp,
  EditTextOp,
  EditTransformOp,
  EditConnectorEndpointsOp,
  SetElementFontOp,
  SetElementParagraphFormatOp,
  FindReplaceOp,
  SetSlideLayoutOp,
  SetSlideSizeOp,
  MasterEditTextOp,
  MasterEditTransformOp,
  MasterEditFillOp,
  MasterEditStrokeOp,
  MasterDeleteElementOp,
  ExportImagesOp,
  ExportPdfOp,
  PrintSlidesOp,
  MenuCommand,
  OpenResult,
  SlidesApi,
  UiTheme,
  SetEffectsPatch,
} from './ipc'

export interface SlidesApiOverrides {
  /** Web-native fullscreen (browser requestFullscreen). */
  setShowFullScreen?: (on: boolean) => Promise<unknown>
  /** Web-native open (browser file picker → temp file → open-path). */
  openPptx?: (fitWidthPx: number) => Promise<unknown>
  /** Web-native image insert (browser file picker → add-image-bytes). */
  insertImage?: (slideIndex: number, fitWidthPx: number) => Promise<unknown>
  /** Web-native media insert (browser file picker → add-media-bytes). */
  insertMedia?: (slideIndex: number, kind: 'video' | 'audio', fitWidthPx: number) => Promise<unknown>
  /** Web-native 3D insert (browser file picker → add-image-bytes). */
  insertModel3d?: (slideIndex: number, fitWidthPx: number) => Promise<unknown>
  /** Web-native export dir picker (temp dir path). */
  pickExportDir?: () => Promise<string | null>
  /** Web-native export PDF path picker (temp file path). */
  pickExportPdfPath?: (defaultName: string) => Promise<string | null>
  /** Web-native image export (write to temp dir, then download the PNGs). */
  exportImages?: (op: ExportImagesOp) => Promise<unknown>
  /** Web-native PDF export (write to temp file, then download the PDF). */
  exportPdf?: (op: ExportPdfOp) => Promise<unknown>
  /** Web-native print (browser print dialog). */
  printSlides?: (op: PrintSlidesOp) => Promise<unknown>
  /** Web-native clipboard (navigator.clipboard). */
  clipboardExternal?: () => Promise<unknown>
  /** Web-native font install (FontFace). */
  fontInstallLocal?: () => Promise<unknown>
  /** Web-native attachment picker (browser file input → temp files → files-add). */
  pickAttachments?: () => Promise<unknown>
  /** Web-native dropped-file path resolution. */
  getPathForFile?: (file: File) => string
}

export function createSlidesApi(t: IpcTransport, overrides: SlidesApiOverrides = {}): SlidesApi {
  return {
  getLanguage: () => t.invoke('app:get-language'),
  onLanguageChanged: (handler) =>
    t.on('app:language-changed', (lang) =>
      handler(lang as 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'),
    ),
  getTheme: () => t.invoke('app:get-theme'),
  onThemeChanged: (handler) => t.on('app:theme-changed', (theme) => handler(theme as UiTheme)),
  onChromePressed: (handler) => t.on('app:chrome-pressed', () => handler()),
  setShowFullScreen: overrides.setShowFullScreen ?? ((on) => t.invoke('slides:show-fullscreen', on)),
  privateFontFaces: () => t.invoke('slides:private-font-faces'),
  privateFontData: (id) => t.invoke('slides:private-font-data', id),
  fontCatalog: () => t.invoke('slides:font-catalog'),
  fontDownload: (family) => t.invoke('slides:font-download', family),
  fontInstallLocal: overrides.fontInstallLocal ?? (() => t.invoke('slides:font-install-local')),
  fontMissing: () => t.invoke('slides:font-missing'),
  onFontsChanged: (handler) => t.on('slides:fonts-changed', () => handler()),
  openPptx: overrides.openPptx ?? ((fitWidthPx) => t.invoke('slides:open', fitWidthPx)),
  openPptxPath: (path, fitWidthPx) => t.invoke('slides:open-path', path, fitWidthPx),
  consumePendingOpen: (fitWidthPx) => t.invoke('slides:consume-pending-open', fitWidthPx),
  newBlank: (fitWidthPx) => t.invoke('slides:new-blank', fitWidthPx),
  landGeneratedPages: (
    pageMarkers: string[],
    fitWidthPx: number,
    mode?: 'replace' | 'append' | 'replace_at' | 'insert_at',
    atIndex?: number,
    deckName?: string,
  ) =>
    t.invoke(
      'slides:land-generated-pages',
      pageMarkers,
      fitWidthPx,
      mode,
      atIndex,
      deckName,
    ),
  cloudGenStatus: () => t.invoke('slides:cloud-gen-status'),
  cloudGeneratePage: (op: {
    brief: string
    title?: string
    styleSkill?: string
    deckContext?: Record<string, unknown>
    images?: { url: string; caption?: string }[]
    width?: number
    height?: number
  }) => t.invoke('slides:cloud-page-generate', op),
  localGeneratePage: (op: { specJson: string }) =>
    t.invoke('slides:local-page-generate', op),
  editText: (op: EditTextOp) => t.invoke('slides:edit-text', op),
  setElementFont: (op: SetElementFontOp) => t.invoke('slides:set-element-font', op),
  setElementParagraphFormat: (op: SetElementParagraphFormatOp) =>
    t.invoke('slides:set-element-paragraph-format', op),
  findReplace: (op: FindReplaceOp) => t.invoke('slides:find-replace', op),
  setSlideLayout: (op: SetSlideLayoutOp) => t.invoke('slides:set-slide-layout', op),
  setSlideSize: (op: SetSlideSizeOp) => t.invoke('slides:set-slide-size', op),
  getSlideSize: () => t.invoke('slides:get-slide-size'),
  editTransform: (op: EditTransformOp) => t.invoke('slides:edit-transform', op),
  editConnectorEndpoints: (op: EditConnectorEndpointsOp) =>
    t.invoke('slides:edit-connector-endpoints', op),
  editPictureSrcRect: (op: EditPictureSrcRectOp) =>
    t.invoke('slides:edit-picture-src-rect', op),
  editPictureOpacity: (op: EditPictureOpacityOp) =>
    t.invoke('slides:edit-picture-opacity', op),
  editImageFill: (op: EditFillImageOp) => t.invoke('slides:edit-image-fill', op),
  changeShape: (op: { slideIndex: number; sourceId: string; prst: string; groupId?: string }) =>
    t.invoke('slides:change-shape', op),
  setShapeAdjust: (op: {
    slideIndex: number
    sourceId: string
    adjust: Record<string, number>
    groupId?: string
    preview?: boolean
  }) => t.invoke('slides:set-shape-adjust', op),
  setTextAnchor: (op: {
    slideIndex: number
    sourceId: string
    anchor: 'top' | 'middle' | 'bottom'
  }) => t.invoke('slides:set-text-anchor', op),
  setTextBodyProps: (op: {
    slideIndex: number
    sourceId: string
    props: {
      vert?: 'horz' | 'eaVert' | 'vert' | 'vert270' | 'wordArtVert'
      autofit?: 'none' | 'shrink' | 'resize'
      insets?: Partial<{ l: number; t: number; r: number; b: number }>
      wrap?: boolean
    }
  }) => t.invoke('slides:set-text-body-props', op),
  setEffects: (op: { slideIndex: number; sourceId: string; effects: SetEffectsPatch }) =>
    t.invoke('slides:set-effects', op),
  clipboardExternal: overrides.clipboardExternal ?? (() => t.invoke('slides:clipboard-external')),
  groupElements: (op: GroupElementsOp) => t.invoke('slides:group-elements', op),
  ungroupElement: (op: UngroupElementOp) => t.invoke('slides:ungroup-element', op),
  batchEditTransform: (op: BatchEditTransformOp) =>
    t.invoke('slides:batch-edit-transform', op),
  getRenderSlides: () => t.invoke('slides:get-render-slides'),
  addElement: (op: AddElementOp) => t.invoke('slides:add-element', op),
  deleteElement: (op: DeleteElementOp) => t.invoke('slides:delete-element', op),
  addSlide: (op: AddSlideOp) => t.invoke('slides:add-slide', op),
  addBlankSlide: (op: AddBlankSlideOp) => t.invoke('slides:add-blank-slide', op),
  addSlideWithLayout: (op: AddSlideWithLayoutOp) =>
    t.invoke('slides:add-slide-with-layout', op),
  getLayouts: () => t.invoke('slides:get-layouts'),
  masterEnter: (fitWidthPx: number) => t.invoke('slides:master-enter', fitWidthPx),
  masterOpen: (partPath: string) => t.invoke('slides:master-open', partPath),
  masterClose: () => t.invoke('slides:master-close'),
  masterEditText: (op: MasterEditTextOp) => t.invoke('slides:master-edit-text', op),
  masterEditTransform: (op: MasterEditTransformOp) =>
    t.invoke('slides:master-edit-transform', op),
  masterEditFill: (op: MasterEditFillOp) => t.invoke('slides:master-edit-fill', op),
  masterEditStroke: (op: MasterEditStrokeOp) => t.invoke('slides:master-edit-stroke', op),
  masterDeleteElement: (op: MasterDeleteElementOp) =>
    t.invoke('slides:master-delete-element', op),
  editFill: (op: EditFillOp) => t.invoke('slides:edit-fill', op),
  editStroke: (op: EditStrokeOp) => t.invoke('slides:edit-stroke', op),
  flipElements: (op: FlipElementOp) => t.invoke('slides:flip-elements', op),
  editBackground: (op: EditBackgroundOp) => t.invoke('slides:edit-background', op),
  insertImage: overrides.insertImage ??
    ((slideIndex: number, fitWidthPx: number) =>
      t.invoke('slides:insert-image', slideIndex, fitWidthPx)),
  copySlide: (slideIndex: number, pngBase64?: string) =>
    t.invoke('slides:copy-slide', slideIndex, pngBase64),
  pasteSlide: (op: PasteSlideOp) => t.invoke('slides:paste-slide', op),
  repasteSlide: (op: RepasteSlideOp) => t.invoke('slides:repaste-slide', op),
  hasSlideClipboard: () => t.invoke('slides:has-slide-clipboard'),
  clipboardProbe: () => t.invoke('slides:clipboard-probe'),
  deleteSlide: (slideIndex: number) => t.invoke('slides:delete-slide', slideIndex),
  reorderElement: (op: ReorderElementOp) => t.invoke('slides:reorder-element', op),
  editTableCell: (op: EditTableCellOp) => t.invoke('slides:edit-table-cell', op),
  tableStructure: (op: TableStructureIpcOp) => t.invoke('slides:table-structure', op),
  tableMerge: (op: TableMergeIpcOp) => t.invoke('slides:table-merge', op),
  setTableColWidth: (op: SetTableColWidthOp) =>
    t.invoke('slides:set-table-col-width', op),
  setTableRowHeight: (op: SetTableRowHeightOp) =>
    t.invoke('slides:set-table-row-height', op),
  setTableCellAnchor: (op: SetTableCellAnchorOp) =>
    t.invoke('slides:set-table-cell-anchor', op),
  editTableStyle: (op: EditTableStyleOp) => t.invoke('slides:edit-table-style', op),
  editChart: (op: EditChartOp) => t.invoke('slides:edit-chart', op),
  getChartColorSchemes: () => t.invoke('slides:chart-color-schemes'),
  getChartData: (slideIndex: number, sourceId: string) =>
    t.invoke('slides:get-chart-data', slideIndex, sourceId),
  copyElements: (op: CopyElementsOp) => t.invoke('slides:copy-elements', op),
  pasteElements: (op: PasteElementsOp) => t.invoke('slides:paste-elements', op),
  duplicateElements: (op: DuplicateElementsOp) =>
    t.invoke('slides:duplicate-elements', op),
  addTable: (op: AddTableOp) => t.invoke('slides:add-table', op),
  addInk: (op: AddInkOp) => t.invoke('slides:add-ink', op),
  addChart: (op: AddChartOp) => t.invoke('slides:add-chart', op),
  addSmartArt: (op: AddSmartArtOp) => t.invoke('slides:add-smartart', op),
  addImageBytes: (op: AddImageBytesOp) => t.invoke('slides:add-image-bytes', op),
  replacePictureBytes: (op: ReplacePictureBytesOp) =>
    t.invoke('slides:replace-picture-bytes', op),
  insertMedia: overrides.insertMedia ??
    ((slideIndex: number, kind: 'video' | 'audio', fitWidthPx: number) =>
      t.invoke('slides:insert-media', slideIndex, kind, fitWidthPx)),
  addMediaBytes: (op: AddMediaBytesOp) => t.invoke('slides:add-media-bytes', op),
  getMediaData: (slideIndex: number, sourceId: string) =>
    t.invoke('slides:media-data', slideIndex, sourceId),
  insertModel3d: overrides.insertModel3d ??
    ((slideIndex: number, fitWidthPx: number) =>
      t.invoke('slides:insert-model3d', slideIndex, fitWidthPx)),
  setLink: (op: SetLinkOp) => t.invoke('slides:set-link', op),
  getLink: (slideIndex: number, sourceId: string) =>
    t.invoke('slides:get-link', slideIndex, sourceId),
  getSlideLinks: (slideIndex: number) => t.invoke('slides:get-slide-links', slideIndex),
  getRunLinks: (slideIndex: number) => t.invoke('slides:get-run-links', slideIndex),
  applyHeaderFooter: (op: HeaderFooterOp) => t.invoke('slides:apply-header-footer', op),
  getHeaderFooter: (slideIndex: number) =>
    t.invoke('slides:get-header-footer', slideIndex),
  applyTheme: (op: ApplyThemeOp) => t.invoke('slides:apply-theme', op),
  setTransition: (op: SetTransitionOp) => t.invoke('slides:set-transition', op),
  getTransition: (slideIndex: number) => t.invoke('slides:get-transition', slideIndex),
  setAdvanceTimes: (op: SetAdvanceTimesOp) => t.invoke('slides:set-advance-times', op),
  getAnimations: (slideIndex: number) => t.invoke('slides:get-animations', slideIndex),
  getShapeKeys: (slideIndex: number) => t.invoke('slides:get-shape-keys', slideIndex),
  setAnimations: (op: SetAnimationsOp) => t.invoke('slides:set-animations', op),
  setSlideHidden: (op: SetSlideHiddenOp) => t.invoke('slides:set-hidden', op),
  getSections: () => t.invoke('slides:get-sections'),
  setSections: (sections: SectionInfo[]) => t.invoke('slides:set-sections', sections),
  addSection: (op: AddSectionOp) => t.invoke('slides:add-section', op),
  renameSection: (op: RenameSectionOp) => t.invoke('slides:rename-section', op),
  removeSection: (op: RemoveSectionOp) => t.invoke('slides:remove-section', op),
  moveSection: (op: MoveSectionOp) => t.invoke('slides:move-section', op),
  moveSlide: (op: MoveSlideOp) => t.invoke('slides:move-slide', op),
  getNotes: (slideIndex: number) => t.invoke('slides:get-notes', slideIndex),
  setNotes: (op) => t.invoke('slides:set-notes', op),
  getComments: (slideIndex: number) => t.invoke('slides:get-comments', slideIndex),
  addComment: (op) => t.invoke('slides:add-comment', op),
  deleteComment: (op) => t.invoke('slides:delete-comment', op),
  nativeClipboard: (op: 'cut' | 'copy' | 'paste') =>
    t.invoke('slides:native-clipboard', op),
  beginHistoryBatch: () => t.invoke('slides:history-batch-begin'),
  endHistoryBatch: () => t.invoke('slides:history-batch-end'),
  applyEditScript: (op: ApplyEditScriptOp) => t.invoke('slides:apply-edit-script', op),
  applyTxn: (op: ApplyTxnOp) => t.invoke('slides:apply-txn', op),
  aiSnapshotRestore: (id: number) => t.invoke('slides:ai-snapshot-restore', id),
  undo: () => t.invoke('slides:undo'),
  redo: () => t.invoke('slides:redo'),
  pickExportDir: overrides.pickExportDir ?? (() => t.invoke('slides:pick-export-dir')),
  exportImages: overrides.exportImages ?? ((op: ExportImagesOp) => t.invoke('slides:export-images', op)),
  pickExportPdfPath: overrides.pickExportPdfPath ??
    ((defaultName: string) => t.invoke('slides:pick-export-pdf-path', defaultName)),
  exportPdf: overrides.exportPdf ?? ((op: ExportPdfOp) => t.invoke('slides:export-pdf', op)),
  printSlides: overrides.printSlides ?? ((op: PrintSlidesOp) => t.invoke('slides:print', op)),
  save: () => t.invoke('slides:save'),
  saveAs: (defaultName: string) => t.invoke('slides:save-as', defaultName),
  onCloseSaveRequest: (handler: () => void) =>
    t.on('slides:close-save-request', () => handler()),
  onHistoryChanged: (handler: (state: { canUndo: boolean; canRedo: boolean }) => void) =>
    t.on('slides:history-changed', (state) => handler(state as { canUndo: boolean; canRedo: boolean })),
  onDeckChanged: (
    handler: (state: { slides: RenderSlide[]; size: { cx: number; cy: number } }) => void,
  ) =>
    t.on('slides:deck-changed', (state) =>
      handler(state as { slides: RenderSlide[]; size: { cx: number; cy: number } }),
    ),
  reportCloseSaveResult: (ok: boolean) => t.send('slides:close-save-result', ok === true),
  setAutoSavePref: (on: boolean) => t.send('slides:autosave-pref', on === true),
  isDirty: () => t.invoke('slides:is-dirty'),
  getRecentFiles: () => t.invoke('slides:recent'),
  onMenuCommand: (handler: (command: MenuCommand) => void) =>
    t.on('slides:menu', (cmd) => handler(cmd as MenuCommand)),
  onOpened: (handler: (result: OpenResult) => void) =>
    t.on('slides:opened', (result) => handler(result as OpenResult)),
  onRenamed: (handler: (newPath: string) => void) =>
    t.on('slides:renamed', (newPath) => handler(newPath as string)),
  getAiSettings: () => t.invoke('ai:get-settings'),
  setAiSettings: (settings: AiSettings) => t.invoke('ai:set-settings', settings),
  aiStream: (request: AiStreamRequest) => t.invoke('ai:stream', request),
  aiStreamCancel: (requestId: string) => t.invoke('ai:stream-cancel', requestId),
  aiGskStatus: (withEmail?: boolean) => t.invoke('ai:gsk-status', withEmail),
  aiGskLogin: () => t.invoke('ai:gsk-login'),
  aiLogRunFailure: (entry: AiRunFailure) => t.invoke('ai:log-run-failure', entry),
  webSearch: (query: string, maxResults?: number) =>
    t.invoke('ai:web-search', query, maxResults),
  imageSearch: (query: string, maxResults?: number) =>
    t.invoke('ai:image-search', query, maxResults),
  insertImageUrl: (op: {
    slideIndex: number
    url: string
    xPx: number
    yPx: number
    wPx: number
    hPx: number
    fitWidthPx: number
  }) => t.invoke('ai:insert-image-url', op),
  replacePictureUrl: (op: {
    slideIndex: number
    sourceId: string
    url: string
    keepSrcRect?: boolean
  }) => t.invoke('ai:replace-picture-url', op),
  generateImage: (op: {
    prompt: string
    model?: string
    referenceImageUrls?: string[]
    aspectRatio?: string
    imageSize?: string
  }) => t.invoke('ai:generate-image', op),
  analyzeMedia: (op: { mediaUrls: string[]; requirements: string }) =>
    t.invoke('ai:analyze-media', op),
  gskStatus: () => t.invoke('ai:gsk-status'),
  onAiStream: (handler: (chunk: AiStreamChunk) => void) =>
    t.on('ai:stream-chunk', (chunk) => handler(chunk as AiStreamChunk)),
  saveStyleSidecar: (data: { topic: string; styleSkill: string; createdAt: string }) =>
    t.invoke('ai:save-sidecar', data),
  saveStyleTemplate: (
    name: string,
    data: { topic: string; styleSkill: string; createdAt: string },
  ) => t.invoke('ai:save-style-template', name, data),
  listStyleTemplates: () => t.invoke('ai:list-style-templates'),
  loadStyleTemplate: (name: string) => t.invoke('ai:load-style-template', name),
  presenterStart: () => t.invoke('slides:presenter-start'),
  presenterSync: (state: ShowSyncState) => t.send('slides:presenter-sync', state),
  presenterInk: (ev: ShowInkEvent) => t.send('slides:presenter-ink', ev),
  presenterSwap: () => t.invoke('slides:presenter-swap'),
  presenterEnd: () => t.invoke('slides:presenter-end'),
  audienceReady: () => t.invoke('slides:audience-ready'),
  audienceNav: (action: AudienceNavAction) => t.send('slides:audience-nav', action),
  onShowSync: (handler: (state: ShowSyncState) => void) =>
    t.on('slides:show-sync', (state) => handler(state as ShowSyncState)),
  onShowInk: (handler: (ev: ShowInkEvent) => void) =>
    t.on('slides:show-ink', (ev) => handler(ev as ShowInkEvent)),
  onAudienceNav: (handler: (action: AudienceNavAction) => void) =>
    t.on('slides:audience-nav', (action) => handler(action as AudienceNavAction)),
  }
}
export function createSlidesFilesApi(
  t: IpcTransport,
  overrides: SlidesApiOverrides = {},
): DesktopFilesApi {
  return {
  pickAttachments: overrides.pickAttachments ?? (() => t.invoke('slides:files-pick')),
  addAttachmentPaths: (paths: string[]) => t.invoke('slides:files-add', paths),
  addPastedImage: (data: ArrayBuffer, ext: string) =>
    t.invoke('slides:files-add-pasted-image', data, ext),
  readAttachment: (path: string, offset: number, maxChars: number) =>
    t.invoke('slides:files-read', path, offset, maxChars),
  readAttachmentImage: (path: string) => t.invoke('slides:files-read-image', path),
  getPathForFile: (file: File) => overrides.getPathForFile?.(file) ??
      (() => { throw new Error("WEB_UNSUPPORTED: resolving dropped files needs the desktop file picker") })(),
  }
}
export function createSlidesProjectApi(t: IpcTransport): ProjectApi {
  return {
  resolveChat: (args) => t.invoke('project:resolveChat', args),
  appendChat: (args) => t.invoke('project:appendChat', args),
  loadChat: (args) => t.invoke('project:loadChat', args),
  rebindChat: (args) => t.invoke('project:rebindChat', args),
  // P1 extensions
  listProjects: () => t.invoke('project:list'),
  createProject: (args) => t.invoke('project:create', args),
  renameProject: (args) => t.invoke('project:rename', args),
  deleteProject: (args) => t.invoke('project:delete', args),
  moveFile: (args) => t.invoke('project:moveFile', args),
  getTimeline: (args) => t.invoke('project:timeline', args),
  }
}
