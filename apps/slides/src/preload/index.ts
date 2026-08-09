import { contextBridge, ipcRenderer, webUtils } from 'electron'
import type { IpcRendererEvent } from 'electron'
import type { ProjectApi } from '@genoffice/project-store'
import type {
  AddChartOp,
  AddElementOp,
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
} from '../shared/ipc'

const api: SlidesApi = {
  getLanguage: () => ipcRenderer.invoke('app:get-language'),
  onLanguageChanged: (handler) => {
    const listener = (
      _event: IpcRendererEvent,
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => handler(lang)
    ipcRenderer.on('app:language-changed', listener)
    return () => ipcRenderer.removeListener('app:language-changed', listener)
  },
  getTheme: () => ipcRenderer.invoke('app:get-theme'),
  onThemeChanged: (handler) => {
    const listener = (_event: IpcRendererEvent, theme: UiTheme) => handler(theme)
    ipcRenderer.on('app:theme-changed', listener)
    return () => ipcRenderer.removeListener('app:theme-changed', listener)
  },
  openPptx: (fitWidthPx) => ipcRenderer.invoke('slides:open', fitWidthPx),
  openPptxPath: (path, fitWidthPx) => ipcRenderer.invoke('slides:open-path', path, fitWidthPx),
  consumePendingOpen: (fitWidthPx) => ipcRenderer.invoke('slides:consume-pending-open', fitWidthPx),
  newBlank: (fitWidthPx) => ipcRenderer.invoke('slides:new-blank', fitWidthPx),
  htmlToPptx: (
    pagesHtml: string[],
    fitWidthPx: number,
    mode?: 'replace' | 'append' | 'replace_at' | 'insert_at',
    atIndex?: number,
    deckName?: string,
  ) => ipcRenderer.invoke('slides:html-to-pptx', pagesHtml, fitWidthPx, mode, atIndex, deckName),
  cloudGenStatus: () => ipcRenderer.invoke('slides:cloud-gen-status'),
  cloudGeneratePage: (op: {
    brief: string
    title?: string
    styleSkill?: string
    deckContext?: Record<string, unknown>
    images?: { url: string; caption?: string }[]
    width?: number
    height?: number
  }) => ipcRenderer.invoke('slides:cloud-page-generate', op),
  editText: (op: EditTextOp) => ipcRenderer.invoke('slides:edit-text', op),
  setElementFont: (op: SetElementFontOp) => ipcRenderer.invoke('slides:set-element-font', op),
  setElementParagraphFormat: (op: SetElementParagraphFormatOp) =>
    ipcRenderer.invoke('slides:set-element-paragraph-format', op),
  findReplace: (op: FindReplaceOp) => ipcRenderer.invoke('slides:find-replace', op),
  setSlideLayout: (op: SetSlideLayoutOp) => ipcRenderer.invoke('slides:set-slide-layout', op),
  setSlideSize: (op: SetSlideSizeOp) => ipcRenderer.invoke('slides:set-slide-size', op),
  getSlideSize: () => ipcRenderer.invoke('slides:get-slide-size'),
  editTransform: (op: EditTransformOp) => ipcRenderer.invoke('slides:edit-transform', op),
  editConnectorEndpoints: (op: EditConnectorEndpointsOp) =>
    ipcRenderer.invoke('slides:edit-connector-endpoints', op),
  editPictureSrcRect: (op: EditPictureSrcRectOp) =>
    ipcRenderer.invoke('slides:edit-picture-src-rect', op),
  editPictureOpacity: (op: EditPictureOpacityOp) =>
    ipcRenderer.invoke('slides:edit-picture-opacity', op),
  editImageFill: (op: { slideIndex: number; sourceId: string }) =>
    ipcRenderer.invoke('slides:edit-image-fill', op),
  setTextAnchor: (op: {
    slideIndex: number
    sourceId: string
    anchor: 'top' | 'middle' | 'bottom'
  }) => ipcRenderer.invoke('slides:set-text-anchor', op),
  clipboardExternal: () => ipcRenderer.invoke('slides:clipboard-external'),
  groupElements: (op: GroupElementsOp) => ipcRenderer.invoke('slides:group-elements', op),
  ungroupElement: (op: UngroupElementOp) => ipcRenderer.invoke('slides:ungroup-element', op),
  batchEditTransform: (op: BatchEditTransformOp) =>
    ipcRenderer.invoke('slides:batch-edit-transform', op),
  getRenderSlides: () => ipcRenderer.invoke('slides:get-render-slides'),
  addElement: (op: AddElementOp) => ipcRenderer.invoke('slides:add-element', op),
  deleteElement: (op: DeleteElementOp) => ipcRenderer.invoke('slides:delete-element', op),
  addSlide: (op: AddSlideOp) => ipcRenderer.invoke('slides:add-slide', op),
  addBlankSlide: (op: AddBlankSlideOp) => ipcRenderer.invoke('slides:add-blank-slide', op),
  addSlideWithLayout: (op: AddSlideWithLayoutOp) =>
    ipcRenderer.invoke('slides:add-slide-with-layout', op),
  getLayouts: () => ipcRenderer.invoke('slides:get-layouts'),
  masterEnter: (fitWidthPx: number) => ipcRenderer.invoke('slides:master-enter', fitWidthPx),
  masterOpen: (partPath: string) => ipcRenderer.invoke('slides:master-open', partPath),
  masterClose: () => ipcRenderer.invoke('slides:master-close'),
  masterEditText: (op: MasterEditTextOp) => ipcRenderer.invoke('slides:master-edit-text', op),
  masterEditTransform: (op: MasterEditTransformOp) =>
    ipcRenderer.invoke('slides:master-edit-transform', op),
  masterEditFill: (op: MasterEditFillOp) => ipcRenderer.invoke('slides:master-edit-fill', op),
  masterEditStroke: (op: MasterEditStrokeOp) => ipcRenderer.invoke('slides:master-edit-stroke', op),
  masterDeleteElement: (op: MasterDeleteElementOp) =>
    ipcRenderer.invoke('slides:master-delete-element', op),
  editFill: (op: EditFillOp) => ipcRenderer.invoke('slides:edit-fill', op),
  editStroke: (op: EditStrokeOp) => ipcRenderer.invoke('slides:edit-stroke', op),
  flipElements: (op: FlipElementOp) => ipcRenderer.invoke('slides:flip-elements', op),
  editBackground: (op: EditBackgroundOp) => ipcRenderer.invoke('slides:edit-background', op),
  insertImage: (slideIndex: number, fitWidthPx: number) =>
    ipcRenderer.invoke('slides:insert-image', slideIndex, fitWidthPx),
  copySlide: (slideIndex: number, pngBase64?: string) =>
    ipcRenderer.invoke('slides:copy-slide', slideIndex, pngBase64),
  pasteSlide: (op: PasteSlideOp) => ipcRenderer.invoke('slides:paste-slide', op),
  repasteSlide: (op: RepasteSlideOp) => ipcRenderer.invoke('slides:repaste-slide', op),
  hasSlideClipboard: () => ipcRenderer.invoke('slides:has-slide-clipboard'),
  deleteSlide: (slideIndex: number) => ipcRenderer.invoke('slides:delete-slide', slideIndex),
  reorderElement: (op: ReorderElementOp) => ipcRenderer.invoke('slides:reorder-element', op),
  editTableCell: (op: EditTableCellOp) => ipcRenderer.invoke('slides:edit-table-cell', op),
  tableStructure: (op: TableStructureIpcOp) => ipcRenderer.invoke('slides:table-structure', op),
  tableMerge: (op: TableMergeIpcOp) => ipcRenderer.invoke('slides:table-merge', op),
  setTableColWidth: (op: SetTableColWidthOp) =>
    ipcRenderer.invoke('slides:set-table-col-width', op),
  setTableRowHeight: (op: SetTableRowHeightOp) =>
    ipcRenderer.invoke('slides:set-table-row-height', op),
  setTableCellAnchor: (op: SetTableCellAnchorOp) =>
    ipcRenderer.invoke('slides:set-table-cell-anchor', op),
  editTableStyle: (op: EditTableStyleOp) => ipcRenderer.invoke('slides:edit-table-style', op),
  editChart: (op: EditChartOp) => ipcRenderer.invoke('slides:edit-chart', op),
  getChartColorSchemes: () => ipcRenderer.invoke('slides:chart-color-schemes'),
  getChartData: (slideIndex: number, sourceId: string) =>
    ipcRenderer.invoke('slides:get-chart-data', slideIndex, sourceId),
  copyElements: (op: CopyElementsOp) => ipcRenderer.invoke('slides:copy-elements', op),
  pasteElements: (op: PasteElementsOp) => ipcRenderer.invoke('slides:paste-elements', op),
  duplicateElements: (op: DuplicateElementsOp) =>
    ipcRenderer.invoke('slides:duplicate-elements', op),
  addTable: (op: AddTableOp) => ipcRenderer.invoke('slides:add-table', op),
  addInk: (op: AddInkOp) => ipcRenderer.invoke('slides:add-ink', op),
  addChart: (op: AddChartOp) => ipcRenderer.invoke('slides:add-chart', op),
  addSmartArt: (op: AddSmartArtOp) => ipcRenderer.invoke('slides:add-smartart', op),
  addImageBytes: (op: AddImageBytesOp) => ipcRenderer.invoke('slides:add-image-bytes', op),
  replacePictureBytes: (op: ReplacePictureBytesOp) =>
    ipcRenderer.invoke('slides:replace-picture-bytes', op),
  insertMedia: (slideIndex: number, kind: 'video' | 'audio', fitWidthPx: number) =>
    ipcRenderer.invoke('slides:insert-media', slideIndex, kind, fitWidthPx),
  addMediaBytes: (op: AddMediaBytesOp) => ipcRenderer.invoke('slides:add-media-bytes', op),
  getMediaData: (slideIndex: number, sourceId: string) =>
    ipcRenderer.invoke('slides:media-data', slideIndex, sourceId),
  insertModel3d: (slideIndex: number, fitWidthPx: number) =>
    ipcRenderer.invoke('slides:insert-model3d', slideIndex, fitWidthPx),
  setLink: (op: SetLinkOp) => ipcRenderer.invoke('slides:set-link', op),
  getLink: (slideIndex: number, sourceId: string) =>
    ipcRenderer.invoke('slides:get-link', slideIndex, sourceId),
  getSlideLinks: (slideIndex: number) => ipcRenderer.invoke('slides:get-slide-links', slideIndex),
  getRunLinks: (slideIndex: number) => ipcRenderer.invoke('slides:get-run-links', slideIndex),
  applyHeaderFooter: (op: HeaderFooterOp) => ipcRenderer.invoke('slides:apply-header-footer', op),
  getHeaderFooter: (slideIndex: number) =>
    ipcRenderer.invoke('slides:get-header-footer', slideIndex),
  applyTheme: (op: ApplyThemeOp) => ipcRenderer.invoke('slides:apply-theme', op),
  setTransition: (op: SetTransitionOp) => ipcRenderer.invoke('slides:set-transition', op),
  getTransition: (slideIndex: number) => ipcRenderer.invoke('slides:get-transition', slideIndex),
  setAdvanceTimes: (op: SetAdvanceTimesOp) => ipcRenderer.invoke('slides:set-advance-times', op),
  getAnimations: (slideIndex: number) => ipcRenderer.invoke('slides:get-animations', slideIndex),
  getShapeKeys: (slideIndex: number) => ipcRenderer.invoke('slides:get-shape-keys', slideIndex),
  setAnimations: (op: SetAnimationsOp) => ipcRenderer.invoke('slides:set-animations', op),
  setSlideHidden: (op: SetSlideHiddenOp) => ipcRenderer.invoke('slides:set-hidden', op),
  getSections: () => ipcRenderer.invoke('slides:get-sections'),
  setSections: (sections: SectionInfo[]) => ipcRenderer.invoke('slides:set-sections', sections),
  addSection: (op: AddSectionOp) => ipcRenderer.invoke('slides:add-section', op),
  renameSection: (op: RenameSectionOp) => ipcRenderer.invoke('slides:rename-section', op),
  removeSection: (op: RemoveSectionOp) => ipcRenderer.invoke('slides:remove-section', op),
  moveSection: (op: MoveSectionOp) => ipcRenderer.invoke('slides:move-section', op),
  moveSlide: (op: MoveSlideOp) => ipcRenderer.invoke('slides:move-slide', op),
  getNotes: (slideIndex: number) => ipcRenderer.invoke('slides:get-notes', slideIndex),
  setNotes: (op) => ipcRenderer.invoke('slides:set-notes', op),
  getComments: (slideIndex: number) => ipcRenderer.invoke('slides:get-comments', slideIndex),
  addComment: (op) => ipcRenderer.invoke('slides:add-comment', op),
  deleteComment: (op) => ipcRenderer.invoke('slides:delete-comment', op),
  nativeClipboard: (op: 'cut' | 'copy' | 'paste') =>
    ipcRenderer.invoke('slides:native-clipboard', op),
  beginHistoryBatch: () => ipcRenderer.invoke('slides:history-batch-begin'),
  endHistoryBatch: () => ipcRenderer.invoke('slides:history-batch-end'),
  aiSnapshotRestore: (id: number) => ipcRenderer.invoke('slides:ai-snapshot-restore', id),
  undo: () => ipcRenderer.invoke('slides:undo'),
  redo: () => ipcRenderer.invoke('slides:redo'),
  pickExportDir: () => ipcRenderer.invoke('slides:pick-export-dir'),
  exportImages: (op: ExportImagesOp) => ipcRenderer.invoke('slides:export-images', op),
  pickExportPdfPath: (defaultName: string) =>
    ipcRenderer.invoke('slides:pick-export-pdf-path', defaultName),
  exportPdf: (op: ExportPdfOp) => ipcRenderer.invoke('slides:export-pdf', op),
  printSlides: (op: PrintSlidesOp) => ipcRenderer.invoke('slides:print', op),
  save: () => ipcRenderer.invoke('slides:save'),
  saveAs: (defaultName: string) => ipcRenderer.invoke('slides:save-as', defaultName),
  onCloseSaveRequest: (handler: () => void) => {
    const listener = () => handler()
    ipcRenderer.on('slides:close-save-request', listener)
    return () => ipcRenderer.removeListener('slides:close-save-request', listener)
  },
  reportCloseSaveResult: (ok: boolean) => ipcRenderer.send('slides:close-save-result', ok === true),
  setAutoSavePref: (on: boolean) => ipcRenderer.send('slides:autosave-pref', on === true),
  isDirty: () => ipcRenderer.invoke('slides:is-dirty'),
  getRecentFiles: () => ipcRenderer.invoke('slides:recent'),
  onMenuCommand: (handler: (command: MenuCommand) => void) => {
    const listener = (_e: IpcRendererEvent, cmd: MenuCommand) => handler(cmd)
    ipcRenderer.on('slides:menu', listener)
    return () => ipcRenderer.removeListener('slides:menu', listener)
  },
  onOpened: (handler: (result: OpenResult) => void) => {
    const listener = (_e: IpcRendererEvent, result: OpenResult) => handler(result)
    ipcRenderer.on('slides:opened', listener)
    return () => ipcRenderer.removeListener('slides:opened', listener)
  },
  onRenamed: (handler: (newPath: string) => void) => {
    const listener = (_e: IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on('slides:renamed', listener)
    return () => ipcRenderer.removeListener('slides:renamed', listener)
  },
  getAiSettings: () => ipcRenderer.invoke('ai:get-settings'),
  setAiSettings: (settings: AiSettings) => ipcRenderer.invoke('ai:set-settings', settings),
  aiStream: (request: AiStreamRequest) => ipcRenderer.invoke('ai:stream', request),
  aiStreamCancel: (requestId: string) => ipcRenderer.invoke('ai:stream-cancel', requestId),
  aiGskStatus: (withEmail?: boolean) => ipcRenderer.invoke('ai:gsk-status', withEmail),
  aiGskLogin: () => ipcRenderer.invoke('ai:gsk-login'),
  webSearch: (query: string, maxResults?: number) =>
    ipcRenderer.invoke('ai:web-search', query, maxResults),
  imageSearch: (query: string, maxResults?: number) =>
    ipcRenderer.invoke('ai:image-search', query, maxResults),
  insertImageUrl: (op: {
    slideIndex: number
    url: string
    xPx: number
    yPx: number
    wPx: number
    hPx: number
    fitWidthPx: number
  }) => ipcRenderer.invoke('ai:insert-image-url', op),
  replacePictureUrl: (op: {
    slideIndex: number
    sourceId: string
    url: string
    keepSrcRect?: boolean
  }) => ipcRenderer.invoke('ai:replace-picture-url', op),
  generateImage: (op: {
    prompt: string
    model?: string
    referenceImageUrls?: string[]
    aspectRatio?: string
    imageSize?: string
  }) => ipcRenderer.invoke('ai:generate-image', op),
  analyzeMedia: (op: { mediaUrls: string[]; requirements: string }) =>
    ipcRenderer.invoke('ai:analyze-media', op),
  gskStatus: () => ipcRenderer.invoke('ai:gsk-status'),
  onAiStream: (handler: (chunk: AiStreamChunk) => void) => {
    const listener = (_e: IpcRendererEvent, chunk: AiStreamChunk) => handler(chunk)
    ipcRenderer.on('ai:stream-chunk', listener)
    return () => ipcRenderer.removeListener('ai:stream-chunk', listener)
  },
  saveStyleSidecar: (data: { topic: string; styleSkill: string; createdAt: string }) =>
    ipcRenderer.invoke('ai:save-sidecar', data),
  saveStyleTemplate: (
    name: string,
    data: { topic: string; styleSkill: string; createdAt: string },
  ) => ipcRenderer.invoke('ai:save-style-template', name, data),
  listStyleTemplates: () => ipcRenderer.invoke('ai:list-style-templates'),
  loadStyleTemplate: (name: string) => ipcRenderer.invoke('ai:load-style-template', name),
  presenterStart: () => ipcRenderer.invoke('slides:presenter-start'),
  presenterSync: (state: ShowSyncState) => ipcRenderer.send('slides:presenter-sync', state),
  presenterInk: (ev: ShowInkEvent) => ipcRenderer.send('slides:presenter-ink', ev),
  presenterSwap: () => ipcRenderer.invoke('slides:presenter-swap'),
  presenterEnd: () => ipcRenderer.invoke('slides:presenter-end'),
  audienceReady: () => ipcRenderer.invoke('slides:audience-ready'),
  audienceNav: (action: AudienceNavAction) => ipcRenderer.send('slides:audience-nav', action),
  onShowSync: (handler: (state: ShowSyncState) => void) => {
    const listener = (_e: IpcRendererEvent, state: ShowSyncState) => handler(state)
    ipcRenderer.on('slides:show-sync', listener)
    return () => ipcRenderer.removeListener('slides:show-sync', listener)
  },
  onShowInk: (handler: (ev: ShowInkEvent) => void) => {
    const listener = (_e: IpcRendererEvent, ev: ShowInkEvent) => handler(ev)
    ipcRenderer.on('slides:show-ink', listener)
    return () => ipcRenderer.removeListener('slides:show-ink', listener)
  },
  onAudienceNav: (handler: (action: AudienceNavAction) => void) => {
    const listener = (_e: IpcRendererEvent, action: AudienceNavAction) => handler(action)
    ipcRenderer.on('slides:audience-nav', listener)
    return () => ipcRenderer.removeListener('slides:audience-nav', listener)
  },
}

contextBridge.exposeInMainWorld('slidesApi', api)

// Chat attachment bridge: method names/signatures match the window.desktop attachment subset in docs, so the renderer's files-skill is copied over wholesale
const filesApi: DesktopFilesApi = {
  pickAttachments: () => ipcRenderer.invoke('slides:files-pick'),
  addAttachmentPaths: (paths: string[]) => ipcRenderer.invoke('slides:files-add', paths),
  addPastedImage: (data: ArrayBuffer, ext: string) =>
    ipcRenderer.invoke('slides:files-add-pasted-image', data, ext),
  readAttachment: (path: string, offset: number, maxChars: number) =>
    ipcRenderer.invoke('slides:files-read', path, offset, maxChars),
  readAttachmentImage: (path: string) => ipcRenderer.invoke('slides:files-read-image', path),
  getPathForFile: (file: File) => webUtils.getPathForFile(file),
}

contextBridge.exposeInMainWorld('desktop', filesApi)

const projectApi: ProjectApi = {
  resolveChat: (args) => ipcRenderer.invoke('project:resolveChat', args),
  appendChat: (args) => ipcRenderer.invoke('project:appendChat', args),
  loadChat: (args) => ipcRenderer.invoke('project:loadChat', args),
  rebindChat: (args) => ipcRenderer.invoke('project:rebindChat', args),
  // P1 extensions
  listProjects: () => ipcRenderer.invoke('project:list'),
  createProject: (args) => ipcRenderer.invoke('project:create', args),
  renameProject: (args) => ipcRenderer.invoke('project:rename', args),
  deleteProject: (args) => ipcRenderer.invoke('project:delete', args),
  moveFile: (args) => ipcRenderer.invoke('project:moveFile', args),
  getTimeline: (args) => ipcRenderer.invoke('project:timeline', args),
}
contextBridge.exposeInMainWorld('projectApi', projectApi)
