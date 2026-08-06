import { contextBridge, ipcRenderer, webUtils } from 'electron'

import type {
  AiChatResponse,
  AiSettings,
  AiStreamChunk,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'
import type { ProjectApi } from '@genoffice/project-store'
import type {
  AttachmentAddResult,
  AttachmentImageResult,
  AttachmentMeta,
  AttachmentReadResult,
  DesktopApi,
  ScreenCaptureResult,
  ScreenSourcesResult,
  WorkbookCellStyle,
  WorkbookConditionalRule,
  WorkbookFile,
  WorkbookFormulaCellsRequest,
  WorkbookFormulaCellsResult,
  WorkbookMediaRequest,
  WorkbookMediaResult,
  WorkbookPivotDefinition,
  WorkbookPivotRequest,
  WorkbookRangeRequest,
  WorkbookRangeResult,
  WorkbookRecalcRequest,
  WorkbookRecalcResult,
  WorkbookRichRun,
  WorkbookSaveRequest,
  WorkbookSaveResult,
  WorkbookVisualObject,
  WebSearchResult,
} from '../shared/desktop-api'
import { IPC_CHANNELS } from '../shared/ipc-channels'

const desktopApi: DesktopApi = {
  getLanguage: () => ipcRenderer.invoke('app:get-language'),
  onLanguageChanged(handler) {
    const listener = (
      _event: Electron.IpcRendererEvent,
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => handler(lang)
    ipcRenderer.on('app:language-changed', listener)
    return () => ipcRenderer.removeListener('app:language-changed', listener)
  },
  async selectWorkbook() {
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.selectWorkbook)
    return result === null ? null : parseWorkbookFile(result)
  },
  async readWorkbookRange(request) {
    const validatedRequest = parseRangeRequest(request)
    const result: unknown = await ipcRenderer.invoke(
      IPC_CHANNELS.readWorkbookRange,
      validatedRequest,
    )
    return parseRangeResult(result)
  },
  async readWorkbookFormulas(request) {
    const validatedRequest = parseFormulaCellsRequest(request)
    const result: unknown = await ipcRenderer.invoke(
      IPC_CHANNELS.readWorkbookFormulas,
      validatedRequest,
    )
    return parseFormulaCellsResult(result)
  },
  async recalcWorkbook(request) {
    const validatedRequest = parseRecalcRequest(request)
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.recalcWorkbook, validatedRequest)
    return parseRecalcResult(result)
  },
  async readWorkbookMedia(request) {
    const validatedRequest = parseMediaRequest(request)
    const result: unknown = await ipcRenderer.invoke(
      IPC_CHANNELS.readWorkbookMedia,
      validatedRequest,
    )
    return parseMediaResult(result)
  },
  async readLocalImage(request) {
    if (
      !isRecord(request) ||
      typeof request.path !== 'string' ||
      request.path.length === 0 ||
      request.path.length > 1024
    ) {
      throw new Error('Invalid local image request.')
    }
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.readLocalImage, request)
    if (
      !isRecord(result) ||
      typeof result.mediaType !== 'string' ||
      !['image/png', 'image/jpeg', 'image/gif'].includes(result.mediaType) ||
      typeof result.base64 !== 'string' ||
      result.base64.length === 0
    ) {
      throw new Error('Invalid local image response.')
    }
    return result as { mediaType: 'image/png' | 'image/jpeg' | 'image/gif'; base64: string }
  },
  async captureScreenSources() {
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.captureScreenSources)
    if (
      !isRecord(result) ||
      (result.status !== 'ok' && result.status !== 'denied') ||
      !Array.isArray(result.sources)
    ) {
      throw new Error('Invalid screen sources response.')
    }
    return result as ScreenSourcesResult
  },
  async captureScreenSource(request) {
    if (!isRecord(request) || typeof request.id !== 'string' || request.id.length === 0) {
      throw new Error('Invalid screen capture request.')
    }
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.captureScreenSource, {
      id: request.id,
    })
    if (result === null) return null
    if (
      !isRecord(result) ||
      result.mediaType !== 'image/png' ||
      typeof result.base64 !== 'string' ||
      result.base64.length === 0 ||
      typeof result.width !== 'number' ||
      typeof result.height !== 'number'
    ) {
      throw new Error('Invalid screen capture response.')
    }
    return result as ScreenCaptureResult
  },
  async readPivotDefinition(request) {
    const validatedRequest = parsePivotRequest(request)
    const result: unknown = await ipcRenderer.invoke(
      IPC_CHANNELS.readPivotDefinition,
      validatedRequest,
    )
    return parsePivotDefinitionResult(result)
  },
  async saveWorkbookEdits(request) {
    const validatedRequest = parseSaveRequest(request)
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.saveWorkbook, validatedRequest)
    return parseSaveResult(result)
  },
  async writeWorkbookRecovery(request) {
    const validatedRequest = parseSaveRequest(request)
    const result: unknown = await ipcRenderer.invoke(
      IPC_CHANNELS.writeWorkbookRecovery,
      validatedRequest,
    )
    return { ok: (result as { ok?: unknown } | null)?.ok === true }
  },
  async autoRenameWorkbook(sessionId, baseName) {
    if (!isUuid(sessionId)) throw new Error('Invalid workbook session.')
    if (typeof baseName !== 'string' || baseName.length === 0 || baseName.length > 100) {
      throw new Error('Invalid workbook name.')
    }
    const result: unknown = await ipcRenderer.invoke(
      IPC_CHANNELS.autoRenameWorkbook,
      sessionId,
      baseName,
    )
    if (!isRecord(result) || typeof result.renamed !== 'boolean') {
      throw new Error('Invalid auto-rename response.')
    }
    return result as { renamed: boolean; name?: string }
  },
  async exportPdf(request) {
    if (
      !isRecord(request) ||
      typeof request.fileName !== 'string' ||
      request.fileName.length === 0 ||
      request.fileName.length > 255 ||
      typeof request.html !== 'string' ||
      request.html.length === 0 ||
      request.html.length > 20_000_000 ||
      typeof request.landscape !== 'boolean' ||
      !isPdfPageSize(request.pageSize) ||
      !isRecord(request.margins) ||
      !['top', 'bottom', 'left', 'right'].every((edge) => {
        const value = (request.margins as Record<string, unknown>)[edge]
        return typeof value === 'number' && value >= 0 && value <= 3
      }) ||
      typeof request.scale !== 'number' ||
      request.scale < 0.1 ||
      request.scale > 2
    ) {
      throw new Error('Invalid PDF export request.')
    }
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.exportPdf, request)
    if (
      !isRecord(result) ||
      typeof result.canceled !== 'boolean' ||
      (result.canceled === false && typeof result.path !== 'string')
    ) {
      throw new Error('Invalid PDF export response.')
    }
    return result as { canceled: true } | { canceled: false; path: string }
  },
  async closeWorkbook(sessionId) {
    if (!isUuid(sessionId)) throw new Error('Invalid workbook session.')
    await ipcRenderer.invoke(IPC_CHANNELS.closeWorkbook, sessionId)
  },
  async openExternal(url) {
    if (typeof url !== 'string' || !/^https?:\/\//.test(url)) {
      throw new Error('Only http(s) links can be opened.')
    }
    await ipcRenderer.invoke(IPC_CHANNELS.openExternal, url)
  },
  onMenuAction(callback) {
    const listener = (_event: unknown, action: unknown): void => {
      if (
        action === 'open' ||
        action === 'save' ||
        action === 'save-as' ||
        action === 'export-pdf' ||
        action === 'undo' ||
        action === 'redo'
      )
        callback(action)
    }
    ipcRenderer.on(IPC_CHANNELS.menuAction, listener)
    return () => ipcRenderer.removeListener(IPC_CHANNELS.menuAction, listener)
  },
  onWorkbookRenamed(callback) {
    const listener = (_event: unknown, newName: unknown): void => {
      if (typeof newName === 'string' && newName) callback(newName)
    }
    ipcRenderer.on(IPC_CHANNELS.workbookRenamed, listener)
    return () => ipcRenderer.removeListener(IPC_CHANNELS.workbookRenamed, listener)
  },
  notifyPendingEdits(count) {
    if (typeof count !== 'number' || !Number.isFinite(count) || count < 0) return
    ipcRenderer.send(IPC_CHANNELS.pendingEditsChanged, Math.floor(count))
  },
  onCloseSaveRequest(callback) {
    const listener = (): void => callback()
    ipcRenderer.on(IPC_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(IPC_CHANNELS.closeSaveRequest, listener)
  },
  reportCloseSaveResult(ok) {
    ipcRenderer.send(IPC_CHANNELS.closeSaveResult, ok === true)
  },
  async getAiSettings() {
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.aiGetSettings)
    if (!isRecord(result)) throw new Error('Invalid AI settings response.')
    return result as unknown as AiSettings
  },
  async setAiSettings(settings) {
    await ipcRenderer.invoke(IPC_CHANNELS.aiSetSettings, settings)
  },
  async aiChat(request) {
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.aiChat, request)
    if (!isRecord(result) || typeof result.ok !== 'boolean') {
      throw new Error('Invalid AI chat response.')
    }
    return result as unknown as AiChatResponse
  },
  async aiStream(request) {
    await ipcRenderer.invoke(IPC_CHANNELS.aiStream, request)
  },
  async aiStreamCancel(requestId) {
    if (!requestId) throw new Error('Invalid AI stream request id.')
    await ipcRenderer.invoke(IPC_CHANNELS.aiStreamCancel, requestId)
  },
  async aiGskStatus(withEmail) {
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.aiGskStatus, withEmail)
    if (!isRecord(result) || typeof result.loggedIn !== 'boolean') {
      throw new Error('Invalid Genspark account status response.')
    }
    return result as unknown as GenSparkAccountStatus
  },
  async aiGskLogin() {
    await ipcRenderer.invoke(IPC_CHANNELS.aiGskLogin)
  },
  async webSearch(query, maxResults) {
    if (typeof query !== 'string' || !query.trim() || query.length > 512) {
      throw new Error('Invalid search query.')
    }
    const result: unknown = await ipcRenderer.invoke('ai:web-search', query, maxResults)
    if (!isRecord(result) || !Array.isArray(result.results) || typeof result.method !== 'string') {
      throw new Error('Invalid web search response.')
    }
    return result as unknown as WebSearchResult
  },
  onAiStream(callback) {
    const listener = (_event: unknown, chunk: unknown): void => {
      if (
        isRecord(chunk) &&
        typeof chunk.requestId === 'string' &&
        typeof chunk.type === 'string'
      ) {
        callback(chunk as unknown as AiStreamChunk)
      }
    }
    ipcRenderer.on(IPC_CHANNELS.aiStreamChunk, listener)
    return () => ipcRenderer.removeListener(IPC_CHANNELS.aiStreamChunk, listener)
  },
  async consumeNewBlankWorkbook() {
    const result: unknown = await ipcRenderer.invoke('sheets:consume-new-blank')
    return result === true
  },
  async pickAttachments() {
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.filesPick)
    return result === null ? null : parseAttachmentAddResult(result)
  },
  async addAttachmentPaths(paths) {
    if (
      !Array.isArray(paths) ||
      paths.length === 0 ||
      paths.length > 50 ||
      paths.some((p) => typeof p !== 'string' || p.length === 0 || p.length > 1024)
    ) {
      throw new Error('Invalid attachment paths.')
    }
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.filesAdd, paths)
    return parseAttachmentAddResult(result)
  },
  async addPastedImage(data, ext) {
    if (
      !(data instanceof ArrayBuffer) ||
      data.byteLength === 0 ||
      data.byteLength > 64 * 1024 * 1024
    ) {
      throw new Error('Invalid pasted image data.')
    }
    if (typeof ext !== 'string' || ext.length === 0 || ext.length > 8) {
      throw new Error('Invalid pasted image extension.')
    }
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.filesAddPastedImage, data, ext)
    return parseAttachmentAddResult(result)
  },
  async readAttachment(path, offset, maxChars) {
    if (typeof path !== 'string' || path.length === 0 || path.length > 1024) {
      throw new Error('Invalid attachment path.')
    }
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.filesRead, path, offset, maxChars)
    if (
      !isRecord(result) ||
      typeof result.ok !== 'boolean' ||
      !isOptionalString(result.error) ||
      !isOptionalString(result.name) ||
      !isOptionalString(result.text) ||
      (result.totalChars !== undefined && !isNonnegativeInteger(result.totalChars)) ||
      (result.offset !== undefined && !isNonnegativeInteger(result.offset))
    ) {
      throw new Error('Invalid attachment read response.')
    }
    const read: AttachmentReadResult = { ok: result.ok }
    if (result.error !== undefined) read.error = result.error
    if (result.name !== undefined) read.name = result.name
    if (result.text !== undefined) read.text = result.text
    if (result.totalChars !== undefined) read.totalChars = result.totalChars
    if (result.offset !== undefined) read.offset = result.offset
    return read
  },
  async readAttachmentImage(path) {
    if (typeof path !== 'string' || path.length === 0 || path.length > 1024) {
      throw new Error('Invalid attachment path.')
    }
    const result: unknown = await ipcRenderer.invoke(IPC_CHANNELS.filesReadImage, path)
    if (
      !isRecord(result) ||
      typeof result.ok !== 'boolean' ||
      !isOptionalString(result.base64) ||
      !isOptionalString(result.mime) ||
      !isOptionalString(result.error)
    ) {
      throw new Error('Invalid attachment image response.')
    }
    const image: AttachmentImageResult = { ok: result.ok }
    if (result.base64 !== undefined) image.base64 = result.base64
    if (result.mime !== undefined) image.mime = result.mime
    if (result.error !== undefined) image.error = result.error
    return image
  },
  getPathForFile(file) {
    return webUtils.getPathForFile(file)
  },
}

function parseAttachmentAddResult(input: unknown): AttachmentAddResult {
  if (
    !isRecord(input) ||
    !Array.isArray(input.accepted) ||
    input.accepted.length > 50 ||
    !Array.isArray(input.rejected) ||
    input.rejected.length > 50 ||
    input.rejected.some((reason) => typeof reason !== 'string')
  ) {
    throw new Error('Invalid attachment response.')
  }
  const accepted = input.accepted.map((meta): AttachmentMeta => {
    if (
      !isRecord(meta) ||
      typeof meta.path !== 'string' ||
      meta.path.length === 0 ||
      typeof meta.name !== 'string' ||
      meta.name.length === 0 ||
      typeof meta.ext !== 'string' ||
      !isNonnegativeInteger(meta.sizeBytes)
    ) {
      throw new Error('Invalid attachment metadata.')
    }
    return { path: meta.path, name: meta.name, ext: meta.ext, sizeBytes: meta.sizeBytes }
  })
  return { accepted, rejected: input.rejected as string[] }
}

contextBridge.exposeInMainWorld('desktopApi', desktopApi)

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

function parseWorkbookFile(input: unknown): WorkbookFile {
  if (!isRecord(input)) throw new Error('Invalid workbook response.')
  const {
    sessionId,
    name,
    sha256,
    entryCount,
    sheets,
    styles,
    dxfStyles,
    visuals,
    definedNames,
    readOnly,
  } = input
  if (
    !isUuid(sessionId) ||
    typeof name !== 'string' ||
    typeof sha256 !== 'string' ||
    !/^[a-f0-9]{64}$/.test(sha256) ||
    !isNonnegativeInteger(entryCount) ||
    !Array.isArray(sheets) ||
    !Array.isArray(styles) ||
    !Array.isArray(dxfStyles) ||
    !Array.isArray(visuals) ||
    !Array.isArray(definedNames) ||
    typeof readOnly !== 'boolean'
  ) {
    throw new Error('Invalid workbook response.')
  }
  const parsedDefinedNames = definedNames.map((entry) => {
    if (
      !isRecord(entry) ||
      typeof entry.name !== 'string' ||
      entry.name.length === 0 ||
      typeof entry.formula !== 'string' ||
      entry.formula.length === 0 ||
      (entry.sheetIndex !== undefined && !isNonnegativeInteger(entry.sheetIndex))
    ) {
      throw new Error('Invalid workbook defined name.')
    }
    return {
      name: entry.name,
      formula: entry.formula,
      ...(entry.sheetIndex === undefined ? {} : { sheetIndex: entry.sheetIndex }),
    }
  })
  const parsedSheets = sheets.map((sheet) => {
    if (
      !isRecord(sheet) ||
      typeof sheet.id !== 'string' ||
      typeof sheet.name !== 'string' ||
      !isPositiveInteger(sheet.rowCount) ||
      !isPositiveInteger(sheet.columnCount) ||
      !Array.isArray(sheet.columnWidths) ||
      typeof sheet.hidden !== 'boolean' ||
      (sheet.tabColor !== null && typeof sheet.tabColor !== 'string') ||
      typeof sheet.showGridLines !== 'boolean' ||
      !Array.isArray(sheet.tables) ||
      !Array.isArray(sheet.comments) ||
      !Array.isArray(sheet.pivotRanges) ||
      sheet.pivotRanges.length > 1_000
    ) {
      throw new Error('Invalid worksheet metadata.')
    }
    const tables = sheet.tables.map((table) => {
      if (
        !isRecord(table) ||
        !isNonnegativeInteger(table.headerRowCount) ||
        typeof table.showRowStripes !== 'boolean' ||
        typeof table.showColumnStripes !== 'boolean' ||
        !isOptionalString(table.styleName) ||
        !isOptionalString(table.headerFill) ||
        !isOptionalString(table.headerFontColor) ||
        !isOptionalString(table.stripeFill)
      ) {
        throw new Error('Invalid worksheet table.')
      }
      return {
        range: parseCellArea(table.range),
        headerRowCount: table.headerRowCount,
        showRowStripes: table.showRowStripes,
        showColumnStripes: table.showColumnStripes,
        ...(table.styleName === undefined ? {} : { styleName: table.styleName }),
        ...(table.headerFill === undefined ? {} : { headerFill: table.headerFill }),
        ...(table.headerFontColor === undefined ? {} : { headerFontColor: table.headerFontColor }),
        ...(table.stripeFill === undefined ? {} : { stripeFill: table.stripeFill }),
      }
    })
    const comments = sheet.comments.map((comment) => {
      if (
        !isRecord(comment) ||
        !isNonnegativeInteger(comment.row) ||
        !isNonnegativeInteger(comment.column) ||
        typeof comment.author !== 'string' ||
        typeof comment.text !== 'string'
      ) {
        throw new Error('Invalid worksheet comment.')
      }
      return {
        row: comment.row,
        column: comment.column,
        author: comment.author,
        text: comment.text,
      }
    })
    const columnWidths = sheet.columnWidths.map((columnWidth) => {
      if (
        !isRecord(columnWidth) ||
        !isNonnegativeInteger(columnWidth.startColumn) ||
        !isNonnegativeInteger(columnWidth.endColumn) ||
        columnWidth.startColumn > columnWidth.endColumn ||
        (columnWidth.width !== undefined &&
          (typeof columnWidth.width !== 'number' ||
            !Number.isFinite(columnWidth.width) ||
            columnWidth.width < 0)) ||
        typeof columnWidth.hidden !== 'boolean' ||
        (columnWidth.outlineLevel !== undefined &&
          (!isNonnegativeInteger(columnWidth.outlineLevel) || columnWidth.outlineLevel > 7)) ||
        (columnWidth.collapsed !== undefined && typeof columnWidth.collapsed !== 'boolean')
      ) {
        throw new Error('Invalid worksheet column width.')
      }
      return {
        startColumn: columnWidth.startColumn,
        endColumn: columnWidth.endColumn,
        hidden: columnWidth.hidden,
        ...(columnWidth.width === undefined ? {} : { width: columnWidth.width }),
        ...(columnWidth.outlineLevel === undefined
          ? {}
          : { outlineLevel: columnWidth.outlineLevel }),
        ...(columnWidth.collapsed === undefined ? {} : { collapsed: columnWidth.collapsed }),
      }
    })
    return {
      id: sheet.id,
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      columnWidths,
      // Degrade instead of rejecting the workbook: 0 / malformed defaults
      // mean "use the built-in size".
      defaultRowHeight: normalizedDefaultSize(sheet.defaultRowHeight),
      defaultColumnWidth: normalizedDefaultSize(sheet.defaultColumnWidth),
      freeze: parseFreeze(sheet.freeze),
      hidden: sheet.hidden,
      tabColor: sheet.tabColor,
      showGridLines: sheet.showGridLines,
      showFormulas: sheet.showFormulas === true,
      tables,
      comments,
      pivotRanges: sheet.pivotRanges.map(parseCellArea),
      pivotTables: parsePivotTableInfos(sheet.pivotTables ?? []),
      sparklines: parseSparklineGroups(sheet.sparklines ?? []),
    }
  })
  if (parsedSheets.length === 0) throw new Error('Workbook contains no worksheets.')
  return {
    sessionId,
    name,
    sha256,
    entryCount,
    sheets: parsedSheets,
    styles: styles.map(parseCellStyle),
    dxfStyles: dxfStyles.map(parseCellStyle),
    visuals: visuals.map(parseVisualObject),
    definedNames: parsedDefinedNames,
    readOnly,
  }
}

function parseRangeRequest(input: WorkbookRangeRequest): WorkbookRangeRequest {
  if (
    !isRecord(input) ||
    !isUuid(input.sessionId) ||
    typeof input.sheetId !== 'string' ||
    !isRecord(input.range)
  ) {
    throw new Error('Invalid workbook range request.')
  }
  const { startRow, endRow, startColumn, endColumn } = input.range
  if (
    !isNonnegativeInteger(startRow) ||
    !isNonnegativeInteger(endRow) ||
    !isNonnegativeInteger(startColumn) ||
    !isNonnegativeInteger(endColumn) ||
    startRow > endRow ||
    startColumn > endColumn ||
    (endRow - startRow + 1) * (endColumn - startColumn + 1) > 20_000
  ) {
    throw new Error('Invalid workbook range request.')
  }
  return input
}

function parseCellRecord(cell: unknown): WorkbookRangeResult['cells'][number] {
  if (
    !isRecord(cell) ||
    !isNonnegativeInteger(cell.row) ||
    !isNonnegativeInteger(cell.column) ||
    !isCellScalar(cell.value) ||
    (cell.formula !== undefined && typeof cell.formula !== 'string') ||
    (cell.styleIndex !== undefined && !isNonnegativeInteger(cell.styleIndex))
  ) {
    throw new Error('Invalid workbook cell response.')
  }
  const rich = cell.rich === undefined ? undefined : parseRichRuns(cell.rich)
  return {
    row: cell.row,
    column: cell.column,
    value: cell.value,
    ...(cell.formula === undefined ? {} : { formula: cell.formula }),
    ...(cell.styleIndex === undefined ? {} : { styleIndex: cell.styleIndex }),
    ...(rich === undefined ? {} : { rich }),
  }
}

function parseFormulaCellsRequest(input: unknown): WorkbookFormulaCellsRequest {
  if (
    !isRecord(input) ||
    !isUuid(input.sessionId) ||
    typeof input.sheetId !== 'string' ||
    input.sheetId.length === 0
  ) {
    throw new Error('Invalid formula-cells request.')
  }
  return { sessionId: input.sessionId, sheetId: input.sheetId }
}

function parseFormulaCellsResult(input: unknown): WorkbookFormulaCellsResult {
  if (
    !isRecord(input) ||
    !Array.isArray(input.cells) ||
    input.cells.length > 100_000 ||
    typeof input.indexingComplete !== 'boolean' ||
    typeof input.truncated !== 'boolean'
  ) {
    throw new Error('Invalid formula-cells response.')
  }
  return {
    cells: input.cells.map(parseCellRecord),
    indexingComplete: input.indexingComplete,
    truncated: input.truncated,
  }
}

function parseRecalcRequest(input: unknown): WorkbookRecalcRequest {
  if (
    !isRecord(input) ||
    !isUuid(input.sessionId) ||
    !Array.isArray(input.edits) ||
    input.edits.length > 10_000 ||
    !Array.isArray(input.reads) ||
    input.reads.length > 200
  ) {
    throw new Error('Invalid workbook recalc request.')
  }
  let readCells = 0
  const reads = input.reads.map((read) => {
    if (!isRecord(read) || typeof read.sheetId !== 'string' || read.sheetId.length === 0) {
      throw new Error('Invalid workbook recalc request.')
    }
    const range = read.range
    if (
      !isRecord(range) ||
      !isNonnegativeInteger(range.startRow) ||
      !isNonnegativeInteger(range.endRow) ||
      !isNonnegativeInteger(range.startColumn) ||
      !isNonnegativeInteger(range.endColumn) ||
      range.startRow > range.endRow ||
      range.startColumn > range.endColumn
    ) {
      throw new Error('Invalid workbook recalc request.')
    }
    readCells += (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)
    return {
      sheetId: read.sheetId,
      range: {
        startRow: range.startRow,
        endRow: range.endRow,
        startColumn: range.startColumn,
        endColumn: range.endColumn,
      },
    }
  })
  if (readCells > 20_000) throw new Error('Invalid workbook recalc request.')
  const edits = input.edits.map((edit) => {
    if (
      !isRecord(edit) ||
      typeof edit.sheetId !== 'string' ||
      edit.sheetId.length === 0 ||
      !isNonnegativeInteger(edit.row) ||
      edit.row > 1_048_575 ||
      !isNonnegativeInteger(edit.column) ||
      edit.column > 16_383 ||
      typeof edit.input !== 'string' ||
      edit.input.length > 32_767
    ) {
      throw new Error('Invalid workbook recalc request.')
    }
    return { sheetId: edit.sheetId, row: edit.row, column: edit.column, input: edit.input }
  })
  return { sessionId: input.sessionId, edits, reads }
}

function parseRecalcResult(input: unknown): WorkbookRecalcResult {
  if (!isRecord(input) || !Array.isArray(input.cells) || input.cells.length > 20_000) {
    throw new Error('Invalid workbook recalc response.')
  }
  const cells = input.cells.map((cell) => {
    if (
      !isRecord(cell) ||
      typeof cell.sheetId !== 'string' ||
      cell.sheetId.length === 0 ||
      !isNonnegativeInteger(cell.row) ||
      !isNonnegativeInteger(cell.column) ||
      typeof cell.formatted !== 'string' ||
      (cell.number !== undefined &&
        (typeof cell.number !== 'number' || !Number.isFinite(cell.number))) ||
      typeof cell.isFormula !== 'boolean'
    ) {
      throw new Error('Invalid workbook recalc response.')
    }
    return {
      sheetId: cell.sheetId,
      row: cell.row,
      column: cell.column,
      formatted: cell.formatted,
      ...(cell.number === undefined ? {} : { number: cell.number }),
      isFormula: cell.isFormula,
    }
  })
  return { cells }
}

function parseRangeResult(input: unknown): WorkbookRangeResult {
  if (
    !isRecord(input) ||
    !Array.isArray(input.cells) ||
    input.cells.length > 20_000 ||
    !Array.isArray(input.rows) ||
    input.rows.length > 20_000 ||
    !Array.isArray(input.merges) ||
    input.merges.length > 20_000 ||
    !Array.isArray(input.hyperlinks) ||
    input.hyperlinks.length > 20_000 ||
    !Array.isArray(input.conditionalRules) ||
    input.conditionalRules.length > 20_000 ||
    (input.indexedThroughRow !== null && !isNonnegativeInteger(input.indexedThroughRow)) ||
    typeof input.indexingComplete !== 'boolean'
  ) {
    throw new Error('Invalid workbook range response.')
  }
  const cells = input.cells.map(parseCellRecord)
  const rows = input.rows.map((row) => {
    if (
      !isRecord(row) ||
      !isNonnegativeInteger(row.row) ||
      (row.height !== undefined &&
        (typeof row.height !== 'number' || !Number.isFinite(row.height) || row.height < 0)) ||
      typeof row.hidden !== 'boolean' ||
      (row.outlineLevel !== undefined &&
        (!isNonnegativeInteger(row.outlineLevel) || row.outlineLevel > 7)) ||
      (row.collapsed !== undefined && typeof row.collapsed !== 'boolean')
    ) {
      throw new Error('Invalid workbook row response.')
    }
    return {
      row: row.row,
      hidden: row.hidden,
      ...(row.height === undefined ? {} : { height: row.height }),
      ...(row.outlineLevel === undefined ? {} : { outlineLevel: row.outlineLevel }),
      ...(row.collapsed === undefined ? {} : { collapsed: row.collapsed }),
    }
  })
  const merges = input.merges.map((merge) => {
    if (
      !isRecord(merge) ||
      !isNonnegativeInteger(merge.startRow) ||
      !isNonnegativeInteger(merge.startColumn) ||
      !isNonnegativeInteger(merge.endRow) ||
      !isNonnegativeInteger(merge.endColumn) ||
      merge.startRow > merge.endRow ||
      merge.startColumn > merge.endColumn
    ) {
      throw new Error('Invalid workbook merge response.')
    }
    return {
      startRow: merge.startRow,
      startColumn: merge.startColumn,
      endRow: merge.endRow,
      endColumn: merge.endColumn,
    }
  })
  const hyperlinks = input.hyperlinks.map((link) => {
    if (
      !isRecord(link) ||
      !isNonnegativeInteger(link.row) ||
      !isNonnegativeInteger(link.column) ||
      typeof link.target !== 'string'
    ) {
      throw new Error('Invalid workbook hyperlink response.')
    }
    return { row: link.row, column: link.column, target: link.target }
  })
  const conditionalRules = input.conditionalRules.map(parseConditionalRule)
  if (
    !Array.isArray(input.dataValidations) ||
    input.dataValidations.length > 20_000 ||
    (input.autoFilter !== null && !isRecord(input.autoFilter))
  ) {
    throw new Error('Invalid workbook range response.')
  }
  const dataValidations = input.dataValidations.map((rule) => {
    if (
      !isRecord(rule) ||
      !Array.isArray(rule.ranges) ||
      typeof rule.ruleType !== 'string' ||
      !isOptionalString(rule.operator) ||
      !Array.isArray(rule.formulas) ||
      rule.formulas.some((formula) => typeof formula !== 'string') ||
      typeof rule.allowBlank !== 'boolean' ||
      typeof rule.suppressDropdown !== 'boolean' ||
      typeof rule.showInputMessage !== 'boolean' ||
      typeof rule.showErrorMessage !== 'boolean' ||
      !isOptionalString(rule.errorStyle) ||
      !isOptionalString(rule.errorTitle) ||
      !isOptionalString(rule.error) ||
      !isOptionalString(rule.promptTitle) ||
      !isOptionalString(rule.prompt)
    ) {
      throw new Error('Invalid workbook data validation.')
    }
    return {
      ranges: rule.ranges.map(parseCellArea),
      ruleType: rule.ruleType,
      formulas: rule.formulas as string[],
      allowBlank: rule.allowBlank,
      suppressDropdown: rule.suppressDropdown,
      showInputMessage: rule.showInputMessage,
      showErrorMessage: rule.showErrorMessage,
      ...(rule.operator === undefined ? {} : { operator: rule.operator }),
      ...(rule.errorStyle === undefined ? {} : { errorStyle: rule.errorStyle }),
      ...(rule.errorTitle === undefined ? {} : { errorTitle: rule.errorTitle }),
      ...(rule.error === undefined ? {} : { error: rule.error }),
      ...(rule.promptTitle === undefined ? {} : { promptTitle: rule.promptTitle }),
      ...(rule.prompt === undefined ? {} : { prompt: rule.prompt }),
    }
  })
  let sheetProtection: { protected: boolean; hasPassword: boolean } | null = null
  if (input.sheetProtection !== null && input.sheetProtection !== undefined) {
    const protection = input.sheetProtection
    if (
      !isRecord(protection) ||
      typeof protection.protected !== 'boolean' ||
      typeof protection.hasPassword !== 'boolean'
    ) {
      throw new Error('Invalid workbook sheet protection.')
    }
    sheetProtection = { protected: protection.protected, hasPassword: protection.hasPassword }
  }
  return {
    cells,
    rows,
    merges,
    hyperlinks,
    conditionalRules,
    autoFilter: input.autoFilter === null ? null : parseCellArea(input.autoFilter),
    dataValidations,
    sheetProtection,
    indexedThroughRow: input.indexedThroughRow,
    indexingComplete: input.indexingComplete,
  }
}

function parseRichRuns(input: unknown): WorkbookRichRun[] {
  if (!Array.isArray(input)) throw new Error('Invalid workbook rich text response.')
  return input.map((run) => {
    if (
      !isRecord(run) ||
      typeof run.text !== 'string' ||
      typeof run.bold !== 'boolean' ||
      typeof run.italic !== 'boolean' ||
      typeof run.underline !== 'boolean' ||
      typeof run.strikethrough !== 'boolean' ||
      !isOptionalString(run.color) ||
      (run.size !== undefined &&
        (typeof run.size !== 'number' || !Number.isFinite(run.size) || run.size <= 0)) ||
      !isOptionalString(run.family)
    ) {
      throw new Error('Invalid workbook rich text response.')
    }
    return {
      text: run.text,
      bold: run.bold,
      italic: run.italic,
      underline: run.underline,
      strikethrough: run.strikethrough,
      ...(run.color === undefined ? {} : { color: run.color }),
      ...(run.size === undefined ? {} : { size: run.size }),
      ...(run.family === undefined ? {} : { family: run.family }),
    }
  })
}

function parseCellArea(input: unknown): WorkbookConditionalRule['ranges'][number] {
  if (
    !isRecord(input) ||
    !isNonnegativeInteger(input.startRow) ||
    !isNonnegativeInteger(input.startColumn) ||
    !isNonnegativeInteger(input.endRow) ||
    !isNonnegativeInteger(input.endColumn) ||
    input.startRow > input.endRow ||
    input.startColumn > input.endColumn
  ) {
    throw new Error('Invalid workbook cell area.')
  }
  return {
    startRow: input.startRow,
    startColumn: input.startColumn,
    endRow: input.endRow,
    endColumn: input.endColumn,
  }
}

function parseConditionalRule(input: unknown): WorkbookConditionalRule {
  if (
    !isRecord(input) ||
    !Array.isArray(input.ranges) ||
    typeof input.ruleType !== 'string' ||
    !isOptionalString(input.operator) ||
    !Array.isArray(input.formulas) ||
    input.formulas.some((formula) => typeof formula !== 'string') ||
    !isOptionalString(input.text) ||
    (input.dxfIndex !== undefined && !isNonnegativeInteger(input.dxfIndex)) ||
    typeof input.priority !== 'number' ||
    !Number.isInteger(input.priority) ||
    (input.rank !== undefined && !isNonnegativeInteger(input.rank)) ||
    typeof input.percent !== 'boolean' ||
    typeof input.bottom !== 'boolean' ||
    !Array.isArray(input.cfvos) ||
    !Array.isArray(input.colors) ||
    input.colors.some((color) => typeof color !== 'string')
  ) {
    throw new Error('Invalid workbook conditional rule.')
  }
  const cfvos = input.cfvos.map((cfvo) => {
    if (
      !isRecord(cfvo) ||
      typeof cfvo.kind !== 'string' ||
      !isOptionalString(cfvo.value) ||
      (cfvo.gte !== undefined && typeof cfvo.gte !== 'boolean')
    ) {
      throw new Error('Invalid workbook conditional rule value.')
    }
    return {
      kind: cfvo.kind,
      ...(cfvo.value === undefined ? {} : { value: cfvo.value }),
      ...(cfvo.gte === undefined ? {} : { gte: cfvo.gte }),
    }
  })
  if (
    !isOptionalString(input.iconSetName) ||
    typeof input.iconReverse !== 'boolean' ||
    typeof input.showValue !== 'boolean'
  ) {
    throw new Error('Invalid workbook conditional rule.')
  }
  return {
    ranges: input.ranges.map(parseCellArea),
    ruleType: input.ruleType,
    formulas: input.formulas as string[],
    priority: input.priority,
    percent: input.percent,
    bottom: input.bottom,
    cfvos,
    colors: input.colors as string[],
    iconReverse: input.iconReverse,
    showValue: input.showValue,
    ...(input.iconSetName === undefined ? {} : { iconSetName: input.iconSetName }),
    ...(input.operator === undefined ? {} : { operator: input.operator }),
    ...(input.text === undefined ? {} : { text: input.text }),
    ...(input.dxfIndex === undefined ? {} : { dxfIndex: input.dxfIndex }),
    ...(input.rank === undefined ? {} : { rank: input.rank }),
  }
}

function parseSaveRequest(input: WorkbookSaveRequest): WorkbookSaveRequest {
  if (
    !isRecord(input) ||
    !isUuid(input.sessionId) ||
    (input.mode !== 'save' && input.mode !== 'save-as') ||
    !Array.isArray(input.edits) ||
    input.edits.length > 10_000 ||
    !Array.isArray(input.structuralOps) ||
    input.structuralOps.length > 1_000 ||
    !Array.isArray(input.chartEdits) ||
    input.chartEdits.length > 100 ||
    !Array.isArray(input.visualEdits) ||
    input.visualEdits.length > 100 ||
    !Array.isArray(input.visualAdditions) ||
    input.visualAdditions.length > 100 ||
    !Array.isArray(input.tableAdditions) ||
    input.tableAdditions.length > 50 ||
    !Array.isArray(input.pivotAdditions) ||
    input.pivotAdditions.length > 20 ||
    !Array.isArray(input.sheetOps) ||
    input.sheetOps.length > 100 ||
    !Array.isArray(input.sheetOrder) ||
    input.sheetOrder.length > 1_000 ||
    input.sheetOrder.some((sheetId) => typeof sheetId !== 'string' || sheetId.length === 0) ||
    !Array.isArray(input.filterStates) ||
    input.filterStates.length > 1_000 ||
    !Array.isArray(input.hyperlinkEdits) ||
    input.hyperlinkEdits.length > 1_000 ||
    !Array.isArray(input.cfStates) ||
    input.cfStates.length > 1_000 ||
    !Array.isArray(input.dvStates) ||
    input.dvStates.length > 1_000 ||
    !Array.isArray(input.pageSetupStates) ||
    input.pageSetupStates.length > 1_000 ||
    !Array.isArray(input.noteStates) ||
    input.noteStates.length > 1_000 ||
    !Array.isArray(input.pivotCacheRefreshPaths) ||
    input.pivotCacheRefreshPaths.length > 100 ||
    input.pivotCacheRefreshPaths.some((path) => typeof path !== 'string') ||
    !Array.isArray(input.pivotRefreshUpdates) ||
    input.pivotRefreshUpdates.length > 100 ||
    input.pivotRefreshUpdates.some(
      (update) =>
        !isRecord(update) ||
        typeof update.cachePath !== 'string' ||
        typeof update.sheetId !== 'string' ||
        update.sheetId.length === 0 ||
        typeof update.newOutputRef !== 'string' ||
        update.newOutputRef.length > 64 ||
        // relayout's full shape is authoritatively validated by main's zod schema.
        (update.relayout !== undefined && !isRecord(update.relayout)),
    ) ||
    !Array.isArray(input.sheetProtections) ||
    input.sheetProtections.length > 1_000 ||
    input.sheetProtections.some(
      (state) =>
        !isRecord(state) ||
        typeof state.sheetId !== 'string' ||
        state.sheetId.length === 0 ||
        typeof state.protected !== 'boolean',
    ) ||
    !isDefinedNamesState(input.definedNamesState) ||
    (input.edits.length === 0 &&
      input.structuralOps.length === 0 &&
      input.chartEdits.length === 0 &&
      input.visualEdits.length === 0 &&
      input.visualAdditions.length === 0 &&
      input.tableAdditions.length === 0 &&
      input.pivotAdditions.length === 0 &&
      input.sheetOps.length === 0 &&
      input.filterStates.length === 0 &&
      input.hyperlinkEdits.length === 0 &&
      input.cfStates.length === 0 &&
      input.dvStates.length === 0 &&
      input.pageSetupStates.length === 0 &&
      input.noteStates.length === 0 &&
      input.pivotCacheRefreshPaths.length === 0 &&
      input.pivotRefreshUpdates.length === 0 &&
      input.sheetProtections.length === 0 &&
      input.definedNamesState === null) ||
    (input.sheetOps.length > 0 && input.sheetOrder.length === 0)
  ) {
    throw new Error('Invalid workbook save request.')
  }
  for (const state of input.dvStates) {
    if (
      !isRecord(state) ||
      typeof state.sheetId !== 'string' ||
      state.sheetId.length === 0 ||
      !Array.isArray(state.rules) ||
      state.rules.length > 500
    ) {
      throw new Error('Invalid workbook data-validation state.')
    }
    for (const rule of state.rules) {
      if (
        !isRecord(rule) ||
        !isRecord(rule.rule) ||
        !Array.isArray(rule.ranges) ||
        rule.ranges.length === 0 ||
        rule.ranges.length > 100
      ) {
        throw new Error('Invalid workbook data-validation rule.')
      }
      for (const range of rule.ranges) parseCellArea(range)
    }
  }
  for (const state of input.pageSetupStates) {
    if (!isPageSetupState(state)) throw new Error('Invalid workbook page-setup state.')
  }
  for (const table of input.tableAdditions) {
    if (
      !isRecord(table) ||
      typeof table.sheetId !== 'string' ||
      table.sheetId.length === 0 ||
      typeof table.name !== 'string' ||
      table.name.length === 0 ||
      table.name.length > 255 ||
      !Array.isArray(table.columnNames) ||
      table.columnNames.length === 0 ||
      table.columnNames.length > 1_000 ||
      table.columnNames.some((name) => typeof name !== 'string' || name.length > 255) ||
      (table.style !== undefined &&
        (typeof table.style !== 'string' ||
          !/^TableStyle(?:Light|Medium|Dark)[1-9][0-9]?$/.test(table.style))) ||
      typeof table.bandedRows !== 'boolean'
    ) {
      throw new Error('Invalid workbook table addition.')
    }
    parseCellArea(table.area)
  }
  for (const pivot of input.pivotAdditions) {
    if (
      !isRecord(pivot) ||
      typeof pivot.sheetId !== 'string' ||
      pivot.sheetId.length === 0 ||
      typeof pivot.sourceSheetId !== 'string' ||
      pivot.sourceSheetId.length === 0 ||
      typeof pivot.name !== 'string' ||
      pivot.name.length === 0 ||
      pivot.name.length > 255 ||
      !Array.isArray(pivot.fieldNames) ||
      pivot.fieldNames.length === 0 ||
      pivot.fieldNames.length > 200 ||
      pivot.fieldNames.some(
        (name) => typeof name !== 'string' || name.length === 0 || name.length > 255,
      ) ||
      !Array.isArray(pivot.rowFieldIndices) ||
      pivot.rowFieldIndices.length === 0 ||
      pivot.rowFieldIndices.length > 8 ||
      pivot.rowFieldIndices.some((i) => !isNonnegativeInteger(i)) ||
      (pivot.rowLevelItems !== undefined &&
        (!Array.isArray(pivot.rowLevelItems) ||
          pivot.rowLevelItems.length > 8 ||
          pivot.rowLevelItems.some(
            (level) =>
              !Array.isArray(level) ||
              level.length > 10_000 ||
              level.some((item) => typeof item !== 'string' || item.length > 255),
          ))) ||
      (pivot.rowLines !== undefined &&
        (!Array.isArray(pivot.rowLines) ||
          pivot.rowLines.length > 20_000 ||
          pivot.rowLines.some(
            (line) =>
              !isRecord(line) ||
              !['data', 'default'].includes(String(line.t)) ||
              !Array.isArray(line.members) ||
              line.members.length === 0 ||
              line.members.length > 8 ||
              line.members.some((member) => !isNonnegativeInteger(member)),
          ))) ||
      (pivot.columnFieldIndex !== undefined && !isNonnegativeInteger(pivot.columnFieldIndex)) ||
      (pivot.pageFieldIndices !== undefined &&
        (!Array.isArray(pivot.pageFieldIndices) ||
          pivot.pageFieldIndices.length > 4 ||
          pivot.pageFieldIndices.some((i) => !isNonnegativeInteger(i)))) ||
      !Array.isArray(pivot.rowItems) ||
      pivot.rowItems.length === 0 ||
      pivot.rowItems.length > 10_000 ||
      pivot.rowItems.some((item) => typeof item !== 'string' || item.length > 255) ||
      (pivot.columnItems !== undefined &&
        (!Array.isArray(pivot.columnItems) ||
          pivot.columnItems.length > 1_000 ||
          pivot.columnItems.some((item) => typeof item !== 'string' || item.length > 255))) ||
      !Array.isArray(pivot.values) ||
      pivot.values.length === 0 ||
      pivot.values.length > 8 ||
      pivot.values.some(
        (value) =>
          !isRecord(value) ||
          !isNonnegativeInteger(value.fieldIndex) ||
          !['sum', 'count', 'average', 'max', 'min'].includes(String(value.agg)) ||
          (value.showDataAs !== undefined &&
            !['percentOfTotal', 'percentOfRow', 'percentOfCol'].includes(String(value.showDataAs))),
      )
    ) {
      throw new Error('Invalid workbook pivot addition.')
    }
    parseCellArea(pivot.sourceArea)
    parseCellArea(pivot.location)
  }
  for (const state of input.cfStates) {
    if (
      !isRecord(state) ||
      typeof state.sheetId !== 'string' ||
      state.sheetId.length === 0 ||
      !Array.isArray(state.rules) ||
      state.rules.length > 500
    ) {
      throw new Error('Invalid workbook conditional-formatting state.')
    }
    for (const rule of state.rules) {
      if (
        !isRecord(rule) ||
        typeof rule.stopIfTrue !== 'boolean' ||
        !isRecord(rule.rule) ||
        !Array.isArray(rule.ranges) ||
        rule.ranges.length === 0 ||
        rule.ranges.length > 100
      ) {
        throw new Error('Invalid workbook conditional-formatting rule.')
      }
      for (const range of rule.ranges) parseCellArea(range)
    }
  }
  for (const link of input.hyperlinkEdits) {
    if (
      !isRecord(link) ||
      typeof link.sheetId !== 'string' ||
      link.sheetId.length === 0 ||
      !isNonnegativeInteger(link.row) ||
      !isNonnegativeInteger(link.column) ||
      (link.target !== null &&
        (typeof link.target !== 'string' || link.target.length === 0 || link.target.length > 2083))
    ) {
      throw new Error('Invalid workbook hyperlink edit.')
    }
  }
  for (const state of input.filterStates) {
    if (
      !isRecord(state) ||
      typeof state.sheetId !== 'string' ||
      state.sheetId.length === 0 ||
      !Array.isArray(state.hiddenRows) ||
      state.hiddenRows.length > 100_000 ||
      state.hiddenRows.some((row) => !isNonnegativeInteger(row))
    ) {
      throw new Error('Invalid workbook filter state.')
    }
    parseCellArea(state.visibilityRange)
    if (state.filter === null) continue
    if (
      !isRecord(state.filter) ||
      !Array.isArray(state.filter.columns) ||
      state.filter.columns.length > 1_000
    ) {
      throw new Error('Invalid workbook filter state.')
    }
    parseCellArea(state.filter.range)
    for (const column of state.filter.columns) {
      if (!isFilterColumn(column)) throw new Error('Invalid workbook filter column.')
    }
  }
  for (const op of input.sheetOps) {
    if (!isRecord(op)) throw new Error('Invalid workbook sheet operation.')
    if (op.kind === 'reorder-sheets') continue
    if (typeof op.sheetId !== 'string' || op.sheetId.length === 0) {
      throw new Error('Invalid workbook sheet operation.')
    }
    if (op.kind === 'rename-sheet') {
      if (!isSheetName(op.newName)) throw new Error('Invalid workbook sheet operation.')
    } else if (op.kind === 'add-sheet') {
      if (!isSheetName(op.name)) throw new Error('Invalid workbook sheet operation.')
    } else if (op.kind === 'duplicate-sheet') {
      if (
        !isSheetName(op.name) ||
        typeof op.sourceSheetId !== 'string' ||
        op.sourceSheetId.length === 0
      ) {
        throw new Error('Invalid workbook sheet operation.')
      }
    } else if (op.kind === 'set-sheet-hidden') {
      if (typeof op.hidden !== 'boolean') throw new Error('Invalid workbook sheet operation.')
    } else if (op.kind !== 'remove-sheet') {
      throw new Error('Invalid workbook sheet operation.')
    }
  }
  for (const chartEdit of input.chartEdits) {
    if (
      !isRecord(chartEdit) ||
      typeof chartEdit.chartPath !== 'string' ||
      !/^xl\/charts\/[A-Za-z0-9._-]+\.xml$/.test(chartEdit.chartPath) ||
      (chartEdit.title !== undefined &&
        (typeof chartEdit.title !== 'string' || chartEdit.title.length > 255)) ||
      (chartEdit.chartType !== undefined &&
        !['column', 'bar', 'line', 'area'].includes(String(chartEdit.chartType))) ||
      (chartEdit.seriesColors !== undefined && !isSeriesColors(chartEdit.seriesColors))
    ) {
      throw new Error('Invalid workbook chart edit.')
    }
  }
  for (const visualEdit of input.visualEdits) {
    if (
      !isRecord(visualEdit) ||
      typeof visualEdit.drawingPath !== 'string' ||
      !/^xl\/drawings\/[A-Za-z0-9._/-]+\.xml$/.test(visualEdit.drawingPath) ||
      !isNonnegativeInteger(visualEdit.drawingIndex) ||
      visualEdit.drawingIndex > 10_000 ||
      (visualEdit.remove !== undefined && visualEdit.remove !== true) ||
      (visualEdit.remove === undefined && !isRecord(visualEdit.anchor))
    ) {
      throw new Error('Invalid workbook visual edit.')
    }
    if (visualEdit.anchor !== undefined) {
      if (!isRecord(visualEdit.anchor)) throw new Error('Invalid workbook visual edit.')
      parseDrawingAnchor(visualEdit.anchor)
    }
  }
  for (const op of input.structuralOps) {
    if (!isRecord(op) || typeof op.sheetId !== 'string' || op.sheetId.length === 0) {
      throw new Error('Invalid workbook structural operation.')
    }
    if ('range' in op) {
      if (op.kind !== 'merge-cells' && op.kind !== 'unmerge-cells') {
        throw new Error('Invalid workbook structural operation.')
      }
      parseCellArea(op.range)
      continue
    }
    if ('size' in op || 'hidden' in op || 'level' in op) {
      const { start, end } = op
      if (
        !isNonnegativeInteger(start) ||
        !isNonnegativeInteger(end) ||
        end < start ||
        end - start >= 100_000
      ) {
        throw new Error('Invalid workbook structural operation.')
      }
      if ('size' in op) {
        if (
          (op.kind !== 'set-row-size' && op.kind !== 'set-col-size') ||
          (op.size !== null &&
            (typeof op.size !== 'number' ||
              !Number.isFinite(op.size) ||
              op.size <= 0 ||
              op.size > 500))
        ) {
          throw new Error('Invalid workbook structural operation.')
        }
      } else if ('level' in op) {
        if (
          (op.kind !== 'set-rows-outline' && op.kind !== 'set-cols-outline') ||
          !isNonnegativeInteger(op.level) ||
          op.level > 7 ||
          (op.collapsed !== undefined && typeof op.collapsed !== 'boolean')
        ) {
          throw new Error('Invalid workbook structural operation.')
        }
      } else if (
        (op.kind !== 'set-rows-hidden' && op.kind !== 'set-cols-hidden') ||
        typeof op.hidden !== 'boolean'
      ) {
        throw new Error('Invalid workbook structural operation.')
      }
      continue
    }
    const { index, count } = op
    if (
      !['insert-rows', 'remove-rows', 'insert-cols', 'remove-cols'].includes(String(op.kind)) ||
      !isNonnegativeInteger(index) ||
      !isNonnegativeInteger(count) ||
      count === 0 ||
      count > 10_000
    ) {
      throw new Error('Invalid workbook structural operation.')
    }
  }
  for (const edit of input.edits) {
    if (
      !isRecord(edit) ||
      typeof edit.sheetId !== 'string' ||
      edit.sheetId.length === 0 ||
      !isNonnegativeInteger(edit.row) ||
      !isNonnegativeInteger(edit.column) ||
      typeof edit.writeValue !== 'boolean' ||
      !isCellScalar(edit.value) ||
      (edit.formula !== undefined &&
        (typeof edit.formula !== 'string' ||
          edit.formula.length === 0 ||
          edit.formula.length > 8_192)) ||
      (edit.style !== undefined && !isStyleEdit(edit.style)) ||
      (edit.styleReset !== undefined && typeof edit.styleReset !== 'boolean') ||
      (!edit.writeValue && edit.style === undefined && edit.styleReset !== true)
    ) {
      throw new Error('Invalid workbook cell edit.')
    }
    if (edit.rich !== undefined) parseRichRuns(edit.rich)
  }
  return input
}

const FILTER_OPERATORS = [
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
]

function isFilterColumn(input: unknown): boolean {
  if (!isRecord(input) || !isNonnegativeInteger(input.colId)) return false
  if (
    input.values !== undefined &&
    (!Array.isArray(input.values) ||
      input.values.length > 10_000 ||
      input.values.some((value) => typeof value !== 'string' || value.length > 255))
  ) {
    return false
  }
  if (input.blank !== undefined && typeof input.blank !== 'boolean') return false
  if (input.customs !== undefined) {
    if (!isRecord(input.customs)) return false
    if (input.customs.and !== undefined && typeof input.customs.and !== 'boolean') return false
    const filters = input.customs.filters
    if (!Array.isArray(filters) || filters.length === 0 || filters.length > 2) return false
    for (const custom of filters) {
      if (
        !isRecord(custom) ||
        (typeof custom.val !== 'string' && typeof custom.val !== 'number') ||
        (custom.operator !== undefined && !FILTER_OPERATORS.includes(String(custom.operator)))
      ) {
        return false
      }
    }
  }
  return input.values !== undefined || input.blank !== undefined || input.customs !== undefined
}

function isPdfPageSize(input: unknown): boolean {
  if (typeof input === 'string') {
    return ['A3', 'A4', 'A5', 'Legal', 'Letter', 'Tabloid'].includes(input)
  }
  return (
    isRecord(input) &&
    typeof input.width === 'number' &&
    input.width > 0 &&
    input.width <= 100 &&
    typeof input.height === 'number' &&
    input.height > 0 &&
    input.height <= 100
  )
}

function isPageSetupState(input: unknown): boolean {
  if (!isRecord(input) || typeof input.sheetId !== 'string' || input.sheetId.length === 0) {
    return false
  }
  const isBoundedInt = (value: unknown, min: number, max: number) =>
    typeof value === 'number' && Number.isInteger(value) && value >= min && value <= max
  if (
    input.orientation !== undefined &&
    input.orientation !== 'portrait' &&
    input.orientation !== 'landscape'
  )
    return false
  if (input.paperSize !== undefined && !isBoundedInt(input.paperSize, 1, 118)) return false
  if (input.scale !== undefined && !isBoundedInt(input.scale, 10, 400)) return false
  if (input.fitToWidth !== undefined && !isBoundedInt(input.fitToWidth, 0, 1_000)) return false
  if (input.fitToHeight !== undefined && !isBoundedInt(input.fitToHeight, 0, 1_000)) return false
  if (input.margins !== undefined && !['normal', 'wide', 'narrow'].includes(String(input.margins)))
    return false
  for (const key of [
    'printGridlines',
    'printHeadings',
    'showGridlines',
    'showFormulas',
    'fitToPage',
  ]) {
    if (input[key] !== undefined && typeof input[key] !== 'boolean') return false
  }
  if (
    input.printArea !== undefined &&
    input.printArea !== null &&
    (typeof input.printArea !== 'string' ||
      input.printArea.length === 0 ||
      input.printArea.length > 255)
  )
    return false
  if (
    input.printTitles !== undefined &&
    input.printTitles !== null &&
    (typeof input.printTitles !== 'string' || !/^\d{1,7}:\d{1,7}$/.test(input.printTitles))
  )
    return false
  return Object.keys(input).length > 1
}

function isSheetName(input: unknown): input is string {
  return (
    typeof input === 'string' &&
    input.length >= 1 &&
    input.length <= 31 &&
    !/[\\/?*[\]:]/.test(input) &&
    !input.startsWith("'") &&
    !input.endsWith("'")
  )
}

function isSeriesColors(input: unknown): boolean {
  if (!isRecord(input)) return false
  return Object.entries(input).every(
    ([key, value]) =>
      /^[0-9]{1,3}$/.test(key) && typeof value === 'string' && /^#[0-9A-Fa-f]{6}$/.test(value),
  )
}

const BORDER_STYLE_NAMES = [
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
]

function isStyleEdit(input: unknown): boolean {
  if (!isRecord(input)) return false
  const isHexColor = (value: unknown): boolean =>
    typeof value === 'string' && /^#[0-9A-Fa-f]{6}$/.test(value)
  const isOptionalBoolean = (value: unknown): boolean =>
    value === undefined || typeof value === 'boolean'
  const isBorderEdge = (value: unknown): boolean =>
    value === undefined ||
    value === null ||
    (isRecord(value) &&
      BORDER_STYLE_NAMES.includes(String(value.style)) &&
      (value.color === undefined || isHexColor(value.color)) &&
      Object.keys(value).every((key) => ['style', 'color'].includes(key)))
  if (
    ![input.borderTop, input.borderBottom, input.borderLeft, input.borderRight].every(isBorderEdge)
  ) {
    return false
  }
  return (
    isOptionalBoolean(input.bold) &&
    isOptionalBoolean(input.italic) &&
    isOptionalBoolean(input.underline) &&
    isOptionalBoolean(input.strikethrough) &&
    isOptionalBoolean(input.wrapText) &&
    (input.fontFamily === undefined ||
      (typeof input.fontFamily === 'string' &&
        input.fontFamily.length > 0 &&
        input.fontFamily.length <= 128)) &&
    (input.fontSize === undefined ||
      (typeof input.fontSize === 'number' &&
        Number.isFinite(input.fontSize) &&
        input.fontSize > 0 &&
        input.fontSize <= 409)) &&
    (input.fontColor === undefined || input.fontColor === null || isHexColor(input.fontColor)) &&
    (input.fillColor === undefined || input.fillColor === null || isHexColor(input.fillColor)) &&
    (input.horizontalAlignment === undefined ||
      ['left', 'center', 'right', 'justify', 'distributed'].includes(
        String(input.horizontalAlignment),
      )) &&
    (input.verticalAlignment === undefined ||
      ['top', 'center', 'bottom'].includes(String(input.verticalAlignment))) &&
    (input.numberFormat === undefined ||
      (typeof input.numberFormat === 'string' &&
        input.numberFormat.length > 0 &&
        input.numberFormat.length <= 255)) &&
    (input.underlineStyle === undefined ||
      ['single', 'double'].includes(String(input.underlineStyle))) &&
    (input.textRotation === undefined ||
      (typeof input.textRotation === 'number' &&
        Number.isInteger(input.textRotation) &&
        ((input.textRotation >= 0 && input.textRotation <= 180) || input.textRotation === 255))) &&
    (input.indent === undefined ||
      (typeof input.indent === 'number' &&
        Number.isInteger(input.indent) &&
        input.indent >= 0 &&
        input.indent <= 250)) &&
    isOptionalBoolean(input.protectionLocked) &&
    isOptionalBoolean(input.protectionHidden) &&
    Object.keys(input).every((key) =>
      [
        'bold',
        'italic',
        'underline',
        'underlineStyle',
        'strikethrough',
        'fontFamily',
        'fontSize',
        'fontColor',
        'fillColor',
        'horizontalAlignment',
        'verticalAlignment',
        'wrapText',
        'numberFormat',
        'textRotation',
        'indent',
        'protectionLocked',
        'protectionHidden',
        'borderTop',
        'borderBottom',
        'borderLeft',
        'borderRight',
      ].includes(key),
    )
  )
}

function isDefinedNamesState(input: unknown): boolean {
  if (input === null) return true
  if (!isRecord(input)) return false
  const { names, preserveNames } = input
  if (
    !Array.isArray(names) ||
    names.length > 2_000 ||
    !Array.isArray(preserveNames) ||
    preserveNames.length > 2_000 ||
    preserveNames.some((name) => typeof name !== 'string' || name.length === 0)
  ) {
    return false
  }
  return names.every(
    (entry) =>
      isRecord(entry) &&
      typeof entry.name === 'string' &&
      entry.name.length >= 1 &&
      entry.name.length <= 255 &&
      typeof entry.formula === 'string' &&
      entry.formula.length >= 1 &&
      entry.formula.length <= 8_192 &&
      (entry.sheetIndex === undefined || isNonnegativeInteger(entry.sheetIndex)),
  )
}

function parseSaveResult(input: unknown): WorkbookSaveResult {
  if (!isRecord(input)) throw new Error('Invalid workbook save response.')
  if (input.canceled === true) return { canceled: true }
  if (
    input.canceled !== false ||
    !Array.isArray(input.touchedEntries) ||
    input.touchedEntries.length > 10_000 ||
    input.touchedEntries.some((entry) => typeof entry !== 'string')
  ) {
    throw new Error('Invalid workbook save response.')
  }
  return {
    canceled: false,
    file: parseWorkbookFile(input.file),
    touchedEntries: input.touchedEntries as string[],
  }
}

function parseMediaRequest(input: WorkbookMediaRequest): WorkbookMediaRequest {
  if (!isRecord(input) || !isUuid(input.sessionId) || typeof input.visualId !== 'string') {
    throw new Error('Invalid workbook media request.')
  }
  return input
}

function parseMediaResult(input: unknown): WorkbookMediaResult {
  if (
    !isRecord(input) ||
    typeof input.mediaType !== 'string' ||
    !input.mediaType.startsWith('image/') ||
    typeof input.base64 !== 'string' ||
    input.base64.length === 0
  ) {
    throw new Error('Invalid workbook media response.')
  }
  return { mediaType: input.mediaType, base64: input.base64 }
}

function parseSparklineGroups(input: unknown): {
  type: 'line' | 'column' | 'stacked'
  color?: string
  negativeColor?: string
  cells: { cell: string; sourceRef: string }[]
}[] {
  if (!Array.isArray(input) || input.length > 100) {
    throw new Error('Invalid worksheet sparkline metadata.')
  }
  return input.map((entry) => {
    if (
      !isRecord(entry) ||
      (entry.type !== 'line' && entry.type !== 'column' && entry.type !== 'stacked') ||
      !Array.isArray(entry.cells) ||
      entry.cells.length > 500 ||
      (entry.color !== undefined && typeof entry.color !== 'string') ||
      (entry.negativeColor !== undefined && typeof entry.negativeColor !== 'string')
    ) {
      throw new Error('Invalid worksheet sparkline metadata.')
    }
    return {
      type: entry.type,
      ...(entry.color === undefined ? {} : { color: entry.color }),
      ...(entry.negativeColor === undefined ? {} : { negativeColor: entry.negativeColor }),
      cells: entry.cells.map((cell) => {
        if (
          !isRecord(cell) ||
          typeof cell.cell !== 'string' ||
          cell.cell.length < 2 ||
          typeof cell.sourceRef !== 'string' ||
          cell.sourceRef.length === 0
        ) {
          throw new Error('Invalid worksheet sparkline metadata.')
        }
        return { cell: cell.cell, sourceRef: cell.sourceRef }
      }),
    }
  })
}

function parsePivotTableInfos(
  input: unknown,
): { path: string; cachePath: string | null; outputRef: string }[] {
  if (!Array.isArray(input) || input.length > 100) {
    throw new Error('Invalid worksheet pivot metadata.')
  }
  return input.map((entry) => {
    if (
      !isRecord(entry) ||
      typeof entry.path !== 'string' ||
      entry.path.length === 0 ||
      (entry.cachePath !== null && typeof entry.cachePath !== 'string') ||
      typeof entry.outputRef !== 'string' ||
      entry.outputRef.length === 0
    ) {
      throw new Error('Invalid worksheet pivot metadata.')
    }
    return { path: entry.path, cachePath: entry.cachePath, outputRef: entry.outputRef }
  })
}

function parsePivotRequest(input: WorkbookPivotRequest): WorkbookPivotRequest {
  const isPartPath = (value: unknown): value is string =>
    typeof value === 'string' && value.length <= 256 && /^xl\/[A-Za-z0-9._/-]+\.xml$/.test(value)
  if (
    !isRecord(input) ||
    !isUuid(input.sessionId) ||
    !isPartPath(input.path) ||
    !isPartPath(input.cachePath)
  ) {
    throw new Error('Invalid pivot definition request.')
  }
  return input
}

/// Structural check only — the main process zod-validates the full shape;
/// the recompute engine fails closed on any bad member data.
function parsePivotDefinitionResult(input: unknown): WorkbookPivotDefinition {
  if (
    !isRecord(input) ||
    typeof input.outputRef !== 'string' ||
    typeof input.firstDataRow !== 'number' ||
    typeof input.firstDataCol !== 'number' ||
    typeof input.sourceSheet !== 'string' ||
    typeof input.sourceRef !== 'string' ||
    !Array.isArray(input.fields) ||
    input.fields.length > 1_000 ||
    !Array.isArray(input.fieldItems) ||
    input.fieldItems.length > 1_000 ||
    !Array.isArray(input.rowFields) ||
    input.rowFields.length > 64 ||
    !Array.isArray(input.colFields) ||
    input.colFields.length > 64 ||
    !Array.isArray(input.rowLines) ||
    input.rowLines.length > 100_000 ||
    !Array.isArray(input.colLines) ||
    input.colLines.length > 10_000 ||
    !Array.isArray(input.dataFields) ||
    input.dataFields.length > 256 ||
    !Array.isArray(input.pageFields) ||
    input.pageFields.length > 256 ||
    !Array.isArray(input.unsupported) ||
    input.unsupported.length > 100 ||
    input.unsupported.some((reason) => typeof reason !== 'string')
  ) {
    throw new Error('Invalid pivot definition response.')
  }
  return input as unknown as WorkbookPivotDefinition
}

function parseCellStyle(input: unknown): WorkbookCellStyle {
  if (
    !isRecord(input) ||
    typeof input.bold !== 'boolean' ||
    typeof input.italic !== 'boolean' ||
    typeof input.underline !== 'boolean' ||
    typeof input.strikethrough !== 'boolean' ||
    typeof input.wrapText !== 'boolean' ||
    !isOptionalString(input.fontFamily) ||
    (input.fontSize !== undefined &&
      (typeof input.fontSize !== 'number' ||
        !Number.isFinite(input.fontSize) ||
        input.fontSize <= 0)) ||
    !isOptionalString(input.fontColor) ||
    !isOptionalString(input.fillColor) ||
    !isOptionalString(input.horizontalAlignment) ||
    !isOptionalString(input.verticalAlignment) ||
    (input.indent !== undefined && !isNonnegativeInteger(input.indent)) ||
    !isOptionalString(input.numberFormat)
  ) {
    throw new Error('Invalid workbook style response.')
  }
  const borderTop = parseBorderEdge(input.borderTop)
  const borderBottom = parseBorderEdge(input.borderBottom)
  const borderLeft = parseBorderEdge(input.borderLeft)
  const borderRight = parseBorderEdge(input.borderRight)
  const borderDiagonal = parseBorderEdge(input.borderDiagonal)
  if (typeof input.diagonalUp !== 'boolean' || typeof input.diagonalDown !== 'boolean') {
    throw new Error('Invalid workbook style response.')
  }
  return {
    bold: input.bold,
    italic: input.italic,
    underline: input.underline,
    strikethrough: input.strikethrough,
    wrapText: input.wrapText,
    diagonalUp: input.diagonalUp,
    diagonalDown: input.diagonalDown,
    ...(input.fontFamily === undefined ? {} : { fontFamily: input.fontFamily }),
    ...(input.fontSize === undefined ? {} : { fontSize: input.fontSize }),
    ...(input.fontColor === undefined ? {} : { fontColor: input.fontColor }),
    ...(input.fillColor === undefined ? {} : { fillColor: input.fillColor }),
    ...(input.horizontalAlignment === undefined
      ? {}
      : { horizontalAlignment: input.horizontalAlignment }),
    ...(input.verticalAlignment === undefined
      ? {}
      : { verticalAlignment: input.verticalAlignment }),
    ...(input.indent === undefined ? {} : { indent: input.indent }),
    ...(input.numberFormat === undefined ? {} : { numberFormat: input.numberFormat }),
    ...(borderTop === undefined ? {} : { borderTop }),
    ...(borderBottom === undefined ? {} : { borderBottom }),
    ...(borderLeft === undefined ? {} : { borderLeft }),
    ...(borderRight === undefined ? {} : { borderRight }),
    ...(borderDiagonal === undefined ? {} : { borderDiagonal }),
  }
}

function parseFreeze(input: unknown): { frozenColumns: number; frozenRows: number } | null {
  if (input === null || input === undefined) return null
  if (
    !isRecord(input) ||
    !isNonnegativeInteger(input.frozenColumns) ||
    !isNonnegativeInteger(input.frozenRows)
  ) {
    throw new Error('Invalid worksheet freeze pane.')
  }
  return { frozenColumns: input.frozenColumns, frozenRows: input.frozenRows }
}

function parseBorderEdge(input: unknown): { style: string; color?: string } | undefined {
  if (input === undefined) return undefined
  if (
    !isRecord(input) ||
    typeof input.style !== 'string' ||
    input.style.length === 0 ||
    !isOptionalString(input.color)
  ) {
    throw new Error('Invalid workbook border response.')
  }
  return {
    style: input.style,
    ...(input.color === undefined ? {} : { color: input.color }),
  }
}

function parseVisualObject(input: unknown): WorkbookVisualObject {
  if (
    !isRecord(input) ||
    typeof input.id !== 'string' ||
    typeof input.sheetId !== 'string' ||
    !['chart', 'image', 'shape'].includes(String(input.kind)) ||
    !isRecord(input.anchor)
  ) {
    throw new Error('Invalid workbook visual response.')
  }
  const anchor = parseDrawingAnchor(input.anchor)
  const chart = input.chart === undefined ? undefined : parseChart(input.chart)
  if (input.kind === 'chart' && chart === undefined) {
    throw new Error('Workbook chart has no chart metadata.')
  }
  if (
    !isOptionalString(input.shapeType) ||
    !isOptionalString(input.fillColor) ||
    !isOptionalString(input.text) ||
    (input.rotation !== undefined &&
      (typeof input.rotation !== 'number' || !Number.isFinite(input.rotation))) ||
    (input.drawingPath !== undefined &&
      (typeof input.drawingPath !== 'string' ||
        !/^xl\/drawings\/[A-Za-z0-9._/-]+\.xml$/.test(input.drawingPath))) ||
    (input.drawingIndex !== undefined &&
      (!isNonnegativeInteger(input.drawingIndex) || input.drawingIndex > 10_000))
  ) {
    throw new Error('Invalid workbook visual response.')
  }
  return {
    id: input.id,
    sheetId: input.sheetId,
    kind: input.kind as 'chart' | 'image' | 'shape',
    anchor,
    ...(chart === undefined ? {} : { chart }),
    ...(typeof input.chartPath === 'string' ? { chartPath: input.chartPath } : {}),
    ...(typeof input.mediaPath === 'string' ? { mediaPath: input.mediaPath } : {}),
    ...(typeof input.mediaType === 'string' ? { mediaType: input.mediaType } : {}),
    ...(typeof input.name === 'string' ? { name: input.name } : {}),
    ...(input.shapeType === undefined ? {} : { shapeType: input.shapeType }),
    ...(input.fillColor === undefined ? {} : { fillColor: input.fillColor }),
    ...(input.text === undefined ? {} : { text: input.text }),
    ...(input.rotation === undefined ? {} : { rotation: input.rotation }),
    ...(input.drawingPath === undefined ? {} : { drawingPath: input.drawingPath }),
    ...(input.drawingIndex === undefined ? {} : { drawingIndex: input.drawingIndex }),
  }
}

function parseDrawingAnchor(input: Record<string, unknown>): WorkbookVisualObject['anchor'] {
  const values = [input.fromRow, input.fromColumn, input.toRow, input.toColumn]
  const offsets = [
    input.fromRowOffset,
    input.fromColumnOffset,
    input.toRowOffset,
    input.toColumnOffset,
  ]
  if (
    values.some((value) => !isNonnegativeInteger(value)) ||
    offsets.some((value) => typeof value !== 'number' || !Number.isInteger(value))
  ) {
    throw new Error('Invalid workbook drawing anchor.')
  }
  return {
    fromRow: input.fromRow as number,
    fromColumn: input.fromColumn as number,
    toRow: input.toRow as number,
    toColumn: input.toColumn as number,
    fromRowOffset: input.fromRowOffset as number,
    fromColumnOffset: input.fromColumnOffset as number,
    toRowOffset: input.toRowOffset as number,
    toColumnOffset: input.toColumnOffset as number,
  }
}

const CHART_LEGENDS = ['none', 'right', 'bottom', 'top', 'left'] as const
const CHART_DATA_LABELS = ['none', 'value', 'percent', 'category-percent'] as const
const CHART_DATA_LABEL_POSITIONS = ['center', 'inside-end', 'outside-end'] as const
const CHART_GROUPINGS = ['clustered', 'stacked', 'percentStacked', 'standard'] as const

function parseChart(input: unknown): NonNullable<WorkbookVisualObject['chart']> {
  if (
    !isRecord(input) ||
    !Array.isArray(input.chartTypes) ||
    input.chartTypes.some((value) => typeof value !== 'string') ||
    typeof input.title !== 'string' ||
    !Array.isArray(input.series) ||
    !isOptionalString(input.barDirection) ||
    !isOptionalEnum(input.legend, CHART_LEGENDS) ||
    !isOptionalEnum(input.dataLabels, CHART_DATA_LABELS) ||
    !isOptionalEnum(input.dataLabelPosition, CHART_DATA_LABEL_POSITIONS) ||
    !isOptionalString(input.dataLabelFormat) ||
    !isOptionalEnum(input.grouping, CHART_GROUPINGS) ||
    (input.gridlines !== undefined && typeof input.gridlines !== 'boolean') ||
    !isOptionalFiniteNumber(input.gapWidthPct) ||
    !isOptionalFiniteNumber(input.holeSizePct)
  ) {
    throw new Error('Invalid workbook chart response.')
  }
  const axisTitles = parseChartAxisTitles(input.axisTitles)
  const valueAxis = parseChartValueAxis(input.valueAxis)
  const series = input.series.map((entry) => {
    if (
      !isRecord(entry) ||
      typeof entry.name !== 'string' ||
      !Array.isArray(entry.categories) ||
      entry.categories.some((value) => typeof value !== 'string') ||
      !Array.isArray(entry.values) ||
      entry.values.some((value) => typeof value !== 'number' || !Number.isFinite(value)) ||
      !isOptionalString(entry.numberFormat) ||
      !isOptionalString(entry.color) ||
      !isOptionalString(entry.trendline) ||
      !isOptionalString(entry.valuesRef) ||
      !isOptionalString(entry.categoriesRef) ||
      !isOptionalFiniteNumber(entry.explosionPct)
    ) {
      throw new Error('Invalid workbook chart series.')
    }
    const pointColors = parseChartPointColors(entry.pointColors)
    const pointExplosions = parseChartPointExplosions(entry.pointExplosions)
    return {
      name: entry.name,
      categories: entry.categories,
      values: entry.values,
      ...(entry.numberFormat === undefined ? {} : { numberFormat: entry.numberFormat }),
      ...(entry.color === undefined ? {} : { color: entry.color }),
      ...(entry.trendline === undefined ? {} : { trendline: entry.trendline }),
      ...(entry.valuesRef === undefined ? {} : { valuesRef: entry.valuesRef }),
      ...(entry.categoriesRef === undefined ? {} : { categoriesRef: entry.categoriesRef }),
      ...(pointColors === undefined ? {} : { pointColors }),
      ...(entry.explosionPct === undefined ? {} : { explosionPct: entry.explosionPct }),
      ...(pointExplosions === undefined ? {} : { pointExplosions }),
    }
  })
  return {
    chartTypes: input.chartTypes,
    title: input.title,
    series,
    ...(input.barDirection === undefined ? {} : { barDirection: input.barDirection }),
    ...(input.legend === undefined ? {} : { legend: input.legend }),
    ...(axisTitles === undefined ? {} : { axisTitles }),
    ...(input.dataLabels === undefined ? {} : { dataLabels: input.dataLabels }),
    ...(input.dataLabelPosition === undefined
      ? {}
      : { dataLabelPosition: input.dataLabelPosition }),
    ...(input.dataLabelFormat === undefined ? {} : { dataLabelFormat: input.dataLabelFormat }),
    ...(input.grouping === undefined ? {} : { grouping: input.grouping }),
    ...(input.gridlines === undefined ? {} : { gridlines: input.gridlines }),
    ...(valueAxis === undefined ? {} : { valueAxis }),
    ...(input.gapWidthPct === undefined ? {} : { gapWidthPct: input.gapWidthPct }),
    ...(input.holeSizePct === undefined ? {} : { holeSizePct: input.holeSizePct }),
  }
}

function parseChartAxisTitles(
  input: unknown,
): { category?: string | null; value?: string | null } | undefined {
  if (input === undefined) return undefined
  if (
    !isRecord(input) ||
    !isOptionalNullableString(input.category) ||
    !isOptionalNullableString(input.value)
  ) {
    throw new Error('Invalid workbook chart axis titles.')
  }
  return {
    ...(input.category === undefined ? {} : { category: input.category }),
    ...(input.value === undefined ? {} : { value: input.value }),
  }
}

function parseChartValueAxis(input: unknown): { min?: number; max?: number } | undefined {
  if (input === undefined) return undefined
  if (
    !isRecord(input) ||
    !isOptionalFiniteNumber(input.min) ||
    !isOptionalFiniteNumber(input.max)
  ) {
    throw new Error('Invalid workbook chart value axis.')
  }
  return {
    ...(input.min === undefined ? {} : { min: input.min }),
    ...(input.max === undefined ? {} : { max: input.max }),
  }
}

function parseChartPointColors(
  input: unknown,
): Array<{ index: number; color: string }> | undefined {
  if (input === undefined) return undefined
  if (!Array.isArray(input)) throw new Error('Invalid workbook chart point colors.')
  return input.map((entry) => {
    if (!isRecord(entry) || !isNonnegativeInteger(entry.index) || typeof entry.color !== 'string') {
      throw new Error('Invalid workbook chart point colors.')
    }
    return { index: entry.index, color: entry.color }
  })
}

function parseChartPointExplosions(
  input: unknown,
): Array<{ index: number; pct: number }> | undefined {
  if (input === undefined) return undefined
  if (!Array.isArray(input)) throw new Error('Invalid workbook chart point explosions.')
  return input.map((entry) => {
    if (
      !isRecord(entry) ||
      !isNonnegativeInteger(entry.index) ||
      typeof entry.pct !== 'number' ||
      !Number.isFinite(entry.pct)
    ) {
      throw new Error('Invalid workbook chart point explosions.')
    }
    return { index: entry.index, pct: entry.pct }
  })
}

function isRecord(input: unknown): input is Record<string, unknown> {
  return typeof input === 'object' && input !== null && !Array.isArray(input)
}

function isUuid(input: unknown): input is string {
  return (
    typeof input === 'string' &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(input)
  )
}

function isNonnegativeInteger(input: unknown): input is number {
  return typeof input === 'number' && Number.isInteger(input) && input >= 0
}

/// Sheet default row height / column width: finite positive number, or null
/// for "no default — use the built-in size" (0 and junk degrade to null).
function normalizedDefaultSize(input: unknown): number | null {
  return typeof input === 'number' && Number.isFinite(input) && input > 0 ? input : null
}

function isPositiveInteger(input: unknown): input is number {
  return isNonnegativeInteger(input) && input > 0
}

function isCellScalar(input: unknown): input is string | number | boolean | null {
  return input === null || ['string', 'number', 'boolean'].includes(typeof input)
}

function isOptionalString(input: unknown): input is string | undefined {
  return input === undefined || typeof input === 'string'
}

function isOptionalNullableString(input: unknown): input is string | null | undefined {
  return input === undefined || input === null || typeof input === 'string'
}

function isOptionalFiniteNumber(input: unknown): input is number | undefined {
  return input === undefined || (typeof input === 'number' && Number.isFinite(input))
}

function isOptionalEnum<T extends string>(
  input: unknown,
  values: readonly T[],
): input is T | undefined {
  return input === undefined || (values as readonly unknown[]).includes(input)
}
