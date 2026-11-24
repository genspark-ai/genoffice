import { createHash, randomUUID } from 'node:crypto'
import { handOffBytes } from './byte-handoff'
import {
  appendFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rm,
  stat,
  unlink,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, extname, isAbsolute, join } from 'node:path'
import { moveFile, renamedDocPath } from './doc-move'
import { pathToFileURL } from 'node:url'
import {
  BrowserWindow,
  Menu,
  WebContentsView,
  app,
  clipboard,
  dialog,
  ipcMain,
  nativeImage,
  net,
  screen,
  session,
  shell,
  webContents,
} from 'electron'
import {
  abortOnDestroyed,
  appMenuLabels,
  appMenuTemplate,
  buildPrintableHtml,
  configuredDefaultSaveDir,
  contextMenuLabels,
  fetchRemoteImage,
  installContextMenu,
  setContextMenuInterceptor,
  installNavigationGuard,
  isHeadlessMode,
  printHtmlToPdf,
  safeExternalUrl,
  saveAsSuggestion,
  saveImageFromUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  aboutMenuItem,
  checkUpdatesMenuItem,
  toggleDevToolsItem,
  settingsMenuItem,
  windowMenuTemplate,
  type HeadlessExportFormat,
  type HeadlessExportTarget,
  installRendererProtocol,
  registerRendererScheme,
  rendererUrl,
  MAX_REMOTE_IMAGE_BYTES,
  readBodyCapped,
  writeJsonAtomic,
  restrictFileToOwner,
  isBesideDocument,
  sharedPathGrants,
  recoveryCopyDiskNewer,
  readLegacyImportSkip,
  writeLegacyImportSkip,
  tileWindows,
} from '@genoffice/electron-utils'
import {
  configureMetricsCache,
  familyAliases,
  familyVerticalMetrics,
} from '@genoffice/font-metrics'
import { createI18n, getUiLang, normalizeLang, setUiLang } from '@genoffice/i18n'
import { ProjectStore } from '@genoffice/project-store'
import type {
  IpcMainInvokeEvent,
  MenuItem,
  MenuItemConstructorOptions,
  OpenDialogOptions,
  SaveDialogOptions,
  WebContents,
} from 'electron'
import { parseFileToText } from '@genoffice/file-parse'
import { convertLegacyDocToDocx } from '../../../../packages/doc2docx/src'
import { convertHtmlToDocx } from '../../../../packages/html2docx/src'
import { ElectronBrowserDriver } from '../../../../packages/html2docx/src/drivers/electron'
import { convertOdtToDocx } from '../../../../packages/odf2ooxml/src'
import { convertHwpToDocx } from '../../../../packages/hwp2docx/src'
import {
  AiCreditsError,
  AiTimeoutError,
  isAiNetworkError,
  isAiOverloadedError,
  isAiRegionBlockedError,
  chatForProvider,
  defaultAiSettings,
  activeProvider,
  testMediaProvider,
  type AiMediaProviderConfig,
  type AiMediaProviderId,
  type AiSearchProviderId,
  resolveAiSettings,
  maxOutputTokensOf,
  sanitizeAiSettings,
  sanitizeCliPath,
  setAiUserAgent,
  setRescueFetch,
  resolveProfileCredential,
  streamForProvider,
  type AiChatRequest,
  type AiSettings,
  type AiStreamChunk,
  type AiStreamRequest,
  type GenSparkAccountStatus,
  type LegacyAiSettings,
} from '@genoffice/ai-provider'
import { listCodexModels, shutdownCodexAppServers } from '@genoffice/ai-provider/codex-app-server'
import { startAiTrace } from '@genoffice/ai-provider/trace'
import { listCustomModelsForIpc } from '@genoffice/ai-provider/custom-models'
import {
  ensureGenofficeLogin,
  noteGskCallFailure,
  setGenofficeLoginOpener,
  gskApiKey,
  generateImageTool,
  testSearchProvider,
  gskLoginInfo,
  hasGskAuth,
  webSearchTool,
  imageSearchTool,
  analyzeMediaTool,
  documentMediaRoots,
} from '@genoffice/ai-search'
import type {
  AiDocContent,
  DocFileInfo,
  DocMoveResult,
  AttachmentAddResult,
  AttachmentImageResult,
  AttachmentMeta,
  AttachmentReadResult,
  ContextMenuRequest,
  CreateDocumentRequest,
  CreateDocumentResult,
  DecryptOpenResult,
  DocsTabInfo,
  MenuCommand,
  MergeOutputRequest,
  OpenDocxResult,
  OrganizerFileResult,
  PickDocxResult,
  SpellLanguages,
  TemplateEntry,
  TemplatePreviewRect,
} from '../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS, MARKDOWN_MEDIA_DIR_TOKEN, type MarkdownImage } from '../shared/ipc'
import {
  COMMANDS,
  toCommandMenuState,
  type CommandDef,
  type CommandId,
  type CommandMenuState,
} from '../shared/commands'
import { DOCS_MENU, type MenuNode, type MenuSlot, type MenuTop } from '../shared/menu-layout'
import {
  bindingId,
  keepsAccelerator,
  menuAccelerator,
  sanitizeKeyOverrides,
  type KeyOverrides,
} from '../shared/keybindings'
import { MERGE_DATA_SOURCE_EXTENSIONS, readMergeDataSource } from './merge-data-source'
import { mergeOutputFileName, mergeOutputPath } from './merge-output'
import { ClickClaims } from '../shared/context-menu-claims'
import { SAVE_AS_EXPORT_FILTERS, saveAsExportFormat } from '../shared/save-as-format'
import {
  OPENABLE_DOC_EXTENSIONS,
  docImportKindForPath,
  findOpenableDocPath,
  isOpenableDocPath,
} from '../shared/open-file'
import { atomicWriteFile, looksLikeZip } from './atomic-write'
import { templateThumbnail, withTemplateThumbnail } from '@genoffice/docx-engine'
import { uniquePathIn } from './file-name'
import { mapBounded } from './map-bounded'
// the shell imports it from here
export { uniquePathIn }
import {
  adoptLazyMediaHashes,
  forgetLazyMediaOwner,
  materializeLazyDocx,
  moveLazyMediaSource,
  openLazyDocx,
  pointLazyMediaAt,
  readLazyMedia,
  registerLazyMediaProtocol,
} from './lazy-media'
import { inlineLazyMediaInHtml } from './lazy-media-inline'
import {
  commitDocPasswordSave,
  currentDocPasswordIntentRevision,
  decryptDocx,
  decryptRecoveryCopy,
  discardDocPasswordIntents,
  DocxDecryptError,
  docPasswordFor,
  encryptDocx,
  forgetDocPasswords,
  isEncryptedDocx,
  markDiskEncrypted,
  prepareRecoveryDocx,
  rememberDocPassword,
  renameDocPassword,
  setDocPassword,
  snapshotDocPassword,
} from './docx-encryption'
import { isExternallyModified, type DiskFileState } from './external-change'
import { copyImageDisplaySize, validCopyImageDataUrl } from './copy-image-guard'
import {
  printJobOptions,
  printPaperOption,
  printScaleOption,
  validPrintGeometry,
  validPrintPaper,
} from './print-args'
import { printPdfViaCups } from '@genoffice/electron-utils/print-cups'
import { printOutcome } from '@genoffice/electron-utils/print-job'
import { addPdfInternalLinks } from './pdf-links'
import {
  assignStarredGroup,
  dropStarredItems,
  readStarredItems,
  renameStarredItem,
  starredGroupMap,
  starredGroupNames,
  toggleStarredItem,
  writeStarredItems,
} from './starred-files'
import { initDocsAutoUpdater } from './updater'
import { registerZoteroIpc, teardownZoteroIpc } from './zotero-ipc'

/**
 * Docs main-process logic as an embeddable module: no top-level side effects.
 * Standalone mode (apps/docs entry) calls startDocsStandalone(); the unified
 * shell (apps/shell) instead calls configureDocsRuntime() + registerDocsIpc()
 * + createDocsWindow() and owns the app lifecycle itself.
 */

const isDev = !!process.env.ELECTRON_RENDERER_URL

const tMain = createI18n({
  zh: {
    dlgOpenDoc: '打开文档',
    filterWord: 'Word 文档',
    odtImportMsg: '转换 .{ext} 文档？',
    odtImportDetail:
      'GenOffice 只编辑 .docx，不编辑 ODF。正文、标题、列表、表格和图片会转换过来；脚注、批注和文本框不会完整保留。原文件不会被修改，转换出的 .docx 是一份独立副本，之后的保存都写入它。', // public-hygiene: allow
    hwpImportDetail:
      'GenOffice 只编辑 .docx，不编辑 HWP。正文、表格、图片、页眉和页脚会转换过来；公式、脚注和嵌入对象不会转换，自动编号会变成普通文本。原文件不会被修改，转换出的 .docx 是一份独立副本，之后的保存都写入它。', // public-hygiene: allow
    odtImportConvertBtn: '转换并打开',
    importSkipConfirm: '下次直接转换，不再提示',
    docImportMsg: '转换 .{ext} 文档？',
    docImportDetail:
      'GenOffice 只编辑 .docx，不编辑 Word 97-2003 旧版 .doc。正文、标题、列表、表格和图片会转换过来；脚注、批注和文本框不会完整保留。原文件不会被修改，转换出的 .docx 是一份独立副本，之后的保存都写入它。', // public-hygiene: allow
    docImportConvertBtn: '转换并打开',
    dlgSaveAs: '另存为',
    closeUnsavedMsg: '此文档有未保存的更改。',
    closeUnsavedDetail: '关闭前是否保存？',
    closeNoReplyMsg: '文档没有响应,可能有未保存的更改。',
    closeNoReplyDetail: '仍要关闭吗?未保存的更改将丢失。',
    btnCloseAnyway: '仍要关闭',
    autosaveFoundTitle: '发现自动恢复版本',
    autosaveDiskNewer: '注意：此恢复副本保存后，磁盘上的文件又被修改过。恢复将替换这些较新的更改。',
    autosaveFoundBody: '上次会话有未保存的更改。要恢复自动保存的版本吗?',
    autosaveRestore: '恢复',
    autosaveDiscard: '放弃',
    btnDontSave: '不保存',
    btnCancel: '取消',
    extModifiedMsg: '文件已被其他程序修改。',
    extModifiedDetail: '仍要保存并覆盖磁盘上的更改吗?',
    btnOverwrite: '覆盖',
    dlgInsertImage: '插入图片',
    filterImages: '图片',
    dlgMergeDataSource: '选择收件人列表',
    filterMergeDataSource: 'Excel 或 CSV',
    dlgAddAttachment: '添加附件',
    filterSupported: '支持的文件',
    filterAll: '所有文件',
    dlgExportPdf: '导出为 PDF',
    dlgExportHtml: '导出为 HTML',
    dlgExportMarkdown: '导出为 Markdown',
    dlgPickExportDir: '选择导出目录',
    errUnsupportedExt: '暂不支持 .{ext} 类型',
    errNotFile: '不是文件',
    errTooLarge: '超过 {mb}MB 上限',
    errTemplateProtected: '此模板受密码保护，请通过“文件 > 打开”打开',
    errImageTooLarge: '图片超过 5MB 上限',
    errUnreadable: '无法读取',
    errFileTooLarge: '文件超过大小上限',
    errParseFailed: '文件解析失败',
    errImageNoText: '图片附件不提供文本,已作为图像随用户消息发送,直接看图即可',
    errNotImage: '不是支持的图片类型',
    errGskNotLoggedIn: '未登录 Genspark:请点击下方「登录 Genspark」完成登录后重试',
    errNoApiKey: '未配置 {provider} 的 API Key',
    errAiRegion: '该模型的厂商不向当前地区提供服务，请换一个模型或切换网络',
    errAiBusy: 'AI 服务当前繁忙，请稍后重试',
    errNoModel: '未配置模型名称',
    menuFile: '文件',
    menuNewDoc: '新建文档',
    menuNewWindow: '新建窗口',
    menuOpen: '打开…',
    menuOpenRecent: '打开最近的文档',
    menuNoRecent: '无最近文档',
    menuRecentMore: '更多…',
    menuClose: '关闭',
    menuSave: '保存',
    menuSaveAs: '另存为…',
    menuPageSetup: '页面设置…',
    menuExportPdf: '导出为 PDF…',
    menuExportHtml: '导出为 HTML…',
    menuExportMarkdown: '导出为 Markdown…',
    menuExportImages: '导出为图片…',
    menuPrint: '打印…',
    menuEdit: '编辑',
    menuUndo: '撤销',
    menuRedo: '重做',
    menuRepeat: '重复',
    menuCantRepeat: '无法重复',
    menuCut: '剪切',
    menuCopy: '复制',
    menuPaste: '粘贴',
    menuPasteMatch: '粘贴并匹配格式',
    menuFindReplace: '查找和替换…',
    menuPasteSpecial: '选择性粘贴…',
    menuClear: '清除',
    menuClearFormatting: '清除格式',
    menuClearContents: '内容',
    menuFindMenu: '查找',
    menuFind: '查找…',
    menuFindNext: '查找下一个',
    menuFindPrevious: '查找上一个',
    menuReplace: '替换…',
    menuAdvancedFind: '高级查找和替换…',
    menuLinks: '链接…',
    menuSelectData: '选择数据…',
    menuToggleDrawing: '切换绘图',
    menuNewFromTemplate: '根据模板新建…',
    menuSaveAsTemplate: '另存为模板…',
    menuMove: '移动…',
    menuRename: '重命名…',
    menuAlwaysOpenReadOnly: '始终以只读方式打开',
    menuReduceFileSize: '减小文件大小…',
    menuProperties: '属性…',
    menuGoTo: '定位…',
    menuSelectAll: '全选',
    menuView: '视图',
    menuZoom: '缩放',
    menuZoomIn: '放大',
    menuZoomOut: '缩小',
    menuZoom100: '实际大小 (100%)',
    menuPrintLayout: '页面视图',
    menuWebLayout: 'Web 版式视图',
    menuOutline: '大纲',
    menuDraft: '草稿',
    menuFocus: '专注',
    menuRibbon: '功能区',
    menuSidebar: '边栏',
    menuThumbnail: '缩略图',
    menuNavigation: '导航',
    menuReviewing: '审阅',
    menuSidebarSearch: '搜索',
    menuRuler: '标尺',
    menuHeaderFooter: '页眉和页脚',
    menuFootnotes: '脚注',
    menuMarkup: '标记',
    menuRevealFormatting: '显示格式',
    menuZoomDialog: '显示比例…',
    menuGridlines: '网格线',
    menuPageWidth: '页宽',
    menuWholePage: '整页',
    menuMultiPage: '多页',
    menuAiSidebar: 'AI 侧栏',
    menuDarkMode: '深色模式',
    menuFullscreen: '进入全屏',
    menuInsert: '插入',
    menuInsertTable: '表格…',
    menuInsertImage: '图片…',
    menuInsertPageBreak: '分页符',
    menuInsertLink: '超链接…',
    menuInsertEquation: '公式…',
    menuPictures: '图片',
    menuPictureFromFile: '来自文件的图片…',
    menuShape: '形状',
    menuShapeRectangle: '矩形',
    menuShapeRoundedRectangle: '圆角矩形',
    menuShapeTriangle: '三角形',
    menuShapeOval: '椭圆',
    menuShapeLine: '直线',
    menuChart: '图表',
    menuChartColumn: '柱形图',
    menuChartBar: '条形图',
    menuChartLine: '折线图',
    menuChartArea: '面积图',
    menuChartPie: '饼图',
    menuChartTreemap: '树状图',
    menuChartSunburst: '旭日图',
    menuChartHistogram: '直方图',
    menuChartPareto: '排列图',
    menuChartBoxWhisker: '箱形图',
    menuChartScatter: 'XY (散点图)',
    menuChartWaterfall: '瀑布图',
    menuChartFunnel: '漏斗图',
    menuChartStock: '股价图',
    menuChartSurface: '曲面图',
    menuChartRadar: '雷达图',
    menuIcons: '图标…',
    menuBookmark: '书签…',
    menuCrossReference: '交叉引用…',
    menuBreak: '分隔符',
    menuColumnBreak: '分栏符',
    menuSectionNextPage: '分节符(下一页)',
    menuSectionContinuous: '分节符(连续)',
    menuSectionOddPage: '分节符(奇数页)',
    menuSectionEvenPage: '分节符(偶数页)',
    menuFootnote: '脚注…',
    menuCaption: '题注…',
    menuIndexAndTables: '索引和目录…',
    menuWatermark: '水印…',
    menuPageNumbers: '页码…',
    menuTextBox: '文本框',
    menuVerticalTextBox: '绘制竖排文本框',
    menuAutoText: '自动图文集',
    menuAutoTextDialog: '自动图文集…',
    menuAutoTextNew: '新建…',
    menuDateAndTime: '日期和时间…',
    menuObject: '对象…',
    menuInsertFile: '文件…',
    menuField: '域…',
    menuAdvancedSymbol: '高级符号…',
    menuComment: '批注',
    menuFormat: '格式',
    menuBold: '加粗',
    menuItalic: '斜体',
    menuUnderline: '下划线',
    menuAlign: '对齐',
    menuAlignLeft: '左对齐',
    menuAlignCenter: '居中',
    menuAlignRight: '右对齐',
    menuAlignJustify: '两端对齐',
    menuFont: '字体…',
    menuParagraph: '段落…',
    menuTextEffects: '文本效果…',
    menuDocument: '文档…',
    menuBulletsAndNumbering: '项目符号和编号…',
    menuBordersAndShading: '边框和底纹…',
    menuColumns: '分栏…',
    menuTabs: '制表位…',
    menuDropCap: '首字下沉…',
    menuTextDirection: '文字方向…',
    menuChangeCase: '更改大小写…',
    menuAsianLayout: '中文版式',
    menuPhoneticGuide: '拼音指南…',
    menuEncloseCharacters: '带圈字符…',
    menuHorizontalInVertical: '纵横混排…',
    menuCombineCharacters: '合并字符…',
    menuTwoLinesInOne: '双行合一…',
    menuFitText: '调整宽度…',
    menuStyle: '样式…',
    menuFontSubstitution: '字体替换…',
    menuFormatObject: '设置对象格式…',
    menuEquationOptions: '公式选项…',
    menuSelectionPane: '选择窗格…',
    menuTools: '工具',
    menuTable: '表格',
    menuTableInsert: '插入',
    menuTableInsertTable: '表格…',
    menuTableColsLeft: '在左侧插入列',
    menuTableColsRight: '在右侧插入列',
    menuTableRowsAbove: '在上方插入行',
    menuTableRowsBelow: '在下方插入行',
    menuTableCells: '单元格…',
    menuTableDelete: '删除',
    menuTableDeleteTable: '表格',
    menuTableDeleteColumns: '列',
    menuTableDeleteRows: '行',
    menuTableSelect: '选择',
    menuTableSelectCell: '单元格',
    menuTableSelectColumn: '列',
    menuTableSelectRow: '行',
    menuTableSelectTable: '表格',
    menuTableMergeCells: '合并单元格',
    menuTableSplitCells: '拆分单元格…',
    menuTableSplitTable: '拆分表格',
    menuTableAutoFit: '自动调整',
    menuTableAutoFitContents: '根据内容自动调整表格',
    menuTableAutoFitWindow: '根据窗口自动调整表格',
    menuTableFixedWidth: '固定列宽',
    menuTableDistributeRows: '平均分布各行',
    menuTableDistributeColumns: '平均分布各列',
    menuTableRepeatHeader: '重复标题行',
    menuTableConvert: '转换',
    menuTableConvertTextToTable: '文本转换成表格…',
    menuTableConvertTableToText: '表格转换成文本…',
    menuTableSort: '排序…',
    menuTableFormula: '公式…',
    menuTableGridlines: '查看网格线',
    menuTableProperties: '表格属性…',
    menuWordCount: '字数统计…',
    menuAutoCorrect: '自动更正选项…',
    menuSpellingAndGrammar: '拼写和语法',
    menuSpellingDialog: '拼写和语法…',
    menuEditor: '编辑器',
    menuHideSpellingErrors: '隐藏拼写错误',
    menuHideGrammarErrors: '隐藏语法错误',
    menuResetIgnoredWords: '重置忽略的单词和语法',
    menuThesaurus: '同义词库…',
    menuHyphenation: '断字…',
    menuTranslate: '翻译',
    menuTranslateSelection: '翻译所选内容',
    menuTranslateDocument: '翻译文档',
    menuSetTranslationLanguage: '设置文档翻译语言…',
    menuLanguage: '语言…',
    menuTrackChanges: '修订',
    menuHighlightChanges: '突出显示修订…',
    menuCompareDocuments: '比较文档…',
    menuCombineDocuments: '合并文档…',
    menuProtectDocument: '保护文档…',
    menuEnvelopes: '信封…',
    menuLabels: '标签…',
    menuCustomizeKeyboard: '自定义键盘…',
    menuPreferences: '偏好设置…',
    menuAiProofread: 'AI 校对',
    menuWindow: '窗口',
    menuHelp: '帮助',
    menuShortcuts: '键盘快捷键',
  },
  en: {
    dlgOpenDoc: 'Open Document',
    filterWord: 'Word Documents',
    odtImportMsg: 'Convert this .{ext} document?',
    odtImportDetail:
      'GenOffice edits .docx, not ODF. Body text, headings, lists, tables and images carry over; footnotes, comments and text boxes are not fully preserved. The original file is never changed; the converted .docx is a separate copy that later saves write to.',
    hwpImportDetail:
      'GenOffice edits .docx, not HWP. Body text, tables, images, headers and footers carry over; equations, footnotes and embedded objects are not converted, and automatic numbering becomes plain text. The original file is never changed; the converted .docx is a separate copy that later saves write to.',
    odtImportConvertBtn: 'Convert and open',
    importSkipConfirm: 'Convert without asking next time',
    docImportMsg: 'Convert this .{ext} document?',
    docImportDetail:
      'GenOffice edits .docx, not Word 97-2003 documents. Body text, headings, lists, tables and images carry over; footnotes, comments and text boxes are not fully preserved. The original file is never changed; the converted .docx is a separate copy that later saves write to.',
    docImportConvertBtn: 'Convert and open',
    dlgSaveAs: 'Save As',
    closeUnsavedMsg: 'This document has unsaved changes.',
    closeUnsavedDetail: 'Do you want to save them before closing?',
    closeNoReplyMsg: 'The document is not responding and may have unsaved changes.',
    closeNoReplyDetail: 'Close anyway? Unsaved changes will be lost.',
    btnCloseAnyway: 'Close Anyway',
    autosaveFoundTitle: 'Recovered version found',
    autosaveDiskNewer:
      'Note: the file on disk was modified after this recovery copy was saved. Restoring replaces those newer changes.',
    autosaveFoundBody:
      'There are unsaved changes from your last session. Restore the autosaved version?',
    autosaveRestore: 'Restore',
    autosaveDiscard: 'Discard',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
    extModifiedMsg: 'The file has been modified by another program.',
    extModifiedDetail: 'Save anyway and overwrite the changes on disk?',
    btnOverwrite: 'Overwrite',
    dlgInsertImage: 'Insert Image',
    filterImages: 'Images',
    dlgMergeDataSource: 'Select Recipient List',
    filterMergeDataSource: 'Excel or CSV',
    dlgAddAttachment: 'Add Attachments',
    filterSupported: 'Supported Files',
    filterAll: 'All Files',
    dlgExportPdf: 'Export as PDF',
    dlgExportHtml: 'Export as HTML',
    dlgExportMarkdown: 'Export as Markdown',
    dlgPickExportDir: 'Choose Export Folder',
    errUnsupportedExt: '.{ext} files are not supported',
    errNotFile: 'not a file',
    errTooLarge: 'exceeds the {mb}MB limit',
    errTemplateProtected: 'this template is password protected; open it with File > Open',
    errImageTooLarge: 'image exceeds the 5MB limit',
    errUnreadable: 'cannot be read',
    errFileTooLarge: 'File exceeds the size limit',
    errParseFailed: 'Failed to parse file',
    errImageNoText: 'Image attachments have no text; the image is sent along with the user message',
    errNotImage: 'not a supported image type',
    errGskNotLoggedIn:
      'Not signed in to Genspark: click “Sign in to Genspark” below, sign in, then retry',
    errNoApiKey: 'No API key configured for {provider}',
    errAiRegion:
      "This model's provider does not serve your current region — pick another model or change your network",
    errAiBusy: 'The AI service is busy right now — please try again in a moment',
    errNoModel: 'No model name configured',
    menuFile: 'File',
    menuNewDoc: 'New Document',
    menuNewWindow: 'New Window',
    menuOpen: 'Open…',
    menuOpenRecent: 'Open Recent',
    menuNoRecent: 'No Recent Documents',
    menuRecentMore: 'More…',
    menuClose: 'Close',
    menuSave: 'Save',
    menuSaveAs: 'Save As…',
    menuPageSetup: 'Page Setup…',
    menuExportPdf: 'Export as PDF…',
    menuExportHtml: 'Export as HTML…',
    menuExportMarkdown: 'Export as Markdown…',
    menuExportImages: 'Export as Images…',
    menuPrint: 'Print…',
    menuEdit: 'Edit',
    menuUndo: 'Undo',
    menuRedo: 'Redo',
    menuRepeat: 'Repeat',
    menuCantRepeat: "Can't Repeat",
    menuCut: 'Cut',
    menuCopy: 'Copy',
    menuPaste: 'Paste',
    menuPasteMatch: 'Paste and Match Formatting',
    menuFindReplace: 'Find and Replace…',
    menuPasteSpecial: 'Paste Special…',
    menuClear: 'Clear',
    menuClearFormatting: 'Clear Formatting',
    menuClearContents: 'Contents',
    menuFindMenu: 'Find',
    menuFind: 'Find…',
    menuFindNext: 'Find Next',
    menuFindPrevious: 'Find Previous',
    menuReplace: 'Replace…',
    menuAdvancedFind: 'Advanced Find and Replace…',
    menuLinks: 'Links…',
    menuSelectData: 'Select Data…',
    menuToggleDrawing: 'Toggle Drawing',
    menuNewFromTemplate: 'New from Template…',
    menuSaveAsTemplate: 'Save as Template…',
    menuMove: 'Move…',
    menuRename: 'Rename…',
    menuAlwaysOpenReadOnly: 'Always Open Read-Only',
    menuReduceFileSize: 'Reduce File Size…',
    menuProperties: 'Properties…',
    menuGoTo: 'Go To…',
    menuSelectAll: 'Select All',
    menuView: 'View',
    menuZoom: 'Zoom',
    menuZoomIn: 'Zoom In',
    menuZoomOut: 'Zoom Out',
    menuZoom100: 'Actual Size (100%)',
    menuPrintLayout: 'Print Layout',
    menuWebLayout: 'Web Layout',
    menuOutline: 'Outline',
    menuDraft: 'Draft',
    menuFocus: 'Focus',
    menuRibbon: 'Ribbon',
    menuSidebar: 'Sidebar',
    menuThumbnail: 'Thumbnail',
    menuNavigation: 'Navigation',
    menuReviewing: 'Reviewing',
    menuSidebarSearch: 'Search',
    menuRuler: 'Ruler',
    menuHeaderFooter: 'Header and Footer',
    menuFootnotes: 'Footnotes',
    menuMarkup: 'Markup',
    menuRevealFormatting: 'Reveal Formatting',
    menuZoomDialog: 'Zoom…',
    menuGridlines: 'Gridlines',
    menuPageWidth: 'Page Width',
    menuWholePage: 'Whole Page',
    menuMultiPage: 'Multiple Pages',
    menuAiSidebar: 'AI Sidebar',
    menuDarkMode: 'Dark Mode',
    menuFullscreen: 'Enter Full Screen',
    menuInsert: 'Insert',
    menuInsertTable: 'Table…',
    menuInsertImage: 'Image…',
    menuInsertPageBreak: 'Page Break',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Equation…',
    menuPictures: 'Pictures',
    menuPictureFromFile: 'Picture from File…',
    menuShape: 'Shape',
    menuShapeRectangle: 'Rectangle',
    menuShapeRoundedRectangle: 'Rounded Rectangle',
    menuShapeTriangle: 'Triangle',
    menuShapeOval: 'Oval',
    menuShapeLine: 'Line',
    menuChart: 'Chart',
    menuChartColumn: 'Column',
    menuChartBar: 'Bar',
    menuChartLine: 'Line',
    menuChartArea: 'Area',
    menuChartPie: 'Pie',
    menuChartTreemap: 'Treemap',
    menuChartSunburst: 'Sunburst',
    menuChartHistogram: 'Histogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Box and Whisker',
    menuChartScatter: 'X Y (Scatter)',
    menuChartWaterfall: 'Waterfall',
    menuChartFunnel: 'Funnel',
    menuChartStock: 'Stock',
    menuChartSurface: 'Surface',
    menuChartRadar: 'Radar',
    menuIcons: 'Icons…',
    menuBookmark: 'Bookmark…',
    menuCrossReference: 'Cross-reference…',
    menuBreak: 'Break',
    menuColumnBreak: 'Column Break',
    menuSectionNextPage: 'Section Break (Next Page)',
    menuSectionContinuous: 'Section Break (Continuous)',
    menuSectionOddPage: 'Section Break (Odd Page)',
    menuSectionEvenPage: 'Section Break (Even Page)',
    menuFootnote: 'Footnote…',
    menuCaption: 'Caption…',
    menuIndexAndTables: 'Index and Tables…',
    menuWatermark: 'Watermark…',
    menuPageNumbers: 'Page Numbers…',
    menuTextBox: 'Text Box',
    menuVerticalTextBox: 'Draw Vertical Text Box',
    menuAutoText: 'AutoText',
    menuAutoTextDialog: 'AutoText…',
    menuAutoTextNew: 'New…',
    menuDateAndTime: 'Date and Time…',
    menuObject: 'Object…',
    menuInsertFile: 'File…',
    menuField: 'Field…',
    menuAdvancedSymbol: 'Advanced Symbol…',
    menuComment: 'Comment',
    menuFormat: 'Format',
    menuBold: 'Bold',
    menuItalic: 'Italic',
    menuUnderline: 'Underline',
    menuAlign: 'Align',
    menuAlignLeft: 'Align Left',
    menuAlignCenter: 'Center',
    menuAlignRight: 'Align Right',
    menuAlignJustify: 'Justify',
    menuFont: 'Font…',
    menuParagraph: 'Paragraph…',
    menuTextEffects: 'Text Effects…',
    menuDocument: 'Document…',
    menuBulletsAndNumbering: 'Bullets and Numbering…',
    menuBordersAndShading: 'Borders and Shading…',
    menuColumns: 'Columns…',
    menuTabs: 'Tabs…',
    menuDropCap: 'Drop Cap…',
    menuTextDirection: 'Text Direction…',
    menuChangeCase: 'Change Case…',
    menuAsianLayout: 'Asian Layout',
    menuPhoneticGuide: 'Phonetic Guide…',
    menuEncloseCharacters: 'Enclose Characters…',
    menuHorizontalInVertical: 'Horizontal in Vertical…',
    menuCombineCharacters: 'Combine Characters…',
    menuTwoLinesInOne: 'Two Lines in One…',
    menuFitText: 'Fit Text…',
    menuStyle: 'Style…',
    menuFontSubstitution: 'Font Substitution…',
    menuFormatObject: 'Format Object…',
    menuEquationOptions: 'Equation Options…',
    menuSelectionPane: 'Selection Pane…',
    menuTools: 'Tools',
    menuTable: 'Table',
    menuTableInsert: 'Insert',
    menuTableInsertTable: 'Table…',
    menuTableColsLeft: 'Columns to the Left',
    menuTableColsRight: 'Columns to the Right',
    menuTableRowsAbove: 'Rows Above',
    menuTableRowsBelow: 'Rows Below',
    menuTableCells: 'Cells…',
    menuTableDelete: 'Delete',
    menuTableDeleteTable: 'Table',
    menuTableDeleteColumns: 'Columns',
    menuTableDeleteRows: 'Rows',
    menuTableSelect: 'Select',
    menuTableSelectCell: 'Cell',
    menuTableSelectColumn: 'Column',
    menuTableSelectRow: 'Row',
    menuTableSelectTable: 'Table',
    menuTableMergeCells: 'Merge Cells',
    menuTableSplitCells: 'Split Cells…',
    menuTableSplitTable: 'Split Table',
    menuTableAutoFit: 'AutoFit and Distribute',
    menuTableAutoFitContents: 'AutoFit to Contents',
    menuTableAutoFitWindow: 'AutoFit to Window',
    menuTableFixedWidth: 'Fixed Column Width',
    menuTableDistributeRows: 'Distribute Rows Evenly',
    menuTableDistributeColumns: 'Distribute Columns Evenly',
    menuTableRepeatHeader: 'Repeat Header Rows',
    menuTableConvert: 'Convert',
    menuTableConvertTextToTable: 'Convert Text to Table…',
    menuTableConvertTableToText: 'Convert Table to Text…',
    menuTableSort: 'Sort…',
    menuTableFormula: 'Formula…',
    menuTableGridlines: 'View Gridlines',
    menuTableProperties: 'Table Properties…',
    menuWordCount: 'Word Count…',
    menuAutoCorrect: 'AutoCorrect Options…',
    menuSpellingAndGrammar: 'Spelling and Grammar',
    menuSpellingDialog: 'Spelling and Grammar…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Hide Spelling Errors',
    menuHideGrammarErrors: 'Hide Grammar Errors',
    menuResetIgnoredWords: 'Reset Ignored Words and Grammar',
    menuThesaurus: 'Thesaurus…',
    menuHyphenation: 'Hyphenation…',
    menuTranslate: 'Translate',
    menuTranslateSelection: 'Translate Selection',
    menuTranslateDocument: 'Translate Document',
    menuSetTranslationLanguage: 'Set Document Translation Language…',
    menuLanguage: 'Language…',
    menuTrackChanges: 'Track Changes',
    menuHighlightChanges: 'Highlight Changes…',
    menuCompareDocuments: 'Compare Documents…',
    menuCombineDocuments: 'Combine Documents…',
    menuProtectDocument: 'Protect Document…',
    menuEnvelopes: 'Envelopes…',
    menuLabels: 'Labels…',
    menuCustomizeKeyboard: 'Customize Keyboard…',
    menuPreferences: 'Preferences…',
    menuAiProofread: 'AI Proofread',
    menuWindow: 'Window',
    menuHelp: 'Help',
    menuShortcuts: 'Keyboard Shortcuts',
  },
  vi: {
    dlgOpenDoc: 'Mở tài liệu',
    filterWord: 'Tài liệu Word',
    odtImportMsg: 'Chuyển đổi tài liệu .{ext} này?',
    odtImportDetail:
      'GenOffice chỉnh sửa .docx chứ không chỉnh sửa ODF. Nội dung, tiêu đề, danh sách, bảng và hình ảnh được chuyển đổi; chú thích cuối trang, bình luận và hộp văn bản không được giữ đầy đủ. Tệp gốc không bao giờ bị sửa; bản .docx sau chuyển đổi là một bản sao độc lập để tiếp tục chỉnh sửa và lưu.',
    hwpImportDetail:
      'GenOffice chỉnh sửa .docx chứ không chỉnh sửa HWP. Nội dung, bảng, hình ảnh, đầu trang và chân trang được chuyển đổi; công thức, chú thích cuối trang và đối tượng nhúng không được chuyển đổi, còn đánh số tự động trở thành văn bản thường. Tệp gốc không bao giờ bị sửa; bản .docx sau chuyển đổi là một bản sao độc lập để tiếp tục chỉnh sửa và lưu.',
    odtImportConvertBtn: 'Chuyển đổi và mở',
    importSkipConfirm: 'Lần sau chuyển đổi ngay, không hỏi lại',
    docImportMsg: 'Chuyển đổi tài liệu .{ext} này?',
    docImportDetail:
      'GenOffice chỉnh sửa .docx chứ không chỉnh sửa tài liệu Word 97-2003 (.doc). Nội dung, tiêu đề, danh sách, bảng và hình ảnh được chuyển đổi; chú thích cuối trang, bình luận và hộp văn bản không được giữ đầy đủ. Tệp gốc không bao giờ bị sửa; bản .docx sau chuyển đổi là một bản sao độc lập để tiếp tục chỉnh sửa và lưu.',
    docImportConvertBtn: 'Chuyển đổi và mở',
    dlgSaveAs: 'Lưu dưới dạng',
    closeUnsavedMsg: 'Tài liệu này có những thay đổi chưa được lưu.',
    closeUnsavedDetail: 'Bạn có muốn lưu các thay đổi trước khi đóng không?',
    closeNoReplyMsg: 'Tài liệu không phản hồi và có thể có những thay đổi chưa được lưu.',
    closeNoReplyDetail: 'Vẫn đóng? Các thay đổi chưa lưu sẽ bị mất.',
    btnCloseAnyway: 'Vẫn đóng',
    autosaveFoundTitle: 'Tìm thấy phiên bản khôi phục',
    autosaveDiskNewer:
      'Lưu ý: tệp trên đĩa đã được sửa đổi sau khi bản sao khôi phục này được lưu. Khôi phục sẽ thay thế những thay đổi mới hơn đó.',
    autosaveFoundBody:
      'Có những thay đổi chưa được lưu từ phiên làm việc trước của bạn. Khôi phục phiên bản đã lưu tự động?',
    autosaveRestore: 'Khôi phục',
    autosaveDiscard: 'Bỏ qua',
    btnDontSave: 'Không lưu',
    btnCancel: 'Hủy',
    extModifiedMsg: 'Tệp đã được sửa đổi bởi một chương trình khác.',
    extModifiedDetail: 'Vẫn lưu và ghi đè các thay đổi trên ổ đĩa?',
    btnOverwrite: 'Ghi đè',
    dlgInsertImage: 'Chèn hình ảnh',
    filterImages: 'Hình ảnh',
    dlgMergeDataSource: 'Chọn danh sách người nhận',
    filterMergeDataSource: 'Excel hoặc CSV',
    dlgAddAttachment: 'Thêm tệp đính kèm',
    filterSupported: 'Các tệp được hỗ trợ',
    filterAll: 'Tất cả các tệp',
    dlgExportPdf: 'Xuất dưới dạng PDF',
    dlgExportHtml: 'Xuất dưới dạng HTML',
    dlgExportMarkdown: 'Xuất dưới dạng Markdown',
    dlgPickExportDir: 'Chọn thư mục xuất',
    errUnsupportedExt: 'Tệp .{ext} không được hỗ trợ',
    errNotFile: 'không phải là tệp',
    errTooLarge: 'vượt quá giới hạn {mb}MB',
    errTemplateProtected: 'mẫu này được bảo vệ bằng mật khẩu; hãy mở bằng Tệp > Mở',
    errImageTooLarge: 'hình ảnh vượt quá giới hạn 5MB',
    errUnreadable: 'không thể đọc được',
    errFileTooLarge: 'Tệp vượt quá giới hạn kích thước',
    errParseFailed: 'Không thể phân tích tệp',
    errImageNoText:
      'Tệp đính kèm hình ảnh không có văn bản; hình ảnh được gửi cùng với tin nhắn của người dùng',
    errNotImage: 'loại hình ảnh không được hỗ trợ',
    errGskNotLoggedIn:
      'Chưa đăng nhập vào Genspark: nhấp vào “Đăng nhập vào Genspark” bên dưới, đăng nhập, sau đó thử lại',
    errNoApiKey: 'Chưa cấu hình khóa API cho {provider}',
    errAiRegion:
      'Nhà cung cấp của mô hình này không phục vụ khu vực hiện tại của bạn — hãy chọn mô hình khác hoặc đổi mạng',
    errAiBusy: 'Dịch vụ AI hiện đang bận — vui lòng thử lại sau giây lát',
    errNoModel: 'Chưa cấu hình tên mô hình',
    menuFile: 'Tệp',
    menuNewDoc: 'Tài liệu mới',
    menuNewWindow: 'Cửa sổ mới',
    menuOpen: 'Mở…',
    menuOpenRecent: 'Mở gần đây',
    menuNoRecent: 'Không có tài liệu gần đây',
    menuRecentMore: 'Xem thêm…',
    menuClose: 'Đóng',
    menuSave: 'Lưu',
    menuSaveAs: 'Lưu dưới dạng…',
    menuPageSetup: 'Thiết lập trang…',
    menuExportPdf: 'Xuất dưới dạng PDF…',
    menuExportHtml: 'Xuất dưới dạng HTML…',
    menuExportMarkdown: 'Xuất dưới dạng Markdown…',
    menuExportImages: 'Xuất dưới dạng hình ảnh…',
    menuPrint: 'In…',
    menuEdit: 'Chỉnh sửa',
    menuUndo: 'Hoàn tác',
    menuRedo: 'Làm lại',
    menuRepeat: 'Lặp lại',
    menuCantRepeat: 'Không thể lặp lại',
    menuCut: 'Cắt',
    menuCopy: 'Sao chép',
    menuPaste: 'Dán',
    menuPasteMatch: 'Dán và khớp định dạng',
    menuFindReplace: 'Tìm kiếm và thay thế…',
    menuPasteSpecial: 'Dán Đặc biệt…',
    menuClear: 'Xóa',
    menuClearFormatting: 'Xóa Định dạng',
    menuClearContents: 'Nội dung',
    menuFindMenu: 'Tìm',
    menuFind: 'Tìm…',
    menuFindNext: 'Tìm Tiếp',
    menuFindPrevious: 'Tìm Trước',
    menuReplace: 'Thay thế…',
    menuAdvancedFind: 'Tìm và Thay thế Nâng cao…',
    menuLinks: 'Liên kết…',
    menuSelectData: 'Chọn Dữ liệu…',
    menuToggleDrawing: 'Bật/tắt Vẽ',
    menuNewFromTemplate: 'Mới từ Mẫu…',
    menuSaveAsTemplate: 'Lưu dưới dạng Mẫu…',
    menuMove: 'Di chuyển…',
    menuRename: 'Đổi tên…',
    menuAlwaysOpenReadOnly: 'Luôn Mở ở chế độ Chỉ đọc',
    menuReduceFileSize: 'Giảm Kích cỡ Tệp…',
    menuProperties: 'Thuộc tính…',
    menuGoTo: 'Đi tới…',
    menuSelectAll: 'Chọn tất cả',
    menuView: 'Xem',
    menuZoom: 'Thu phóng',
    menuZoomIn: 'Phóng to',
    menuZoomOut: 'Thu nhỏ',
    menuZoom100: 'Kích thước thực tế (100%)',
    menuPrintLayout: 'Bố trí In',
    menuWebLayout: 'Bố trí Web',
    menuOutline: 'Dàn bài',
    menuDraft: 'Bản nháp',
    menuFocus: 'Tập trung',
    menuRibbon: 'Dải băng',
    menuSidebar: 'Thanh bên',
    menuThumbnail: 'Hình thu nhỏ',
    menuNavigation: 'Dẫn hướng',
    menuReviewing: 'Xem lại',
    menuSidebarSearch: 'Tìm kiếm',
    menuRuler: 'Thước',
    menuHeaderFooter: 'Đầu trang và Chân trang',
    menuFootnotes: 'Cước chú',
    menuMarkup: 'Đánh dấu',
    menuRevealFormatting: 'Hiển thị Định dạng',
    menuZoomDialog: 'Thu phóng…',
    menuGridlines: 'Đường lưới',
    menuPageWidth: 'Chiều rộng trang',
    menuWholePage: 'Toàn bộ trang',
    menuMultiPage: 'Nhiều trang',
    menuAiSidebar: 'Thanh bên AI',
    menuDarkMode: 'Chế độ tối',
    menuFullscreen: 'Vào chế độ toàn màn hình',
    menuInsert: 'Chèn',
    menuInsertTable: 'Bảng…',
    menuInsertImage: 'Hình ảnh…',
    menuInsertPageBreak: 'Ngắt trang',
    menuInsertLink: 'Siêu liên kết…',
    menuInsertEquation: 'Phương trình…',
    menuPictures: 'Ảnh',
    menuPictureFromFile: 'Ảnh từ Tệp…',
    menuShape: 'Hình dạng',
    menuShapeRectangle: 'Hình chữ nhật',
    menuShapeRoundedRectangle: 'Hình chữ nhật bo tròn',
    menuShapeTriangle: 'Hình tam giác',
    menuShapeOval: 'Hình bầu dục',
    menuShapeLine: 'Đường thẳng',
    menuChart: 'Biểu đồ',
    menuChartColumn: 'Cột',
    menuChartBar: 'Thanh',
    menuChartLine: 'Đường',
    menuChartArea: 'Vùng',
    menuChartPie: 'Tròn',
    menuChartTreemap: 'Sơ đồ cây',
    menuChartSunburst: 'Tia mặt trời',
    menuChartHistogram: 'Biểu đồ tần suất',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Hộp và Râu',
    menuChartScatter: 'X Y (Phân tán)',
    menuChartWaterfall: 'Thác nước',
    menuChartFunnel: 'Phễu',
    menuChartStock: 'Chứng khoán',
    menuChartSurface: 'Bề mặt',
    menuChartRadar: 'Radar',
    menuIcons: 'Biểu tượng…',
    menuBookmark: 'Thẻ đánh dấu…',
    menuCrossReference: 'Tham chiếu chéo…',
    menuBreak: 'Ngắt',
    menuColumnBreak: 'Ngắt cột',
    menuSectionNextPage: 'Ngắt phần (Trang tiếp theo)',
    menuSectionContinuous: 'Ngắt phần (Liên tục)',
    menuSectionOddPage: 'Ngắt phần (Trang lẻ)',
    menuSectionEvenPage: 'Ngắt phần (Trang chẵn)',
    menuFootnote: 'Cước chú…',
    menuCaption: 'Chú thích…',
    menuIndexAndTables: 'Chỉ mục và Bảng…',
    menuWatermark: 'Hình nền mờ…',
    menuPageNumbers: 'Số trang…',
    menuTextBox: 'Hộp văn bản',
    menuVerticalTextBox: 'Vẽ hộp văn bản dọc',
    menuAutoText: 'Văn bản tự động',
    menuAutoTextDialog: 'Văn bản tự động…',
    menuAutoTextNew: 'Mới…',
    menuDateAndTime: 'Ngày và Giờ…',
    menuObject: 'Đối tượng…',
    menuInsertFile: 'Tệp…',
    menuField: 'Trường…',
    menuAdvancedSymbol: 'Ký hiệu nâng cao…',
    menuComment: 'Bình luận',
    menuFormat: 'Định dạng',
    menuBold: 'In đậm',
    menuItalic: 'In nghiêng',
    menuUnderline: 'Gạch chân',
    menuAlign: 'Căn chỉnh',
    menuAlignLeft: 'Căn trái',
    menuAlignCenter: 'Căn giữa',
    menuAlignRight: 'Căn phải',
    menuAlignJustify: 'Căn đều',
    menuFont: 'Phông chữ…',
    menuParagraph: 'Đoạn văn…',
    menuTextEffects: 'Hiệu ứng Văn bản…',
    menuDocument: 'Tài liệu…',
    menuBulletsAndNumbering: 'Dấu đầu dòng và Đánh số…',
    menuBordersAndShading: 'Viền và Tô bóng…',
    menuColumns: 'Cột…',
    menuTabs: 'Tab…',
    menuDropCap: 'Chữ hoa đầu đoạn…',
    menuTextDirection: 'Hướng Văn bản…',
    menuChangeCase: 'Đổi Kiểu chữ…',
    menuAsianLayout: 'Bố trí Châu Á',
    menuPhoneticGuide: 'Hướng dẫn Ngữ âm…',
    menuEncloseCharacters: 'Bao Ký tự…',
    menuHorizontalInVertical: 'Ngang trong Dọc…',
    menuCombineCharacters: 'Kết hợp Ký tự…',
    menuTwoLinesInOne: 'Hai Dòng trong Một…',
    menuFitText: 'Vừa Văn bản…',
    menuStyle: 'Kiểu…',
    menuFontSubstitution: 'Thay thế Phông…',
    menuFormatObject: 'Định dạng Đối tượng…',
    menuEquationOptions: 'Tùy chọn Phương trình…',
    menuSelectionPane: 'Ngăn Chọn…',
    menuTools: 'Công cụ',
    menuTable: 'Bảng',
    menuTableInsert: 'Chèn',
    menuTableInsertTable: 'Bảng…',
    menuTableColsLeft: 'Chèn cột bên trái',
    menuTableColsRight: 'Chèn cột bên phải',
    menuTableRowsAbove: 'Chèn hàng phía trên',
    menuTableRowsBelow: 'Chèn hàng phía dưới',
    menuTableCells: 'Ô…',
    menuTableDelete: 'Xóa',
    menuTableDeleteTable: 'Bảng',
    menuTableDeleteColumns: 'Cột',
    menuTableDeleteRows: 'Hàng',
    menuTableSelect: 'Chọn',
    menuTableSelectCell: 'Ô',
    menuTableSelectColumn: 'Cột',
    menuTableSelectRow: 'Hàng',
    menuTableSelectTable: 'Bảng',
    menuTableMergeCells: 'Hợp nhất các ô',
    menuTableSplitCells: 'Tách ô…',
    menuTableSplitTable: 'Tách bảng',
    menuTableAutoFit: 'Tự động điều chỉnh',
    menuTableAutoFitContents: 'Tự động điều chỉnh theo nội dung',
    menuTableAutoFitWindow: 'Tự động điều chỉnh theo cửa sổ',
    menuTableFixedWidth: 'Chiều rộng cột cố định',
    menuTableDistributeRows: 'Phân bố hàng đều nhau',
    menuTableDistributeColumns: 'Phân bố cột đều nhau',
    menuTableRepeatHeader: 'Lặp lại hàng tiêu đề',
    menuTableConvert: 'Chuyển đổi',
    menuTableConvertTextToTable: 'Chuyển văn bản thành bảng…',
    menuTableConvertTableToText: 'Chuyển bảng thành văn bản…',
    menuTableSort: 'Sắp xếp…',
    menuTableFormula: 'Công thức…',
    menuTableGridlines: 'Xem đường lưới',
    menuTableProperties: 'Thuộc tính bảng…',
    menuWordCount: 'Đếm từ…',
    menuAutoCorrect: 'Tùy chọn tự sửa lỗi…',
    menuSpellingAndGrammar: 'Chính tả và Ngữ pháp',
    menuSpellingDialog: 'Chính tả và Ngữ pháp…',
    menuEditor: 'Trình soạn thảo',
    menuHideSpellingErrors: 'Ẩn Lỗi Chính tả',
    menuHideGrammarErrors: 'Ẩn Lỗi Ngữ pháp',
    menuResetIgnoredWords: 'Đặt lại Từ và Ngữ pháp đã Bỏ qua',
    menuThesaurus: 'Từ điển đồng nghĩa…',
    menuHyphenation: 'Gạch nối từ…',
    menuTranslate: 'Dịch',
    menuTranslateSelection: 'Dịch Vùng chọn',
    menuTranslateDocument: 'Dịch Tài liệu',
    menuSetTranslationLanguage: 'Đặt Ngôn ngữ Dịch Tài liệu…',
    menuLanguage: 'Ngôn ngữ…',
    menuTrackChanges: 'Theo dõi Thay đổi',
    menuHighlightChanges: 'Tô sáng Thay đổi…',
    menuCompareDocuments: 'So sánh Tài liệu…',
    menuCombineDocuments: 'Kết hợp Tài liệu…',
    menuProtectDocument: 'Bảo vệ Tài liệu…',
    menuEnvelopes: 'Phong bì…',
    menuLabels: 'Nhãn…',
    menuCustomizeKeyboard: 'Tùy chỉnh Bàn phím…',
    menuPreferences: 'Tùy chọn…',
    menuAiProofread: 'Hiệu đính bằng AI',
    menuWindow: 'Cửa sổ',
    menuHelp: 'Trợ giúp',
    menuShortcuts: 'Phím tắt bàn phím',
  },
  tr: {
    dlgOpenDoc: 'Belge Aç',
    filterWord: 'Word Belgeleri',
    odtImportMsg: 'Bu .{ext} belgesi dönüştürülsün mü?',
    odtImportDetail:
      'GenOffice ODF değil .docx düzenler. Gövde metni, başlıklar, listeler, tablolar ve resimler aktarılır; dipnotlar, açıklamalar ve metin kutuları tam olarak korunmaz. Özgün dosya hiçbir zaman değiştirilmez; dönüştürülen .docx, sonraki kaydetmelerin yazıldığı ayrı bir kopyadır.',
    hwpImportDetail:
      'GenOffice HWP değil .docx düzenler. Gövde metni, tablolar, resimler, üst bilgiler ve alt bilgiler aktarılır; denklemler, dipnotlar ve gömülü nesneler dönüştürülmez, otomatik numaralandırma düz metne dönüşür. Özgün dosya hiçbir zaman değiştirilmez; dönüştürülen .docx, sonraki kaydetmelerin yazıldığı ayrı bir kopyadır.',
    odtImportConvertBtn: 'Dönüştür ve aç',
    importSkipConfirm: 'Bir dahaki sefere sormadan dönüştür',
    docImportMsg: 'Bu .{ext} belgesi dönüştürülsün mü?',
    docImportDetail:
      'GenOffice Word 97-2003 (.doc) belgelerini değil .docx düzenler. Gövde metni, başlıklar, listeler, tablolar ve resimler aktarılır; dipnotlar, açıklamalar ve metin kutuları tam olarak korunmaz. Özgün dosya hiçbir zaman değiştirilmez; dönüştürülen .docx, sonraki kaydetmelerin yazıldığı ayrı bir kopyadır.',
    docImportConvertBtn: 'Dönüştür ve aç',
    dlgSaveAs: 'Farklı Kaydet',
    closeUnsavedMsg: 'Bu belgede kaydedilmemiş değişiklikler var.',
    closeUnsavedDetail: 'Kapatmadan önce kaydetmek ister misiniz?',
    closeNoReplyMsg: 'Belge yanıt vermiyor ve kaydedilmemiş değişiklikler içeriyor olabilir.',
    closeNoReplyDetail: 'Yine de kapatılsın mı? Kaydedilmemiş değişiklikler kaybolacak.',
    btnCloseAnyway: 'Yine de Kapat',
    autosaveFoundTitle: 'Kurtarılan sürüm bulundu',
    autosaveDiskNewer:
      'Not: diskteki dosya bu kurtarma kopyası kaydedildikten sonra değiştirilmiş. Geri yüklemek bu daha yeni değişikliklerin yerini alır.',
    autosaveFoundBody:
      'Son oturumunuzdan kaydedilmemiş değişiklikler var. Otomatik kaydedilen sürüm geri yüklensin mi?',
    autosaveRestore: 'Geri Yükle',
    autosaveDiscard: 'At',
    btnDontSave: 'Kaydetme',
    btnCancel: 'İptal',
    extModifiedMsg: 'Dosya başka bir program tarafından değiştirildi.',
    extModifiedDetail: 'Yine de kaydedilip diskteki değişikliklerin üzerine yazılsın mı?',
    btnOverwrite: 'Üzerine Yaz',
    dlgInsertImage: 'Resim Ekle',
    filterImages: 'Resimler',
    dlgMergeDataSource: 'Alıcı Listesi Seç',
    filterMergeDataSource: 'Excel veya CSV',
    dlgAddAttachment: 'Ek Ekle',
    filterSupported: 'Desteklenen Dosyalar',
    filterAll: 'Tüm Dosyalar',
    dlgExportPdf: 'PDF Olarak Dışa Aktar',
    dlgExportHtml: 'HTML Olarak Dışa Aktar',
    dlgExportMarkdown: 'Markdown Olarak Dışa Aktar',
    dlgPickExportDir: 'Dışa Aktarma Klasörünü Seç',
    errUnsupportedExt: '.{ext} dosyaları desteklenmiyor',
    errNotFile: 'bir dosya değil',
    errTooLarge: '{mb}MB sınırını aşıyor',
    errTemplateProtected: 'bu şablon parola korumalı; Dosya > Aç ile açın',
    errImageTooLarge: 'resim 5MB sınırını aşıyor',
    errUnreadable: 'okunamıyor',
    errFileTooLarge: 'Dosya boyut sınırını aşıyor',
    errParseFailed: 'Dosya ayrıştırılamadı',
    errImageNoText: 'Resim eklerinde metin yok; resim kullanıcı iletisiyle birlikte gönderilir',
    errNotImage: 'desteklenen bir resim türü değil',
    errGskNotLoggedIn:
      "Genspark'a giriş yapılmadı: aşağıdaki “Genspark'a giriş yap” düğmesine tıklayın, giriş yapın ve yeniden deneyin",
    errNoApiKey: '{provider} için API anahtarı yapılandırılmamış',
    errAiRegion:
      'Bu modelin sağlayıcısı bulunduğunuz bölgeye hizmet vermiyor — başka bir model seçin veya ağınızı değiştirin',
    errAiBusy: 'AI hizmeti şu anda yoğun — lütfen biraz sonra yeniden deneyin',
    errNoModel: 'Model adı yapılandırılmamış',
    menuFile: 'Dosya',
    menuNewDoc: 'Yeni Belge',
    menuNewWindow: 'Yeni Pencere',
    menuOpen: 'Aç…',
    menuOpenRecent: 'Son Kullanılanları Aç',
    menuNoRecent: 'Son Kullanılan Belge Yok',
    menuRecentMore: 'Diğer…',
    menuClose: 'Kapat',
    menuSave: 'Kaydet',
    menuSaveAs: 'Farklı Kaydet…',
    menuPageSetup: 'Sayfa Yapısı…',
    menuExportPdf: 'PDF Olarak Dışa Aktar…',
    menuExportHtml: 'HTML Olarak Dışa Aktar…',
    menuExportMarkdown: 'Markdown Olarak Dışa Aktar…',
    menuExportImages: 'Resim Olarak Dışa Aktar…',
    menuPrint: 'Yazdır…',
    menuEdit: 'Düzenle',
    menuUndo: 'Geri Al',
    menuRedo: 'Yinele',
    menuRepeat: 'Tekrarla',
    menuCantRepeat: 'Tekrarlanamıyor',
    menuCut: 'Kes',
    menuCopy: 'Kopyala',
    menuPaste: 'Yapıştır',
    menuPasteMatch: 'Yapıştır ve Biçimlendirmeyi Eşleştir',
    menuFindReplace: 'Bul ve Değiştir…',
    menuPasteSpecial: 'Özel Yapıştır…',
    menuClear: 'Temizle',
    menuClearFormatting: 'Biçimlendirmeyi Temizle',
    menuClearContents: 'İçindekiler',
    menuFindMenu: 'Bul',
    menuFind: 'Bul…',
    menuFindNext: 'Sonrakini Bul',
    menuFindPrevious: 'Öncekini Bul',
    menuReplace: 'Değiştir…',
    menuAdvancedFind: 'Gelişmiş Bul ve Değiştir…',
    menuLinks: 'Bağlantılar…',
    menuSelectData: 'Veri Seç…',
    menuToggleDrawing: 'Çizimi Aç/Kapat',
    menuNewFromTemplate: 'Şablondan Yeni…',
    menuSaveAsTemplate: 'Şablon Olarak Kaydet…',
    menuMove: 'Taşı…',
    menuRename: 'Yeniden Adlandır…',
    menuAlwaysOpenReadOnly: 'Her Zaman Salt Okunur Aç',
    menuReduceFileSize: 'Dosya Boyutunu Küçült…',
    menuProperties: 'Özellikler…',
    menuGoTo: 'Git…',
    menuSelectAll: 'Tümünü Seç',
    menuView: 'Görünüm',
    menuZoom: 'Yakınlaştır',
    menuZoomIn: 'Yakınlaştır',
    menuZoomOut: 'Uzaklaştır',
    menuZoom100: 'Gerçek Boyut (%100)',
    menuPrintLayout: 'Sayfa Düzeni',
    menuWebLayout: 'Web Düzeni',
    menuOutline: 'Anahat',
    menuDraft: 'Taslak',
    menuFocus: 'Odak',
    menuRibbon: 'Şerit',
    menuSidebar: 'Kenar Çubuğu',
    menuThumbnail: 'Küçük Resim',
    menuNavigation: 'Gezinti',
    menuReviewing: 'Gözden Geçirme',
    menuSidebarSearch: 'Ara',
    menuRuler: 'Cetvel',
    menuHeaderFooter: 'Üst Bilgi ve Alt Bilgi',
    menuFootnotes: 'Dipnotlar',
    menuMarkup: 'İşaretleme',
    menuRevealFormatting: 'Biçimlendirmeyi Göster',
    menuZoomDialog: 'Yakınlaştır…',
    menuGridlines: 'Kılavuz Çizgileri',
    menuPageWidth: 'Sayfa Genişliği',
    menuWholePage: 'Tüm Sayfa',
    menuMultiPage: 'Birden Çok Sayfa',
    menuAiSidebar: 'AI Kenar Çubuğu',
    menuDarkMode: 'Koyu Mod',
    menuFullscreen: 'Tam Ekrana Geç',
    menuInsert: 'Ekle',
    menuInsertTable: 'Tablo…',
    menuInsertImage: 'Resim…',
    menuInsertPageBreak: 'Sayfa Sonu',
    menuInsertLink: 'Köprü…',
    menuInsertEquation: 'Denklem…',
    menuPictures: 'Resimler',
    menuPictureFromFile: 'Dosyadan Resim…',
    menuShape: 'Şekil',
    menuShapeRectangle: 'Dikdörtgen',
    menuShapeRoundedRectangle: 'Yuvarlatılmış Dikdörtgen',
    menuShapeTriangle: 'Üçgen',
    menuShapeOval: 'Oval',
    menuShapeLine: 'Çizgi',
    menuChart: 'Grafik',
    menuChartColumn: 'Sütun',
    menuChartBar: 'Çubuk',
    menuChartLine: 'Çizgi',
    menuChartArea: 'Alan',
    menuChartPie: 'Pasta',
    menuChartTreemap: 'Ağaç Haritası',
    menuChartSunburst: 'Güneş Işığı',
    menuChartHistogram: 'Histogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Kutu ve Yatay Çizgi',
    menuChartScatter: 'X Y (Dağılım)',
    menuChartWaterfall: 'Şelale',
    menuChartFunnel: 'Huni',
    menuChartStock: 'Hisse Senedi',
    menuChartSurface: 'Yüzey',
    menuChartRadar: 'Radar',
    menuIcons: 'Simgeler…',
    menuBookmark: 'Yer İşareti…',
    menuCrossReference: 'Çapraz Başvuru…',
    menuBreak: 'Kesme',
    menuColumnBreak: 'Sütun Sonu',
    menuSectionNextPage: 'Bölüm Sonu (Sonraki Sayfa)',
    menuSectionContinuous: 'Bölüm Sonu (Sürekli)',
    menuSectionOddPage: 'Bölüm Sonu (Tek Sayfa)',
    menuSectionEvenPage: 'Bölüm Sonu (Çift Sayfa)',
    menuFootnote: 'Dipnot…',
    menuCaption: 'Resim Yazısı…',
    menuIndexAndTables: 'Dizin ve Tablolar…',
    menuWatermark: 'Filigran…',
    menuPageNumbers: 'Sayfa Numaraları…',
    menuTextBox: 'Metin Kutusu',
    menuVerticalTextBox: 'Dikey Metin Kutusu Çiz',
    menuAutoText: 'Otomatik Metin',
    menuAutoTextDialog: 'Otomatik Metin…',
    menuAutoTextNew: 'Yeni…',
    menuDateAndTime: 'Tarih ve Saat…',
    menuObject: 'Nesne…',
    menuInsertFile: 'Dosya…',
    menuField: 'Alan…',
    menuAdvancedSymbol: 'Gelişmiş Simge…',
    menuComment: 'Açıklama',
    menuFormat: 'Biçim',
    menuBold: 'Kalın',
    menuItalic: 'İtalik',
    menuUnderline: 'Altı Çizili',
    menuAlign: 'Hizala',
    menuAlignLeft: 'Sola Hizala',
    menuAlignCenter: 'Ortala',
    menuAlignRight: 'Sağa Hizala',
    menuAlignJustify: 'İki Yana Yasla',
    menuFont: 'Yazı Tipi…',
    menuParagraph: 'Paragraf…',
    menuTextEffects: 'Metin Efektleri…',
    menuDocument: 'Belge…',
    menuBulletsAndNumbering: 'Madde İşaretleri ve Numaralandırma…',
    menuBordersAndShading: 'Kenarlıklar ve Gölgelendirme…',
    menuColumns: 'Sütunlar…',
    menuTabs: 'Sekmeler…',
    menuDropCap: 'Büyük İlk Harf…',
    menuTextDirection: 'Metin Yönü…',
    menuChangeCase: 'Büyük/Küçük Harf Değiştir…',
    menuAsianLayout: 'Asya Düzeni',
    menuPhoneticGuide: 'Fonetik Kılavuz…',
    menuEncloseCharacters: 'Karakterleri Çevrele…',
    menuHorizontalInVertical: 'Dikeyde Yatay…',
    menuCombineCharacters: 'Karakterleri Birleştir…',
    menuTwoLinesInOne: 'İki Satırı Bire Sığdır…',
    menuFitText: 'Metni Sığdır…',
    menuStyle: 'Stil…',
    menuFontSubstitution: 'Yazı Tipi Değiştirme…',
    menuFormatObject: 'Nesneyi Biçimlendir…',
    menuEquationOptions: 'Denklem Seçenekleri…',
    menuSelectionPane: 'Seçim Bölmesi…',
    menuTools: 'Araçlar',
    menuTable: 'Tablo',
    menuTableInsert: 'Ekle',
    menuTableInsertTable: 'Tablo…',
    menuTableColsLeft: 'Sola Sütun',
    menuTableColsRight: 'Sağa Sütun',
    menuTableRowsAbove: 'Üste Satır',
    menuTableRowsBelow: 'Alta Satır',
    menuTableCells: 'Hücreler…',
    menuTableDelete: 'Sil',
    menuTableDeleteTable: 'Tablo',
    menuTableDeleteColumns: 'Sütunlar',
    menuTableDeleteRows: 'Satırlar',
    menuTableSelect: 'Seç',
    menuTableSelectCell: 'Hücre',
    menuTableSelectColumn: 'Sütun',
    menuTableSelectRow: 'Satır',
    menuTableSelectTable: 'Tablo',
    menuTableMergeCells: 'Hücreleri Birleştir',
    menuTableSplitCells: 'Hücreleri Böl…',
    menuTableSplitTable: 'Tabloyu Böl',
    menuTableAutoFit: 'Otomatik Sığdır ve Dağıt',
    menuTableAutoFitContents: 'İçeriğe Otomatik Sığdır',
    menuTableAutoFitWindow: 'Pencereye Otomatik Sığdır',
    menuTableFixedWidth: 'Sabit Sütun Genişliği',
    menuTableDistributeRows: 'Satırları Eşit Dağıt',
    menuTableDistributeColumns: 'Sütunları Eşit Dağıt',
    menuTableRepeatHeader: 'Başlık Satırlarını Yinele',
    menuTableConvert: 'Dönüştür',
    menuTableConvertTextToTable: 'Metni Tabloya Dönüştür…',
    menuTableConvertTableToText: 'Tabloyu Metne Dönüştür…',
    menuTableSort: 'Sırala…',
    menuTableFormula: 'Formül…',
    menuTableGridlines: 'Kılavuz Çizgilerini Göster',
    menuTableProperties: 'Tablo Özellikleri…',
    menuWordCount: 'Sözcük Sayısı…',
    menuAutoCorrect: 'Otomatik Düzelt Seçenekleri…',
    menuSpellingAndGrammar: 'Yazım ve Dil Bilgisi',
    menuSpellingDialog: 'Yazım ve Dil Bilgisi…',
    menuEditor: 'Düzenleyici',
    menuHideSpellingErrors: 'Yazım Hatalarını Gizle',
    menuHideGrammarErrors: 'Dil Bilgisi Hatalarını Gizle',
    menuResetIgnoredWords: 'Yoksayılan Sözcükleri ve Dil Bilgisini Sıfırla',
    menuThesaurus: 'Eş Anlamlılar…',
    menuHyphenation: 'Heceleme…',
    menuTranslate: 'Çevir',
    menuTranslateSelection: 'Seçimi Çevir',
    menuTranslateDocument: 'Belgeyi Çevir',
    menuSetTranslationLanguage: 'Belge Çeviri Dilini Ayarla…',
    menuLanguage: 'Dil…',
    menuTrackChanges: 'Değişiklikleri İzle',
    menuHighlightChanges: 'Değişiklikleri Vurgula…',
    menuCompareDocuments: 'Belgeleri Karşılaştır…',
    menuCombineDocuments: 'Belgeleri Birleştir…',
    menuProtectDocument: 'Belgeyi Koru…',
    menuEnvelopes: 'Zarflar…',
    menuLabels: 'Etiketler…',
    menuCustomizeKeyboard: 'Klavyeyi Özelleştir…',
    menuPreferences: 'Tercihler…',
    menuAiProofread: 'AI Düzeltme Okuması',
    menuWindow: 'Pencere',
    menuHelp: 'Yardım',
    menuShortcuts: 'Klavye Kısayolları',
  },
  hu: {
    dlgOpenDoc: 'Dokumentum megnyitása',
    filterWord: 'Word-dokumentumok',
    odtImportMsg: 'Átalakítsuk ezt a(z) .{ext} dokumentumot?',
    odtImportDetail:
      'A GenOffice .docx fájlokat szerkeszt, ODF-et nem. A szövegtörzs, a címsorok, a listák, a táblázatok és a képek átmennek; a lábjegyzetek, a megjegyzések és a szövegdobozok nem maradnak meg teljesen. Az eredeti fájl sosem változik; az átalakított .docx egy külön példány, amelybe a későbbi mentések kerülnek.',
    hwpImportDetail:
      'A GenOffice .docx fájlokat szerkeszt, HWP-t nem. A szövegtörzs, a táblázatok, a képek, az élőfejek és az élőlábak átmennek; az egyenletek, a lábjegyzetek és a beágyazott objektumok nem alakulnak át, az automatikus számozás pedig egyszerű szöveggé válik. Az eredeti fájl sosem változik; az átalakított .docx egy külön példány, amelybe a későbbi mentések kerülnek.',
    odtImportConvertBtn: 'Átalakítás és megnyitás',
    importSkipConfirm: 'Átalakítás a továbbiakban kérdés nélkül',
    docImportMsg: 'Átalakítsuk ezt a(z) .{ext} dokumentumot?',
    docImportDetail:
      'A GenOffice .docx fájlokat szerkeszt, Word 97-2003 dokumentumokat nem. A szövegtörzs, a címsorok, a listák, a táblázatok és a képek átmennek; a lábjegyzetek, a megjegyzések és a szövegdobozok nem maradnak meg teljesen. Az eredeti fájl sosem változik; az átalakított .docx egy külön példány, amelybe a későbbi mentések kerülnek.',
    docImportConvertBtn: 'Átalakítás és megnyitás',
    dlgSaveAs: 'Mentés másként',
    closeUnsavedMsg: 'A dokumentum nem mentett módosításokat tartalmaz.',
    closeUnsavedDetail: 'Mented a módosításokat bezárás előtt?',
    closeNoReplyMsg:
      'A dokumentum nem válaszol, és lehet, hogy nem mentett módosításokat tartalmaz.',
    closeNoReplyDetail: 'Mindenképp bezárod? A nem mentett módosítások elvesznek.',
    btnCloseAnyway: 'Bezárás mindenképp',
    autosaveFoundTitle: 'Helyreállított verzió található',
    autosaveDiskNewer:
      'Megjegyzés: a lemezen lévő fájlt a helyreállítási másolat mentése után módosították. A visszaállítás felülírja ezeket az újabb módosításokat.',
    autosaveFoundBody:
      'Nem mentett módosítások maradtak az előző munkamenetből. Visszaállítod az automatikusan mentett verziót?',
    autosaveRestore: 'Visszaállítás',
    autosaveDiscard: 'Elvetés',
    btnDontSave: 'Ne mentsd',
    btnCancel: 'Mégse',
    extModifiedMsg: 'A fájlt egy másik program módosította.',
    extModifiedDetail: 'Mindenképp mented, és felülírod a lemezen lévő módosításokat?',
    btnOverwrite: 'Felülírás',
    dlgInsertImage: 'Kép beszúrása',
    filterImages: 'Képek',
    dlgMergeDataSource: 'Címzettlista kiválasztása',
    filterMergeDataSource: 'Excel vagy CSV',
    dlgAddAttachment: 'Mellékletek hozzáadása',
    filterSupported: 'Támogatott fájlok',
    filterAll: 'Minden fájl',
    dlgExportPdf: 'Exportálás PDF-ként',
    dlgExportHtml: 'Exportálás HTML-ként',
    dlgExportMarkdown: 'Exportálás Markdown-ként',
    dlgPickExportDir: 'Exportmappa kiválasztása',
    errUnsupportedExt: 'A .{ext} fájlok nem támogatottak',
    errNotFile: 'nem fájl',
    errTooLarge: 'túllépi a(z) {mb}MB-os korlátot',
    errTemplateProtected: 'ez a sablon jelszóval védett; nyissa meg a Fájl > Megnyitás paranccsal',
    errImageTooLarge: 'a kép túllépi az 5MB-os korlátot',
    errUnreadable: 'nem olvasható',
    errFileTooLarge: 'A fájl túllépi a méretkorlátot',
    errParseFailed: 'Nem sikerült értelmezni a fájlt',
    errImageNoText:
      'A képmellékletek nem tartalmaznak szöveget; a kép a felhasználó üzenetével együtt lesz elküldve',
    errNotImage: 'nem támogatott képtípus',
    errGskNotLoggedIn:
      'Nem vagy bejelentkezve a Gensparkba: kattints alább a „Bejelentkezés a Gensparkba” gombra, jelentkezz be, majd próbáld újra',
    errNoApiKey: 'Nincs API-kulcs beállítva a következőhöz: {provider}',
    errAiRegion:
      'A modell szolgáltatója nem érhető el a jelenlegi régiódban — válassz másik modellt, vagy válts hálózatot',
    errAiBusy: 'Az AI-szolgáltatás jelenleg elfoglalt — próbáld újra egy pillanat múlva',
    errNoModel: 'Nincs modellnév megadva',
    menuFile: 'Fájl',
    menuNewDoc: 'Új dokumentum',
    menuNewWindow: 'Új ablak',
    menuOpen: 'Megnyitás…',
    menuOpenRecent: 'Legutóbbiak megnyitása',
    menuNoRecent: 'Nincsenek legutóbbi dokumentumok',
    menuRecentMore: 'Továbbiak…',
    menuClose: 'Bezárás',
    menuSave: 'Mentés',
    menuSaveAs: 'Mentés másként…',
    menuPageSetup: 'Oldalbeállítás…',
    menuExportPdf: 'Exportálás PDF-ként…',
    menuExportHtml: 'Exportálás HTML-ként…',
    menuExportMarkdown: 'Exportálás Markdown-ként…',
    menuExportImages: 'Exportálás képekként…',
    menuPrint: 'Nyomtatás…',
    menuEdit: 'Szerkesztés',
    menuUndo: 'Visszavonás',
    menuRedo: 'Újra',
    menuRepeat: 'Ismétlés',
    menuCantRepeat: 'Nem ismételhető',
    menuCut: 'Kivágás',
    menuCopy: 'Másolás',
    menuPaste: 'Beillesztés',
    menuPasteMatch: 'Beillesztés a formátum egyeztetésével',
    menuFindReplace: 'Keresés és csere…',
    menuGoTo: 'Ugrás…',
    menuSelectAll: 'Az összes kijelölése',
    menuView: 'Nézet',
    menuZoom: 'Nagyítás',
    menuZoomIn: 'Nagyítás',
    menuZoomOut: 'Kicsinyítés',
    menuZoom100: 'Tényleges méret (100%)',
    menuPageWidth: 'Oldalszélesség',
    menuWholePage: 'Teljes oldal',
    menuMultiPage: 'Több oldal',
    menuAiSidebar: 'AI-oldalsáv',
    menuDarkMode: 'Sötét mód',
    menuFullscreen: 'Teljes képernyő',
    menuInsert: 'Beszúrás',
    menuInsertTable: 'Táblázat…',
    menuInsertImage: 'Kép…',
    menuInsertPageBreak: 'Oldaltörés',
    menuInsertLink: 'Hivatkozás…',
    menuInsertEquation: 'Egyenlet…',
    menuComment: 'Megjegyzés',
    menuFormat: 'Formázás',
    menuBold: 'Félkövér',
    menuItalic: 'Dőlt',
    menuUnderline: 'Aláhúzott',
    menuAlign: 'Igazítás',
    menuAlignLeft: 'Balra igazítás',
    menuAlignCenter: 'Középre igazítás',
    menuAlignRight: 'Jobbra igazítás',
    menuAlignJustify: 'Sorkizárt',
    menuFont: 'Betűtípus…',
    menuParagraph: 'Bekezdés…',
    menuTools: 'Eszközök',
    menuTable: 'Táblázat',
    menuTableInsert: 'Beszúrás',
    menuTableInsertTable: 'Táblázat…',
    menuTableColsLeft: 'Oszlopok balra',
    menuTableColsRight: 'Oszlopok jobbra',
    menuTableRowsAbove: 'Sorok fölé',
    menuTableRowsBelow: 'Sorok alá',
    menuTableCells: 'Cellák…',
    menuTableDelete: 'Törlés',
    menuTableDeleteTable: 'Táblázat',
    menuTableDeleteColumns: 'Oszlopok',
    menuTableDeleteRows: 'Sorok',
    menuTableSelect: 'Kijelölés',
    menuTableSelectCell: 'Cella',
    menuTableSelectColumn: 'Oszlop',
    menuTableSelectRow: 'Sor',
    menuTableSelectTable: 'Táblázat',
    menuTableMergeCells: 'Cellák egyesítése',
    menuTableSplitCells: 'Cellák felosztása…',
    menuTableSplitTable: 'Táblázat felosztása',
    menuTableAutoFit: 'Automatikus méretezés és elosztás',
    menuTableAutoFitContents: 'Automatikus méretezés a tartalomhoz',
    menuTableAutoFitWindow: 'Automatikus méretezés az ablakhoz',
    menuTableFixedWidth: 'Rögzített oszlopszélesség',
    menuTableDistributeRows: 'Sorok egyenletes elosztása',
    menuTableDistributeColumns: 'Oszlopok egyenletes elosztása',
    menuTableRepeatHeader: 'Fejlécsorok ismétlése',
    menuTableConvert: 'Konvertálás',
    menuTableConvertTextToTable: 'Szöveg táblázattá…',
    menuTableConvertTableToText: 'Táblázat szöveggé…',
    menuTableSort: 'Rendezés…',
    menuTableFormula: 'Képlet…',
    menuTableGridlines: 'Rácsvonalak megjelenítése',
    menuTableProperties: 'Táblázat tulajdonságai…',
    menuWordCount: 'Szavak száma…',
    menuAutoCorrect: 'Automatikus javítás beállításai…',
    menuPreferences: 'Beállítások…',
    menuAiProofread: 'AI nyelvi ellenőrzés',
    menuWindow: 'Ablak',
    menuHelp: 'Súgó',
    menuShortcuts: 'Billentyűparancsok',
    menuPasteSpecial: 'Irányított beillesztés…',
    menuClear: 'Törlés',
    menuClearFormatting: 'Formázás törlése',
    menuClearContents: 'Tartalom',
    menuFindMenu: 'Keresés',
    menuFind: 'Keresés…',
    menuFindNext: 'Következő keresése',
    menuFindPrevious: 'Előző keresése',
    menuReplace: 'Csere…',
    menuAdvancedFind: 'Speciális keresés és csere…',
    menuLinks: 'Csatolások…',
    menuSelectData: 'Adatok kijelölése…',
    menuToggleDrawing: 'Rajzolás be/ki',
    menuNewFromTemplate: 'Új sablonból…',
    menuSaveAsTemplate: 'Mentés sablonként…',
    menuMove: 'Áthelyezés…',
    menuRename: 'Átnevezés…',
    menuAlwaysOpenReadOnly: 'Megnyitás mindig írásvédettként',
    menuReduceFileSize: 'Fájlméret csökkentése…',
    menuProperties: 'Tulajdonságok…',
    menuPrintLayout: 'Nyomtatási elrendezés',
    menuWebLayout: 'Webes elrendezés',
    menuOutline: 'Vázlat',
    menuDraft: 'Piszkozat',
    menuFocus: 'Fókusz',
    menuRibbon: 'Menüszalag',
    menuSidebar: 'Oldalsáv',
    menuThumbnail: 'Miniatűrök',
    menuNavigation: 'Navigáció',
    menuReviewing: 'Véleményezés',
    menuSidebarSearch: 'Keresés',
    menuRuler: 'Vonalzó',
    menuHeaderFooter: 'Élőfej és élőláb',
    menuFootnotes: 'Lábjegyzetek',
    menuMarkup: 'Korrektúra',
    menuRevealFormatting: 'Formázás megjelenítése',
    menuZoomDialog: 'Nagyítás…',
    menuGridlines: 'Rácsvonalak',
    menuPictures: 'Képek',
    menuPictureFromFile: 'Kép fájlból…',
    menuShape: 'Alakzat',
    menuShapeRectangle: 'Téglalap',
    menuShapeRoundedRectangle: 'Lekerekített téglalap',
    menuShapeTriangle: 'Háromszög',
    menuShapeOval: 'Ellipszis',
    menuShapeLine: 'Vonal',
    menuChart: 'Diagram',
    menuChartColumn: 'Oszlop',
    menuChartBar: 'Sáv',
    menuChartLine: 'Vonal',
    menuChartArea: 'Terület',
    menuChartPie: 'Kör',
    menuChartTreemap: 'Fatérkép',
    menuChartSunburst: 'Többszintű gyűrű',
    menuChartHistogram: 'Hisztogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Doboz',
    menuChartScatter: 'Pont (XY)',
    menuChartWaterfall: 'Vízesés',
    menuChartFunnel: 'Tölcsér',
    menuChartStock: 'Árfolyam',
    menuChartSurface: 'Felület',
    menuChartRadar: 'Sugár',
    menuIcons: 'Ikonok…',
    menuBookmark: 'Könyvjelző…',
    menuCrossReference: 'Kereszthivatkozás…',
    menuBreak: 'Törés',
    menuColumnBreak: 'Hasábtörés',
    menuSectionNextPage: 'Szakasztörés (új oldal)',
    menuSectionContinuous: 'Szakasztörés (folyamatos)',
    menuSectionOddPage: 'Szakasztörés (páratlan oldal)',
    menuSectionEvenPage: 'Szakasztörés (páros oldal)',
    menuFootnote: 'Lábjegyzet…',
    menuCaption: 'Képaláírás…',
    menuIndexAndTables: 'Tárgymutató és jegyzékek…',
    menuWatermark: 'Vízjel…',
    menuPageNumbers: 'Oldalszámok…',
    menuTextBox: 'Szövegdoboz',
    menuVerticalTextBox: 'Függőleges szövegdoboz rajzolása',
    menuAutoText: 'Kész szöveg',
    menuAutoTextDialog: 'Kész szöveg…',
    menuAutoTextNew: 'Új…',
    menuDateAndTime: 'Dátum és idő…',
    menuObject: 'Objektum…',
    menuInsertFile: 'Fájl…',
    menuField: 'Mező…',
    menuAdvancedSymbol: 'Speciális szimbólum…',
    menuTextEffects: 'Szövegeffektusok…',
    menuDocument: 'Dokumentum…',
    menuBulletsAndNumbering: 'Felsorolás és számozás…',
    menuBordersAndShading: 'Szegély és mintázat…',
    menuColumns: 'Hasábok…',
    menuTabs: 'Tabulátorok…',
    menuDropCap: 'Iniciálé…',
    menuTextDirection: 'Szövegirány…',
    menuChangeCase: 'Kisbetű-nagybetű…',
    menuAsianLayout: 'Ázsiai elrendezés',
    menuPhoneticGuide: 'Fonetikus útmutató…',
    menuEncloseCharacters: 'Bekeretezett karakterek…',
    menuHorizontalInVertical: 'Vízszintes a függőlegesben…',
    menuCombineCharacters: 'Karakterek kombinálása…',
    menuTwoLinesInOne: 'Két sor egyben…',
    menuFitText: 'Szöveg illesztése…',
    menuStyle: 'Stílus…',
    menuFontSubstitution: 'Betűtípus-helyettesítés…',
    menuFormatObject: 'Objektum formázása…',
    menuEquationOptions: 'Egyenlet beállításai…',
    menuSelectionPane: 'Kijelöléspanel…',
    menuSpellingAndGrammar: 'Helyesírás és nyelvhelyesség',
    menuSpellingDialog: 'Helyesírás és nyelvhelyesség…',
    menuEditor: 'Szerkesztő',
    menuHideSpellingErrors: 'Helyesírási hibák elrejtése',
    menuHideGrammarErrors: 'Nyelvhelyességi hibák elrejtése',
    menuResetIgnoredWords: 'Mellőzött szavak és nyelvhelyességi hibák visszaállítása',
    menuThesaurus: 'Szinonimaszótár…',
    menuHyphenation: 'Elválasztás…',
    menuTranslate: 'Fordítás',
    menuTranslateSelection: 'Kijelölés fordítása',
    menuTranslateDocument: 'Dokumentum fordítása',
    menuSetTranslationLanguage: 'A dokumentum fordítási nyelvének beállítása…',
    menuLanguage: 'Nyelv…',
    menuTrackChanges: 'Változások követése',
    menuHighlightChanges: 'Változások kiemelése…',
    menuCompareDocuments: 'Dokumentumok összehasonlítása…',
    menuCombineDocuments: 'Dokumentumok egyesítése…',
    menuProtectDocument: 'Dokumentum védelme…',
    menuEnvelopes: 'Borítékok…',
    menuLabels: 'Címkék…',
    menuCustomizeKeyboard: 'Billentyűzet testreszabása…',
  },
  ja: {
    dlgOpenDoc: '文書を開く',
    filterWord: 'Word 文書',
    odtImportMsg: 'この .{ext} 文書を変換しますか？',
    odtImportDetail:
      'GenOffice は ODF ではなく .docx を編集します。本文・見出し・リスト・表・画像は変換されますが、脚注・コメント・テキストボックスは完全には保持されません。元のファイルは変更されず、変換後の .docx は独立したコピーとして以後の保存先になります。',
    hwpImportDetail:
      'GenOffice は HWP ではなく .docx を編集します。本文・表・画像・ヘッダー・フッターは変換されますが、数式・脚注・埋め込みオブジェクトは変換されず、自動番号は通常のテキストになります。元のファイルは変更されず、変換後の .docx は独立したコピーとして以後の保存先になります。',
    odtImportConvertBtn: '変換して開く',
    importSkipConfirm: '次回から確認せずに変換する',
    docImportMsg: 'この .{ext} 文書を変換しますか？',
    docImportDetail:
      'GenOffice は Word 97-2003 形式（.doc）ではなく .docx を編集します。本文・見出し・リスト・表・画像は変換されますが、脚注・コメント・テキストボックスは完全には保持されません。元のファイルは変更されず、変換後の .docx は独立したコピーとして以後の保存先になります。',
    docImportConvertBtn: '変換して開く',
    dlgSaveAs: '名前を付けて保存',
    closeUnsavedMsg: 'このドキュメントに未保存の変更があります。',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    closeNoReplyMsg: 'ドキュメントが応答していません。未保存の変更がある可能性があります。',
    closeNoReplyDetail: 'それでも閉じますか?未保存の変更は失われます。',
    btnCloseAnyway: '閉じる',
    autosaveFoundTitle: '自動回復バージョンがあります',
    autosaveDiskNewer:
      '注意: この回復コピーの保存後にディスク上のファイルが変更されています。復元するとその新しい変更は置き換えられます。',
    autosaveFoundBody: '前回のセッションに未保存の変更があります。自動保存版を復元しますか?',
    autosaveRestore: '復元',
    autosaveDiscard: '破棄',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
    extModifiedMsg: 'このファイルは別のプログラムによって変更されています。',
    extModifiedDetail: 'このまま保存してディスク上の変更を上書きしますか?',
    btnOverwrite: '上書き',
    dlgInsertImage: '画像の挿入',
    filterImages: '画像',
    dlgMergeDataSource: '受信者リストを選択',
    filterMergeDataSource: 'Excel または CSV',
    dlgAddAttachment: '添付ファイルの追加',
    filterSupported: 'サポートされているファイル',
    filterAll: 'すべてのファイル',
    dlgExportPdf: 'PDF としてエクスポート',
    dlgExportHtml: 'HTML としてエクスポート',
    dlgExportMarkdown: 'Markdown としてエクスポート',
    dlgPickExportDir: 'エクスポート先フォルダーの選択',
    errUnsupportedExt: '.{ext} 形式には対応していません',
    errNotFile: 'ファイルではありません',
    errTooLarge: '{mb}MB の上限を超えています',
    errTemplateProtected:
      'このテンプレートはパスワードで保護されています。[ファイル] > [開く] で開いてください',
    errImageTooLarge: '画像が 5MB の上限を超えています',
    errUnreadable: '読み取れません',
    errFileTooLarge: 'ファイルがサイズ上限を超えています',
    errParseFailed: 'ファイルの解析に失敗しました',
    errImageNoText:
      '画像の添付ファイルはテキストを提供しません。画像としてユーザーメッセージと一緒に送信されるため、そのまま画像をご確認ください',
    errNotImage: 'サポートされていない画像形式です',
    errGskNotLoggedIn:
      'Genspark にサインインしていません。下の「Genspark にサインイン」からサインインして再試行してください',
    errNoApiKey: '{provider} の API キーが設定されていません',
    errAiRegion:
      'このモデルの提供元は現在の地域ではサービスを提供していません。別のモデルを選ぶか、ネットワークを切り替えてください',
    errAiBusy: 'AI サービスが混み合っています。しばらくしてからもう一度お試しください',
    errNoModel: 'モデル名が設定されていません',
    menuFile: 'ファイル',
    menuNewDoc: '新規文書',
    menuNewWindow: '新規ウィンドウ',
    menuOpen: '開く…',
    menuOpenRecent: '最近使った文書を開く',
    menuNoRecent: '最近使った文書はありません',
    menuRecentMore: 'その他…',
    menuClose: '閉じる',
    menuSave: '保存',
    menuSaveAs: '名前を付けて保存…',
    menuPageSetup: 'ページ設定…',
    menuExportPdf: 'PDF としてエクスポート…',
    menuExportHtml: 'HTML としてエクスポート…',
    menuExportMarkdown: 'Markdown としてエクスポート…',
    menuExportImages: '画像としてエクスポート…',
    menuPrint: '印刷…',
    menuEdit: '編集',
    menuUndo: '元に戻す',
    menuRedo: 'やり直す',
    menuRepeat: '繰り返し',
    menuCantRepeat: '繰り返しできません',
    menuCut: '切り取り',
    menuCopy: 'コピー',
    menuPaste: '貼り付け',
    menuPasteMatch: '貼り付けて書式を合わせる',
    menuFindReplace: '検索と置換…',
    menuPasteSpecial: '形式を選択してペースト…',
    menuClear: 'クリア',
    menuClearFormatting: '書式のクリア',
    menuClearContents: '内容',
    menuFindMenu: '検索',
    menuFind: '検索…',
    menuFindNext: '次を検索',
    menuFindPrevious: '前を検索',
    menuReplace: '置換…',
    menuAdvancedFind: '高度な検索と置換…',
    menuLinks: 'リンクの設定…',
    menuSelectData: 'データの選択…',
    menuToggleDrawing: '描画の切り替え',
    menuNewFromTemplate: 'テンプレートから新規作成…',
    menuSaveAsTemplate: 'テンプレートとして保存…',
    menuMove: '移動…',
    menuRename: '名前の変更…',
    menuAlwaysOpenReadOnly: '常に読み取り専用で開く',
    menuReduceFileSize: 'ファイル サイズの圧縮…',
    menuProperties: 'プロパティ…',
    menuGoTo: 'ジャンプ…',
    menuSelectAll: 'すべて選択',
    menuView: '表示',
    menuZoom: 'ズーム',
    menuZoomIn: '拡大',
    menuZoomOut: '縮小',
    menuZoom100: '実際のサイズ (100%)',
    menuPrintLayout: '印刷レイアウト',
    menuWebLayout: 'Web レイアウト',
    menuOutline: 'アウトライン',
    menuDraft: '下書き',
    menuFocus: 'フォーカス',
    menuRibbon: 'リボン',
    menuSidebar: 'サイドバー',
    menuThumbnail: 'サムネイル',
    menuNavigation: 'ナビゲーション',
    menuReviewing: '校閲',
    menuSidebarSearch: '検索',
    menuRuler: 'ルーラー',
    menuHeaderFooter: 'ヘッダーとフッター',
    menuFootnotes: '脚注',
    menuMarkup: '変更履歴/コメント',
    menuRevealFormatting: '書式の詳細',
    menuZoomDialog: 'ズーム…',
    menuGridlines: 'グリッド線',
    menuPageWidth: 'ページ幅',
    menuWholePage: 'ページ全体',
    menuMultiPage: '複数ページ',
    menuAiSidebar: 'AI サイドバー',
    menuDarkMode: 'ダークモード',
    menuFullscreen: 'フルスクリーンにする',
    menuInsert: '挿入',
    menuInsertTable: '表…',
    menuInsertImage: '画像…',
    menuInsertPageBreak: '改ページ',
    menuInsertLink: 'ハイパーリンク…',
    menuInsertEquation: '数式…',
    menuPictures: '画像',
    menuPictureFromFile: 'ファイルからの画像…',
    menuShape: '図形',
    menuShapeRectangle: '四角形',
    menuShapeRoundedRectangle: '角丸四角形',
    menuShapeTriangle: '三角形',
    menuShapeOval: '楕円',
    menuShapeLine: '直線',
    menuChart: 'グラフ',
    menuChartColumn: '縦棒',
    menuChartBar: '横棒',
    menuChartLine: '折れ線',
    menuChartArea: '面',
    menuChartPie: '円',
    menuChartTreemap: 'ツリーマップ',
    menuChartSunburst: 'サンバースト',
    menuChartHistogram: 'ヒストグラム',
    menuChartPareto: 'パレート図',
    menuChartBoxWhisker: '箱ひげ図',
    menuChartScatter: '散布図',
    menuChartWaterfall: 'ウォーターフォール',
    menuChartFunnel: 'じょうご',
    menuChartStock: '株価',
    menuChartSurface: '等高線',
    menuChartRadar: 'レーダー',
    menuIcons: 'アイコン…',
    menuBookmark: 'ブックマーク…',
    menuCrossReference: '相互参照…',
    menuBreak: '区切り',
    menuColumnBreak: '段区切り',
    menuSectionNextPage: 'セクション区切り (次のページから開始)',
    menuSectionContinuous: 'セクション区切り (現在の位置から開始)',
    menuSectionOddPage: 'セクション区切り (奇数ページから開始)',
    menuSectionEvenPage: 'セクション区切り (偶数ページから開始)',
    menuFootnote: '脚注…',
    menuCaption: '図表番号…',
    menuIndexAndTables: '索引と目次…',
    menuWatermark: '透かし…',
    menuPageNumbers: 'ページ番号…',
    menuTextBox: 'テキスト ボックス',
    menuVerticalTextBox: '縦書きテキスト ボックスの描画',
    menuAutoText: '定型句',
    menuAutoTextDialog: '定型句…',
    menuAutoTextNew: '新規…',
    menuDateAndTime: '日付と時刻…',
    menuObject: 'オブジェクト…',
    menuInsertFile: 'ファイル…',
    menuField: 'フィールド…',
    menuAdvancedSymbol: '記号と特殊文字…',
    menuComment: 'コメント',
    menuFormat: '書式',
    menuBold: '太字',
    menuItalic: '斜体',
    menuUnderline: '下線',
    menuAlign: '配置',
    menuAlignLeft: '左揃え',
    menuAlignCenter: '中央揃え',
    menuAlignRight: '右揃え',
    menuAlignJustify: '両端揃え',
    menuFont: 'フォント…',
    menuParagraph: '段落…',
    menuTextEffects: '文字の効果…',
    menuDocument: '文書のレイアウト…',
    menuBulletsAndNumbering: '箇条書きと段落番号…',
    menuBordersAndShading: '線種とページ罫線と網かけの設定…',
    menuColumns: '段組み…',
    menuTabs: 'タブとリーダー…',
    menuDropCap: 'ドロップ キャップ…',
    menuTextDirection: '縦書きと横書き…',
    menuChangeCase: '文字種の変換…',
    menuAsianLayout: '拡張書式',
    menuPhoneticGuide: 'ルビ…',
    menuEncloseCharacters: '囲い文字…',
    menuHorizontalInVertical: '縦中横…',
    menuCombineCharacters: '組み文字…',
    menuTwoLinesInOne: '割注…',
    menuFitText: '文字の均等割り付け…',
    menuStyle: 'スタイル…',
    menuFontSubstitution: 'フォントの置換…',
    menuFormatObject: 'オブジェクトの書式設定…',
    menuEquationOptions: '数式オプション…',
    menuSelectionPane: '選択ウィンドウ…',
    menuTools: 'ツール',
    menuTable: '表',
    menuTableInsert: '挿入',
    menuTableInsertTable: '表…',
    menuTableColsLeft: '左に列',
    menuTableColsRight: '右に列',
    menuTableRowsAbove: '上に行',
    menuTableRowsBelow: '下に行',
    menuTableCells: 'セル…',
    menuTableDelete: '削除',
    menuTableDeleteTable: '表',
    menuTableDeleteColumns: '列',
    menuTableDeleteRows: '行',
    menuTableSelect: '選択',
    menuTableSelectCell: 'セル',
    menuTableSelectColumn: '列',
    menuTableSelectRow: '行',
    menuTableSelectTable: '表',
    menuTableMergeCells: 'セルの結合',
    menuTableSplitCells: 'セルの分割…',
    menuTableSplitTable: '表の分割',
    menuTableAutoFit: '自動調整',
    menuTableAutoFitContents: '文字列の幅に自動調整',
    menuTableAutoFitWindow: 'ウィンドウ幅に自動調整',
    menuTableFixedWidth: '列の幅を固定する',
    menuTableDistributeRows: '行の高さを揃える',
    menuTableDistributeColumns: '列の幅を揃える',
    menuTableRepeatHeader: 'タイトル行の繰り返し',
    menuTableConvert: '変換',
    menuTableConvertTextToTable: '文字列を表にする…',
    menuTableConvertTableToText: '表の解除…',
    menuTableSort: '並べ替え…',
    menuTableFormula: '計算式…',
    menuTableGridlines: 'グリッド線の表示',
    menuTableProperties: '表のプロパティ…',
    menuWordCount: '文字カウント…',
    menuAutoCorrect: 'オートコレクトのオプション…',
    menuSpellingAndGrammar: 'スペル チェックと文章校正',
    menuSpellingDialog: 'スペル チェックと文章校正…',
    menuEditor: 'エディター',
    menuHideSpellingErrors: 'スペル ミスを表示しない',
    menuHideGrammarErrors: '文章校正エラーを表示しない',
    menuResetIgnoredWords: '無視した単語と文法をリセット',
    menuThesaurus: '類義語辞典…',
    menuHyphenation: 'ハイフネーション…',
    menuTranslate: '翻訳',
    menuTranslateSelection: '選択範囲の翻訳',
    menuTranslateDocument: 'ドキュメントの翻訳',
    menuSetTranslationLanguage: 'ドキュメントの翻訳言語の設定…',
    menuLanguage: '言語…',
    menuTrackChanges: '変更履歴の記録',
    menuHighlightChanges: '変更箇所の表示…',
    menuCompareDocuments: '文書の比較…',
    menuCombineDocuments: '文書の組み込み…',
    menuProtectDocument: '文書の保護…',
    menuEnvelopes: '封筒…',
    menuLabels: 'ラベル…',
    menuCustomizeKeyboard: 'キーボードのユーザー設定…',
    menuPreferences: '環境設定…',
    menuAiProofread: 'AI 校正',
    menuWindow: 'ウィンドウ',
    menuHelp: 'ヘルプ',
    menuShortcuts: 'キーボードショートカット',
  },
  ko: {
    dlgOpenDoc: '문서 열기',
    filterWord: 'Word 문서',
    odtImportMsg: '이 .{ext} 문서를 변환할까요?',
    odtImportDetail:
      'GenOffice은 ODF가 아니라 .docx를 편집합니다. 본문, 제목, 목록, 표, 이미지는 변환되지만 각주, 메모, 텍스트 상자는 완전히 유지되지 않습니다. 원본 파일은 변경되지 않으며, 변환된 .docx는 이후 저장이 기록되는 독립적인 사본입니다.',
    hwpImportDetail:
      'GenOffice은 HWP가 아니라 .docx를 편집합니다. 본문, 표, 그림, 머리말, 꼬리말은 변환되지만 수식, 각주, 삽입 개체는 변환되지 않으며 자동 번호는 일반 텍스트로 바뀝니다. 원본 파일은 변경되지 않으며, 변환된 .docx는 이후 저장이 기록되는 독립적인 사본입니다.',
    odtImportConvertBtn: '변환 후 열기',
    importSkipConfirm: '다음부터 묻지 않고 변환',
    docImportMsg: '이 .{ext} 문서를 변환할까요?',
    docImportDetail:
      'GenOffice은 Word 97-2003(.doc) 문서가 아니라 .docx를 편집합니다. 본문, 제목, 목록, 표, 이미지는 변환되지만 각주, 메모, 텍스트 상자는 완전히 유지되지 않습니다. 원본 파일은 변경되지 않으며, 변환된 .docx는 이후 저장이 기록되는 독립적인 사본입니다.',
    docImportConvertBtn: '변환 후 열기',
    dlgSaveAs: '다른 이름으로 저장',
    closeUnsavedMsg: '이 문서에 저장하지 않은 변경 사항이 있습니다.',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    closeNoReplyMsg: '문서가 응답하지 않으며 저장하지 않은 변경 사항이 있을 수 있습니다.',
    closeNoReplyDetail: '그래도 닫을까요? 저장하지 않은 변경 사항은 사라집니다.',
    btnCloseAnyway: '닫기',
    autosaveFoundTitle: '자동 복구 버전 발견',
    autosaveDiskNewer:
      '참고: 이 복구 사본이 저장된 후 디스크의 파일이 수정되었습니다. 복원하면 그 최신 변경 내용이 대체됩니다.',
    autosaveFoundBody:
      '마지막 세션에 저장되지 않은 변경 내용이 있습니다. 자동 저장 버전을 복원할까요?',
    autosaveRestore: '복원',
    autosaveDiscard: '취소',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
    extModifiedMsg: '이 파일이 다른 프로그램에서 수정되었습니다.',
    extModifiedDetail: '그래도 저장하여 디스크의 변경 사항을 덮어쓸까요?',
    btnOverwrite: '덮어쓰기',
    dlgInsertImage: '그림 삽입',
    filterImages: '그림',
    dlgMergeDataSource: '받는 사람 목록 선택',
    filterMergeDataSource: 'Excel 또는 CSV',
    dlgAddAttachment: '첨부 파일 추가',
    filterSupported: '지원되는 파일',
    filterAll: '모든 파일',
    dlgExportPdf: 'PDF로 내보내기',
    dlgExportHtml: 'HTML로 내보내기',
    dlgExportMarkdown: 'Markdown로 내보내기',
    dlgPickExportDir: '내보낼 폴더 선택',
    errUnsupportedExt: '.{ext} 형식은 지원되지 않습니다',
    errNotFile: '파일이 아닙니다',
    errTooLarge: '{mb}MB 제한을 초과했습니다',
    errTemplateProtected: '이 서식 파일은 암호로 보호되어 있습니다. 파일 > 열기로 여세요',
    errImageTooLarge: '이미지가 5MB 제한을 초과했습니다',
    errUnreadable: '읽을 수 없습니다',
    errFileTooLarge: '파일이 크기 제한을 초과했습니다',
    errParseFailed: '파일을 분석하지 못했습니다',
    errImageNoText:
      '이미지 첨부 파일은 텍스트를 제공하지 않으며, 이미지 형태로 사용자 메시지와 함께 전송되므로 이미지를 직접 확인하면 됩니다',
    errNotImage: '지원되지 않는 이미지 형식입니다',
    errGskNotLoggedIn:
      'Genspark에 로그인되어 있지 않습니다. 아래 "Genspark 로그인"을 눌러 로그인한 뒤 다시 시도하세요',
    errNoApiKey: '{provider}의 API 키가 설정되지 않았습니다',
    errAiRegion:
      '이 모델의 제공업체는 현재 지역에서 서비스를 제공하지 않습니다. 다른 모델을 선택하거나 네트워크를 변경하세요',
    errAiBusy: 'AI 서비스가 혼잡합니다. 잠시 후 다시 시도해 주세요',
    errNoModel: '모델 이름이 설정되지 않았습니다',
    menuFile: '파일',
    menuNewDoc: '새 문서',
    menuNewWindow: '새 창',
    menuOpen: '열기…',
    menuOpenRecent: '최근 문서 열기',
    menuNoRecent: '최근 문서 없음',
    menuRecentMore: '기타…',
    menuClose: '닫기',
    menuSave: '저장',
    menuSaveAs: '다른 이름으로 저장…',
    menuPageSetup: '페이지 설정…',
    menuExportPdf: 'PDF로 내보내기…',
    menuExportHtml: 'HTML로 내보내기…',
    menuExportMarkdown: 'Markdown로 내보내기…',
    menuExportImages: '이미지로 내보내기…',
    menuPrint: '인쇄…',
    menuEdit: '편집',
    menuUndo: '실행 취소',
    menuRedo: '다시 실행',
    menuRepeat: '반복',
    menuCantRepeat: '반복할 수 없음',
    menuCut: '잘라내기',
    menuCopy: '복사',
    menuPaste: '붙여넣기',
    menuPasteMatch: '서식 맞춰 붙여넣기',
    menuFindReplace: '찾기 및 바꾸기…',
    menuPasteSpecial: '선택하여 붙여넣기…',
    menuClear: '지우기',
    menuClearFormatting: '서식 지우기',
    menuClearContents: '내용',
    menuFindMenu: '찾기',
    menuFind: '찾기…',
    menuFindNext: '다음 찾기',
    menuFindPrevious: '이전 찾기',
    menuReplace: '바꾸기…',
    menuAdvancedFind: '고급 찾기 및 바꾸기…',
    menuLinks: '연결…',
    menuSelectData: '데이터 선택…',
    menuToggleDrawing: '그리기 전환',
    menuNewFromTemplate: '서식 파일에서 새로 만들기…',
    menuSaveAsTemplate: '서식 파일로 저장…',
    menuMove: '이동…',
    menuRename: '이름 바꾸기…',
    menuAlwaysOpenReadOnly: '항상 읽기 전용으로 열기',
    menuReduceFileSize: '파일 크기 줄이기…',
    menuProperties: '속성…',
    menuGoTo: '이동…',
    menuSelectAll: '모두 선택',
    menuView: '보기',
    menuZoom: '확대/축소',
    menuZoomIn: '확대',
    menuZoomOut: '축소',
    menuZoom100: '실제 크기(100%)',
    menuPrintLayout: '인쇄 모양',
    menuWebLayout: '웹 모양',
    menuOutline: '개요',
    menuDraft: '초안',
    menuFocus: '포커스',
    menuRibbon: '리본',
    menuSidebar: '사이드바',
    menuThumbnail: '축소판',
    menuNavigation: '탐색',
    menuReviewing: '검토',
    menuSidebarSearch: '검색',
    menuRuler: '눈금자',
    menuHeaderFooter: '머리글 및 바닥글',
    menuFootnotes: '각주',
    menuMarkup: '메모 및 변경 내용',
    menuRevealFormatting: '서식 표시',
    menuZoomDialog: '확대/축소…',
    menuGridlines: '눈금선',
    menuPageWidth: '페이지 너비',
    menuWholePage: '전체 페이지',
    menuMultiPage: '여러 페이지',
    menuAiSidebar: 'AI 사이드바',
    menuDarkMode: '다크 모드',
    menuFullscreen: '전체 화면 시작',
    menuInsert: '삽입',
    menuInsertTable: '표…',
    menuInsertImage: '그림…',
    menuInsertPageBreak: '페이지 나누기',
    menuInsertLink: '하이퍼링크…',
    menuInsertEquation: '수식…',
    menuPictures: '그림',
    menuPictureFromFile: '파일에서 그림…',
    menuShape: '도형',
    menuShapeRectangle: '직사각형',
    menuShapeRoundedRectangle: '둥근 모서리 직사각형',
    menuShapeTriangle: '삼각형',
    menuShapeOval: '타원',
    menuShapeLine: '선',
    menuChart: '차트',
    menuChartColumn: '세로 막대형',
    menuChartBar: '가로 막대형',
    menuChartLine: '꺾은선형',
    menuChartArea: '영역형',
    menuChartPie: '원형',
    menuChartTreemap: '트리맵',
    menuChartSunburst: '선버스트',
    menuChartHistogram: '히스토그램',
    menuChartPareto: '파레토',
    menuChartBoxWhisker: '상자 수염',
    menuChartScatter: '분산형',
    menuChartWaterfall: '폭포',
    menuChartFunnel: '깔때기형',
    menuChartStock: '주식형',
    menuChartSurface: '표면형',
    menuChartRadar: '방사형',
    menuIcons: '아이콘…',
    menuBookmark: '책갈피…',
    menuCrossReference: '상호 참조…',
    menuBreak: '나누기',
    menuColumnBreak: '단 나누기',
    menuSectionNextPage: '구역 나누기(다음 페이지부터)',
    menuSectionContinuous: '구역 나누기(이어서)',
    menuSectionOddPage: '구역 나누기(홀수 페이지부터)',
    menuSectionEvenPage: '구역 나누기(짝수 페이지부터)',
    menuFootnote: '각주…',
    menuCaption: '캡션…',
    menuIndexAndTables: '색인 및 목차…',
    menuWatermark: '워터마크…',
    menuPageNumbers: '페이지 번호…',
    menuTextBox: '텍스트 상자',
    menuVerticalTextBox: '세로 텍스트 상자 그리기',
    menuAutoText: '상용구',
    menuAutoTextDialog: '상용구…',
    menuAutoTextNew: '새로 만들기…',
    menuDateAndTime: '날짜 및 시간…',
    menuObject: '개체…',
    menuInsertFile: '파일…',
    menuField: '필드…',
    menuAdvancedSymbol: '고급 기호…',
    menuComment: '메모',
    menuFormat: '서식',
    menuBold: '굵게',
    menuItalic: '기울임꼴',
    menuUnderline: '밑줄',
    menuAlign: '맞춤',
    menuAlignLeft: '왼쪽 맞춤',
    menuAlignCenter: '가운데 맞춤',
    menuAlignRight: '오른쪽 맞춤',
    menuAlignJustify: '양쪽 맞춤',
    menuFont: '글꼴…',
    menuParagraph: '단락…',
    menuTextEffects: '텍스트 효과…',
    menuDocument: '문서…',
    menuBulletsAndNumbering: '글머리 기호 및 번호 매기기…',
    menuBordersAndShading: '테두리 및 음영…',
    menuColumns: '단…',
    menuTabs: '탭…',
    menuDropCap: '단락의 첫 문자 장식…',
    menuTextDirection: '텍스트 방향…',
    menuChangeCase: '대/소문자 바꾸기…',
    menuAsianLayout: '동아시아 언어 레이아웃',
    menuPhoneticGuide: '윗주 달기…',
    menuEncloseCharacters: '원 문자…',
    menuHorizontalInVertical: '세로 중 가로…',
    menuCombineCharacters: '문자 겹치기…',
    menuTwoLinesInOne: '두 줄을 한 줄로…',
    menuFitText: '글자 맞춤…',
    menuStyle: '스타일…',
    menuFontSubstitution: '글꼴 대체…',
    menuFormatObject: '개체 서식…',
    menuEquationOptions: '수식 옵션…',
    menuSelectionPane: '선택 창…',
    menuTools: '도구',
    menuTable: '표',
    menuTableInsert: '삽입',
    menuTableInsertTable: '표…',
    menuTableColsLeft: '왼쪽에 열',
    menuTableColsRight: '오른쪽에 열',
    menuTableRowsAbove: '위에 행',
    menuTableRowsBelow: '아래에 행',
    menuTableCells: '셀…',
    menuTableDelete: '삭제',
    menuTableDeleteTable: '표',
    menuTableDeleteColumns: '열',
    menuTableDeleteRows: '행',
    menuTableSelect: '선택',
    menuTableSelectCell: '셀',
    menuTableSelectColumn: '열',
    menuTableSelectRow: '행',
    menuTableSelectTable: '표',
    menuTableMergeCells: '셀 병합',
    menuTableSplitCells: '셀 분할…',
    menuTableSplitTable: '표 분할',
    menuTableAutoFit: '자동 맞춤 및 분배',
    menuTableAutoFitContents: '내용에 자동 맞춤',
    menuTableAutoFitWindow: '창에 자동 맞춤',
    menuTableFixedWidth: '열 너비 고정',
    menuTableDistributeRows: '행 균등 분배',
    menuTableDistributeColumns: '열 균등 분배',
    menuTableRepeatHeader: '머리글 행 반복',
    menuTableConvert: '변환',
    menuTableConvertTextToTable: '텍스트를 표로 변환…',
    menuTableConvertTableToText: '표를 텍스트로 변환…',
    menuTableSort: '정렬…',
    menuTableFormula: '수식…',
    menuTableGridlines: '눈금선 보기',
    menuTableProperties: '표 속성…',
    menuWordCount: '단어 개수…',
    menuAutoCorrect: '자동 고침 옵션…',
    menuSpellingAndGrammar: '맞춤법 및 문법',
    menuSpellingDialog: '맞춤법 및 문법…',
    menuEditor: '편집기',
    menuHideSpellingErrors: '맞춤법 오류 숨기기',
    menuHideGrammarErrors: '문법 오류 숨기기',
    menuResetIgnoredWords: '무시한 단어 및 문법 다시 설정',
    menuThesaurus: '동의어 사전…',
    menuHyphenation: '하이픈 넣기…',
    menuTranslate: '번역',
    menuTranslateSelection: '선택 영역 번역',
    menuTranslateDocument: '문서 번역',
    menuSetTranslationLanguage: '문서 번역 언어 설정…',
    menuLanguage: '언어…',
    menuTrackChanges: '변경 내용 추적',
    menuHighlightChanges: '변경 내용 강조…',
    menuCompareDocuments: '문서 비교…',
    menuCombineDocuments: '문서 병합…',
    menuProtectDocument: '문서 보호…',
    menuEnvelopes: '봉투…',
    menuLabels: '레이블…',
    menuCustomizeKeyboard: '키보드 사용자 지정…',
    menuPreferences: '기본 설정…',
    menuAiProofread: 'AI 교정',
    menuWindow: '창',
    menuHelp: '도움말',
    menuShortcuts: '키보드 바로 가기',
  },
  fr: {
    dlgOpenDoc: 'Ouvrir un document',
    filterWord: 'Documents Word',
    odtImportMsg: 'Convertir ce document .{ext} ?',
    odtImportDetail:
      'GenOffice modifie des .docx, pas de l’ODF. Le texte, les titres, les listes, les tableaux et les images sont convertis ; les notes de bas de page, les commentaires et les zones de texte ne sont pas entièrement conservés. Le fichier d’origine n’est jamais modifié ; le .docx converti est une copie distincte où s’effectuent les enregistrements suivants.',
    hwpImportDetail:
      'GenOffice modifie des .docx, pas des fichiers HWP. Le texte, les tableaux, les images, les en-têtes et les pieds de page sont convertis ; les équations, les notes de bas de page et les objets incorporés ne le sont pas, et la numérotation automatique devient du texte simple. Le fichier d’origine n’est jamais modifié ; le .docx converti est une copie distincte où s’effectuent les enregistrements suivants.',
    odtImportConvertBtn: 'Convertir et ouvrir',
    importSkipConfirm: 'La prochaine fois, convertir sans demander',
    docImportMsg: 'Convertir ce document .{ext} ?',
    docImportDetail:
      'GenOffice modifie des .docx, pas de documents Word 97-2003 (.doc). Le texte, les titres, les listes, les tableaux et les images sont convertis ; les notes de bas de page, les commentaires et les zones de texte ne sont pas entièrement conservés. Le fichier d’origine n’est jamais modifié ; le .docx converti est une copie distincte où s’effectuent les enregistrements suivants.',
    docImportConvertBtn: 'Convertir et ouvrir',
    dlgSaveAs: 'Enregistrer sous',
    closeUnsavedMsg: 'Ce document contient des modifications non enregistrées.',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    closeNoReplyMsg:
      'Le document ne répond pas et peut contenir des modifications non enregistrées.',
    closeNoReplyDetail: 'Fermer quand même ? Les modifications non enregistrées seront perdues.',
    btnCloseAnyway: 'Fermer quand même',
    autosaveFoundTitle: 'Version récupérée trouvée',
    autosaveDiskNewer:
      'Remarque : le fichier sur le disque a été modifié après l’enregistrement de cette copie de récupération. La restauration remplacera ces modifications plus récentes.',
    autosaveFoundBody:
      'Des modifications non enregistrées existent. Restaurer la version auto-enregistrée ?',
    autosaveRestore: 'Restaurer',
    autosaveDiscard: 'Ignorer',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
    extModifiedMsg: 'Le fichier a été modifié par un autre programme.',
    extModifiedDetail: 'Enregistrer quand même et écraser les modifications sur le disque ?',
    btnOverwrite: 'Écraser',
    dlgInsertImage: 'Insérer une image',
    filterImages: 'Images',
    dlgMergeDataSource: 'Sélectionner une liste de destinataires',
    filterMergeDataSource: 'Excel ou CSV',
    dlgAddAttachment: 'Ajouter des pièces jointes',
    filterSupported: 'Fichiers pris en charge',
    filterAll: 'Tous les fichiers',
    dlgExportPdf: 'Exporter au format PDF',
    dlgExportHtml: 'Exporter au format HTML',
    dlgExportMarkdown: 'Exporter au format Markdown',
    dlgPickExportDir: "Choisir le dossier d'exportation",
    errUnsupportedExt: 'les fichiers .{ext} ne sont pas pris en charge',
    errNotFile: "n'est pas un fichier",
    errTooLarge: 'dépasse la limite de {mb} Mo',
    errTemplateProtected:
      'ce modèle est protégé par mot de passe ; ouvrez-le avec Fichier > Ouvrir',
    errImageTooLarge: "l'image dépasse la limite de 5 Mo",
    errUnreadable: 'lecture impossible',
    errFileTooLarge: 'Le fichier dépasse la taille maximale',
    errParseFailed: "Échec de l'analyse du fichier",
    errImageNoText:
      "Les pièces jointes image ne fournissent pas de texte ; l'image est envoyée avec le message de l'utilisateur, consultez-la directement",
    errNotImage: "type d'image non pris en charge",
    errGskNotLoggedIn:
      'Non connecté à Genspark : cliquez sur « Se connecter à Genspark » ci-dessous, connectez-vous puis réessayez',
    errNoApiKey: 'Aucune clé API configurée pour {provider}',
    errAiRegion:
      'Le fournisseur de ce modèle ne dessert pas votre région actuelle — choisissez un autre modèle ou changez de réseau',
    errAiBusy: "Le service d'IA est actuellement surchargé — réessayez dans un instant",
    errNoModel: 'Aucun nom de modèle configuré',
    menuFile: 'Fichier',
    menuNewDoc: 'Nouveau document',
    menuNewWindow: 'Nouvelle fenêtre',
    menuOpen: 'Ouvrir…',
    menuOpenRecent: 'Ouvrir un document récent',
    menuNoRecent: 'Aucun document récent',
    menuRecentMore: 'Autres…',
    menuClose: 'Fermer',
    menuSave: 'Enregistrer',
    menuSaveAs: 'Enregistrer sous…',
    menuPageSetup: 'Mise en page…',
    menuExportPdf: 'Exporter au format PDF…',
    menuExportHtml: 'Exporter au format HTML…',
    menuExportMarkdown: 'Exporter au format Markdown…',
    menuExportImages: 'Exporter en images…',
    menuPrint: 'Imprimer…',
    menuEdit: 'Édition',
    menuUndo: 'Annuler',
    menuRedo: 'Rétablir',
    menuRepeat: 'Répéter',
    menuCantRepeat: 'Impossible de répéter',
    menuCut: 'Couper',
    menuCopy: 'Copier',
    menuPaste: 'Coller',
    menuPasteMatch: 'Coller et respecter la mise en forme',
    menuFindReplace: 'Rechercher et remplacer…',
    menuPasteSpecial: 'Collage spécial…',
    menuClear: 'Effacer',
    menuClearFormatting: 'Effacer la mise en forme',
    menuClearContents: 'Contenu',
    menuFindMenu: 'Rechercher',
    menuFind: 'Rechercher…',
    menuFindNext: 'Suivant',
    menuFindPrevious: 'Précédent',
    menuReplace: 'Remplacer…',
    menuAdvancedFind: 'Recherche avancée et remplacement…',
    menuLinks: 'Liaisons…',
    menuSelectData: 'Sélectionner des données…',
    menuToggleDrawing: 'Activer/désactiver le dessin',
    menuNewFromTemplate: 'Nouveau à partir d’un modèle…',
    menuSaveAsTemplate: 'Enregistrer comme modèle…',
    menuMove: 'Déplacer…',
    menuRename: 'Renommer…',
    menuAlwaysOpenReadOnly: 'Toujours ouvrir en lecture seule',
    menuReduceFileSize: 'Réduire la taille du fichier…',
    menuProperties: 'Propriétés…',
    menuGoTo: 'Atteindre…',
    menuSelectAll: 'Tout sélectionner',
    menuView: 'Affichage',
    menuZoom: 'Zoom',
    menuZoomIn: 'Zoom avant',
    menuZoomOut: 'Zoom arrière',
    menuZoom100: 'Taille réelle (100 %)',
    menuPrintLayout: 'Page',
    menuWebLayout: 'Web',
    menuOutline: 'Plan',
    menuDraft: 'Brouillon',
    menuFocus: 'Focus',
    menuRibbon: 'Ruban',
    menuSidebar: 'Barre latérale',
    menuThumbnail: 'Miniature',
    menuNavigation: 'Navigation',
    menuReviewing: 'Révision',
    menuSidebarSearch: 'Rechercher',
    menuRuler: 'Règle',
    menuHeaderFooter: 'En-tête et pied de page',
    menuFootnotes: 'Notes de bas de page',
    menuMarkup: 'Marques',
    menuRevealFormatting: 'Afficher la mise en forme',
    menuZoomDialog: 'Zoom…',
    menuGridlines: 'Quadrillage',
    menuPageWidth: 'Largeur de page',
    menuWholePage: 'Page entière',
    menuMultiPage: 'Plusieurs pages',
    menuAiSidebar: 'Volet IA',
    menuDarkMode: 'Mode sombre',
    menuFullscreen: 'Activer le mode plein écran',
    menuInsert: 'Insertion',
    menuInsertTable: 'Tableau…',
    menuInsertImage: 'Image…',
    menuInsertPageBreak: 'Saut de page',
    menuInsertLink: 'Lien hypertexte…',
    menuInsertEquation: 'Équation…',
    menuPictures: 'Images',
    menuPictureFromFile: 'Image à partir d’un fichier…',
    menuShape: 'Forme',
    menuShapeRectangle: 'Rectangle',
    menuShapeRoundedRectangle: 'Rectangle à coins arrondis',
    menuShapeTriangle: 'Triangle',
    menuShapeOval: 'Ovale',
    menuShapeLine: 'Trait',
    menuChart: 'Graphique',
    menuChartColumn: 'Histogramme',
    menuChartBar: 'Barres',
    menuChartLine: 'Courbes',
    menuChartArea: 'Aires',
    menuChartPie: 'Secteurs',
    menuChartTreemap: 'Compartimentage',
    menuChartSunburst: 'Rayons de soleil',
    menuChartHistogram: 'Histogramme de fréquences',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Zone et valeur',
    menuChartScatter: 'Nuage de points (XY)',
    menuChartWaterfall: 'Cascade',
    menuChartFunnel: 'Entonnoir',
    menuChartStock: 'Boursier',
    menuChartSurface: 'Surface',
    menuChartRadar: 'Radar',
    menuIcons: 'Icônes…',
    menuBookmark: 'Signet…',
    menuCrossReference: 'Renvoi…',
    menuBreak: 'Saut',
    menuColumnBreak: 'Saut de colonne',
    menuSectionNextPage: 'Saut de section (page suivante)',
    menuSectionContinuous: 'Saut de section (continu)',
    menuSectionOddPage: 'Saut de section (page impaire)',
    menuSectionEvenPage: 'Saut de section (page paire)',
    menuFootnote: 'Note de bas de page…',
    menuCaption: 'Légende…',
    menuIndexAndTables: 'Tables et index…',
    menuWatermark: 'Filigrane…',
    menuPageNumbers: 'Numéros de page…',
    menuTextBox: 'Zone de texte',
    menuVerticalTextBox: 'Dessiner une zone de texte verticale',
    menuAutoText: 'Insertion automatique',
    menuAutoTextDialog: 'Insertion automatique…',
    menuAutoTextNew: 'Nouveau…',
    menuDateAndTime: 'Date et heure…',
    menuObject: 'Objet…',
    menuInsertFile: 'Fichier…',
    menuField: 'Champ…',
    menuAdvancedSymbol: 'Caractères spéciaux…',
    menuComment: 'Commentaire',
    menuFormat: 'Format',
    menuBold: 'Gras',
    menuItalic: 'Italique',
    menuUnderline: 'Souligné',
    menuAlign: 'Alignement',
    menuAlignLeft: 'Aligner à gauche',
    menuAlignCenter: 'Centrer',
    menuAlignRight: 'Aligner à droite',
    menuAlignJustify: 'Justifier',
    menuFont: 'Police…',
    menuParagraph: 'Paragraphe…',
    menuTextEffects: 'Effets de texte…',
    menuDocument: 'Document…',
    menuBulletsAndNumbering: 'Puces et numéros…',
    menuBordersAndShading: 'Bordure et trame…',
    menuColumns: 'Colonnes…',
    menuTabs: 'Tabulations…',
    menuDropCap: 'Lettrine…',
    menuTextDirection: 'Orientation du texte…',
    menuChangeCase: 'Modifier la casse…',
    menuAsianLayout: 'Disposition asiatique',
    menuPhoneticGuide: 'Guide phonétique…',
    menuEncloseCharacters: 'Caractères entourés…',
    menuHorizontalInVertical: 'Horizontal dans vertical…',
    menuCombineCharacters: 'Combiner des caractères…',
    menuTwoLinesInOne: 'Deux lignes en une…',
    menuFitText: 'Ajuster le texte…',
    menuStyle: 'Style…',
    menuFontSubstitution: 'Substitution de polices…',
    menuFormatObject: 'Format de l’objet…',
    menuEquationOptions: 'Options d’équation…',
    menuSelectionPane: 'Volet Sélection…',
    menuTools: 'Outils',
    menuTable: 'Tableau',
    menuTableInsert: 'Insérer',
    menuTableInsertTable: 'Tableau…',
    menuTableColsLeft: 'Colonnes à gauche',
    menuTableColsRight: 'Colonnes à droite',
    menuTableRowsAbove: 'Lignes au-dessus',
    menuTableRowsBelow: 'Lignes en dessous',
    menuTableCells: 'Cellules…',
    menuTableDelete: 'Supprimer',
    menuTableDeleteTable: 'Tableau',
    menuTableDeleteColumns: 'Colonnes',
    menuTableDeleteRows: 'Lignes',
    menuTableSelect: 'Sélectionner',
    menuTableSelectCell: 'Cellule',
    menuTableSelectColumn: 'Colonne',
    menuTableSelectRow: 'Ligne',
    menuTableSelectTable: 'Tableau',
    menuTableMergeCells: 'Fusionner les cellules',
    menuTableSplitCells: 'Fractionner les cellules…',
    menuTableSplitTable: 'Fractionner le tableau',
    menuTableAutoFit: 'Ajustement automatique',
    menuTableAutoFitContents: 'Ajuster au contenu',
    menuTableAutoFitWindow: 'Ajuster à la fenêtre',
    menuTableFixedWidth: 'Largeur de colonne fixe',
    menuTableDistributeRows: 'Uniformiser la hauteur des lignes',
    menuTableDistributeColumns: 'Uniformiser la largeur des colonnes',
    menuTableRepeatHeader: 'Répéter les lignes d’en-tête',
    menuTableConvert: 'Convertir',
    menuTableConvertTextToTable: 'Convertir le texte en tableau…',
    menuTableConvertTableToText: 'Convertir le tableau en texte…',
    menuTableSort: 'Trier…',
    menuTableFormula: 'Formule…',
    menuTableGridlines: 'Afficher le quadrillage',
    menuTableProperties: 'Propriétés du tableau…',
    menuWordCount: 'Statistiques…',
    menuAutoCorrect: 'Options de correction automatique…',
    menuSpellingAndGrammar: 'Grammaire et orthographe',
    menuSpellingDialog: 'Grammaire et orthographe…',
    menuEditor: 'Rédacteur',
    menuHideSpellingErrors: 'Masquer les fautes d’orthographe',
    menuHideGrammarErrors: 'Masquer les fautes de grammaire',
    menuResetIgnoredWords: 'Réinitialiser les mots et la grammaire ignorés',
    menuThesaurus: 'Dictionnaire des synonymes…',
    menuHyphenation: 'Coupure de mots…',
    menuTranslate: 'Traduire',
    menuTranslateSelection: 'Traduire la sélection',
    menuTranslateDocument: 'Traduire le document',
    menuSetTranslationLanguage: 'Définir la langue de traduction du document…',
    menuLanguage: 'Langue…',
    menuTrackChanges: 'Suivi des modifications',
    menuHighlightChanges: 'Afficher les modifications…',
    menuCompareDocuments: 'Comparer des documents…',
    menuCombineDocuments: 'Combiner des documents…',
    menuProtectDocument: 'Protéger le document…',
    menuEnvelopes: 'Enveloppes…',
    menuLabels: 'Étiquettes…',
    menuCustomizeKeyboard: 'Personnaliser le clavier…',
    menuPreferences: 'Préférences…',
    menuAiProofread: 'Relecture IA',
    menuWindow: 'Fenêtre',
    menuHelp: 'Aide',
    menuShortcuts: 'Raccourcis clavier',
  },
  de: {
    dlgOpenDoc: 'Dokument öffnen',
    filterWord: 'Word-Dokumente',
    odtImportMsg: 'Dieses .{ext}-Dokument konvertieren?',
    odtImportDetail:
      'GenOffice bearbeitet .docx, nicht ODF. Text, Überschriften, Listen, Tabellen und Bilder werden übernommen; Fußnoten, Kommentare und Textfelder werden nicht vollständig bewahrt. Die Originaldatei bleibt unverändert; die konvertierte .docx-Kopie ist eigenständig, und spätere Speicherungen landen in ihr.',
    hwpImportDetail:
      'GenOffice bearbeitet .docx, nicht HWP. Text, Tabellen, Bilder sowie Kopf- und Fußzeilen werden übernommen; Formeln, Fußnoten und eingebettete Objekte werden nicht konvertiert, und automatische Nummerierung wird zu normalem Text. Die Originaldatei bleibt unverändert; die konvertierte .docx-Kopie ist eigenständig, und spätere Speicherungen landen in ihr.',
    odtImportConvertBtn: 'Konvertieren und öffnen',
    importSkipConfirm: 'Nächstes Mal ohne Nachfrage konvertieren',
    docImportMsg: 'Dieses .{ext}-Dokument konvertieren?',
    docImportDetail:
      'GenOffice bearbeitet .docx, nicht Word 97-2003-Dokumente (.doc). Text, Überschriften, Listen, Tabellen und Bilder werden übernommen; Fußnoten, Kommentare und Textfelder werden nicht vollständig bewahrt. Die Originaldatei bleibt unverändert; die konvertierte .docx-Kopie ist eigenständig, und spätere Speicherungen landen in ihr.',
    docImportConvertBtn: 'Konvertieren und öffnen',
    dlgSaveAs: 'Speichern unter',
    closeUnsavedMsg: 'Dieses Dokument enthält ungespeicherte Änderungen.',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    closeNoReplyMsg:
      'Das Dokument reagiert nicht und enthält möglicherweise ungespeicherte Änderungen.',
    closeNoReplyDetail: 'Trotzdem schließen? Ungespeicherte Änderungen gehen verloren.',
    btnCloseAnyway: 'Trotzdem schließen',
    autosaveFoundTitle: 'Wiederhergestellte Version gefunden',
    autosaveDiskNewer:
      'Hinweis: Die Datei auf dem Datenträger wurde nach dem Speichern dieser Wiederherstellungskopie geändert. Beim Wiederherstellen werden diese neueren Änderungen ersetzt.',
    autosaveFoundBody:
      'Es gibt ungespeicherte Änderungen. Automatisch gespeicherte Version wiederherstellen?',
    autosaveRestore: 'Wiederherstellen',
    autosaveDiscard: 'Verwerfen',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
    extModifiedMsg: 'Die Datei wurde von einem anderen Programm geändert.',
    extModifiedDetail: 'Trotzdem speichern und die Änderungen auf dem Datenträger überschreiben?',
    btnOverwrite: 'Überschreiben',
    dlgInsertImage: 'Bild einfügen',
    filterImages: 'Bilder',
    dlgMergeDataSource: 'Empfängerliste auswählen',
    filterMergeDataSource: 'Excel oder CSV',
    dlgAddAttachment: 'Anlagen hinzufügen',
    filterSupported: 'Unterstützte Dateien',
    filterAll: 'Alle Dateien',
    dlgExportPdf: 'Als PDF exportieren',
    dlgExportHtml: 'Als HTML exportieren',
    dlgExportMarkdown: 'Als Markdown exportieren',
    dlgPickExportDir: 'Exportordner auswählen',
    errUnsupportedExt: '.{ext}-Dateien werden nicht unterstützt',
    errNotFile: 'keine Datei',
    errTooLarge: 'überschreitet das Limit von {mb} MB',
    errTemplateProtected: 'Diese Vorlage ist kennwortgeschützt; öffnen Sie sie über Datei > Öffnen',
    errImageTooLarge: 'Bild überschreitet das Limit von 5 MB',
    errUnreadable: 'kann nicht gelesen werden',
    errFileTooLarge: 'Datei überschreitet die maximale Größe',
    errParseFailed: 'Datei konnte nicht analysiert werden',
    errImageNoText:
      'Bildanlagen liefern keinen Text; das Bild wird mit der Benutzernachricht gesendet und kann direkt betrachtet werden',
    errNotImage: 'kein unterstütztes Bildformat',
    errGskNotLoggedIn:
      'Nicht bei Genspark angemeldet: Klicken Sie unten auf „Bei Genspark anmelden“, melden Sie sich an und versuchen Sie es erneut',
    errNoApiKey: 'Kein API-Schlüssel für {provider} konfiguriert',
    errAiRegion:
      'Der Anbieter dieses Modells bedient Ihre aktuelle Region nicht — wählen Sie ein anderes Modell oder wechseln Sie das Netzwerk',
    errAiBusy: 'Der KI-Dienst ist derzeit überlastet — bitte gleich erneut versuchen',
    errNoModel: 'Kein Modellname konfiguriert',
    menuFile: 'Datei',
    menuNewDoc: 'Neues Dokument',
    menuNewWindow: 'Neues Fenster',
    menuOpen: 'Öffnen…',
    menuOpenRecent: 'Zuletzt verwendete Dokumente',
    menuNoRecent: 'Keine zuletzt verwendeten Dokumente',
    menuRecentMore: 'Weitere…',
    menuClose: 'Schließen',
    menuSave: 'Speichern',
    menuSaveAs: 'Speichern unter…',
    menuPageSetup: 'Seite einrichten…',
    menuExportPdf: 'Als PDF exportieren…',
    menuExportHtml: 'Als HTML exportieren…',
    menuExportMarkdown: 'Als Markdown exportieren…',
    menuExportImages: 'Als Bilder exportieren…',
    menuPrint: 'Drucken…',
    menuEdit: 'Bearbeiten',
    menuUndo: 'Rückgängig',
    menuRedo: 'Wiederholen',
    menuRepeat: 'Befehl wiederholen',
    menuCantRepeat: 'Wiederholen nicht möglich',
    menuCut: 'Ausschneiden',
    menuCopy: 'Kopieren',
    menuPaste: 'Einfügen',
    menuPasteMatch: 'Einfügen und Formatierung anpassen',
    menuFindReplace: 'Suchen und Ersetzen…',
    menuPasteSpecial: 'Inhalte einfügen…',
    menuClear: 'Löschen',
    menuClearFormatting: 'Formatierung löschen',
    menuClearContents: 'Inhalt',
    menuFindMenu: 'Suchen',
    menuFind: 'Suchen…',
    menuFindNext: 'Weitersuchen',
    menuFindPrevious: 'Vorheriges suchen',
    menuReplace: 'Ersetzen…',
    menuAdvancedFind: 'Erweitertes Suchen und Ersetzen…',
    menuLinks: 'Verknüpfungen…',
    menuSelectData: 'Daten auswählen…',
    menuToggleDrawing: 'Zeichnung umschalten',
    menuNewFromTemplate: 'Neu aus Vorlage…',
    menuSaveAsTemplate: 'Als Vorlage speichern…',
    menuMove: 'Verschieben…',
    menuRename: 'Umbenennen…',
    menuAlwaysOpenReadOnly: 'Immer schreibgeschützt öffnen',
    menuReduceFileSize: 'Dateigröße verringern…',
    menuProperties: 'Eigenschaften…',
    menuGoTo: 'Gehe zu…',
    menuSelectAll: 'Alles auswählen',
    menuView: 'Ansicht',
    menuZoom: 'Zoom',
    menuZoomIn: 'Vergrößern',
    menuZoomOut: 'Verkleinern',
    menuZoom100: 'Originalgröße (100 %)',
    menuPrintLayout: 'Drucklayout',
    menuWebLayout: 'Weblayout',
    menuOutline: 'Gliederung',
    menuDraft: 'Entwurf',
    menuFocus: 'Fokus',
    menuRibbon: 'Menüband',
    menuSidebar: 'Randleiste',
    menuThumbnail: 'Miniaturansicht',
    menuNavigation: 'Navigation',
    menuReviewing: 'Überarbeiten',
    menuSidebarSearch: 'Suchen',
    menuRuler: 'Lineal',
    menuHeaderFooter: 'Kopf- und Fußzeile',
    menuFootnotes: 'Fußnoten',
    menuMarkup: 'Markup',
    menuRevealFormatting: 'Formatierung anzeigen',
    menuZoomDialog: 'Zoom…',
    menuGridlines: 'Gitternetzlinien',
    menuPageWidth: 'Seitenbreite',
    menuWholePage: 'Ganze Seite',
    menuMultiPage: 'Mehrere Seiten',
    menuAiSidebar: 'KI-Seitenleiste',
    menuDarkMode: 'Dunkelmodus',
    menuFullscreen: 'Vollbild ein',
    menuInsert: 'Einfügen',
    menuInsertTable: 'Tabelle…',
    menuInsertImage: 'Bild…',
    menuInsertPageBreak: 'Seitenumbruch',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Formel…',
    menuPictures: 'Bilder',
    menuPictureFromFile: 'Bild aus Datei…',
    menuShape: 'Form',
    menuShapeRectangle: 'Rechteck',
    menuShapeRoundedRectangle: 'Abgerundetes Rechteck',
    menuShapeTriangle: 'Dreieck',
    menuShapeOval: 'Ellipse',
    menuShapeLine: 'Linie',
    menuChart: 'Diagramm',
    menuChartColumn: 'Säule',
    menuChartBar: 'Balken',
    menuChartLine: 'Linie',
    menuChartArea: 'Fläche',
    menuChartPie: 'Kreis',
    menuChartTreemap: 'Treemap',
    menuChartSunburst: 'Sunburst',
    menuChartHistogram: 'Histogramm',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Kastengrafik',
    menuChartScatter: 'Punkt (XY)',
    menuChartWaterfall: 'Wasserfall',
    menuChartFunnel: 'Trichter',
    menuChartStock: 'Kurs',
    menuChartSurface: 'Oberfläche',
    menuChartRadar: 'Netz',
    menuIcons: 'Symbole…',
    menuBookmark: 'Textmarke…',
    menuCrossReference: 'Querverweis…',
    menuBreak: 'Umbruch',
    menuColumnBreak: 'Spaltenumbruch',
    menuSectionNextPage: 'Abschnittsumbruch (Nächste Seite)',
    menuSectionContinuous: 'Abschnittsumbruch (Fortlaufend)',
    menuSectionOddPage: 'Abschnittsumbruch (Ungerade Seite)',
    menuSectionEvenPage: 'Abschnittsumbruch (Gerade Seite)',
    menuFootnote: 'Fußnote…',
    menuCaption: 'Beschriftung…',
    menuIndexAndTables: 'Index und Verzeichnisse…',
    menuWatermark: 'Wasserzeichen…',
    menuPageNumbers: 'Seitenzahlen…',
    menuTextBox: 'Textfeld',
    menuVerticalTextBox: 'Vertikales Textfeld zeichnen',
    menuAutoText: 'AutoText',
    menuAutoTextDialog: 'AutoText…',
    menuAutoTextNew: 'Neu…',
    menuDateAndTime: 'Datum und Uhrzeit…',
    menuObject: 'Objekt…',
    menuInsertFile: 'Datei…',
    menuField: 'Feld…',
    menuAdvancedSymbol: 'Erweitertes Symbol…',
    menuComment: 'Kommentar',
    menuFormat: 'Format',
    menuBold: 'Fett',
    menuItalic: 'Kursiv',
    menuUnderline: 'Unterstrichen',
    menuAlign: 'Ausrichten',
    menuAlignLeft: 'Linksbündig',
    menuAlignCenter: 'Zentriert',
    menuAlignRight: 'Rechtsbündig',
    menuAlignJustify: 'Blocksatz',
    menuFont: 'Schriftart…',
    menuParagraph: 'Absatz…',
    menuTextEffects: 'Texteffekte…',
    menuDocument: 'Dokument…',
    menuBulletsAndNumbering: 'Aufzählungszeichen und Nummerierung…',
    menuBordersAndShading: 'Rahmen und Schattierung…',
    menuColumns: 'Spalten…',
    menuTabs: 'Tabstopps…',
    menuDropCap: 'Initiale…',
    menuTextDirection: 'Textrichtung…',
    menuChangeCase: 'Groß-/Kleinschreibung…',
    menuAsianLayout: 'Asiatisches Layout',
    menuPhoneticGuide: 'Phonetische Leitzeichen…',
    menuEncloseCharacters: 'Eingeschlossene Zeichen…',
    menuHorizontalInVertical: 'Horizontal in vertikal…',
    menuCombineCharacters: 'Zeichen kombinieren…',
    menuTwoLinesInOne: 'Zwei Zeilen in einer…',
    menuFitText: 'Text einpassen…',
    menuStyle: 'Formatvorlage…',
    menuFontSubstitution: 'Schriftartersetzung…',
    menuFormatObject: 'Objekt formatieren…',
    menuEquationOptions: 'Formeloptionen…',
    menuSelectionPane: 'Auswahlbereich…',
    menuTools: 'Extras',
    menuTable: 'Tabelle',
    menuTableInsert: 'Einfügen',
    menuTableInsertTable: 'Tabelle…',
    menuTableColsLeft: 'Spalten links',
    menuTableColsRight: 'Spalten rechts',
    menuTableRowsAbove: 'Zeilen oberhalb',
    menuTableRowsBelow: 'Zeilen unterhalb',
    menuTableCells: 'Zellen…',
    menuTableDelete: 'Löschen',
    menuTableDeleteTable: 'Tabelle',
    menuTableDeleteColumns: 'Spalten',
    menuTableDeleteRows: 'Zeilen',
    menuTableSelect: 'Auswählen',
    menuTableSelectCell: 'Zelle',
    menuTableSelectColumn: 'Spalte',
    menuTableSelectRow: 'Zeile',
    menuTableSelectTable: 'Tabelle',
    menuTableMergeCells: 'Zellen verbinden',
    menuTableSplitCells: 'Zellen teilen…',
    menuTableSplitTable: 'Tabelle teilen',
    menuTableAutoFit: 'AutoAnpassen und Verteilen',
    menuTableAutoFitContents: 'An Inhalt anpassen',
    menuTableAutoFitWindow: 'An Fenster anpassen',
    menuTableFixedWidth: 'Feste Spaltenbreite',
    menuTableDistributeRows: 'Zeilen gleichmäßig verteilen',
    menuTableDistributeColumns: 'Spalten gleichmäßig verteilen',
    menuTableRepeatHeader: 'Überschriftenzeilen wiederholen',
    menuTableConvert: 'Umwandeln',
    menuTableConvertTextToTable: 'Text in Tabelle umwandeln…',
    menuTableConvertTableToText: 'Tabelle in Text umwandeln…',
    menuTableSort: 'Sortieren…',
    menuTableFormula: 'Formel…',
    menuTableGridlines: 'Gitternetzlinien anzeigen',
    menuTableProperties: 'Tabelleneigenschaften…',
    menuWordCount: 'Wörter zählen…',
    menuAutoCorrect: 'AutoKorrektur-Optionen…',
    menuSpellingAndGrammar: 'Rechtschreibung und Grammatik',
    menuSpellingDialog: 'Rechtschreibung und Grammatik…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Rechtschreibfehler ausblenden',
    menuHideGrammarErrors: 'Grammatikfehler ausblenden',
    menuResetIgnoredWords: 'Ignorierte Wörter und Grammatik zurücksetzen',
    menuThesaurus: 'Thesaurus…',
    menuHyphenation: 'Silbentrennung…',
    menuTranslate: 'Übersetzen',
    menuTranslateSelection: 'Auswahl übersetzen',
    menuTranslateDocument: 'Dokument übersetzen',
    menuSetTranslationLanguage: 'Übersetzungssprache für Dokument festlegen…',
    menuLanguage: 'Sprache…',
    menuTrackChanges: 'Änderungen nachverfolgen',
    menuHighlightChanges: 'Änderungen hervorheben…',
    menuCompareDocuments: 'Dokumente vergleichen…',
    menuCombineDocuments: 'Dokumente kombinieren…',
    menuProtectDocument: 'Dokument schützen…',
    menuEnvelopes: 'Umschläge…',
    menuLabels: 'Etiketten…',
    menuCustomizeKeyboard: 'Tastatur anpassen…',
    menuPreferences: 'Einstellungen…',
    menuAiProofread: 'KI-Korrektur',
    menuWindow: 'Fenster',
    menuHelp: 'Hilfe',
    menuShortcuts: 'Tastenkombinationen',
  },
  es: {
    dlgOpenDoc: 'Abrir documento',
    filterWord: 'Documentos de Word',
    odtImportMsg: '¿Convertir este documento .{ext}?',
    odtImportDetail:
      'GenOffice edita .docx, no ODF. El texto, los títulos, las listas, las tablas y las imágenes se convierten; las notas al pie, los comentarios y los cuadros de texto no se conservan por completo. El archivo original nunca se modifica; el .docx convertido es una copia independiente en la que se guardan los cambios posteriores.',
    hwpImportDetail:
      'GenOffice edita .docx, no HWP. El texto, las tablas, las imágenes, los encabezados y los pies de página se convierten; las ecuaciones, las notas al pie y los objetos incrustados no, y la numeración automática pasa a ser texto normal. El archivo original nunca se modifica; el .docx convertido es una copia independiente en la que se guardan los cambios posteriores.',
    odtImportConvertBtn: 'Convertir y abrir',
    importSkipConfirm: 'La próxima vez, convertir sin preguntar',
    docImportMsg: '¿Convertir este documento .{ext}?',
    docImportDetail:
      'GenOffice edita .docx, no documentos de Word 97-2003 (.doc). El texto, los títulos, las listas, las tablas y las imágenes se convierten; las notas al pie, los comentarios y los cuadros de texto no se conservan por completo. El archivo original nunca se modifica; el .docx convertido es una copia independiente en la que se guardan los cambios posteriores.',
    docImportConvertBtn: 'Convertir y abrir',
    dlgSaveAs: 'Guardar como',
    closeUnsavedMsg: 'Este documento tiene cambios sin guardar.',
    closeUnsavedDetail: '¿Desea guardarlos antes de cerrar?',
    closeNoReplyMsg: 'El documento no responde y puede tener cambios sin guardar.',
    closeNoReplyDetail: '¿Cerrar de todos modos? Los cambios sin guardar se perderán.',
    btnCloseAnyway: 'Cerrar de todos modos',
    autosaveFoundTitle: 'Se encontró una versión recuperada',
    autosaveDiskNewer:
      'Nota: el archivo en disco se modificó después de guardar esta copia de recuperación. Restaurar reemplazará esos cambios más recientes.',
    autosaveFoundBody:
      'Hay cambios sin guardar de la última sesión. ¿Restaurar la versión autoguardada?',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
    extModifiedMsg: 'El archivo ha sido modificado por otro programa.',
    extModifiedDetail: '¿Guardar de todos modos y sobrescribir los cambios en el disco?',
    btnOverwrite: 'Sobrescribir',
    dlgInsertImage: 'Insertar imagen',
    filterImages: 'Imágenes',
    dlgMergeDataSource: 'Seleccionar lista de destinatarios',
    filterMergeDataSource: 'Excel o CSV',
    dlgAddAttachment: 'Agregar datos adjuntos',
    filterSupported: 'Archivos compatibles',
    filterAll: 'Todos los archivos',
    dlgExportPdf: 'Exportar como PDF',
    dlgExportHtml: 'Exportar como HTML',
    dlgExportMarkdown: 'Exportar como Markdown',
    dlgPickExportDir: 'Elegir carpeta de exportación',
    errUnsupportedExt: 'los archivos .{ext} no son compatibles',
    errNotFile: 'no es un archivo',
    errTooLarge: 'supera el límite de {mb} MB',
    errTemplateProtected:
      'esta plantilla está protegida con contraseña; ábrala con Archivo > Abrir',
    errImageTooLarge: 'la imagen supera el límite de 5 MB',
    errUnreadable: 'no se puede leer',
    errFileTooLarge: 'El archivo supera el tamaño máximo',
    errParseFailed: 'No se pudo analizar el archivo',
    errImageNoText:
      'Las imágenes adjuntas no proporcionan texto; la imagen se envía junto con el mensaje del usuario, puedes verla directamente',
    errNotImage: 'no es un tipo de imagen compatible',
    errGskNotLoggedIn:
      'No has iniciado sesión en Genspark: pulsa «Iniciar sesión en Genspark» abajo, inicia sesión y vuelve a intentarlo',
    errNoApiKey: 'No hay clave de API configurada para {provider}',
    errAiRegion:
      'El proveedor de este modelo no da servicio en tu región actual — elige otro modelo o cambia de red',
    errAiBusy:
      'El servicio de IA está saturado en este momento; inténtalo de nuevo en unos instantes',
    errNoModel: 'No se ha configurado el nombre del modelo',
    menuFile: 'Archivo',
    menuNewDoc: 'Nuevo documento',
    menuNewWindow: 'Nueva ventana',
    menuOpen: 'Abrir…',
    menuOpenRecent: 'Abrir recientes',
    menuNoRecent: 'No hay documentos recientes',
    menuRecentMore: 'Más…',
    menuClose: 'Cerrar',
    menuSave: 'Guardar',
    menuSaveAs: 'Guardar como…',
    menuPageSetup: 'Configurar página…',
    menuExportPdf: 'Exportar como PDF…',
    menuExportHtml: 'Exportar como HTML…',
    menuExportMarkdown: 'Exportar como Markdown…',
    menuExportImages: 'Exportar como imágenes…',
    menuPrint: 'Imprimir…',
    menuEdit: 'Edición',
    menuUndo: 'Deshacer',
    menuRedo: 'Rehacer',
    menuRepeat: 'Repetir',
    menuCantRepeat: 'No se puede repetir',
    menuCut: 'Cortar',
    menuCopy: 'Copiar',
    menuPaste: 'Pegar',
    menuPasteMatch: 'Pegar y hacer coincidir formato',
    menuFindReplace: 'Buscar y reemplazar…',
    menuPasteSpecial: 'Pegado especial…',
    menuClear: 'Borrar',
    menuClearFormatting: 'Borrar formato',
    menuClearContents: 'Contenido',
    menuFindMenu: 'Buscar',
    menuFind: 'Buscar…',
    menuFindNext: 'Buscar siguiente',
    menuFindPrevious: 'Buscar anterior',
    menuReplace: 'Reemplazar…',
    menuAdvancedFind: 'Búsqueda avanzada y reemplazo…',
    menuLinks: 'Vínculos…',
    menuSelectData: 'Seleccionar datos…',
    menuToggleDrawing: 'Alternar dibujo',
    menuNewFromTemplate: 'Nuevo a partir de plantilla…',
    menuSaveAsTemplate: 'Guardar como plantilla…',
    menuMove: 'Mover…',
    menuRename: 'Cambiar nombre…',
    menuAlwaysOpenReadOnly: 'Abrir siempre como solo lectura',
    menuReduceFileSize: 'Reducir tamaño de archivo…',
    menuProperties: 'Propiedades…',
    menuGoTo: 'Ir a…',
    menuSelectAll: 'Seleccionar todo',
    menuView: 'Ver',
    menuZoom: 'Zoom',
    menuZoomIn: 'Acercar',
    menuZoomOut: 'Alejar',
    menuZoom100: 'Tamaño real (100 %)',
    menuPrintLayout: 'Diseño de impresión',
    menuWebLayout: 'Diseño web',
    menuOutline: 'Esquema',
    menuDraft: 'Borrador',
    menuFocus: 'Enfoque',
    menuRibbon: 'Cinta de opciones',
    menuSidebar: 'Barra lateral',
    menuThumbnail: 'Miniatura',
    menuNavigation: 'Navegación',
    menuReviewing: 'Revisión',
    menuSidebarSearch: 'Buscar',
    menuRuler: 'Regla',
    menuHeaderFooter: 'Encabezado y pie de página',
    menuFootnotes: 'Notas al pie',
    menuMarkup: 'Marcas',
    menuRevealFormatting: 'Mostrar formato',
    menuZoomDialog: 'Zoom…',
    menuGridlines: 'Líneas de cuadrícula',
    menuPageWidth: 'Ancho de página',
    menuWholePage: 'Página completa',
    menuMultiPage: 'Varias páginas',
    menuAiSidebar: 'Barra lateral de IA',
    menuDarkMode: 'Modo oscuro',
    menuFullscreen: 'Usar pantalla completa',
    menuInsert: 'Insertar',
    menuInsertTable: 'Tabla…',
    menuInsertImage: 'Imagen…',
    menuInsertPageBreak: 'Salto de página',
    menuInsertLink: 'Hipervínculo…',
    menuInsertEquation: 'Ecuación…',
    menuPictures: 'Imágenes',
    menuPictureFromFile: 'Imagen desde archivo…',
    menuShape: 'Forma',
    menuShapeRectangle: 'Rectángulo',
    menuShapeRoundedRectangle: 'Rectángulo redondeado',
    menuShapeTriangle: 'Triángulo',
    menuShapeOval: 'Elipse',
    menuShapeLine: 'Línea',
    menuChart: 'Gráfico',
    menuChartColumn: 'Columnas',
    menuChartBar: 'Barras',
    menuChartLine: 'Líneas',
    menuChartArea: 'Áreas',
    menuChartPie: 'Circular',
    menuChartTreemap: 'Rectángulos',
    menuChartSunburst: 'Proyección solar',
    menuChartHistogram: 'Histograma',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Cajas y bigotes',
    menuChartScatter: 'X Y (dispersión)',
    menuChartWaterfall: 'Cascada',
    menuChartFunnel: 'Embudo',
    menuChartStock: 'Cotizaciones',
    menuChartSurface: 'Superficie',
    menuChartRadar: 'Radial',
    menuIcons: 'Iconos…',
    menuBookmark: 'Marcador…',
    menuCrossReference: 'Referencia cruzada…',
    menuBreak: 'Salto',
    menuColumnBreak: 'Salto de columna',
    menuSectionNextPage: 'Salto de sección (página siguiente)',
    menuSectionContinuous: 'Salto de sección (continuo)',
    menuSectionOddPage: 'Salto de sección (página impar)',
    menuSectionEvenPage: 'Salto de sección (página par)',
    menuFootnote: 'Nota al pie…',
    menuCaption: 'Título…',
    menuIndexAndTables: 'Índice y tablas…',
    menuWatermark: 'Marca de agua…',
    menuPageNumbers: 'Números de página…',
    menuTextBox: 'Cuadro de texto',
    menuVerticalTextBox: 'Dibujar cuadro de texto vertical',
    menuAutoText: 'Autotexto',
    menuAutoTextDialog: 'Autotexto…',
    menuAutoTextNew: 'Nuevo…',
    menuDateAndTime: 'Fecha y hora…',
    menuObject: 'Objeto…',
    menuInsertFile: 'Archivo…',
    menuField: 'Campo…',
    menuAdvancedSymbol: 'Símbolo avanzado…',
    menuComment: 'Comentario',
    menuFormat: 'Formato',
    menuBold: 'Negrita',
    menuItalic: 'Cursiva',
    menuUnderline: 'Subrayado',
    menuAlign: 'Alinear',
    menuAlignLeft: 'Alinear a la izquierda',
    menuAlignCenter: 'Centrar',
    menuAlignRight: 'Alinear a la derecha',
    menuAlignJustify: 'Justificar',
    menuFont: 'Fuente…',
    menuParagraph: 'Párrafo…',
    menuTextEffects: 'Efectos de texto…',
    menuDocument: 'Documento…',
    menuBulletsAndNumbering: 'Numeración y viñetas…',
    menuBordersAndShading: 'Bordes y sombreado…',
    menuColumns: 'Columnas…',
    menuTabs: 'Tabulaciones…',
    menuDropCap: 'Letra capital…',
    menuTextDirection: 'Dirección del texto…',
    menuChangeCase: 'Cambiar mayúsculas y minúsculas…',
    menuAsianLayout: 'Diseño asiático',
    menuPhoneticGuide: 'Guía fonética…',
    menuEncloseCharacters: 'Caracteres en círculo…',
    menuHorizontalInVertical: 'Horizontal en vertical…',
    menuCombineCharacters: 'Combinar caracteres…',
    menuTwoLinesInOne: 'Dos líneas en una…',
    menuFitText: 'Ajustar texto…',
    menuStyle: 'Estilo…',
    menuFontSubstitution: 'Sustitución de fuentes…',
    menuFormatObject: 'Formato de objeto…',
    menuEquationOptions: 'Opciones de ecuación…',
    menuSelectionPane: 'Panel de selección…',
    menuTools: 'Herramientas',
    menuTable: 'Tabla',
    menuTableInsert: 'Insertar',
    menuTableInsertTable: 'Tabla…',
    menuTableColsLeft: 'Columnas a la izquierda',
    menuTableColsRight: 'Columnas a la derecha',
    menuTableRowsAbove: 'Filas arriba',
    menuTableRowsBelow: 'Filas abajo',
    menuTableCells: 'Celdas…',
    menuTableDelete: 'Eliminar',
    menuTableDeleteTable: 'Tabla',
    menuTableDeleteColumns: 'Columnas',
    menuTableDeleteRows: 'Filas',
    menuTableSelect: 'Seleccionar',
    menuTableSelectCell: 'Celda',
    menuTableSelectColumn: 'Columna',
    menuTableSelectRow: 'Fila',
    menuTableSelectTable: 'Tabla',
    menuTableMergeCells: 'Combinar celdas',
    menuTableSplitCells: 'Dividir celdas…',
    menuTableSplitTable: 'Dividir tabla',
    menuTableAutoFit: 'Autoajustar y distribuir',
    menuTableAutoFitContents: 'Autoajustar al contenido',
    menuTableAutoFitWindow: 'Autoajustar a la ventana',
    menuTableFixedWidth: 'Ancho de columna fijo',
    menuTableDistributeRows: 'Distribuir filas uniformemente',
    menuTableDistributeColumns: 'Distribuir columnas uniformemente',
    menuTableRepeatHeader: 'Repetir filas de encabezado',
    menuTableConvert: 'Convertir',
    menuTableConvertTextToTable: 'Convertir texto en tabla…',
    menuTableConvertTableToText: 'Convertir tabla en texto…',
    menuTableSort: 'Ordenar…',
    menuTableFormula: 'Fórmula…',
    menuTableGridlines: 'Ver líneas de cuadrícula',
    menuTableProperties: 'Propiedades de tabla…',
    menuWordCount: 'Contar palabras…',
    menuAutoCorrect: 'Opciones de autocorrección…',
    menuSpellingAndGrammar: 'Ortografía y gramática',
    menuSpellingDialog: 'Ortografía y gramática…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Ocultar errores de ortografía',
    menuHideGrammarErrors: 'Ocultar errores gramaticales',
    menuResetIgnoredWords: 'Restablecer palabras y gramática omitidas',
    menuThesaurus: 'Sinónimos…',
    menuHyphenation: 'Guiones…',
    menuTranslate: 'Traducir',
    menuTranslateSelection: 'Traducir selección',
    menuTranslateDocument: 'Traducir documento',
    menuSetTranslationLanguage: 'Establecer idioma de traducción del documento…',
    menuLanguage: 'Idioma…',
    menuTrackChanges: 'Control de cambios',
    menuHighlightChanges: 'Resaltar cambios…',
    menuCompareDocuments: 'Comparar documentos…',
    menuCombineDocuments: 'Combinar documentos…',
    menuProtectDocument: 'Proteger documento…',
    menuEnvelopes: 'Sobres…',
    menuLabels: 'Etiquetas…',
    menuCustomizeKeyboard: 'Personalizar teclado…',
    menuPreferences: 'Preferencias…',
    menuAiProofread: 'Corrección con IA',
    menuWindow: 'Ventana',
    menuHelp: 'Ayuda',
    menuShortcuts: 'Atajos de teclado',
  },
  th: {
    dlgOpenDoc: 'เปิดเอกสาร',
    filterWord: 'เอกสาร Word',
    odtImportMsg: 'แปลงเอกสาร .{ext} นี้ไหม?',
    odtImportDetail:
      'GenOffice แก้ไข .docx ไม่ใช่ ODF ข้อความ หัวข้อ รายการ ตาราง และรูปภาพจะถูกแปลง ส่วนเชิงอรรถ ความคิดเห็น และกล่องข้อความจะไม่ถูกเก็บไว้ทั้งหมด ไฟล์ต้นทางจะไม่ถูกแก้ไข และไฟล์ .docx ที่แปลงแล้วเป็นสำเนาแยกต่างหากที่การบันทึกครั้งต่อไปจะเขียนลง',
    hwpImportDetail:
      'GenOffice แก้ไข .docx ไม่ใช่ HWP ข้อความ ตาราง รูปภาพ หัวกระดาษ และท้ายกระดาษจะถูกแปลง ส่วนสมการ เชิงอรรถ และวัตถุที่ฝังไว้จะไม่ถูกแปลง และลำดับเลขอัตโนมัติจะกลายเป็นข้อความธรรมดา ไฟล์ต้นทางจะไม่ถูกแก้ไข และไฟล์ .docx ที่แปลงแล้วเป็นสำเนาแยกต่างหากที่การบันทึกครั้งต่อไปจะเขียนลง',
    odtImportConvertBtn: 'แปลงแล้วเปิด',
    importSkipConfirm: 'ครั้งถัดไปแปลงทันทีโดยไม่ต้องถาม',
    docImportMsg: 'แปลงเอกสาร .{ext} นี้ไหม?',
    docImportDetail:
      'GenOffice แก้ไข .docx ไม่ใช่เอกสาร Word 97-2003 (.doc) ข้อความ หัวข้อ รายการ ตาราง และรูปภาพจะถูกแปลง ส่วนเชิงอรรถ ความคิดเห็น และกล่องข้อความจะไม่ถูกเก็บไว้ทั้งหมด ไฟล์ต้นทางจะไม่ถูกแก้ไข และไฟล์ .docx ที่แปลงแล้วเป็นสำเนาแยกต่างหากที่การบันทึกครั้งต่อไปจะเขียนลง',
    docImportConvertBtn: 'แปลงแล้วเปิด',
    dlgSaveAs: 'บันทึกเป็น',
    closeUnsavedMsg: 'เอกสารนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    closeNoReplyMsg: 'เอกสารไม่ตอบสนองและอาจมีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeNoReplyDetail: 'ปิดต่อไปหรือไม่ การเปลี่ยนแปลงที่ยังไม่ได้บันทึกจะหายไป',
    btnCloseAnyway: 'ปิดต่อไป',
    autosaveFoundTitle: 'พบเวอร์ชันกู้คืนอัตโนมัติ',
    autosaveDiskNewer:
      'หมายเหตุ: ไฟล์บนดิสก์ถูกแก้ไขหลังจากบันทึกสำเนากู้คืนนี้ การกู้คืนจะแทนที่การเปลี่ยนแปลงที่ใหม่กว่าเหล่านั้น',
    autosaveFoundBody: 'มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึกจากครั้งก่อน ต้องการกู้คืนหรือไม่?',
    autosaveRestore: 'กู้คืน',
    autosaveDiscard: 'ละทิ้ง',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
    extModifiedMsg: 'ไฟล์ถูกแก้ไขโดยโปรแกรมอื่น',
    extModifiedDetail: 'บันทึกต่อไปและเขียนทับการเปลี่ยนแปลงบนดิสก์หรือไม่?',
    btnOverwrite: 'เขียนทับ',
    dlgInsertImage: 'แทรกรูปภาพ',
    filterImages: 'รูปภาพ',
    dlgMergeDataSource: 'เลือกรายชื่อผู้รับ',
    filterMergeDataSource: 'Excel หรือ CSV',
    dlgAddAttachment: 'เพิ่มสิ่งที่แนบ',
    filterSupported: 'ไฟล์ที่รองรับ',
    filterAll: 'ไฟล์ทั้งหมด',
    dlgExportPdf: 'ส่งออกเป็น PDF',
    dlgExportHtml: 'ส่งออกเป็น HTML',
    dlgExportMarkdown: 'ส่งออกเป็น Markdown',
    dlgPickExportDir: 'เลือกโฟลเดอร์ส่งออก',
    errUnsupportedExt: 'ไม่รองรับไฟล์ .{ext}',
    errNotFile: 'ไม่ใช่ไฟล์',
    errTooLarge: 'เกินขีดจำกัด {mb}MB',
    errTemplateProtected: 'เทมเพลตนี้มีรหัสผ่านป้องกัน โปรดเปิดด้วย ไฟล์ > เปิด',
    errImageTooLarge: 'รูปภาพเกินขีดจำกัด 5MB',
    errUnreadable: 'ไม่สามารถอ่านได้',
    errFileTooLarge: 'ไฟล์เกินขนาดสูงสุด',
    errParseFailed: 'แยกวิเคราะห์ไฟล์ไม่สำเร็จ',
    errImageNoText:
      'สิ่งที่แนบเป็นรูปภาพไม่มีข้อความ รูปจะถูกส่งไปพร้อมข้อความของผู้ใช้ ดูรูปได้โดยตรง',
    errNotImage: 'ไม่ใช่ชนิดรูปภาพที่รองรับ',
    errGskNotLoggedIn:
      'ยังไม่ได้ลงชื่อเข้าใช้ Genspark: แตะ “ลงชื่อเข้าใช้ Genspark” ด้านล่าง แล้วลองอีกครั้ง',
    errNoApiKey: 'ยังไม่ได้ตั้งค่า API Key ของ {provider}',
    errAiRegion:
      'ผู้ให้บริการของโมเดลนี้ไม่ให้บริการในภูมิภาคปัจจุบันของคุณ — โปรดเลือกโมเดลอื่นหรือเปลี่ยนเครือข่าย',
    errAiBusy: 'บริการ AI มีผู้ใช้งานจำนวนมากในขณะนี้ โปรดลองอีกครั้งในอีกสักครู่',
    errNoModel: 'ยังไม่ได้ตั้งค่าชื่อโมเดล',
    menuFile: 'ไฟล์',
    menuNewDoc: 'เอกสารใหม่',
    menuNewWindow: 'หน้าต่างใหม่',
    menuOpen: 'เปิด…',
    menuOpenRecent: 'เปิดเอกสารล่าสุด',
    menuNoRecent: 'ไม่มีเอกสารล่าสุด',
    menuRecentMore: 'เพิ่มเติม…',
    menuClose: 'ปิด',
    menuSave: 'บันทึก',
    menuSaveAs: 'บันทึกเป็น…',
    menuPageSetup: 'ตั้งค่าหน้ากระดาษ…',
    menuExportPdf: 'ส่งออกเป็น PDF…',
    menuExportHtml: 'ส่งออกเป็น HTML…',
    menuExportMarkdown: 'ส่งออกเป็น Markdown…',
    menuExportImages: 'ส่งออกเป็นรูปภาพ…',
    menuPrint: 'พิมพ์…',
    menuEdit: 'แก้ไข',
    menuUndo: 'เลิกทำ',
    menuRedo: 'ทำซ้ำ',
    menuRepeat: 'ทำซ้ำคำสั่ง',
    menuCantRepeat: 'ทำซ้ำไม่ได้',
    menuCut: 'ตัด',
    menuCopy: 'คัดลอก',
    menuPaste: 'วาง',
    menuPasteMatch: 'วางและจับคู่การจัดรูปแบบ',
    menuFindReplace: 'ค้นหาและแทนที่…',
    menuPasteSpecial: 'วางแบบพิเศษ…',
    menuClear: 'ล้าง',
    menuClearFormatting: 'ล้างการจัดรูปแบบ',
    menuClearContents: 'เนื้อหา',
    menuFindMenu: 'ค้นหา',
    menuFind: 'ค้นหา…',
    menuFindNext: 'ค้นหาถัดไป',
    menuFindPrevious: 'ค้นหาก่อนหน้า',
    menuReplace: 'แทนที่…',
    menuAdvancedFind: 'การค้นหาและแทนที่ขั้นสูง…',
    menuLinks: 'ลิงก์…',
    menuSelectData: 'เลือกข้อมูล…',
    menuToggleDrawing: 'สลับการวาด',
    menuNewFromTemplate: 'สร้างจากเทมเพลต…',
    menuSaveAsTemplate: 'บันทึกเป็นเทมเพลต…',
    menuMove: 'ย้าย…',
    menuRename: 'เปลี่ยนชื่อ…',
    menuAlwaysOpenReadOnly: 'เปิดแบบอ่านอย่างเดียวเสมอ',
    menuReduceFileSize: 'ลดขนาดไฟล์…',
    menuProperties: 'คุณสมบัติ…',
    menuGoTo: 'ไปที่…',
    menuSelectAll: 'เลือกทั้งหมด',
    menuView: 'มุมมอง',
    menuZoom: 'ซูม',
    menuZoomIn: 'ขยาย',
    menuZoomOut: 'ย่อ',
    menuZoom100: 'ขนาดจริง (100%)',
    menuPrintLayout: 'เค้าโครงเหมือนพิมพ์',
    menuWebLayout: 'เค้าโครงเว็บ',
    menuOutline: 'เค้าร่าง',
    menuDraft: 'แบบร่าง',
    menuFocus: 'โฟกัส',
    menuRibbon: 'ริบบอน',
    menuSidebar: 'แถบด้านข้าง',
    menuThumbnail: 'รูปขนาดย่อ',
    menuNavigation: 'การนำทาง',
    menuReviewing: 'การรีวิว',
    menuSidebarSearch: 'ค้นหา',
    menuRuler: 'ไม้บรรทัด',
    menuHeaderFooter: 'หัวกระดาษและท้ายกระดาษ',
    menuFootnotes: 'เชิงอรรถ',
    menuMarkup: 'มาร์กอัป',
    menuRevealFormatting: 'แสดงการจัดรูปแบบ',
    menuZoomDialog: 'ย่อ/ขยาย…',
    menuGridlines: 'เส้นตาราง',
    menuPageWidth: 'ความกว้างของหน้า',
    menuWholePage: 'ทั้งหน้า',
    menuMultiPage: 'หลายหน้า',
    menuAiSidebar: 'แถบข้าง AI',
    menuDarkMode: 'โหมดมืด',
    menuFullscreen: 'เข้าสู่โหมดเต็มหน้าจอ',
    menuInsert: 'แทรก',
    menuInsertTable: 'ตาราง…',
    menuInsertImage: 'รูปภาพ…',
    menuInsertPageBreak: 'ตัวแบ่งหน้า',
    menuInsertLink: 'ไฮเปอร์ลิงก์…',
    menuInsertEquation: 'สมการ…',
    menuPictures: 'รูปภาพ',
    menuPictureFromFile: 'รูปภาพจากไฟล์…',
    menuShape: 'รูปร่าง',
    menuShapeRectangle: 'สี่เหลี่ยมผืนผ้า',
    menuShapeRoundedRectangle: 'สี่เหลี่ยมผืนผ้ามุมมน',
    menuShapeTriangle: 'สามเหลี่ยม',
    menuShapeOval: 'วงรี',
    menuShapeLine: 'เส้น',
    menuChart: 'แผนภูมิ',
    menuChartColumn: 'คอลัมน์',
    menuChartBar: 'แท่ง',
    menuChartLine: 'เส้น',
    menuChartArea: 'พื้นที่',
    menuChartPie: 'วงกลม',
    menuChartTreemap: 'แผนภูมิต้นไม้',
    menuChartSunburst: 'ซันเบิร์สต์',
    menuChartHistogram: 'ฮิสโทแกรม',
    menuChartPareto: 'พาเรโต',
    menuChartBoxWhisker: 'กล่องและหนวด',
    menuChartScatter: 'X Y (กระจาย)',
    menuChartWaterfall: 'น้ำตก',
    menuChartFunnel: 'กรวย',
    menuChartStock: 'หุ้น',
    menuChartSurface: 'พื้นผิว',
    menuChartRadar: 'เรดาร์',
    menuIcons: 'ไอคอน…',
    menuBookmark: 'บุ๊กมาร์ก…',
    menuCrossReference: 'การอ้างอิงโยง…',
    menuBreak: 'ตัวแบ่ง',
    menuColumnBreak: 'ตัวแบ่งคอลัมน์',
    menuSectionNextPage: 'ตัวแบ่งส่วน (หน้าถัดไป)',
    menuSectionContinuous: 'ตัวแบ่งส่วน (ต่อเนื่อง)',
    menuSectionOddPage: 'ตัวแบ่งส่วน (หน้าคี่)',
    menuSectionEvenPage: 'ตัวแบ่งส่วน (หน้าคู่)',
    menuFootnote: 'เชิงอรรถ…',
    menuCaption: 'คำอธิบายภาพ…',
    menuIndexAndTables: 'ดัชนีและสารบัญ…',
    menuWatermark: 'ลายน้ำ…',
    menuPageNumbers: 'หมายเลขหน้า…',
    menuTextBox: 'กล่องข้อความ',
    menuVerticalTextBox: 'วาดกล่องข้อความแนวตั้ง',
    menuAutoText: 'ข้อความอัตโนมัติ',
    menuAutoTextDialog: 'ข้อความอัตโนมัติ…',
    menuAutoTextNew: 'สร้าง…',
    menuDateAndTime: 'วันที่และเวลา…',
    menuObject: 'วัตถุ…',
    menuInsertFile: 'ไฟล์…',
    menuField: 'เขตข้อมูล…',
    menuAdvancedSymbol: 'สัญลักษณ์ขั้นสูง…',
    menuComment: 'ข้อคิดเห็น',
    menuFormat: 'รูปแบบ',
    menuBold: 'ตัวหนา',
    menuItalic: 'ตัวเอียง',
    menuUnderline: 'ขีดเส้นใต้',
    menuAlign: 'จัดแนว',
    menuAlignLeft: 'จัดชิดซ้าย',
    menuAlignCenter: 'จัดกึ่งกลาง',
    menuAlignRight: 'จัดชิดขวา',
    menuAlignJustify: 'จัดชิดขอบ',
    menuFont: 'ฟอนต์…',
    menuParagraph: 'ย่อหน้า…',
    menuTextEffects: 'เอฟเฟ็กต์ข้อความ…',
    menuDocument: 'เอกสาร…',
    menuBulletsAndNumbering: 'สัญลักษณ์แสดงหัวข้อย่อยและลำดับเลข…',
    menuBordersAndShading: 'เส้นขอบและแรเงา…',
    menuColumns: 'คอลัมน์…',
    menuTabs: 'แท็บ…',
    menuDropCap: 'ตัวอักษรขึ้นต้นขนาดใหญ่…',
    menuTextDirection: 'ทิศทางข้อความ…',
    menuChangeCase: 'เปลี่ยนตัวพิมพ์…',
    menuAsianLayout: 'เค้าโครงแบบเอเชีย',
    menuPhoneticGuide: 'คำแนะนำการออกเสียง…',
    menuEncloseCharacters: 'ล้อมรอบอักขระ…',
    menuHorizontalInVertical: 'แนวนอนในแนวตั้ง…',
    menuCombineCharacters: 'รวมอักขระ…',
    menuTwoLinesInOne: 'สองบรรทัดในหนึ่ง…',
    menuFitText: 'จัดข้อความให้พอดี…',
    menuStyle: 'สไตล์…',
    menuFontSubstitution: 'การแทนที่ฟอนต์…',
    menuFormatObject: 'จัดรูปแบบวัตถุ…',
    menuEquationOptions: 'ตัวเลือกสมการ…',
    menuSelectionPane: 'บานหน้าต่างส่วนที่เลือก…',
    menuTools: 'เครื่องมือ',
    menuTable: 'ตาราง',
    menuTableInsert: 'แทรก',
    menuTableInsertTable: 'ตาราง…',
    menuTableColsLeft: 'คอลัมน์ทางซ้าย',
    menuTableColsRight: 'คอลัมน์ทางขวา',
    menuTableRowsAbove: 'แถวด้านบน',
    menuTableRowsBelow: 'แถวด้านล่าง',
    menuTableCells: 'เซลล์…',
    menuTableDelete: 'ลบ',
    menuTableDeleteTable: 'ตาราง',
    menuTableDeleteColumns: 'คอลัมน์',
    menuTableDeleteRows: 'แถว',
    menuTableSelect: 'เลือก',
    menuTableSelectCell: 'เซลล์',
    menuTableSelectColumn: 'คอลัมน์',
    menuTableSelectRow: 'แถว',
    menuTableSelectTable: 'ตาราง',
    menuTableMergeCells: 'ผสานเซลล์',
    menuTableSplitCells: 'แยกเซลล์…',
    menuTableSplitTable: 'แยกตาราง',
    menuTableAutoFit: 'ปรับพอดีอัตโนมัติ',
    menuTableAutoFitContents: 'ปรับพอดีกับเนื้อหา',
    menuTableAutoFitWindow: 'ปรับพอดีกับหน้าต่าง',
    menuTableFixedWidth: 'ความกว้างคอลัมน์คงที่',
    menuTableDistributeRows: 'กระจายแถวเท่ากัน',
    menuTableDistributeColumns: 'กระจายคอลัมน์เท่ากัน',
    menuTableRepeatHeader: 'ทำซ้ำแถวส่วนหัว',
    menuTableConvert: 'แปลง',
    menuTableConvertTextToTable: 'แปลงข้อความเป็นตาราง…',
    menuTableConvertTableToText: 'แปลงตารางเป็นข้อความ…',
    menuTableSort: 'เรียงลำดับ…',
    menuTableFormula: 'สูตร…',
    menuTableGridlines: 'แสดงเส้นตาราง',
    menuTableProperties: 'คุณสมบัติตาราง…',
    menuWordCount: 'นับจำนวนคำ…',
    menuAutoCorrect: 'ตัวเลือกการแก้ไขอัตโนมัติ…',
    menuSpellingAndGrammar: 'การสะกดและไวยากรณ์',
    menuSpellingDialog: 'การสะกดและไวยากรณ์…',
    menuEditor: 'ตัวแก้ไข',
    menuHideSpellingErrors: 'ซ่อนข้อผิดพลาดการสะกด',
    menuHideGrammarErrors: 'ซ่อนข้อผิดพลาดไวยากรณ์',
    menuResetIgnoredWords: 'รีเซ็ตคำและไวยากรณ์ที่ละเว้น',
    menuThesaurus: 'อรรถาภิธาน…',
    menuHyphenation: 'การใส่ยัติภังค์…',
    menuTranslate: 'แปล',
    menuTranslateSelection: 'แปลส่วนที่เลือก',
    menuTranslateDocument: 'แปลเอกสาร',
    menuSetTranslationLanguage: 'ตั้งค่าภาษาการแปลเอกสาร…',
    menuLanguage: 'ภาษา…',
    menuTrackChanges: 'ติดตามการเปลี่ยนแปลง',
    menuHighlightChanges: 'เน้นการเปลี่ยนแปลง…',
    menuCompareDocuments: 'เปรียบเทียบเอกสาร…',
    menuCombineDocuments: 'รวมเอกสาร…',
    menuProtectDocument: 'ป้องกันเอกสาร…',
    menuEnvelopes: 'ซองจดหมาย…',
    menuLabels: 'ป้ายผนึก…',
    menuCustomizeKeyboard: 'กำหนดคีย์บอร์ดเอง…',
    menuPreferences: 'การตั้งค่า…',
    menuAiProofread: 'พิสูจน์อักษรด้วย AI',
    menuWindow: 'หน้าต่าง',
    menuHelp: 'วิธีใช้',
    menuShortcuts: 'แป้นพิมพ์ลัด',
  },
  id: {
    dlgOpenDoc: 'Buka Dokumen',
    filterWord: 'Dokumen Word',
    odtImportMsg: 'Konversi dokumen .{ext} ini?',
    odtImportDetail:
      'GenOffice mengedit .docx, bukan ODF. Teks isi, judul, daftar, tabel, dan gambar ikut dikonversi; catatan kaki, komentar, dan kotak teks tidak sepenuhnya dipertahankan. File asli tidak pernah diubah; salinan .docx hasil konversi berdiri sendiri dan penyimpanan berikutnya ditulis ke sana.',
    hwpImportDetail:
      'GenOffice mengedit .docx, bukan HWP. Teks isi, tabel, gambar, header, dan footer ikut dikonversi; persamaan, catatan kaki, dan objek tersemat tidak dikonversi, dan penomoran otomatis menjadi teks biasa. File asli tidak pernah diubah; salinan .docx hasil konversi berdiri sendiri dan penyimpanan berikutnya ditulis ke sana.',
    odtImportConvertBtn: 'Konversi dan buka',
    importSkipConfirm: 'Lain kali langsung konversi tanpa bertanya',
    docImportMsg: 'Konversi dokumen .{ext} ini?',
    docImportDetail:
      'GenOffice mengedit .docx, bukan dokumen Word 97-2003 (.doc). Teks isi, judul, daftar, tabel, dan gambar ikut dikonversi; catatan kaki, komentar, dan kotak teks tidak sepenuhnya dipertahankan. File asli tidak pernah diubah; salinan .docx hasil konversi berdiri sendiri dan penyimpanan berikutnya ditulis ke sana.',
    docImportConvertBtn: 'Konversi dan buka',
    dlgSaveAs: 'Simpan Sebagai',
    closeUnsavedMsg: 'Dokumen ini memiliki perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    closeNoReplyMsg: 'Dokumen tidak merespons dan mungkin memiliki perubahan yang belum disimpan.',
    closeNoReplyDetail: 'Tetap tutup? Perubahan yang belum disimpan akan hilang.',
    btnCloseAnyway: 'Tetap Tutup',
    autosaveFoundTitle: 'Versi pemulihan ditemukan',
    autosaveDiskNewer:
      'Catatan: file di disk diubah setelah salinan pemulihan ini disimpan. Memulihkan akan menggantikan perubahan yang lebih baru tersebut.',
    autosaveFoundBody:
      'Ada perubahan yang belum disimpan dari sesi terakhir. Pulihkan versi tersimpan otomatis?',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    extModifiedMsg: 'File telah diubah oleh program lain.',
    extModifiedDetail: 'Tetap simpan dan timpa perubahan di disk?',
    btnOverwrite: 'Timpa',
    dlgInsertImage: 'Sisipkan Gambar',
    filterImages: 'Gambar',
    dlgMergeDataSource: 'Pilih daftar penerima',
    filterMergeDataSource: 'Excel atau CSV',
    dlgAddAttachment: 'Tambahkan Lampiran',
    filterSupported: 'File yang Didukung',
    filterAll: 'Semua File',
    dlgExportPdf: 'Ekspor sebagai PDF',
    dlgExportHtml: 'Ekspor sebagai HTML',
    dlgExportMarkdown: 'Ekspor sebagai Markdown',
    dlgPickExportDir: 'Pilih Folder Ekspor',
    errUnsupportedExt: 'file .{ext} tidak didukung',
    errNotFile: 'bukan file',
    errTooLarge: 'melebihi batas {mb}MB',
    errTemplateProtected: 'templat ini dilindungi kata sandi; buka melalui File > Buka',
    errImageTooLarge: 'gambar melebihi batas 5MB',
    errUnreadable: 'tidak dapat dibaca',
    errFileTooLarge: 'File melebihi batas ukuran',
    errParseFailed: 'Gagal mengurai file',
    errImageNoText:
      'Lampiran gambar tidak menyediakan teks; gambar dikirim bersama pesan pengguna dan dapat dilihat langsung',
    errNotImage: 'bukan jenis gambar yang didukung',
    errGskNotLoggedIn: 'Belum masuk ke Genspark: klik “Masuk ke Genspark” di bawah, lalu coba lagi',
    errNoApiKey: 'API Key untuk {provider} belum dikonfigurasi',
    errAiRegion:
      'Penyedia model ini tidak melayani wilayah Anda saat ini — pilih model lain atau ganti jaringan',
    errAiBusy: 'Layanan AI sedang sibuk — silakan coba lagi sebentar lagi',
    errNoModel: 'Nama model belum dikonfigurasi',
    menuFile: 'File',
    menuNewDoc: 'Dokumen Baru',
    menuNewWindow: 'Jendela Baru',
    menuOpen: 'Buka…',
    menuOpenRecent: 'Buka Terbaru',
    menuNoRecent: 'Tidak Ada Dokumen Terbaru',
    menuRecentMore: 'Lainnya…',
    menuClose: 'Tutup',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuPageSetup: 'Penyetelan Halaman…',
    menuExportPdf: 'Ekspor sebagai PDF…',
    menuExportHtml: 'Ekspor sebagai HTML…',
    menuExportMarkdown: 'Ekspor sebagai Markdown…',
    menuExportImages: 'Ekspor sebagai gambar…',
    menuPrint: 'Cetak…',
    menuEdit: 'Edit',
    menuUndo: 'Urungkan',
    menuRedo: 'Ulangi',
    menuRepeat: 'Ulangi Perintah',
    menuCantRepeat: 'Tidak Dapat Mengulang',
    menuCut: 'Potong',
    menuCopy: 'Salin',
    menuPaste: 'Tempel',
    menuPasteMatch: 'Tempel dan Cocokkan Format',
    menuFindReplace: 'Temukan dan Ganti…',
    menuPasteSpecial: 'Tempel Spesial…',
    menuClear: 'Hapus',
    menuClearFormatting: 'Hapus Pemformatan',
    menuClearContents: 'Konten',
    menuFindMenu: 'Temukan',
    menuFind: 'Temukan…',
    menuFindNext: 'Temukan Berikutnya',
    menuFindPrevious: 'Temukan Sebelumnya',
    menuReplace: 'Ganti…',
    menuAdvancedFind: 'Temukan dan Ganti Tingkat Lanjut…',
    menuLinks: 'Tautan…',
    menuSelectData: 'Pilih Data…',
    menuToggleDrawing: 'Alihkan Gambar',
    menuNewFromTemplate: 'Baru dari Templat…',
    menuSaveAsTemplate: 'Simpan sebagai Templat…',
    menuMove: 'Pindahkan…',
    menuRename: 'Ganti Nama…',
    menuAlwaysOpenReadOnly: 'Selalu Buka Baca-Saja',
    menuReduceFileSize: 'Kurangi Ukuran File…',
    menuProperties: 'Properti…',
    menuGoTo: 'Pergi ke…',
    menuSelectAll: 'Pilih Semua',
    menuView: 'Tampilan',
    menuZoom: 'Zoom',
    menuZoomIn: 'Perbesar',
    menuZoomOut: 'Perkecil',
    menuZoom100: 'Ukuran Sebenarnya (100%)',
    menuPrintLayout: 'Tata Letak Cetak',
    menuWebLayout: 'Tata Letak Web',
    menuOutline: 'Kerangka',
    menuDraft: 'Draf',
    menuFocus: 'Fokus',
    menuRibbon: 'Pita',
    menuSidebar: 'Bilah Sisi',
    menuThumbnail: 'Gambar Mini',
    menuNavigation: 'Navigasi',
    menuReviewing: 'Peninjauan',
    menuSidebarSearch: 'Cari',
    menuRuler: 'Penggaris',
    menuHeaderFooter: 'Header dan Footer',
    menuFootnotes: 'Catatan Kaki',
    menuMarkup: 'Markup',
    menuRevealFormatting: 'Tampilkan Pemformatan',
    menuZoomDialog: 'Zoom…',
    menuGridlines: 'Garis Kisi',
    menuPageWidth: 'Lebar Halaman',
    menuWholePage: 'Seluruh Halaman',
    menuMultiPage: 'Beberapa halaman',
    menuAiSidebar: 'Bilah Samping AI',
    menuDarkMode: 'Mode Gelap',
    menuFullscreen: 'Masuk Layar Penuh',
    menuInsert: 'Sisipkan',
    menuInsertTable: 'Tabel…',
    menuInsertImage: 'Gambar…',
    menuInsertPageBreak: 'Pemisah Halaman',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Persamaan…',
    menuPictures: 'Gambar',
    menuPictureFromFile: 'Gambar dari File…',
    menuShape: 'Bentuk',
    menuShapeRectangle: 'Persegi Panjang',
    menuShapeRoundedRectangle: 'Persegi Panjang Bersudut Bulat',
    menuShapeTriangle: 'Segitiga',
    menuShapeOval: 'Oval',
    menuShapeLine: 'Garis',
    menuChart: 'Bagan',
    menuChartColumn: 'Kolom',
    menuChartBar: 'Batang',
    menuChartLine: 'Garis',
    menuChartArea: 'Area',
    menuChartPie: 'Pai',
    menuChartTreemap: 'Peta Hierarki',
    menuChartSunburst: 'Semburan Matahari',
    menuChartHistogram: 'Histogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Kotak dan Kumis',
    menuChartScatter: 'X Y (Sebar)',
    menuChartWaterfall: 'Air Terjun',
    menuChartFunnel: 'Corong',
    menuChartStock: 'Saham',
    menuChartSurface: 'Permukaan',
    menuChartRadar: 'Radar',
    menuIcons: 'Ikon…',
    menuBookmark: 'Bookmark…',
    menuCrossReference: 'Referensi Silang…',
    menuBreak: 'Hentian',
    menuColumnBreak: 'Hentian Kolom',
    menuSectionNextPage: 'Hentian Bagian (Halaman Berikutnya)',
    menuSectionContinuous: 'Hentian Bagian (Berkelanjutan)',
    menuSectionOddPage: 'Hentian Bagian (Halaman Ganjil)',
    menuSectionEvenPage: 'Hentian Bagian (Halaman Genap)',
    menuFootnote: 'Catatan Kaki…',
    menuCaption: 'Keterangan…',
    menuIndexAndTables: 'Indeks dan Tabel…',
    menuWatermark: 'Cap Air…',
    menuPageNumbers: 'Nomor Halaman…',
    menuTextBox: 'Kotak Teks',
    menuVerticalTextBox: 'Gambar Kotak Teks Vertikal',
    menuAutoText: 'TeksOtomatis',
    menuAutoTextDialog: 'TeksOtomatis…',
    menuAutoTextNew: 'Baru…',
    menuDateAndTime: 'Tanggal dan Waktu…',
    menuObject: 'Objek…',
    menuInsertFile: 'File…',
    menuField: 'Bidang…',
    menuAdvancedSymbol: 'Simbol Tingkat Lanjut…',
    menuComment: 'Komentar',
    menuFormat: 'Format',
    menuBold: 'Tebal',
    menuItalic: 'Miring',
    menuUnderline: 'Garis Bawah',
    menuAlign: 'Perataan',
    menuAlignLeft: 'Rata Kiri',
    menuAlignCenter: 'Rata Tengah',
    menuAlignRight: 'Rata Kanan',
    menuAlignJustify: 'Rata Kiri Kanan',
    menuFont: 'Font…',
    menuParagraph: 'Paragraf…',
    menuTextEffects: 'Efek Teks…',
    menuDocument: 'Dokumen…',
    menuBulletsAndNumbering: 'Butir dan Penomoran…',
    menuBordersAndShading: 'Batas dan Bayangan…',
    menuColumns: 'Kolom…',
    menuTabs: 'Tab…',
    menuDropCap: 'Drop Cap…',
    menuTextDirection: 'Arah Teks…',
    menuChangeCase: 'Ubah Huruf Besar/Kecil…',
    menuAsianLayout: 'Tata Letak Asia',
    menuPhoneticGuide: 'Panduan Fonetik…',
    menuEncloseCharacters: 'Lingkupi Karakter…',
    menuHorizontalInVertical: 'Horizontal dalam Vertikal…',
    menuCombineCharacters: 'Gabungkan Karakter…',
    menuTwoLinesInOne: 'Dua Baris dalam Satu…',
    menuFitText: 'Paskan Teks…',
    menuStyle: 'Gaya…',
    menuFontSubstitution: 'Substitusi Font…',
    menuFormatObject: 'Format Objek…',
    menuEquationOptions: 'Opsi Persamaan…',
    menuSelectionPane: 'Panel Pilihan…',
    menuTools: 'Alat',
    menuTable: 'Tabel',
    menuTableInsert: 'Sisipkan',
    menuTableInsertTable: 'Tabel…',
    menuTableColsLeft: 'Kolom di Kiri',
    menuTableColsRight: 'Kolom di Kanan',
    menuTableRowsAbove: 'Baris di Atas',
    menuTableRowsBelow: 'Baris di Bawah',
    menuTableCells: 'Sel…',
    menuTableDelete: 'Hapus',
    menuTableDeleteTable: 'Tabel',
    menuTableDeleteColumns: 'Kolom',
    menuTableDeleteRows: 'Baris',
    menuTableSelect: 'Pilih',
    menuTableSelectCell: 'Sel',
    menuTableSelectColumn: 'Kolom',
    menuTableSelectRow: 'Baris',
    menuTableSelectTable: 'Tabel',
    menuTableMergeCells: 'Gabungkan Sel',
    menuTableSplitCells: 'Pisahkan Sel…',
    menuTableSplitTable: 'Pisahkan Tabel',
    menuTableAutoFit: 'Paskan Otomatis',
    menuTableAutoFitContents: 'Paskan ke Konten',
    menuTableAutoFitWindow: 'Paskan ke Jendela',
    menuTableFixedWidth: 'Lebar Kolom Tetap',
    menuTableDistributeRows: 'Distribusikan Baris Merata',
    menuTableDistributeColumns: 'Distribusikan Kolom Merata',
    menuTableRepeatHeader: 'Ulangi Baris Header',
    menuTableConvert: 'Konversi',
    menuTableConvertTextToTable: 'Konversi Teks ke Tabel…',
    menuTableConvertTableToText: 'Konversi Tabel ke Teks…',
    menuTableSort: 'Urutkan…',
    menuTableFormula: 'Rumus…',
    menuTableGridlines: 'Lihat Garis Kisi',
    menuTableProperties: 'Properti Tabel…',
    menuWordCount: 'Hitungan Kata…',
    menuAutoCorrect: 'Opsi Koreksi Otomatis…',
    menuSpellingAndGrammar: 'Ejaan dan Tata Bahasa',
    menuSpellingDialog: 'Ejaan dan Tata Bahasa…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Sembunyikan Kesalahan Ejaan',
    menuHideGrammarErrors: 'Sembunyikan Kesalahan Tata Bahasa',
    menuResetIgnoredWords: 'Atur Ulang Kata dan Tata Bahasa yang Diabaikan',
    menuThesaurus: 'Tesaurus…',
    menuHyphenation: 'Pemenggalan Kata…',
    menuTranslate: 'Terjemahkan',
    menuTranslateSelection: 'Terjemahkan Pilihan',
    menuTranslateDocument: 'Terjemahkan Dokumen',
    menuSetTranslationLanguage: 'Atur Bahasa Terjemahan Dokumen…',
    menuLanguage: 'Bahasa…',
    menuTrackChanges: 'Lacak Perubahan',
    menuHighlightChanges: 'Sorot Perubahan…',
    menuCompareDocuments: 'Bandingkan Dokumen…',
    menuCombineDocuments: 'Gabungkan Dokumen…',
    menuProtectDocument: 'Proteksi Dokumen…',
    menuEnvelopes: 'Amplop…',
    menuLabels: 'Label…',
    menuCustomizeKeyboard: 'Kustomisasi Keyboard…',
    menuPreferences: 'Preferensi…',
    menuAiProofread: 'Koreksi AI',
    menuWindow: 'Jendela',
    menuHelp: 'Bantuan',
    menuShortcuts: 'Pintasan Papan Ketik',
  },
  ru: {
    dlgOpenDoc: 'Открыть документ',
    filterWord: 'Документы Word',
    odtImportMsg: 'Преобразовать этот документ .{ext}?',
    odtImportDetail:
      'GenOffice редактирует .docx, а не ODF. Текст, заголовки, списки, таблицы и изображения переносятся; сноски, примечания и надписи сохраняются не полностью. Исходный файл никогда не изменяется; преобразованный .docx — отдельная копия, в которую записывается дальнейшее сохранение.',
    hwpImportDetail:
      'GenOffice редактирует .docx, а не HWP. Текст, таблицы, изображения и колонтитулы переносятся; формулы, сноски и внедрённые объекты не преобразуются, а автоматическая нумерация становится обычным текстом. Исходный файл никогда не изменяется; преобразованный .docx — отдельная копия, в которую записывается дальнейшее сохранение.',
    odtImportConvertBtn: 'Преобразовать и открыть',
    importSkipConfirm: 'В следующий раз конвертировать без запроса',
    docImportMsg: 'Преобразовать этот документ .{ext}?',
    docImportDetail:
      'GenOffice редактирует .docx, а не документы Word 97-2003 (.doc). Текст, заголовки, списки, таблицы и изображения переносятся; сноски, примечания и надписи сохраняются не полностью. Исходный файл никогда не изменяется; преобразованный .docx — отдельная копия, в которую записывается дальнейшее сохранение.',
    docImportConvertBtn: 'Преобразовать и открыть',
    dlgSaveAs: 'Сохранить как',
    closeUnsavedMsg: 'В этом документе есть несохранённые изменения.',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    closeNoReplyMsg: 'Документ не отвечает; возможно, есть несохранённые изменения.',
    closeNoReplyDetail: 'Всё равно закрыть? Несохранённые изменения будут потеряны.',
    btnCloseAnyway: 'Закрыть всё равно',
    autosaveFoundTitle: 'Найдена восстановленная версия',
    autosaveDiskNewer:
      'Внимание: файл на диске был изменён после сохранения этой копии восстановления. Восстановление заменит эти более новые изменения.',
    autosaveFoundBody:
      'Есть несохранённые изменения из прошлого сеанса. Восстановить автосохранённую версию?',
    autosaveRestore: 'Восстановить',
    autosaveDiscard: 'Отклонить',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
    extModifiedMsg: 'Файл был изменён другой программой.',
    extModifiedDetail: 'Всё равно сохранить и перезаписать изменения на диске?',
    btnOverwrite: 'Перезаписать',
    dlgInsertImage: 'Вставить рисунок',
    filterImages: 'Изображения',
    dlgMergeDataSource: 'Выбрать список получателей',
    filterMergeDataSource: 'Excel или CSV',
    dlgAddAttachment: 'Добавить вложения',
    filterSupported: 'Поддерживаемые файлы',
    filterAll: 'Все файлы',
    dlgExportPdf: 'Экспорт в PDF',
    dlgExportHtml: 'Экспорт в HTML',
    dlgExportMarkdown: 'Экспорт в Markdown',
    dlgPickExportDir: 'Выбор папки для экспорта',
    errUnsupportedExt: 'файлы .{ext} не поддерживаются',
    errNotFile: 'не является файлом',
    errTooLarge: 'превышает лимит {mb} МБ',
    errTemplateProtected: 'этот шаблон защищён паролем; откройте его через «Файл > Открыть»',
    errImageTooLarge: 'изображение превышает лимит 5 МБ',
    errUnreadable: 'не удается прочитать',
    errFileTooLarge: 'Файл превышает максимальный размер',
    errParseFailed: 'Не удалось разобрать файл',
    errImageNoText:
      'Вложенные изображения не содержат текста; изображение отправляется вместе с сообщением пользователя, смотрите его напрямую',
    errNotImage: 'неподдерживаемый тип изображения',
    errGskNotLoggedIn:
      'Вы не вошли в Genspark: нажмите «Войти в Genspark» ниже, войдите и повторите попытку',
    errNoApiKey: 'API-ключ для {provider} не настроен',
    errAiRegion:
      'Поставщик этой модели не обслуживает ваш текущий регион — выберите другую модель или смените сеть',
    errAiBusy: 'Сервис ИИ сейчас перегружен — повторите попытку чуть позже',
    errNoModel: 'Не указано имя модели',
    menuFile: 'Файл',
    menuNewDoc: 'Создать документ',
    menuNewWindow: 'Новое окно',
    menuOpen: 'Открыть…',
    menuOpenRecent: 'Открыть последние',
    menuNoRecent: 'Нет последних документов',
    menuRecentMore: 'Другие…',
    menuClose: 'Закрыть',
    menuSave: 'Сохранить',
    menuSaveAs: 'Сохранить как…',
    menuPageSetup: 'Параметры страницы…',
    menuExportPdf: 'Экспорт в PDF…',
    menuExportHtml: 'Экспорт в HTML…',
    menuExportMarkdown: 'Экспорт в Markdown…',
    menuExportImages: 'Экспорт в изображения…',
    menuPrint: 'Печать…',
    menuEdit: 'Правка',
    menuUndo: 'Отменить',
    menuRedo: 'Повторить',
    menuRepeat: 'Повторить команду',
    menuCantRepeat: 'Нельзя повторить',
    menuCut: 'Вырезать',
    menuCopy: 'Копировать',
    menuPaste: 'Вставить',
    menuPasteMatch: 'Вставить с согласованием форматирования',
    menuFindReplace: 'Найти и заменить…',
    menuPasteSpecial: 'Специальная вставка…',
    menuClear: 'Очистить',
    menuClearFormatting: 'Очистить формат',
    menuClearContents: 'Содержимое',
    menuFindMenu: 'Найти',
    menuFind: 'Найти…',
    menuFindNext: 'Найти далее',
    menuFindPrevious: 'Найти предыдущее',
    menuReplace: 'Заменить…',
    menuAdvancedFind: 'Расширенный поиск и замена…',
    menuLinks: 'Связи…',
    menuSelectData: 'Выбрать данные…',
    menuToggleDrawing: 'Переключить рисование',
    menuNewFromTemplate: 'Создать из шаблона…',
    menuSaveAsTemplate: 'Сохранить как шаблон…',
    menuMove: 'Переместить…',
    menuRename: 'Переименовать…',
    menuAlwaysOpenReadOnly: 'Всегда открывать только для чтения',
    menuReduceFileSize: 'Уменьшить размер файла…',
    menuProperties: 'Свойства…',
    menuGoTo: 'Перейти…',
    menuSelectAll: 'Выделить все',
    menuView: 'Вид',
    menuZoom: 'Масштаб',
    menuZoomIn: 'Увеличить',
    menuZoomOut: 'Уменьшить',
    menuZoom100: 'Фактический размер (100%)',
    menuPrintLayout: 'Разметка страницы',
    menuWebLayout: 'Веб-документ',
    menuOutline: 'Структура',
    menuDraft: 'Черновик',
    menuFocus: 'Фокусировка',
    menuRibbon: 'Лента',
    menuSidebar: 'Боковая панель',
    menuThumbnail: 'Эскизы',
    menuNavigation: 'Навигация',
    menuReviewing: 'Рецензирование',
    menuSidebarSearch: 'Поиск',
    menuRuler: 'Линейка',
    menuHeaderFooter: 'Колонтитулы',
    menuFootnotes: 'Сноски',
    menuMarkup: 'Исправления',
    menuRevealFormatting: 'Показать форматирование',
    menuZoomDialog: 'Масштаб…',
    menuGridlines: 'Сетка',
    menuPageWidth: 'По ширине страницы',
    menuWholePage: 'Страница целиком',
    menuMultiPage: 'Несколько страниц',
    menuAiSidebar: 'Боковая панель ИИ',
    menuDarkMode: 'Темный режим',
    menuFullscreen: 'Перейти в полноэкранный режим',
    menuInsert: 'Вставка',
    menuInsertTable: 'Таблица…',
    menuInsertImage: 'Рисунок…',
    menuInsertPageBreak: 'Разрыв страницы',
    menuInsertLink: 'Гиперссылка…',
    menuInsertEquation: 'Уравнение…',
    menuPictures: 'Рисунки',
    menuPictureFromFile: 'Рисунок из файла…',
    menuShape: 'Фигура',
    menuShapeRectangle: 'Прямоугольник',
    menuShapeRoundedRectangle: 'Скругленный прямоугольник',
    menuShapeTriangle: 'Треугольник',
    menuShapeOval: 'Овал',
    menuShapeLine: 'Линия',
    menuChart: 'Диаграмма',
    menuChartColumn: 'Гистограмма',
    menuChartBar: 'Линейчатая',
    menuChartLine: 'График',
    menuChartArea: 'С областями',
    menuChartPie: 'Круговая',
    menuChartTreemap: 'Дерево',
    menuChartSunburst: 'Солнечные лучи',
    menuChartHistogram: 'Гистограмма частот',
    menuChartPareto: 'Парето',
    menuChartBoxWhisker: 'Ящик с усами',
    menuChartScatter: 'Точечная (X Y)',
    menuChartWaterfall: 'Каскадная',
    menuChartFunnel: 'Воронка',
    menuChartStock: 'Биржевая',
    menuChartSurface: 'Поверхность',
    menuChartRadar: 'Лепестковая',
    menuIcons: 'Значки…',
    menuBookmark: 'Закладка…',
    menuCrossReference: 'Перекрестная ссылка…',
    menuBreak: 'Разрыв',
    menuColumnBreak: 'Разрыв столбца',
    menuSectionNextPage: 'Разрыв раздела (со следующей страницы)',
    menuSectionContinuous: 'Разрыв раздела (на текущей странице)',
    menuSectionOddPage: 'Разрыв раздела (с нечетной страницы)',
    menuSectionEvenPage: 'Разрыв раздела (с четной страницы)',
    menuFootnote: 'Сноска…',
    menuCaption: 'Название…',
    menuIndexAndTables: 'Оглавление и указатели…',
    menuWatermark: 'Подложка…',
    menuPageNumbers: 'Номера страниц…',
    menuTextBox: 'Надпись',
    menuVerticalTextBox: 'Нарисовать вертикальную надпись',
    menuAutoText: 'Автотекст',
    menuAutoTextDialog: 'Автотекст…',
    menuAutoTextNew: 'Создать…',
    menuDateAndTime: 'Дата и время…',
    menuObject: 'Объект…',
    menuInsertFile: 'Файл…',
    menuField: 'Поле…',
    menuAdvancedSymbol: 'Дополнительные символы…',
    menuComment: 'Примечание',
    menuFormat: 'Формат',
    menuBold: 'Полужирный',
    menuItalic: 'Курсив',
    menuUnderline: 'Подчеркнутый',
    menuAlign: 'Выравнивание',
    menuAlignLeft: 'По левому краю',
    menuAlignCenter: 'По центру',
    menuAlignRight: 'По правому краю',
    menuAlignJustify: 'По ширине',
    menuFont: 'Шрифт…',
    menuParagraph: 'Абзац…',
    menuTextEffects: 'Текстовые эффекты…',
    menuDocument: 'Документ…',
    menuBulletsAndNumbering: 'Маркеры и нумерация…',
    menuBordersAndShading: 'Границы и заливка…',
    menuColumns: 'Колонки…',
    menuTabs: 'Табуляция…',
    menuDropCap: 'Буквица…',
    menuTextDirection: 'Направление текста…',
    menuChangeCase: 'Регистр…',
    menuAsianLayout: 'Азиатский макет',
    menuPhoneticGuide: 'Фонетическое руководство…',
    menuEncloseCharacters: 'Обвести знаки…',
    menuHorizontalInVertical: 'Горизонтально в вертикальном…',
    menuCombineCharacters: 'Объединить знаки…',
    menuTwoLinesInOne: 'Две строки в одной…',
    menuFitText: 'Вписать текст…',
    menuStyle: 'Стиль…',
    menuFontSubstitution: 'Подстановка шрифтов…',
    menuFormatObject: 'Формат объекта…',
    menuEquationOptions: 'Параметры уравнения…',
    menuSelectionPane: 'Область выделения…',
    menuTools: 'Сервис',
    menuTable: 'Таблица',
    menuTableInsert: 'Вставить',
    menuTableInsertTable: 'Таблица…',
    menuTableColsLeft: 'Столбцы слева',
    menuTableColsRight: 'Столбцы справа',
    menuTableRowsAbove: 'Строки выше',
    menuTableRowsBelow: 'Строки ниже',
    menuTableCells: 'Ячейки…',
    menuTableDelete: 'Удалить',
    menuTableDeleteTable: 'Таблицу',
    menuTableDeleteColumns: 'Столбцы',
    menuTableDeleteRows: 'Строки',
    menuTableSelect: 'Выделить',
    menuTableSelectCell: 'Ячейку',
    menuTableSelectColumn: 'Столбец',
    menuTableSelectRow: 'Строку',
    menuTableSelectTable: 'Таблицу',
    menuTableMergeCells: 'Объединить ячейки',
    menuTableSplitCells: 'Разделить ячейки…',
    menuTableSplitTable: 'Разделить таблицу',
    menuTableAutoFit: 'Автоподбор',
    menuTableAutoFitContents: 'По содержимому',
    menuTableAutoFitWindow: 'По ширине окна',
    menuTableFixedWidth: 'Фиксированная ширина столбца',
    menuTableDistributeRows: 'Выровнять высоту строк',
    menuTableDistributeColumns: 'Выровнять ширину столбцов',
    menuTableRepeatHeader: 'Повторять строки заголовков',
    menuTableConvert: 'Преобразовать',
    menuTableConvertTextToTable: 'Преобразовать текст в таблицу…',
    menuTableConvertTableToText: 'Преобразовать таблицу в текст…',
    menuTableSort: 'Сортировка…',
    menuTableFormula: 'Формула…',
    menuTableGridlines: 'Отображать сетку',
    menuTableProperties: 'Свойства таблицы…',
    menuWordCount: 'Статистика…',
    menuAutoCorrect: 'Параметры автозамены…',
    menuSpellingAndGrammar: 'Правописание',
    menuSpellingDialog: 'Правописание…',
    menuEditor: 'Редактор',
    menuHideSpellingErrors: 'Скрыть орфографические ошибки',
    menuHideGrammarErrors: 'Скрыть грамматические ошибки',
    menuResetIgnoredWords: 'Сбросить пропущенные слова и грамматику',
    menuThesaurus: 'Тезаурус…',
    menuHyphenation: 'Расстановка переносов…',
    menuTranslate: 'Перевод',
    menuTranslateSelection: 'Перевести выделенный фрагмент',
    menuTranslateDocument: 'Перевести документ',
    menuSetTranslationLanguage: 'Задать язык перевода документа…',
    menuLanguage: 'Язык…',
    menuTrackChanges: 'Исправления',
    menuHighlightChanges: 'Выделить исправления…',
    menuCompareDocuments: 'Сравнить документы…',
    menuCombineDocuments: 'Объединить документы…',
    menuProtectDocument: 'Защитить документ…',
    menuEnvelopes: 'Конверты…',
    menuLabels: 'Наклейки…',
    menuCustomizeKeyboard: 'Настройка клавиатуры…',
    menuPreferences: 'Параметры…',
    menuAiProofread: 'ИИ-корректура',
    menuWindow: 'Окно',
    menuHelp: 'Справка',
    menuShortcuts: 'Сочетания клавиш',
  },
  ar: {
    dlgOpenDoc: 'فتح مستند',
    filterWord: 'مستندات Word',
    odtImportMsg: 'تحويل مستند .{ext} هذا؟',
    odtImportDetail:
      'يحرّر GenOffice ملفات ‎.docx ولا يحرّر ODF. يُحوَّل النص والعناوين والقوائم والجداول والصور؛ أما الحواشي السفلية والتعليقات ومربعات النص فلا تُحفظ بالكامل. لا يتغيّر الملف الأصلي أبدًا؛ والملف ‎.docx المحوَّل نسخة مستقلة تُكتب فيها عمليات الحفظ اللاحقة.',
    hwpImportDetail:
      'يحرّر GenOffice ملفات ‎.docx ولا يحرّر HWP. يُحوَّل النص والجداول والصور والرؤوس والتذييلات؛ أما المعادلات والحواشي السفلية والكائنات المضمّنة فلا تُحوَّل، ويصبح الترقيم التلقائي نصًا عاديًا. لا يتغيّر الملف الأصلي أبدًا؛ والملف ‎.docx المحوَّل نسخة مستقلة تُكتب فيها عمليات الحفظ اللاحقة.',
    odtImportConvertBtn: 'تحويل وفتح',
    importSkipConfirm: 'في المرة القادمة، حوّل دون سؤال',
    docImportMsg: 'تحويل مستند .{ext} هذا؟',
    docImportDetail:
      'يحرّر GenOffice ملفات ‎.docx ولا يحرّر مستندات Word 97-2003 (‎.doc). يُحوَّل النص والعناوين والقوائم والجداول والصور؛ أما الحواشي السفلية والتعليقات ومربعات النص فلا تُحفظ بالكامل. لا يتغيّر الملف الأصلي أبدًا؛ والملف ‎.docx المحوَّل نسخة مستقلة تُكتب فيها عمليات الحفظ اللاحقة.',
    docImportConvertBtn: 'تحويل وفتح',
    dlgSaveAs: 'حفظ باسم',
    closeUnsavedMsg: 'يحتوي هذا المستند على تغييرات غير محفوظة.',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    closeNoReplyMsg: 'المستند لا يستجيب وقد يحتوي على تغييرات غير محفوظة.',
    closeNoReplyDetail: 'هل تريد الإغلاق رغم ذلك؟ ستفقد التغييرات غير المحفوظة.',
    btnCloseAnyway: 'إغلاق على أي حال',
    autosaveFoundTitle: 'تم العثور على نسخة مستردة',
    autosaveDiskNewer:
      'ملاحظة: تم تعديل الملف على القرص بعد حفظ نسخة الاسترداد هذه. ستؤدي الاستعادة إلى استبدال تلك التغييرات الأحدث.',
    autosaveFoundBody:
      'توجد تغييرات غير محفوظة من الجلسة الأخيرة. هل تريد استعادة النسخة المحفوظة تلقائيًا؟',
    autosaveRestore: 'استعادة',
    autosaveDiscard: 'تجاهل',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
    extModifiedMsg: 'تم تعديل الملف بواسطة برنامج آخر.',
    extModifiedDetail: 'هل تريد الحفظ على أي حال والكتابة فوق التغييرات على القرص؟',
    btnOverwrite: 'استبدال',
    dlgInsertImage: 'إدراج صورة',
    filterImages: 'الصور',
    dlgMergeDataSource: 'اختيار قائمة المستلمين',
    filterMergeDataSource: 'Excel أو CSV',
    dlgAddAttachment: 'إضافة مرفقات',
    filterSupported: 'الملفات المدعومة',
    filterAll: 'كل الملفات',
    dlgExportPdf: 'تصدير بتنسيق PDF',
    dlgExportHtml: 'تصدير بتنسيق HTML',
    dlgExportMarkdown: 'تصدير بتنسيق Markdown',
    dlgPickExportDir: 'اختيار مجلد التصدير',
    errUnsupportedExt: 'ملفات .{ext} غير مدعومة',
    errNotFile: 'ليس ملفًا',
    errTooLarge: 'يتجاوز الحد {mb}MB',
    errTemplateProtected: 'هذا القالب محمي بكلمة مرور؛ افتحه من ملف > فتح',
    errImageTooLarge: 'الصورة تتجاوز حد 5MB',
    errUnreadable: 'تعذرت القراءة',
    errFileTooLarge: 'الملف يتجاوز الحد الأقصى للحجم',
    errParseFailed: 'فشل تحليل الملف',
    errImageNoText:
      'مرفقات الصور لا توفر نصًا؛ تُرسل الصورة مع رسالة المستخدم ويمكن الاطلاع عليها مباشرة',
    errNotImage: 'ليس نوع صورة مدعومًا',
    errGskNotLoggedIn:
      'لم تسجّل الدخول إلى Genspark: انقر على «تسجيل الدخول إلى Genspark» أدناه ثم أعد المحاولة',
    errNoApiKey: 'لم يتم تكوين مفتاح API لـ {provider}',
    errAiRegion:
      'مزوّد هذا النموذج لا يقدم الخدمة في منطقتك الحالية — اختر نموذجًا آخر أو غيّر الشبكة',
    errAiBusy: 'خدمة الذكاء الاصطناعي مشغولة حاليًا — يرجى المحاولة مرة أخرى بعد قليل',
    errNoModel: 'لم يتم تكوين اسم النموذج',
    menuFile: 'ملف',
    menuNewDoc: 'مستند جديد',
    menuNewWindow: 'نافذة جديدة',
    menuOpen: 'فتح…',
    menuOpenRecent: 'فتح الأخيرة',
    menuNoRecent: 'لا توجد مستندات أخيرة',
    menuRecentMore: 'المزيد…',
    menuClose: 'إغلاق',
    menuSave: 'حفظ',
    menuSaveAs: 'حفظ باسم…',
    menuPageSetup: 'إعداد الصفحة…',
    menuExportPdf: 'تصدير بتنسيق PDF…',
    menuExportHtml: 'تصدير بتنسيق HTML…',
    menuExportMarkdown: 'تصدير بتنسيق Markdown…',
    menuExportImages: 'تصدير كصور…',
    menuPrint: 'طباعة…',
    menuEdit: 'تحرير',
    menuUndo: 'تراجع',
    menuRedo: 'إعادة',
    menuRepeat: 'تكرار',
    menuCantRepeat: 'يتعذر التكرار',
    menuCut: 'قص',
    menuCopy: 'نسخ',
    menuPaste: 'لصق',
    menuPasteMatch: 'لصق ومطابقة التنسيق',
    menuFindReplace: 'بحث واستبدال…',
    menuPasteSpecial: 'لصق خاص…',
    menuClear: 'مسح',
    menuClearFormatting: 'مسح التنسيق',
    menuClearContents: 'المحتويات',
    menuFindMenu: 'بحث',
    menuFind: 'بحث…',
    menuFindNext: 'بحث عن التالي',
    menuFindPrevious: 'بحث عن السابق',
    menuReplace: 'استبدال…',
    menuAdvancedFind: 'بحث واستبدال متقدم…',
    menuLinks: 'الارتباطات…',
    menuSelectData: 'تحديد البيانات…',
    menuToggleDrawing: 'تبديل الرسم',
    menuNewFromTemplate: 'جديد من قالب…',
    menuSaveAsTemplate: 'حفظ كقالب…',
    menuMove: 'نقل…',
    menuRename: 'إعادة تسمية…',
    menuAlwaysOpenReadOnly: 'الفتح دائماً للقراءة فقط',
    menuReduceFileSize: 'تقليل حجم الملف…',
    menuProperties: 'خصائص…',
    menuGoTo: 'الانتقال إلى…',
    menuSelectAll: 'تحديد الكل',
    menuView: 'عرض',
    menuZoom: 'تكبير/تصغير',
    menuZoomIn: 'تكبير',
    menuZoomOut: 'تصغير',
    menuZoom100: 'الحجم الفعلي (100%)',
    menuPrintLayout: 'تخطيط الطباعة',
    menuWebLayout: 'تخطيط ويب',
    menuOutline: 'مخطط تفصيلي',
    menuDraft: 'مسودة',
    menuFocus: 'التركيز',
    menuRibbon: 'الشريط',
    menuSidebar: 'الشريط الجانبي',
    menuThumbnail: 'صورة مصغرة',
    menuNavigation: 'التنقل',
    menuReviewing: 'المراجعة',
    menuSidebarSearch: 'بحث',
    menuRuler: 'المسطرة',
    menuHeaderFooter: 'رأس وتذييل الصفحة',
    menuFootnotes: 'الحواشي السفلية',
    menuMarkup: 'العلامات',
    menuRevealFormatting: 'إظهار التنسيق',
    menuZoomDialog: 'تكبير/تصغير…',
    menuGridlines: 'خطوط الشبكة',
    menuPageWidth: 'عرض الصفحة',
    menuWholePage: 'صفحة كاملة',
    menuMultiPage: 'صفحات متعددة',
    menuAiSidebar: 'الشريط الجانبي للذكاء الاصطناعي',
    menuDarkMode: 'الوضع الداكن',
    menuFullscreen: 'الدخول إلى ملء الشاشة',
    menuInsert: 'إدراج',
    menuInsertTable: 'جدول…',
    menuInsertImage: 'صورة…',
    menuInsertPageBreak: 'فاصل صفحات',
    menuInsertLink: 'ارتباط تشعبي…',
    menuInsertEquation: 'معادلة…',
    menuPictures: 'الصور',
    menuPictureFromFile: 'صورة من ملف…',
    menuShape: 'شكل',
    menuShapeRectangle: 'مستطيل',
    menuShapeRoundedRectangle: 'مستطيل مستدير الزوايا',
    menuShapeTriangle: 'مثلث',
    menuShapeOval: 'شكل بيضاوي',
    menuShapeLine: 'خط',
    menuChart: 'مخطط',
    menuChartColumn: 'عمودي',
    menuChartBar: 'شريطي',
    menuChartLine: 'خطي',
    menuChartArea: 'مساحي',
    menuChartPie: 'دائري',
    menuChartTreemap: 'مخطط هيكلي',
    menuChartSunburst: 'مخطط الشمس المشرقة',
    menuChartHistogram: 'مدرج تكراري',
    menuChartPareto: 'باريتو',
    menuChartBoxWhisker: 'مخطط الصندوق',
    menuChartScatter: 'س ص (مبعثر)',
    menuChartWaterfall: 'شلالي',
    menuChartFunnel: 'قمعي',
    menuChartStock: 'أسهم',
    menuChartSurface: 'سطحي',
    menuChartRadar: 'نسيجي',
    menuIcons: 'أيقونات…',
    menuBookmark: 'إشارة مرجعية…',
    menuCrossReference: 'إسناد ترافقي…',
    menuBreak: 'فاصل',
    menuColumnBreak: 'فاصل أعمدة',
    menuSectionNextPage: 'فاصل مقطعي (الصفحة التالية)',
    menuSectionContinuous: 'فاصل مقطعي (مستمر)',
    menuSectionOddPage: 'فاصل مقطعي (صفحة فردية)',
    menuSectionEvenPage: 'فاصل مقطعي (صفحة زوجية)',
    menuFootnote: 'حاشية سفلية…',
    menuCaption: 'تسمية توضيحية…',
    menuIndexAndTables: 'فهارس وجداول…',
    menuWatermark: 'علامة مائية…',
    menuPageNumbers: 'أرقام الصفحات…',
    menuTextBox: 'مربع نص',
    menuVerticalTextBox: 'رسم مربع نص عمودي',
    menuAutoText: 'نص تلقائي',
    menuAutoTextDialog: 'نص تلقائي…',
    menuAutoTextNew: 'جديد…',
    menuDateAndTime: 'التاريخ والوقت…',
    menuObject: 'كائن…',
    menuInsertFile: 'ملف…',
    menuField: 'حقل…',
    menuAdvancedSymbol: 'رمز متقدم…',
    menuComment: 'تعليق',
    menuFormat: 'تنسيق',
    menuBold: 'غامق',
    menuItalic: 'مائل',
    menuUnderline: 'تسطير',
    menuAlign: 'محاذاة',
    menuAlignLeft: 'محاذاة إلى اليسار',
    menuAlignCenter: 'توسيط',
    menuAlignRight: 'محاذاة إلى اليمين',
    menuAlignJustify: 'ضبط',
    menuFont: 'الخط…',
    menuParagraph: 'فقرة…',
    menuTextEffects: 'تأثيرات النص…',
    menuDocument: 'المستند…',
    menuBulletsAndNumbering: 'تعداد نقطي ورقمي…',
    menuBordersAndShading: 'حدود وتظليل…',
    menuColumns: 'أعمدة…',
    menuTabs: 'علامات الجدولة…',
    menuDropCap: 'إسقاط الأحرف الاستهلالية…',
    menuTextDirection: 'اتجاه النص…',
    menuChangeCase: 'تغيير حالة الأحرف…',
    menuAsianLayout: 'تخطيط آسيوي',
    menuPhoneticGuide: 'دليل صوتي…',
    menuEncloseCharacters: 'إحاطة الأحرف…',
    menuHorizontalInVertical: 'أفقي داخل عمودي…',
    menuCombineCharacters: 'دمج الأحرف…',
    menuTwoLinesInOne: 'سطران في سطر واحد…',
    menuFitText: 'احتواء النص…',
    menuStyle: 'نمط…',
    menuFontSubstitution: 'استبدال الخطوط…',
    menuFormatObject: 'تنسيق الكائن…',
    menuEquationOptions: 'خيارات المعادلة…',
    menuSelectionPane: 'جزء التحديد…',
    menuTools: 'أدوات',
    menuTable: 'جدول',
    menuTableInsert: 'إدراج',
    menuTableInsertTable: 'جدول…',
    menuTableColsLeft: 'أعمدة إلى اليسار',
    menuTableColsRight: 'أعمدة إلى اليمين',
    menuTableRowsAbove: 'صفوف أعلى',
    menuTableRowsBelow: 'صفوف أسفل',
    menuTableCells: 'خلايا…',
    menuTableDelete: 'حذف',
    menuTableDeleteTable: 'الجدول',
    menuTableDeleteColumns: 'الأعمدة',
    menuTableDeleteRows: 'الصفوف',
    menuTableSelect: 'تحديد',
    menuTableSelectCell: 'الخلية',
    menuTableSelectColumn: 'العمود',
    menuTableSelectRow: 'الصف',
    menuTableSelectTable: 'الجدول',
    menuTableMergeCells: 'دمج الخلايا',
    menuTableSplitCells: 'تقسيم الخلايا…',
    menuTableSplitTable: 'تقسيم الجدول',
    menuTableAutoFit: 'احتواء تلقائي',
    menuTableAutoFitContents: 'احتواء تلقائي للمحتويات',
    menuTableAutoFitWindow: 'احتواء تلقائي للنافذة',
    menuTableFixedWidth: 'عرض عمود ثابت',
    menuTableDistributeRows: 'توزيع الصفوف بالتساوي',
    menuTableDistributeColumns: 'توزيع الأعمدة بالتساوي',
    menuTableRepeatHeader: 'تكرار صفوف الرأس',
    menuTableConvert: 'تحويل',
    menuTableConvertTextToTable: 'تحويل النص إلى جدول…',
    menuTableConvertTableToText: 'تحويل الجدول إلى نص…',
    menuTableSort: 'فرز…',
    menuTableFormula: 'صيغة…',
    menuTableGridlines: 'عرض خطوط الشبكة',
    menuTableProperties: 'خصائص الجدول…',
    menuWordCount: 'عدد الكلمات…',
    menuAutoCorrect: 'خيارات التصحيح التلقائي…',
    menuSpellingAndGrammar: 'التدقيق الإملائي والنحوي',
    menuSpellingDialog: 'التدقيق الإملائي والنحوي…',
    menuEditor: 'المحرر',
    menuHideSpellingErrors: 'إخفاء الأخطاء الإملائية',
    menuHideGrammarErrors: 'إخفاء الأخطاء النحوية',
    menuResetIgnoredWords: 'إعادة تعيين الكلمات والقواعد التي تم تجاهلها',
    menuThesaurus: 'قاموس المرادفات…',
    menuHyphenation: 'الواصلة…',
    menuTranslate: 'ترجمة',
    menuTranslateSelection: 'ترجمة التحديد',
    menuTranslateDocument: 'ترجمة المستند',
    menuSetTranslationLanguage: 'تعيين لغة ترجمة المستند…',
    menuLanguage: 'اللغة…',
    menuTrackChanges: 'تعقب التغييرات',
    menuHighlightChanges: 'تمييز التغييرات…',
    menuCompareDocuments: 'مقارنة المستندات…',
    menuCombineDocuments: 'دمج المستندات…',
    menuProtectDocument: 'حماية المستند…',
    menuEnvelopes: 'مغلفات…',
    menuLabels: 'تسميات…',
    menuCustomizeKeyboard: 'تخصيص لوحة المفاتيح…',
    menuPreferences: 'التفضيلات…',
    menuAiProofread: 'تدقيق بالذكاء الاصطناعي',
    menuWindow: 'نافذة',
    menuHelp: 'تعليمات',
    menuShortcuts: 'اختصارات لوحة المفاتيح',
  },
  pt: {
    dlgOpenDoc: 'Abrir Documento',
    filterWord: 'Documentos do Word',
    odtImportMsg: 'Converter este documento .{ext}?',
    odtImportDetail:
      'O GenOffice edita .docx, não ODF. Texto, títulos, listas, tabelas e imagens são convertidos; notas de rodapé, comentários e caixas de texto não são preservados por completo. O arquivo original nunca é alterado; o .docx convertido é uma cópia independente em que os salvamentos seguintes são gravados.',
    hwpImportDetail:
      'O GenOffice edita .docx, não HWP. Texto, tabelas, imagens, cabeçalhos e rodapés são convertidos; equações, notas de rodapé e objetos incorporados não são, e a numeração automática vira texto comum. O arquivo original nunca é alterado; o .docx convertido é uma cópia independente em que os salvamentos seguintes são gravados.',
    odtImportConvertBtn: 'Converter e abrir',
    importSkipConfirm: 'Da próxima vez, converter sem perguntar',
    docImportMsg: 'Converter este documento .{ext}?',
    docImportDetail:
      'O GenOffice edita .docx, não documentos do Word 97-2003 (.doc). Texto, títulos, listas, tabelas e imagens são convertidos; notas de rodapé, comentários e caixas de texto não são preservados por completo. O arquivo original nunca é alterado; o .docx convertido é uma cópia independente em que os salvamentos seguintes são gravados.',
    docImportConvertBtn: 'Converter e abrir',
    dlgSaveAs: 'Salvar Como',
    closeUnsavedMsg: 'Este documento tem alterações não salvas.',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    closeNoReplyMsg: 'O documento não está respondendo e pode ter alterações não salvas.',
    closeNoReplyDetail: 'Fechar mesmo assim? As alterações não salvas serão perdidas.',
    btnCloseAnyway: 'Fechar mesmo assim',
    autosaveFoundTitle: 'Versão recuperada encontrada',
    autosaveDiskNewer:
      'Nota: o arquivo no disco foi modificado depois que esta cópia de recuperação foi salva. Restaurar substituirá essas alterações mais recentes.',
    autosaveFoundBody:
      'Há alterações não salvas da sua última sessão. Restaurar a versão salva automaticamente?',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
    extModifiedMsg: 'O arquivo foi modificado por outro programa.',
    extModifiedDetail: 'Salvar mesmo assim e sobrescrever as alterações no disco?',
    btnOverwrite: 'Sobrescrever',
    dlgInsertImage: 'Inserir Imagem',
    filterImages: 'Imagens',
    dlgMergeDataSource: 'Selecionar lista de destinatários',
    filterMergeDataSource: 'Excel ou CSV',
    dlgAddAttachment: 'Adicionar Anexos',
    filterSupported: 'Arquivos Compatíveis',
    filterAll: 'Todos os Arquivos',
    dlgExportPdf: 'Exportar como PDF',
    dlgExportHtml: 'Exportar como HTML',
    dlgExportMarkdown: 'Exportar como Markdown',
    dlgPickExportDir: 'Escolher Pasta de Exportação',
    errUnsupportedExt: 'arquivos .{ext} não são suportados',
    errNotFile: 'não é um arquivo',
    errTooLarge: 'excede o limite de {mb}MB',
    errTemplateProtected: 'este modelo está protegido por senha; abra-o com Arquivo > Abrir',
    errImageTooLarge: 'a imagem excede o limite de 5MB',
    errUnreadable: 'não é possível ler',
    errFileTooLarge: 'O arquivo excede o limite de tamanho',
    errParseFailed: 'Falha ao analisar o arquivo',
    errImageNoText:
      'Anexos de imagem não fornecem texto; a imagem é enviada junto com a mensagem do usuário, basta vê-la diretamente',
    errNotImage: 'não é um tipo de imagem suportado',
    errGskNotLoggedIn:
      'Não conectado ao Genspark: clique em “Entrar no Genspark” abaixo, entre e tente novamente',
    errNoApiKey: 'Nenhuma chave de API configurada para {provider}',
    errAiRegion:
      'O fornecedor deste modelo não atende à sua região atual — escolha outro modelo ou mude de rede',
    errAiBusy: 'O serviço de IA está sobrecarregado no momento — tente novamente em instantes',
    errNoModel: 'Nenhum nome de modelo configurado',
    menuFile: 'Arquivo',
    menuNewDoc: 'Novo Documento',
    menuNewWindow: 'Nova Janela',
    menuOpen: 'Abrir…',
    menuOpenRecent: 'Abrir Recente',
    menuNoRecent: 'Nenhum Documento Recente',
    menuRecentMore: 'Mais…',
    menuClose: 'Fechar',
    menuSave: 'Salvar',
    menuSaveAs: 'Salvar Como…',
    menuPageSetup: 'Configurar Página…',
    menuExportPdf: 'Exportar como PDF…',
    menuExportHtml: 'Exportar como HTML…',
    menuExportMarkdown: 'Exportar como Markdown…',
    menuExportImages: 'Exportar como imagens…',
    menuPrint: 'Imprimir…',
    menuEdit: 'Editar',
    menuUndo: 'Desfazer',
    menuRedo: 'Refazer',
    menuRepeat: 'Repetir',
    menuCantRepeat: 'Não é Possível Repetir',
    menuCut: 'Recortar',
    menuCopy: 'Copiar',
    menuPaste: 'Colar',
    menuPasteMatch: 'Colar com a Mesma Formatação',
    menuFindReplace: 'Localizar e Substituir…',
    menuPasteSpecial: 'Colar Especial…',
    menuClear: 'Limpar',
    menuClearFormatting: 'Limpar Formatação',
    menuClearContents: 'Conteúdo',
    menuFindMenu: 'Localizar',
    menuFind: 'Localizar…',
    menuFindNext: 'Localizar Próxima',
    menuFindPrevious: 'Localizar Anterior',
    menuReplace: 'Substituir…',
    menuAdvancedFind: 'Localização Avançada e Substituição…',
    menuLinks: 'Vínculos…',
    menuSelectData: 'Selecionar Dados…',
    menuToggleDrawing: 'Alternar Desenho',
    menuNewFromTemplate: 'Novo a partir de Modelo…',
    menuSaveAsTemplate: 'Salvar como Modelo…',
    menuMove: 'Mover…',
    menuRename: 'Renomear…',
    menuAlwaysOpenReadOnly: 'Sempre Abrir como Somente Leitura',
    menuReduceFileSize: 'Reduzir Tamanho do Arquivo…',
    menuProperties: 'Propriedades…',
    menuGoTo: 'Ir para…',
    menuSelectAll: 'Selecionar Tudo',
    menuView: 'Exibir',
    menuZoom: 'Zoom',
    menuZoomIn: 'Ampliar',
    menuZoomOut: 'Reduzir',
    menuZoom100: 'Tamanho Real (100%)',
    menuPrintLayout: 'Layout de Impressão',
    menuWebLayout: 'Layout da Web',
    menuOutline: 'Estrutura de Tópicos',
    menuDraft: 'Rascunho',
    menuFocus: 'Foco',
    menuRibbon: 'Faixa de Opções',
    menuSidebar: 'Barra Lateral',
    menuThumbnail: 'Miniatura',
    menuNavigation: 'Navegação',
    menuReviewing: 'Revisão',
    menuSidebarSearch: 'Pesquisar',
    menuRuler: 'Régua',
    menuHeaderFooter: 'Cabeçalho e Rodapé',
    menuFootnotes: 'Notas de Rodapé',
    menuMarkup: 'Marcação',
    menuRevealFormatting: 'Revelar Formatação',
    menuZoomDialog: 'Zoom…',
    menuGridlines: 'Linhas de Grade',
    menuPageWidth: 'Largura da Página',
    menuWholePage: 'Página Inteira',
    menuMultiPage: 'Várias páginas',
    menuAiSidebar: 'Barra Lateral de IA',
    menuDarkMode: 'Modo Escuro',
    menuFullscreen: 'Entrar em Tela Cheia',
    menuInsert: 'Inserir',
    menuInsertTable: 'Tabela…',
    menuInsertImage: 'Imagem…',
    menuInsertPageBreak: 'Quebra de Página',
    menuInsertLink: 'Hiperlink…',
    menuInsertEquation: 'Equação…',
    menuPictures: 'Imagens',
    menuPictureFromFile: 'Imagem do Arquivo…',
    menuShape: 'Forma',
    menuShapeRectangle: 'Retângulo',
    menuShapeRoundedRectangle: 'Retângulo Arredondado',
    menuShapeTriangle: 'Triângulo',
    menuShapeOval: 'Elipse',
    menuShapeLine: 'Linha',
    menuChart: 'Gráfico',
    menuChartColumn: 'Colunas',
    menuChartBar: 'Barras',
    menuChartLine: 'Linhas',
    menuChartArea: 'Área',
    menuChartPie: 'Pizza',
    menuChartTreemap: 'Mapa de Árvore',
    menuChartSunburst: 'Explosão Solar',
    menuChartHistogram: 'Histograma',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Caixa e Caixa Estreita',
    menuChartScatter: 'X Y (Dispersão)',
    menuChartWaterfall: 'Cascata',
    menuChartFunnel: 'Funil',
    menuChartStock: 'Ações',
    menuChartSurface: 'Superfície',
    menuChartRadar: 'Radar',
    menuIcons: 'Ícones…',
    menuBookmark: 'Indicador…',
    menuCrossReference: 'Referência Cruzada…',
    menuBreak: 'Quebra',
    menuColumnBreak: 'Quebra de Coluna',
    menuSectionNextPage: 'Quebra de Seção (Próxima Página)',
    menuSectionContinuous: 'Quebra de Seção (Contínua)',
    menuSectionOddPage: 'Quebra de Seção (Página Ímpar)',
    menuSectionEvenPage: 'Quebra de Seção (Página Par)',
    menuFootnote: 'Nota de Rodapé…',
    menuCaption: 'Legenda…',
    menuIndexAndTables: 'Índices…',
    menuWatermark: 'Marca-d’água…',
    menuPageNumbers: 'Números de Página…',
    menuTextBox: 'Caixa de Texto',
    menuVerticalTextBox: 'Desenhar Caixa de Texto Vertical',
    menuAutoText: 'AutoTexto',
    menuAutoTextDialog: 'AutoTexto…',
    menuAutoTextNew: 'Novo…',
    menuDateAndTime: 'Data e Hora…',
    menuObject: 'Objeto…',
    menuInsertFile: 'Arquivo…',
    menuField: 'Campo…',
    menuAdvancedSymbol: 'Símbolo Avançado…',
    menuComment: 'Comentário',
    menuFormat: 'Formatar',
    menuBold: 'Negrito',
    menuItalic: 'Itálico',
    menuUnderline: 'Sublinhado',
    menuAlign: 'Alinhar',
    menuAlignLeft: 'Alinhar à Esquerda',
    menuAlignCenter: 'Centralizar',
    menuAlignRight: 'Alinhar à Direita',
    menuAlignJustify: 'Justificar',
    menuFont: 'Fonte…',
    menuParagraph: 'Parágrafo…',
    menuTextEffects: 'Efeitos de Texto…',
    menuDocument: 'Documento…',
    menuBulletsAndNumbering: 'Marcadores e Numeração…',
    menuBordersAndShading: 'Bordas e Sombreamento…',
    menuColumns: 'Colunas…',
    menuTabs: 'Tabulações…',
    menuDropCap: 'Capitular…',
    menuTextDirection: 'Direção do Texto…',
    menuChangeCase: 'Maiúsculas e Minúsculas…',
    menuAsianLayout: 'Layout Asiático',
    menuPhoneticGuide: 'Guia Fonético…',
    menuEncloseCharacters: 'Caracteres Circunscritos…',
    menuHorizontalInVertical: 'Horizontal em Vertical…',
    menuCombineCharacters: 'Combinar Caracteres…',
    menuTwoLinesInOne: 'Duas Linhas em Uma…',
    menuFitText: 'Ajustar Texto…',
    menuStyle: 'Estilo…',
    menuFontSubstitution: 'Substituição de Fontes…',
    menuFormatObject: 'Formatar Objeto…',
    menuEquationOptions: 'Opções de Equação…',
    menuSelectionPane: 'Painel de Seleção…',
    menuTools: 'Ferramentas',
    menuTable: 'Tabela',
    menuTableInsert: 'Inserir',
    menuTableInsertTable: 'Tabela…',
    menuTableColsLeft: 'Colunas à esquerda',
    menuTableColsRight: 'Colunas à direita',
    menuTableRowsAbove: 'Linhas acima',
    menuTableRowsBelow: 'Linhas abaixo',
    menuTableCells: 'Células…',
    menuTableDelete: 'Excluir',
    menuTableDeleteTable: 'Tabela',
    menuTableDeleteColumns: 'Colunas',
    menuTableDeleteRows: 'Linhas',
    menuTableSelect: 'Selecionar',
    menuTableSelectCell: 'Célula',
    menuTableSelectColumn: 'Coluna',
    menuTableSelectRow: 'Linha',
    menuTableSelectTable: 'Tabela',
    menuTableMergeCells: 'Mesclar células',
    menuTableSplitCells: 'Dividir células…',
    menuTableSplitTable: 'Dividir tabela',
    menuTableAutoFit: 'AutoAjuste',
    menuTableAutoFitContents: 'AutoAjustar ao conteúdo',
    menuTableAutoFitWindow: 'AutoAjustar à janela',
    menuTableFixedWidth: 'Largura de coluna fixa',
    menuTableDistributeRows: 'Distribuir linhas uniformemente',
    menuTableDistributeColumns: 'Distribuir colunas uniformemente',
    menuTableRepeatHeader: 'Repetir linhas de cabeçalho',
    menuTableConvert: 'Converter',
    menuTableConvertTextToTable: 'Converter texto em tabela…',
    menuTableConvertTableToText: 'Converter tabela em texto…',
    menuTableSort: 'Classificar…',
    menuTableFormula: 'Fórmula…',
    menuTableGridlines: 'Ver linhas de grade',
    menuTableProperties: 'Propriedades da tabela…',
    menuWordCount: 'Contagem de Palavras…',
    menuAutoCorrect: 'Opções de Correção Automática…',
    menuSpellingAndGrammar: 'Ortografia e Gramática',
    menuSpellingDialog: 'Ortografia e Gramática…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Ocultar Erros de Ortografia',
    menuHideGrammarErrors: 'Ocultar Erros Gramaticais',
    menuResetIgnoredWords: 'Redefinir Palavras e Gramática Ignoradas',
    menuThesaurus: 'Dicionário de Sinônimos…',
    menuHyphenation: 'Hifenização…',
    menuTranslate: 'Traduzir',
    menuTranslateSelection: 'Traduzir Seleção',
    menuTranslateDocument: 'Traduzir Documento',
    menuSetTranslationLanguage: 'Definir Idioma de Tradução do Documento…',
    menuLanguage: 'Idioma…',
    menuTrackChanges: 'Controlar Alterações',
    menuHighlightChanges: 'Realçar Alterações…',
    menuCompareDocuments: 'Comparar Documentos…',
    menuCombineDocuments: 'Combinar Documentos…',
    menuProtectDocument: 'Proteger Documento…',
    menuEnvelopes: 'Envelopes…',
    menuLabels: 'Etiquetas…',
    menuCustomizeKeyboard: 'Personalizar Teclado…',
    menuPreferences: 'Preferências…',
    menuAiProofread: 'Revisão com IA',
    menuWindow: 'Janela',
    menuHelp: 'Ajuda',
    menuShortcuts: 'Atalhos de Teclado',
  },
  it: {
    dlgOpenDoc: 'Apri documento',
    filterWord: 'Documenti Word',
    odtImportMsg: 'Convertire questo documento .{ext}?',
    odtImportDetail:
      'GenOffice modifica .docx, non ODF. Testo, titoli, elenchi, tabelle e immagini vengono convertiti; note a piè di pagina, commenti e caselle di testo non sono conservati completamente. Il file originale non viene mai modificato; il .docx convertito è una copia indipendente su cui vengono salvate le modifiche successive.',
    hwpImportDetail:
      'GenOffice modifica .docx, non HWP. Testo, tabelle, immagini, intestazioni e piè di pagina vengono convertiti; equazioni, note a piè di pagina e oggetti incorporati no, e la numerazione automatica diventa testo semplice. Il file originale non viene mai modificato; il .docx convertito è una copia indipendente su cui vengono salvate le modifiche successive.',
    odtImportConvertBtn: 'Converti e apri',
    importSkipConfirm: 'La prossima volta converti senza chiedere',
    docImportMsg: 'Convertire questo documento .{ext}?',
    docImportDetail:
      'GenOffice modifica .docx, non documenti Word 97-2003 (.doc). Testo, titoli, elenchi, tabelle e immagini vengono convertiti; note a piè di pagina, commenti e caselle di testo non sono conservati completamente. Il file originale non viene mai modificato; il .docx convertito è una copia indipendente su cui vengono salvate le modifiche successive.',
    docImportConvertBtn: 'Converti e apri',
    dlgSaveAs: 'Salva con nome',
    closeUnsavedMsg: 'Questo documento contiene modifiche non salvate.',
    closeUnsavedDetail: 'Salvarle prima di chiudere?',
    closeNoReplyMsg: 'Il documento non risponde e potrebbe avere modifiche non salvate.',
    closeNoReplyDetail: 'Chiudere comunque? Le modifiche non salvate andranno perse.',
    btnCloseAnyway: 'Chiudi comunque',
    autosaveFoundTitle: 'Trovata versione recuperata',
    autosaveDiskNewer:
      'Nota: il file su disco è stato modificato dopo il salvataggio di questa copia di ripristino. Il ripristino sostituirà tali modifiche più recenti.',
    autosaveFoundBody:
      "Ci sono modifiche non salvate dall'ultima sessione. Ripristinare la versione salvata automaticamente?",
    autosaveRestore: 'Ripristina',
    autosaveDiscard: 'Ignora',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
    extModifiedMsg: 'Il file è stato modificato da un altro programma.',
    extModifiedDetail: 'Salvare comunque e sovrascrivere le modifiche sul disco?',
    btnOverwrite: 'Sovrascrivi',
    dlgInsertImage: 'Inserisci immagine',
    filterImages: 'Immagini',
    dlgMergeDataSource: 'Seleziona elenco destinatari',
    filterMergeDataSource: 'Excel o CSV',
    dlgAddAttachment: 'Aggiungi allegati',
    filterSupported: 'File supportati',
    filterAll: 'Tutti i file',
    dlgExportPdf: 'Esporta come PDF',
    dlgExportHtml: 'Esporta come HTML',
    dlgExportMarkdown: 'Esporta come Markdown',
    dlgPickExportDir: 'Scegli la cartella di esportazione',
    errUnsupportedExt: 'i file .{ext} non sono supportati',
    errNotFile: 'non è un file',
    errTooLarge: 'supera il limite di {mb} MB',
    errTemplateProtected: 'questo modello è protetto da password; aprirlo con File > Apri',
    errImageTooLarge: "l'immagine supera il limite di 5 MB",
    errUnreadable: 'impossibile leggere',
    errFileTooLarge: 'Il file supera il limite di dimensione',
    errParseFailed: 'Impossibile analizzare il file',
    errImageNoText:
      "Gli allegati immagine non forniscono testo; l'immagine viene inviata insieme al messaggio dell'utente, basta guardarla direttamente",
    errNotImage: 'tipo di immagine non supportato',
    errGskNotLoggedIn:
      'Accesso a Genspark non effettuato: fai clic su “Accedi a Genspark” qui sotto, accedi e riprova',
    errNoApiKey: 'Nessuna chiave API configurata per {provider}',
    errAiRegion:
      'Il fornitore di questo modello non serve la tua regione attuale — scegli un altro modello o cambia rete',
    errAiBusy: 'Il servizio IA è momentaneamente sovraccarico — riprova tra poco',
    errNoModel: 'Nessun nome di modello configurato',
    menuFile: 'File',
    menuNewDoc: 'Nuovo documento',
    menuNewWindow: 'Nuova finestra',
    menuOpen: 'Apri…',
    menuOpenRecent: 'Apri recenti',
    menuNoRecent: 'Nessun documento recente',
    menuRecentMore: 'Altro…',
    menuClose: 'Chiudi',
    menuSave: 'Salva',
    menuSaveAs: 'Salva con nome…',
    menuPageSetup: 'Imposta pagina…',
    menuExportPdf: 'Esporta come PDF…',
    menuExportHtml: 'Esporta come HTML…',
    menuExportMarkdown: 'Esporta come Markdown…',
    menuExportImages: 'Esporta come immagini…',
    menuPrint: 'Stampa…',
    menuEdit: 'Modifica',
    menuUndo: 'Annulla',
    menuRedo: 'Ripeti',
    menuRepeat: 'Ripeti comando',
    menuCantRepeat: 'Impossibile ripetere',
    menuCut: 'Taglia',
    menuCopy: 'Copia',
    menuPaste: 'Incolla',
    menuPasteMatch: 'Incolla e adatta la formattazione',
    menuFindReplace: 'Trova e sostituisci…',
    menuPasteSpecial: 'Incolla speciale…',
    menuClear: 'Cancella',
    menuClearFormatting: 'Cancella formattazione',
    menuClearContents: 'Contenuto',
    menuFindMenu: 'Trova',
    menuFind: 'Trova…',
    menuFindNext: 'Trova successivo',
    menuFindPrevious: 'Trova precedente',
    menuReplace: 'Sostituisci…',
    menuAdvancedFind: 'Trova e sostituisci avanzato…',
    menuLinks: 'Collegamenti…',
    menuSelectData: 'Seleziona dati…',
    menuToggleDrawing: 'Attiva/disattiva disegno',
    menuNewFromTemplate: 'Nuovo da modello…',
    menuSaveAsTemplate: 'Salva come modello…',
    menuMove: 'Sposta…',
    menuRename: 'Rinomina…',
    menuAlwaysOpenReadOnly: 'Apri sempre in sola lettura',
    menuReduceFileSize: 'Riduci dimensioni file…',
    menuProperties: 'Proprietà…',
    menuGoTo: 'Vai a…',
    menuSelectAll: 'Seleziona tutto',
    menuView: 'Visualizza',
    menuZoom: 'Zoom',
    menuZoomIn: 'Ingrandisci',
    menuZoomOut: 'Riduci',
    menuZoom100: 'Dimensioni effettive (100%)',
    menuPrintLayout: 'Layout di stampa',
    menuWebLayout: 'Layout Web',
    menuOutline: 'Struttura',
    menuDraft: 'Bozza',
    menuFocus: 'Messa a fuoco',
    menuRibbon: 'Barra multifunzione',
    menuSidebar: 'Barra laterale',
    menuThumbnail: 'Anteprima',
    menuNavigation: 'Navigazione',
    menuReviewing: 'Revisione',
    menuSidebarSearch: 'Cerca',
    menuRuler: 'Righello',
    menuHeaderFooter: 'Intestazione e piè di pagina',
    menuFootnotes: 'Note a piè di pagina',
    menuMarkup: 'Commenti',
    menuRevealFormatting: 'Rivela formattazione',
    menuZoomDialog: 'Zoom…',
    menuGridlines: 'Griglia',
    menuPageWidth: 'Larghezza pagina',
    menuWholePage: 'Pagina intera',
    menuMultiPage: 'Più pagine',
    menuAiSidebar: 'Barra laterale IA',
    menuDarkMode: 'Modalità scura',
    menuFullscreen: 'Attiva schermo intero',
    menuInsert: 'Inserisci',
    menuInsertTable: 'Tabella…',
    menuInsertImage: 'Immagine…',
    menuInsertPageBreak: 'Interruzione di pagina',
    menuInsertLink: 'Collegamento ipertestuale…',
    menuInsertEquation: 'Equazione…',
    menuPictures: 'Immagini',
    menuPictureFromFile: 'Immagine da file…',
    menuShape: 'Forma',
    menuShapeRectangle: 'Rettangolo',
    menuShapeRoundedRectangle: 'Rettangolo arrotondato',
    menuShapeTriangle: 'Triangolo',
    menuShapeOval: 'Ovale',
    menuShapeLine: 'Linea',
    menuChart: 'Grafico',
    menuChartColumn: 'Istogramma',
    menuChartBar: 'Barre',
    menuChartLine: 'Linee',
    menuChartArea: 'Area',
    menuChartPie: 'Torta',
    menuChartTreemap: 'Mappa ad albero',
    menuChartSunburst: 'Radiale',
    menuChartHistogram: 'Istogramma di frequenza',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Scatola e baffi',
    menuChartScatter: 'X Y (dispersione)',
    menuChartWaterfall: 'Cascata',
    menuChartFunnel: 'Imbuto',
    menuChartStock: 'Azionario',
    menuChartSurface: 'Superficie',
    menuChartRadar: 'Radar',
    menuIcons: 'Icone…',
    menuBookmark: 'Segnalibro…',
    menuCrossReference: 'Riferimento incrociato…',
    menuBreak: 'Interruzione',
    menuColumnBreak: 'Interruzione di colonna',
    menuSectionNextPage: 'Interruzione di sezione (pagina successiva)',
    menuSectionContinuous: 'Interruzione di sezione (continua)',
    menuSectionOddPage: 'Interruzione di sezione (pagina dispari)',
    menuSectionEvenPage: 'Interruzione di sezione (pagina pari)',
    menuFootnote: 'Nota a piè di pagina…',
    menuCaption: 'Didascalia…',
    menuIndexAndTables: 'Indici e sommario…',
    menuWatermark: 'Filigrana…',
    menuPageNumbers: 'Numeri di pagina…',
    menuTextBox: 'Casella di testo',
    menuVerticalTextBox: 'Disegna casella di testo verticale',
    menuAutoText: 'Glossario',
    menuAutoTextDialog: 'Glossario…',
    menuAutoTextNew: 'Nuovo…',
    menuDateAndTime: 'Data e ora…',
    menuObject: 'Oggetto…',
    menuInsertFile: 'File…',
    menuField: 'Campo…',
    menuAdvancedSymbol: 'Simbolo avanzato…',
    menuComment: 'Commento',
    menuFormat: 'Formato',
    menuBold: 'Grassetto',
    menuItalic: 'Corsivo',
    menuUnderline: 'Sottolineato',
    menuAlign: 'Allinea',
    menuAlignLeft: 'Allinea a sinistra',
    menuAlignCenter: 'Centra',
    menuAlignRight: 'Allinea a destra',
    menuAlignJustify: 'Giustifica',
    menuFont: 'Carattere…',
    menuParagraph: 'Paragrafo…',
    menuTextEffects: 'Effetti di testo…',
    menuDocument: 'Documento…',
    menuBulletsAndNumbering: 'Elenchi puntati e numerati…',
    menuBordersAndShading: 'Bordi e sfondo…',
    menuColumns: 'Colonne…',
    menuTabs: 'Tabulazioni…',
    menuDropCap: 'Capolettera…',
    menuTextDirection: 'Orientamento testo…',
    menuChangeCase: 'Maiuscole/minuscole…',
    menuAsianLayout: 'Layout asiatico',
    menuPhoneticGuide: 'Guida fonetica…',
    menuEncloseCharacters: 'Caratteri racchiusi…',
    menuHorizontalInVertical: 'Orizzontale in verticale…',
    menuCombineCharacters: 'Combina caratteri…',
    menuTwoLinesInOne: 'Due righe in una…',
    menuFitText: 'Adatta testo…',
    menuStyle: 'Stile…',
    menuFontSubstitution: 'Sostituzione tipi di carattere…',
    menuFormatObject: 'Formato oggetto…',
    menuEquationOptions: 'Opzioni equazione…',
    menuSelectionPane: 'Riquadro di selezione…',
    menuTools: 'Strumenti',
    menuTable: 'Tabella',
    menuTableInsert: 'Inserisci',
    menuTableInsertTable: 'Tabella…',
    menuTableColsLeft: 'Colonne a sinistra',
    menuTableColsRight: 'Colonne a destra',
    menuTableRowsAbove: 'Righe sopra',
    menuTableRowsBelow: 'Righe sotto',
    menuTableCells: 'Celle…',
    menuTableDelete: 'Elimina',
    menuTableDeleteTable: 'Tabella',
    menuTableDeleteColumns: 'Colonne',
    menuTableDeleteRows: 'Righe',
    menuTableSelect: 'Seleziona',
    menuTableSelectCell: 'Cella',
    menuTableSelectColumn: 'Colonna',
    menuTableSelectRow: 'Riga',
    menuTableSelectTable: 'Tabella',
    menuTableMergeCells: 'Unisci celle',
    menuTableSplitCells: 'Dividi celle…',
    menuTableSplitTable: 'Dividi tabella',
    menuTableAutoFit: 'Adatta e distribuisci',
    menuTableAutoFitContents: 'Adatta al contenuto',
    menuTableAutoFitWindow: 'Adatta alla finestra',
    menuTableFixedWidth: 'Larghezza colonna fissa',
    menuTableDistributeRows: 'Distribuisci righe uniformemente',
    menuTableDistributeColumns: 'Distribuisci colonne uniformemente',
    menuTableRepeatHeader: 'Ripeti righe di intestazione',
    menuTableConvert: 'Converti',
    menuTableConvertTextToTable: 'Converti il testo in tabella…',
    menuTableConvertTableToText: 'Converti la tabella in testo…',
    menuTableSort: 'Ordina…',
    menuTableFormula: 'Formula…',
    menuTableGridlines: 'Mostra griglia',
    menuTableProperties: 'Proprietà tabella…',
    menuWordCount: 'Conteggio parole…',
    menuAutoCorrect: 'Opzioni correzione automatica…',
    menuSpellingAndGrammar: 'Ortografia e grammatica',
    menuSpellingDialog: 'Ortografia e grammatica…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Nascondi errori di ortografia',
    menuHideGrammarErrors: 'Nascondi errori grammaticali',
    menuResetIgnoredWords: 'Reimposta parole e grammatica ignorate',
    menuThesaurus: 'Thesaurus…',
    menuHyphenation: 'Sillabazione…',
    menuTranslate: 'Traduci',
    menuTranslateSelection: 'Traduci selezione',
    menuTranslateDocument: 'Traduci documento',
    menuSetTranslationLanguage: 'Imposta lingua di traduzione del documento…',
    menuLanguage: 'Lingua…',
    menuTrackChanges: 'Revisioni',
    menuHighlightChanges: 'Evidenzia revisioni…',
    menuCompareDocuments: 'Confronta documenti…',
    menuCombineDocuments: 'Combina documenti…',
    menuProtectDocument: 'Proteggi documento…',
    menuEnvelopes: 'Buste…',
    menuLabels: 'Etichette…',
    menuCustomizeKeyboard: 'Personalizza tastiera…',
    menuPreferences: 'Preferenze…',
    menuAiProofread: 'Correzione IA',
    menuWindow: 'Finestra',
    menuHelp: 'Aiuto',
    menuShortcuts: 'Scelte rapide da tastiera',
  },
  pl: {
    dlgOpenDoc: 'Otwórz dokument',
    filterWord: 'Dokumenty programu Word',
    odtImportMsg: 'Skonwertować ten dokument .{ext}?',
    odtImportDetail:
      'GenOffice edytuje .docx, a nie ODF. Tekst, nagłówki, listy, tabele i obrazy są konwertowane; przypisy, komentarze i pola tekstowe nie są zachowywane w całości. Plik oryginalny nigdy nie jest zmieniany; skonwertowany .docx to niezależna kopia, do której trafiają kolejne zapisy.',
    hwpImportDetail:
      'GenOffice edytuje .docx, a nie HWP. Tekst, tabele, obrazy, nagłówki i stopki są konwertowane; równania, przypisy i obiekty osadzone nie są, a automatyczna numeracja staje się zwykłym tekstem. Plik oryginalny nigdy nie jest zmieniany; skonwertowany .docx to niezależna kopia, do której trafiają kolejne zapisy.',
    odtImportConvertBtn: 'Konwertuj i otwórz',
    importSkipConfirm: 'Następnym razem konwertuj bez pytania',
    docImportMsg: 'Skonwertować ten dokument .{ext}?',
    docImportDetail:
      'GenOffice edytuje .docx, a nie dokumenty Word 97-2003 (.doc). Tekst, nagłówki, listy, tabele i obrazy są konwertowane; przypisy, komentarze i pola tekstowe nie są zachowywane w całości. Plik oryginalny nigdy nie jest zmieniany; skonwertowany .docx to niezależna kopia, do której trafiają kolejne zapisy.',
    docImportConvertBtn: 'Konwertuj i otwórz',
    dlgSaveAs: 'Zapisz jako',
    closeUnsavedMsg: 'Ten dokument zawiera niezapisane zmiany.',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    closeNoReplyMsg: 'Dokument nie odpowiada i może zawierać niezapisane zmiany.',
    closeNoReplyDetail: 'Zamknąć mimo to? Niezapisane zmiany zostaną utracone.',
    btnCloseAnyway: 'Zamknij mimo to',
    autosaveFoundTitle: 'Znaleziono odzyskaną wersję',
    autosaveDiskNewer:
      'Uwaga: plik na dysku został zmodyfikowany po zapisaniu tej kopii odzyskiwania. Przywrócenie zastąpi te nowsze zmiany.',
    autosaveFoundBody:
      'Istnieją niezapisane zmiany z ostatniej sesji. Przywrócić wersję zapisaną automatycznie?',
    autosaveRestore: 'Przywróć',
    autosaveDiscard: 'Odrzuć',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
    extModifiedMsg: 'Plik został zmodyfikowany przez inny program.',
    extModifiedDetail: 'Zapisać mimo to i nadpisać zmiany na dysku?',
    btnOverwrite: 'Nadpisz',
    dlgInsertImage: 'Wstaw obraz',
    filterImages: 'Obrazy',
    dlgMergeDataSource: 'Wybierz listę odbiorców',
    filterMergeDataSource: 'Excel lub CSV',
    dlgAddAttachment: 'Dodaj załączniki',
    filterSupported: 'Obsługiwane pliki',
    filterAll: 'Wszystkie pliki',
    dlgExportPdf: 'Eksportuj jako PDF',
    dlgExportHtml: 'Eksportuj jako HTML',
    dlgExportMarkdown: 'Eksportuj jako Markdown',
    dlgPickExportDir: 'Wybierz folder eksportu',
    errUnsupportedExt: 'pliki .{ext} nie są obsługiwane',
    errNotFile: 'to nie jest plik',
    errTooLarge: 'przekracza limit {mb} MB',
    errTemplateProtected: 'ten szablon jest chroniony hasłem; otwórz go przez Plik > Otwórz',
    errImageTooLarge: 'obraz przekracza limit 5 MB',
    errUnreadable: 'nie można odczytać',
    errFileTooLarge: 'Plik przekracza limit rozmiaru',
    errParseFailed: 'Nie udało się przeanalizować pliku',
    errImageNoText:
      'Załączniki graficzne nie zawierają tekstu; obraz jest wysyłany razem z wiadomością użytkownika, wystarczy na niego spojrzeć',
    errNotImage: 'nieobsługiwany typ obrazu',
    errGskNotLoggedIn:
      'Nie zalogowano do Genspark: kliknij „Zaloguj się do Genspark” poniżej, zaloguj się i spróbuj ponownie',
    errNoApiKey: 'Nie skonfigurowano klucza API dla {provider}',
    errAiRegion:
      'Dostawca tego modelu nie obsługuje Twojego obecnego regionu — wybierz inny model lub zmień sieć',
    errAiBusy: 'Usługa AI jest obecnie przeciążona — spróbuj ponownie za chwilę',
    errNoModel: 'Nie skonfigurowano nazwy modelu',
    menuFile: 'Plik',
    menuNewDoc: 'Nowy dokument',
    menuNewWindow: 'Nowe okno',
    menuOpen: 'Otwórz…',
    menuOpenRecent: 'Otwórz ostatnie',
    menuNoRecent: 'Brak ostatnich dokumentów',
    menuRecentMore: 'Więcej…',
    menuClose: 'Zamknij',
    menuSave: 'Zapisz',
    menuSaveAs: 'Zapisz jako…',
    menuPageSetup: 'Ustawienia strony…',
    menuExportPdf: 'Eksportuj jako PDF…',
    menuExportHtml: 'Eksportuj jako HTML…',
    menuExportMarkdown: 'Eksportuj jako Markdown…',
    menuExportImages: 'Eksportuj jako obrazy…',
    menuPrint: 'Drukuj…',
    menuEdit: 'Edycja',
    menuUndo: 'Cofnij',
    menuRedo: 'Ponów',
    menuRepeat: 'Powtórz',
    menuCantRepeat: 'Nie można powtórzyć',
    menuCut: 'Wytnij',
    menuCopy: 'Kopiuj',
    menuPaste: 'Wklej',
    menuPasteMatch: 'Wklej i dopasuj formatowanie',
    menuFindReplace: 'Znajdź i zamień…',
    menuPasteSpecial: 'Wklej specjalnie…',
    menuClear: 'Wyczyść',
    menuClearFormatting: 'Wyczyść formatowanie',
    menuClearContents: 'Zawartość',
    menuFindMenu: 'Znajdź',
    menuFind: 'Znajdź…',
    menuFindNext: 'Znajdź następny',
    menuFindPrevious: 'Znajdź poprzedni',
    menuReplace: 'Zamień…',
    menuAdvancedFind: 'Zaawansowane znajdowanie i zamienianie…',
    menuLinks: 'Łącza…',
    menuSelectData: 'Zaznacz dane…',
    menuToggleDrawing: 'Przełącz rysowanie',
    menuNewFromTemplate: 'Nowy z szablonu…',
    menuSaveAsTemplate: 'Zapisz jako szablon…',
    menuMove: 'Przenieś…',
    menuRename: 'Zmień nazwę…',
    menuAlwaysOpenReadOnly: 'Zawsze otwieraj tylko do odczytu',
    menuReduceFileSize: 'Zmniejsz rozmiar pliku…',
    menuProperties: 'Właściwości…',
    menuGoTo: 'Przejdź do…',
    menuSelectAll: 'Zaznacz wszystko',
    menuView: 'Widok',
    menuZoom: 'Powiększenie',
    menuZoomIn: 'Powiększ',
    menuZoomOut: 'Pomniejsz',
    menuZoom100: 'Rzeczywisty rozmiar (100%)',
    menuPrintLayout: 'Układ wydruku',
    menuWebLayout: 'Układ sieci Web',
    menuOutline: 'Konspekt',
    menuDraft: 'Wersja robocza',
    menuFocus: 'Fokus',
    menuRibbon: 'Wstążka',
    menuSidebar: 'Pasek boczny',
    menuThumbnail: 'Miniatura',
    menuNavigation: 'Nawigacja',
    menuReviewing: 'Recenzja',
    menuSidebarSearch: 'Wyszukaj',
    menuRuler: 'Linijka',
    menuHeaderFooter: 'Nagłówek i stopka',
    menuFootnotes: 'Przypisy dolne',
    menuMarkup: 'Adiustacja',
    menuRevealFormatting: 'Pokaż formatowanie',
    menuZoomDialog: 'Powiększenie…',
    menuGridlines: 'Linie siatki',
    menuPageWidth: 'Szerokość strony',
    menuWholePage: 'Cała strona',
    menuMultiPage: 'Wiele stron',
    menuAiSidebar: 'Pasek boczny AI',
    menuDarkMode: 'Tryb ciemny',
    menuFullscreen: 'Przejdź do pełnego ekranu',
    menuInsert: 'Wstaw',
    menuInsertTable: 'Tabela…',
    menuInsertImage: 'Obraz…',
    menuInsertPageBreak: 'Podział strony',
    menuInsertLink: 'Hiperłącze…',
    menuInsertEquation: 'Równanie…',
    menuPictures: 'Obrazy',
    menuPictureFromFile: 'Obraz z pliku…',
    menuShape: 'Kształt',
    menuShapeRectangle: 'Prostokąt',
    menuShapeRoundedRectangle: 'Prostokąt zaokrąglony',
    menuShapeTriangle: 'Trójkąt',
    menuShapeOval: 'Owal',
    menuShapeLine: 'Linia',
    menuChart: 'Wykres',
    menuChartColumn: 'Kolumnowy',
    menuChartBar: 'Słupkowy',
    menuChartLine: 'Liniowy',
    menuChartArea: 'Warstwowy',
    menuChartPie: 'Kołowy',
    menuChartTreemap: 'Mapa drzewa',
    menuChartSunburst: 'Pierścieniowy rozbłysk',
    menuChartHistogram: 'Histogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Skrzynka i wąsy',
    menuChartScatter: 'X Y (punktowy)',
    menuChartWaterfall: 'Kaskadowy',
    menuChartFunnel: 'Lejkowy',
    menuChartStock: 'Giełdowy',
    menuChartSurface: 'Powierzchniowy',
    menuChartRadar: 'Radarowy',
    menuIcons: 'Ikony…',
    menuBookmark: 'Zakładka…',
    menuCrossReference: 'Odsyłacz…',
    menuBreak: 'Podział',
    menuColumnBreak: 'Podział kolumny',
    menuSectionNextPage: 'Podział sekcji (następna strona)',
    menuSectionContinuous: 'Podział sekcji (ciągły)',
    menuSectionOddPage: 'Podział sekcji (strona nieparzysta)',
    menuSectionEvenPage: 'Podział sekcji (strona parzysta)',
    menuFootnote: 'Przypis dolny…',
    menuCaption: 'Podpis…',
    menuIndexAndTables: 'Indeks i spisy…',
    menuWatermark: 'Znak wodny…',
    menuPageNumbers: 'Numery stron…',
    menuTextBox: 'Pole tekstowe',
    menuVerticalTextBox: 'Rysuj pionowe pole tekstowe',
    menuAutoText: 'Autotekst',
    menuAutoTextDialog: 'Autotekst…',
    menuAutoTextNew: 'Nowy…',
    menuDateAndTime: 'Data i godzina…',
    menuObject: 'Obiekt…',
    menuInsertFile: 'Plik…',
    menuField: 'Pole…',
    menuAdvancedSymbol: 'Symbol zaawansowany…',
    menuComment: 'Komentarz',
    menuFormat: 'Formatowanie',
    menuBold: 'Pogrubienie',
    menuItalic: 'Kursywa',
    menuUnderline: 'Podkreślenie',
    menuAlign: 'Wyrównaj',
    menuAlignLeft: 'Wyrównaj do lewej',
    menuAlignCenter: 'Wyśrodkuj',
    menuAlignRight: 'Wyrównaj do prawej',
    menuAlignJustify: 'Wyjustuj',
    menuFont: 'Czcionka…',
    menuParagraph: 'Akapit…',
    menuTextEffects: 'Efekty tekstowe…',
    menuDocument: 'Dokument…',
    menuBulletsAndNumbering: 'Punktory i numeracja…',
    menuBordersAndShading: 'Obramowanie i cieniowanie…',
    menuColumns: 'Kolumny…',
    menuTabs: 'Tabulatory…',
    menuDropCap: 'Inicjał…',
    menuTextDirection: 'Kierunek tekstu…',
    menuChangeCase: 'Zmień wielkość liter…',
    menuAsianLayout: 'Układ azjatycki',
    menuPhoneticGuide: 'Przewodnik fonetyczny…',
    menuEncloseCharacters: 'Otocz znaki…',
    menuHorizontalInVertical: 'Poziomo w pionie…',
    menuCombineCharacters: 'Połącz znaki…',
    menuTwoLinesInOne: 'Dwa wiersze w jednym…',
    menuFitText: 'Dopasuj tekst…',
    menuStyle: 'Styl…',
    menuFontSubstitution: 'Zastępowanie czcionek…',
    menuFormatObject: 'Formatuj obiekt…',
    menuEquationOptions: 'Opcje równań…',
    menuSelectionPane: 'Okienko zaznaczenia…',
    menuTools: 'Narzędzia',
    menuTable: 'Tabela',
    menuTableInsert: 'Wstaw',
    menuTableInsertTable: 'Tabela…',
    menuTableColsLeft: 'Kolumny z lewej',
    menuTableColsRight: 'Kolumny z prawej',
    menuTableRowsAbove: 'Wiersze powyżej',
    menuTableRowsBelow: 'Wiersze poniżej',
    menuTableCells: 'Komórki…',
    menuTableDelete: 'Usuń',
    menuTableDeleteTable: 'Tabela',
    menuTableDeleteColumns: 'Kolumny',
    menuTableDeleteRows: 'Wiersze',
    menuTableSelect: 'Zaznacz',
    menuTableSelectCell: 'Komórka',
    menuTableSelectColumn: 'Kolumna',
    menuTableSelectRow: 'Wiersz',
    menuTableSelectTable: 'Tabela',
    menuTableMergeCells: 'Scal komórki',
    menuTableSplitCells: 'Podziel komórki…',
    menuTableSplitTable: 'Podziel tabelę',
    menuTableAutoFit: 'Autodopasowanie',
    menuTableAutoFitContents: 'Dopasuj do zawartości',
    menuTableAutoFitWindow: 'Dopasuj do okna',
    menuTableFixedWidth: 'Stała szerokość kolumny',
    menuTableDistributeRows: 'Rozłóż wiersze równomiernie',
    menuTableDistributeColumns: 'Rozłóż kolumny równomiernie',
    menuTableRepeatHeader: 'Powtórz wiersze nagłówka',
    menuTableConvert: 'Konwertuj',
    menuTableConvertTextToTable: 'Konwertuj tekst na tabelę…',
    menuTableConvertTableToText: 'Konwertuj tabelę na tekst…',
    menuTableSort: 'Sortuj…',
    menuTableFormula: 'Formuła…',
    menuTableGridlines: 'Wyświetl linie siatki',
    menuTableProperties: 'Właściwości tabeli…',
    menuWordCount: 'Statystyka wyrazów…',
    menuAutoCorrect: 'Opcje Autokorekty…',
    menuSpellingAndGrammar: 'Pisownia i gramatyka',
    menuSpellingDialog: 'Pisownia i gramatyka…',
    menuEditor: 'Redaktor',
    menuHideSpellingErrors: 'Ukryj błędy pisowni',
    menuHideGrammarErrors: 'Ukryj błędy gramatyczne',
    menuResetIgnoredWords: 'Resetuj zignorowane wyrazy i gramatykę',
    menuThesaurus: 'Tezaurus…',
    menuHyphenation: 'Dzielenie wyrazów…',
    menuTranslate: 'Tłumacz',
    menuTranslateSelection: 'Tłumacz zaznaczenie',
    menuTranslateDocument: 'Tłumacz dokument',
    menuSetTranslationLanguage: 'Ustaw język tłumaczenia dokumentu…',
    menuLanguage: 'Język…',
    menuTrackChanges: 'Śledź zmiany',
    menuHighlightChanges: 'Wyróżnij zmiany…',
    menuCompareDocuments: 'Porównaj dokumenty…',
    menuCombineDocuments: 'Połącz dokumenty…',
    menuProtectDocument: 'Chroń dokument…',
    menuEnvelopes: 'Koperty…',
    menuLabels: 'Etykiety…',
    menuCustomizeKeyboard: 'Dostosuj klawiaturę…',
    menuPreferences: 'Preferencje…',
    menuAiProofread: 'Korekta AI',
    menuWindow: 'Okno',
    menuHelp: 'Pomoc',
    menuShortcuts: 'Skróty klawiaturowe',
  },
  cs: {
    dlgOpenDoc: 'Otevřít dokument',
    filterWord: 'Dokumenty Wordu',
    odtImportMsg: 'Převést tento dokument .{ext}?',
    odtImportDetail:
      'GenOffice upravuje .docx, nikoli ODF. Text, nadpisy, seznamy, tabulky a obrázky se převedou; poznámky pod čarou, komentáře a textová pole se nepřevezmou úplně. Původní soubor se nikdy nemění; převedený .docx je samostatná kopie, do níž se dalším ukládáním zapisuje.',
    hwpImportDetail:
      'GenOffice upravuje .docx, nikoli HWP. Text, tabulky, obrázky, záhlaví a zápatí se převedou; rovnice, poznámky pod čarou a vložené objekty se nepřevedou a automatické číslování se změní na prostý text. Původní soubor se nikdy nemění; převedený .docx je samostatná kopie, do níž se dalším ukládáním zapisuje.',
    odtImportConvertBtn: 'Převést a otevřít',
    importSkipConfirm: 'Příště převést bez ptaní',
    docImportMsg: 'Převést tento dokument .{ext}?',
    docImportDetail:
      'GenOffice upravuje .docx, nikoli dokumenty Word 97-2003 (.doc). Text, nadpisy, seznamy, tabulky a obrázky se převedou; poznámky pod čarou, komentáře a textová pole se nepřevezmou úplně. Původní soubor se nikdy nemění; převedený .docx je samostatná kopie, do níž se dalším ukládáním zapisuje.',
    docImportConvertBtn: 'Převést a otevřít',
    dlgSaveAs: 'Uložit jako',
    closeUnsavedMsg: 'Tento dokument obsahuje neuložené změny.',
    closeUnsavedDetail: 'Chcete je před zavřením uložit?',
    closeNoReplyMsg: 'Dokument neodpovídá a může obsahovat neuložené změny.',
    closeNoReplyDetail: 'Přesto zavřít? Neuložené změny budou ztraceny.',
    btnCloseAnyway: 'Přesto zavřít',
    autosaveFoundTitle: 'Nalezena obnovená verze',
    autosaveDiskNewer:
      'Poznámka: soubor na disku byl změněn po uložení této kopie pro obnovení. Obnovení tyto novější změny nahradí.',
    autosaveFoundBody:
      'Z poslední relace existují neuložené změny. Obnovit automaticky uloženou verzi?',
    autosaveRestore: 'Obnovit',
    autosaveDiscard: 'Zahodit',
    btnDontSave: 'Neukládat',
    btnCancel: 'Zrušit',
    extModifiedMsg: 'Soubor byl změněn jiným programem.',
    extModifiedDetail: 'Přesto uložit a přepsat změny na disku?',
    btnOverwrite: 'Přepsat',
    dlgInsertImage: 'Vložit obrázek',
    filterImages: 'Obrázky',
    dlgMergeDataSource: 'Vybrat seznam příjemců',
    filterMergeDataSource: 'Excel nebo CSV',
    dlgAddAttachment: 'Přidat přílohy',
    filterSupported: 'Podporované soubory',
    filterAll: 'Všechny soubory',
    dlgExportPdf: 'Exportovat jako PDF',
    dlgExportHtml: 'Exportovat jako HTML',
    dlgExportMarkdown: 'Exportovat jako Markdown',
    dlgPickExportDir: 'Zvolte složku pro export',
    errUnsupportedExt: 'soubory .{ext} nejsou podporovány',
    errNotFile: 'není soubor',
    errTooLarge: 'překračuje limit {mb} MB',
    errTemplateProtected: 'tato šablona je chráněna heslem; otevřete ji přes Soubor > Otevřít',
    errImageTooLarge: 'obrázek překračuje limit 5 MB',
    errUnreadable: 'nelze přečíst',
    errFileTooLarge: 'Soubor překračuje limit velikosti',
    errParseFailed: 'Soubor se nepodařilo zpracovat',
    errImageNoText:
      'Obrázkové přílohy neobsahují text; obrázek se odesílá spolu se zprávou uživatele',
    errNotImage: 'nepodporovaný typ obrázku',
    errGskNotLoggedIn:
      'Nejste přihlášeni do Genspark: klikněte níže na „Přihlásit se do Genspark“, přihlaste se a zkuste to znovu',
    errNoApiKey: 'Pro {provider} není nakonfigurován žádný klíč API',
    errAiRegion:
      'Poskytovatel tohoto modelu neposkytuje služby ve vaší aktuální oblasti — vyberte jiný model nebo změňte síť',
    errAiBusy: 'Služba AI je právě zaneprázdněna — zkuste to prosím za chvíli znovu',
    errNoModel: 'Není nakonfigurován název modelu',
    menuFile: 'Soubor',
    menuNewDoc: 'Nový dokument',
    menuNewWindow: 'Nové okno',
    menuOpen: 'Otevřít…',
    menuOpenRecent: 'Otevřít poslední',
    menuNoRecent: 'Žádné poslední dokumenty',
    menuRecentMore: 'Další…',
    menuClose: 'Zavřít',
    menuSave: 'Uložit',
    menuSaveAs: 'Uložit jako…',
    menuPageSetup: 'Vzhled stránky…',
    menuExportPdf: 'Exportovat jako PDF…',
    menuExportHtml: 'Exportovat jako HTML…',
    menuExportMarkdown: 'Exportovat jako Markdown…',
    menuExportImages: 'Exportovat jako obrázky…',
    menuPrint: 'Tisk…',
    menuEdit: 'Úpravy',
    menuUndo: 'Zpět',
    menuRedo: 'Znovu',
    menuRepeat: 'Opakovat',
    menuCantRepeat: 'Nelze opakovat',
    menuCut: 'Vyjmout',
    menuCopy: 'Kopírovat',
    menuPaste: 'Vložit',
    menuPasteMatch: 'Vložit a přizpůsobit formátování',
    menuFindReplace: 'Najít a nahradit…',
    menuPasteSpecial: 'Vložit jinak…',
    menuClear: 'Vymazat',
    menuClearFormatting: 'Vymazat formátování',
    menuClearContents: 'Obsah',
    menuFindMenu: 'Najít',
    menuFind: 'Najít…',
    menuFindNext: 'Najít další',
    menuFindPrevious: 'Najít předchozí',
    menuReplace: 'Nahradit…',
    menuAdvancedFind: 'Rozšířené hledání a nahrazování…',
    menuLinks: 'Propojení…',
    menuSelectData: 'Vybrat data…',
    menuToggleDrawing: 'Přepnout kreslení',
    menuNewFromTemplate: 'Nový ze šablony…',
    menuSaveAsTemplate: 'Uložit jako šablonu…',
    menuMove: 'Přesunout…',
    menuRename: 'Přejmenovat…',
    menuAlwaysOpenReadOnly: 'Vždy otevírat jen pro čtení',
    menuReduceFileSize: 'Zmenšit velikost souboru…',
    menuProperties: 'Vlastnosti…',
    menuGoTo: 'Přejít na…',
    menuSelectAll: 'Vybrat vše',
    menuView: 'Zobrazení',
    menuZoom: 'Lupa',
    menuZoomIn: 'Zvětšit',
    menuZoomOut: 'Zmenšit',
    menuZoom100: 'Skutečná velikost (100 %)',
    menuPrintLayout: 'Rozložení při tisku',
    menuWebLayout: 'Rozložení webové stránky',
    menuOutline: 'Osnova',
    menuDraft: 'Koncept',
    menuFocus: 'Fokus',
    menuRibbon: 'Pás karet',
    menuSidebar: 'Boční panel',
    menuThumbnail: 'Miniatura',
    menuNavigation: 'Navigace',
    menuReviewing: 'Revize',
    menuSidebarSearch: 'Hledat',
    menuRuler: 'Pravítko',
    menuHeaderFooter: 'Záhlaví a zápatí',
    menuFootnotes: 'Poznámky pod čarou',
    menuMarkup: 'Revize a komentáře',
    menuRevealFormatting: 'Zobrazit formátování',
    menuZoomDialog: 'Lupa…',
    menuGridlines: 'Mřížka',
    menuPageWidth: 'Šířka stránky',
    menuWholePage: 'Celá stránka',
    menuMultiPage: 'Více stránek',
    menuAiSidebar: 'Boční panel AI',
    menuDarkMode: 'Tmavý režim',
    menuFullscreen: 'Přejít na celou obrazovku',
    menuInsert: 'Vložení',
    menuInsertTable: 'Tabulka…',
    menuInsertImage: 'Obrázek…',
    menuInsertPageBreak: 'Konec stránky',
    menuInsertLink: 'Hypertextový odkaz…',
    menuInsertEquation: 'Rovnice…',
    menuPictures: 'Obrázky',
    menuPictureFromFile: 'Obrázek ze souboru…',
    menuShape: 'Obrazec',
    menuShapeRectangle: 'Obdélník',
    menuShapeRoundedRectangle: 'Zaoblený obdélník',
    menuShapeTriangle: 'Trojúhelník',
    menuShapeOval: 'Ovál',
    menuShapeLine: 'Čára',
    menuChart: 'Graf',
    menuChartColumn: 'Sloupcový',
    menuChartBar: 'Pruhový',
    menuChartLine: 'Spojnicový',
    menuChartArea: 'Plošný',
    menuChartPie: 'Výsečový',
    menuChartTreemap: 'Stromová mapa',
    menuChartSunburst: 'Vícevrstvý prstencový',
    menuChartHistogram: 'Histogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Krabicový',
    menuChartScatter: 'XY bodový',
    menuChartWaterfall: 'Vodopádový',
    menuChartFunnel: 'Trychtýřový',
    menuChartStock: 'Burzovní',
    menuChartSurface: 'Povrchový',
    menuChartRadar: 'Paprskový',
    menuIcons: 'Ikony…',
    menuBookmark: 'Záložka…',
    menuCrossReference: 'Křížový odkaz…',
    menuBreak: 'Konec',
    menuColumnBreak: 'Konec sloupce',
    menuSectionNextPage: 'Konec oddílu (další stránka)',
    menuSectionContinuous: 'Konec oddílu (průběžně)',
    menuSectionOddPage: 'Konec oddílu (lichá stránka)',
    menuSectionEvenPage: 'Konec oddílu (sudá stránka)',
    menuFootnote: 'Poznámka pod čarou…',
    menuCaption: 'Titulek…',
    menuIndexAndTables: 'Rejstřík a seznamy…',
    menuWatermark: 'Vodoznak…',
    menuPageNumbers: 'Čísla stránek…',
    menuTextBox: 'Textové pole',
    menuVerticalTextBox: 'Nakreslit svislé textové pole',
    menuAutoText: 'Automatický text',
    menuAutoTextDialog: 'Automatický text…',
    menuAutoTextNew: 'Nový…',
    menuDateAndTime: 'Datum a čas…',
    menuObject: 'Objekt…',
    menuInsertFile: 'Soubor…',
    menuField: 'Pole…',
    menuAdvancedSymbol: 'Rozšířený symbol…',
    menuComment: 'Komentář',
    menuFormat: 'Formát',
    menuBold: 'Tučné',
    menuItalic: 'Kurzíva',
    menuUnderline: 'Podtržení',
    menuAlign: 'Zarovnat',
    menuAlignLeft: 'Zarovnat vlevo',
    menuAlignCenter: 'Zarovnat na střed',
    menuAlignRight: 'Zarovnat vpravo',
    menuAlignJustify: 'Zarovnat do bloku',
    menuFont: 'Písmo…',
    menuParagraph: 'Odstavec…',
    menuTextEffects: 'Textové efekty…',
    menuDocument: 'Dokument…',
    menuBulletsAndNumbering: 'Odrážky a číslování…',
    menuBordersAndShading: 'Ohraničení a stínování…',
    menuColumns: 'Sloupce…',
    menuTabs: 'Tabulátory…',
    menuDropCap: 'Iniciála…',
    menuTextDirection: 'Směr textu…',
    menuChangeCase: 'Velká písmena…',
    menuAsianLayout: 'Asijské rozložení',
    menuPhoneticGuide: 'Fonetický průvodce…',
    menuEncloseCharacters: 'Znaky v kroužku…',
    menuHorizontalInVertical: 'Vodorovně ve svislém…',
    menuCombineCharacters: 'Kombinovat znaky…',
    menuTwoLinesInOne: 'Dva řádky v jednom…',
    menuFitText: 'Přizpůsobit text…',
    menuStyle: 'Styl…',
    menuFontSubstitution: 'Nahrazení písem…',
    menuFormatObject: 'Formát objektu…',
    menuEquationOptions: 'Možnosti rovnice…',
    menuSelectionPane: 'Podokno výběru…',
    menuTools: 'Nástroje',
    menuTable: 'Tabulka',
    menuTableInsert: 'Vložit',
    menuTableInsertTable: 'Tabulka…',
    menuTableColsLeft: 'Sloupce vlevo',
    menuTableColsRight: 'Sloupce vpravo',
    menuTableRowsAbove: 'Řádky nad',
    menuTableRowsBelow: 'Řádky pod',
    menuTableCells: 'Buňky…',
    menuTableDelete: 'Odstranit',
    menuTableDeleteTable: 'Tabulka',
    menuTableDeleteColumns: 'Sloupce',
    menuTableDeleteRows: 'Řádky',
    menuTableSelect: 'Vybrat',
    menuTableSelectCell: 'Buňka',
    menuTableSelectColumn: 'Sloupec',
    menuTableSelectRow: 'Řádek',
    menuTableSelectTable: 'Tabulka',
    menuTableMergeCells: 'Sloučit buňky',
    menuTableSplitCells: 'Rozdělit buňky…',
    menuTableSplitTable: 'Rozdělit tabulku',
    menuTableAutoFit: 'Přizpůsobit',
    menuTableAutoFitContents: 'Přizpůsobit obsahu',
    menuTableAutoFitWindow: 'Přizpůsobit oknu',
    menuTableFixedWidth: 'Pevná šířka sloupce',
    menuTableDistributeRows: 'Rozdělit řádky rovnoměrně',
    menuTableDistributeColumns: 'Rozdělit sloupce rovnoměrně',
    menuTableRepeatHeader: 'Opakovat řádky záhlaví',
    menuTableConvert: 'Převést',
    menuTableConvertTextToTable: 'Převést text na tabulku…',
    menuTableConvertTableToText: 'Převést tabulku na text…',
    menuTableSort: 'Seřadit…',
    menuTableFormula: 'Vzorec…',
    menuTableGridlines: 'Zobrazit mřížku',
    menuTableProperties: 'Vlastnosti tabulky…',
    menuWordCount: 'Počet slov…',
    menuAutoCorrect: 'Možnosti automatických oprav…',
    menuSpellingAndGrammar: 'Pravopis a gramatika',
    menuSpellingDialog: 'Pravopis a gramatika…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Skrýt pravopisné chyby',
    menuHideGrammarErrors: 'Skrýt gramatické chyby',
    menuResetIgnoredWords: 'Obnovit ignorovaná slova a gramatiku',
    menuThesaurus: 'Tezaurus…',
    menuHyphenation: 'Dělení slov…',
    menuTranslate: 'Přeložit',
    menuTranslateSelection: 'Přeložit výběr',
    menuTranslateDocument: 'Přeložit dokument',
    menuSetTranslationLanguage: 'Nastavit jazyk překladu dokumentu…',
    menuLanguage: 'Jazyk…',
    menuTrackChanges: 'Sledování změn',
    menuHighlightChanges: 'Zvýraznit změny…',
    menuCompareDocuments: 'Porovnat dokumenty…',
    menuCombineDocuments: 'Sloučit dokumenty…',
    menuProtectDocument: 'Zamknout dokument…',
    menuEnvelopes: 'Obálky…',
    menuLabels: 'Štítky…',
    menuCustomizeKeyboard: 'Vlastní klávesnice…',
    menuPreferences: 'Předvolby…',
    menuAiProofread: 'Korektura AI',
    menuWindow: 'Okno',
    menuHelp: 'Nápověda',
    menuShortcuts: 'Klávesové zkratky',
  },
  nl: {
    dlgOpenDoc: 'Document openen',
    filterWord: 'Word-documenten',
    odtImportMsg: 'Dit .{ext}-document converteren?',
    odtImportDetail:
      'GenOffice bewerkt .docx, niet ODF. Tekst, koppen, lijsten, tabellen en afbeeldingen worden omgezet; voetnoten, opmerkingen en tekstvakken niet volledig. Het oorspronkelijke bestand wordt nooit gewijzigd; de omgezette .docx is een zelfstandige kopie waarin volgende opslaan-acties terechtkomen.',
    hwpImportDetail:
      'GenOffice bewerkt .docx, niet HWP. Tekst, tabellen, afbeeldingen, kop- en voetteksten worden omgezet; vergelijkingen, voetnoten en ingesloten objecten niet, en automatische nummering wordt gewone tekst. Het oorspronkelijke bestand wordt nooit gewijzigd; de omgezette .docx is een zelfstandige kopie waarin volgende opslaan-acties terechtkomen.',
    odtImportConvertBtn: 'Converteren en openen',
    importSkipConfirm: 'Volgende keer converteren zonder te vragen',
    docImportMsg: 'Dit .{ext}-document converteren?',
    docImportDetail:
      'GenOffice bewerkt .docx, geen Word 97-2003-documenten (.doc). Tekst, koppen, lijsten, tabellen en afbeeldingen worden omgezet; voetnoten, opmerkingen en tekstvakken niet volledig. Het oorspronkelijke bestand wordt nooit gewijzigd; de omgezette .docx is een zelfstandige kopie waarin volgende opslaan-acties terechtkomen.',
    docImportConvertBtn: 'Converteren en openen',
    dlgSaveAs: 'Opslaan als',
    closeUnsavedMsg: 'Dit document bevat niet-opgeslagen wijzigingen.',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    closeNoReplyMsg: 'Het document reageert niet en heeft mogelijk niet-opgeslagen wijzigingen.',
    closeNoReplyDetail: 'Toch sluiten? Niet-opgeslagen wijzigingen gaan verloren.',
    btnCloseAnyway: 'Toch sluiten',
    autosaveFoundTitle: 'Herstelde versie gevonden',
    autosaveDiskNewer:
      'Let op: het bestand op schijf is gewijzigd nadat deze herstelkopie is opgeslagen. Herstellen vervangt die nieuwere wijzigingen.',
    autosaveFoundBody:
      'Er zijn niet-opgeslagen wijzigingen van uw laatste sessie. De automatisch opgeslagen versie herstellen?',
    autosaveRestore: 'Herstellen',
    autosaveDiscard: 'Negeren',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
    extModifiedMsg: 'Het bestand is door een ander programma gewijzigd.',
    extModifiedDetail: 'Toch opslaan en de wijzigingen op schijf overschrijven?',
    btnOverwrite: 'Overschrijven',
    dlgInsertImage: 'Afbeelding invoegen',
    filterImages: 'Afbeeldingen',
    dlgMergeDataSource: 'Ontvangerslijst selecteren',
    filterMergeDataSource: 'Excel of CSV',
    dlgAddAttachment: 'Bijlagen toevoegen',
    filterSupported: 'Ondersteunde bestanden',
    filterAll: 'Alle bestanden',
    dlgExportPdf: 'Exporteren als PDF',
    dlgExportHtml: 'Exporteren als HTML',
    dlgExportMarkdown: 'Exporteren als Markdown',
    dlgPickExportDir: 'Exportmap kiezen',
    errUnsupportedExt: '.{ext}-bestanden worden niet ondersteund',
    errNotFile: 'geen bestand',
    errTooLarge: 'overschrijdt de limiet van {mb} MB',
    errTemplateProtected:
      'deze sjabloon is beveiligd met een wachtwoord; open deze via Bestand > Openen',
    errImageTooLarge: 'afbeelding overschrijdt de limiet van 5 MB',
    errUnreadable: 'kan niet worden gelezen',
    errFileTooLarge: 'Bestand overschrijdt de maximale grootte',
    errParseFailed: 'Kan bestand niet parseren',
    errImageNoText:
      'Afbeeldingsbijlagen bevatten geen tekst; de afbeelding wordt samen met het gebruikersbericht verzonden en kan direct worden bekeken',
    errNotImage: 'geen ondersteund afbeeldingstype',
    errGskNotLoggedIn:
      'Niet aangemeld bij Genspark: klik hieronder op “Aanmelden bij Genspark”, meld u aan en probeer het opnieuw',
    errNoApiKey: 'Geen API-sleutel geconfigureerd voor {provider}',
    errAiRegion:
      'De aanbieder van dit model bedient uw huidige regio niet — kies een ander model of wissel van netwerk',
    errAiBusy: 'De AI-service is momenteel overbelast — probeer het zo opnieuw',
    errNoModel: 'Geen modelnaam geconfigureerd',
    menuFile: 'Bestand',
    menuNewDoc: 'Nieuw document',
    menuNewWindow: 'Nieuw venster',
    menuOpen: 'Openen…',
    menuOpenRecent: 'Recent geopend',
    menuNoRecent: 'Geen recente documenten',
    menuRecentMore: 'Meer…',
    menuClose: 'Sluiten',
    menuSave: 'Opslaan',
    menuSaveAs: 'Opslaan als…',
    menuPageSetup: 'Pagina-instelling…',
    menuExportPdf: 'Exporteren als PDF…',
    menuExportHtml: 'Exporteren als HTML…',
    menuExportMarkdown: 'Exporteren als Markdown…',
    menuExportImages: 'Exporteren als afbeeldingen…',
    menuPrint: 'Afdrukken…',
    menuEdit: 'Bewerken',
    menuUndo: 'Ongedaan maken',
    menuRedo: 'Opnieuw',
    menuRepeat: 'Herhalen',
    menuCantRepeat: 'Kan niet herhalen',
    menuCut: 'Knippen',
    menuCopy: 'Kopiëren',
    menuPaste: 'Plakken',
    menuPasteMatch: 'Plakken met bijpassende opmaak',
    menuFindReplace: 'Zoeken en vervangen…',
    menuPasteSpecial: 'Plakken speciaal…',
    menuClear: 'Wissen',
    menuClearFormatting: 'Opmaak wissen',
    menuClearContents: 'Inhoud',
    menuFindMenu: 'Zoeken',
    menuFind: 'Zoeken…',
    menuFindNext: 'Volgende zoeken',
    menuFindPrevious: 'Vorige zoeken',
    menuReplace: 'Vervangen…',
    menuAdvancedFind: 'Geavanceerd zoeken en vervangen…',
    menuLinks: 'Koppelingen…',
    menuSelectData: 'Gegevens selecteren…',
    menuToggleDrawing: 'Tekenen in-/uitschakelen',
    menuNewFromTemplate: 'Nieuw op basis van sjabloon…',
    menuSaveAsTemplate: 'Opslaan als sjabloon…',
    menuMove: 'Verplaatsen…',
    menuRename: 'Naam wijzigen…',
    menuAlwaysOpenReadOnly: 'Altijd alleen-lezen openen',
    menuReduceFileSize: 'Bestandsgrootte verkleinen…',
    menuProperties: 'Eigenschappen…',
    menuGoTo: 'Ga naar…',
    menuSelectAll: 'Alles selecteren',
    menuView: 'Beeld',
    menuZoom: 'Zoomen',
    menuZoomIn: 'Inzoomen',
    menuZoomOut: 'Uitzoomen',
    menuZoom100: 'Ware grootte (100%)',
    menuPrintLayout: 'Afdrukweergave',
    menuWebLayout: 'Weblay-out',
    menuOutline: 'Overzicht',
    menuDraft: 'Concept',
    menuFocus: 'Focus',
    menuRibbon: 'Lint',
    menuSidebar: 'Zijbalk',
    menuThumbnail: 'Miniatuur',
    menuNavigation: 'Navigatie',
    menuReviewing: 'Controleren',
    menuSidebarSearch: 'Zoeken',
    menuRuler: 'Liniaal',
    menuHeaderFooter: 'Koptekst en voettekst',
    menuFootnotes: 'Voetnoten',
    menuMarkup: 'Markeringen',
    menuRevealFormatting: 'Opmaak weergeven',
    menuZoomDialog: 'Zoomen…',
    menuGridlines: 'Rasterlijnen',
    menuPageWidth: 'Paginabreedte',
    menuWholePage: 'Hele pagina',
    menuMultiPage: "Meerdere pagina's",
    menuAiSidebar: 'AI-zijbalk',
    menuDarkMode: 'Donkere modus',
    menuFullscreen: 'Schermvullende weergave',
    menuInsert: 'Invoegen',
    menuInsertTable: 'Tabel…',
    menuInsertImage: 'Afbeelding…',
    menuInsertPageBreak: 'Pagina-einde',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Vergelijking…',
    menuPictures: 'Afbeeldingen',
    menuPictureFromFile: 'Afbeelding uit bestand…',
    menuShape: 'Vorm',
    menuShapeRectangle: 'Rechthoek',
    menuShapeRoundedRectangle: 'Afgeronde rechthoek',
    menuShapeTriangle: 'Driehoek',
    menuShapeOval: 'Ovaal',
    menuShapeLine: 'Lijn',
    menuChart: 'Grafiek',
    menuChartColumn: 'Kolom',
    menuChartBar: 'Staaf',
    menuChartLine: 'Lijn',
    menuChartArea: 'Vlak',
    menuChartPie: 'Cirkel',
    menuChartTreemap: 'Treemap',
    menuChartSunburst: 'Zonnestraal',
    menuChartHistogram: 'Histogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Box-and-whisker',
    menuChartScatter: 'Spreiding (X Y)',
    menuChartWaterfall: 'Waterval',
    menuChartFunnel: 'Trechter',
    menuChartStock: 'Koers',
    menuChartSurface: 'Oppervlak',
    menuChartRadar: 'Radar',
    menuIcons: 'Pictogrammen…',
    menuBookmark: 'Bladwijzer…',
    menuCrossReference: 'Kruisverwijzing…',
    menuBreak: 'Einde',
    menuColumnBreak: 'Kolomeinde',
    menuSectionNextPage: 'Sectie-einde (volgende pagina)',
    menuSectionContinuous: 'Sectie-einde (doorlopend)',
    menuSectionOddPage: 'Sectie-einde (oneven pagina)',
    menuSectionEvenPage: 'Sectie-einde (even pagina)',
    menuFootnote: 'Voetnoot…',
    menuCaption: 'Bijschrift…',
    menuIndexAndTables: 'Index en inhoudsopgave…',
    menuWatermark: 'Watermerk…',
    menuPageNumbers: 'Paginanummers…',
    menuTextBox: 'Tekstvak',
    menuVerticalTextBox: 'Verticaal tekstvak tekenen',
    menuAutoText: 'AutoTekst',
    menuAutoTextDialog: 'AutoTekst…',
    menuAutoTextNew: 'Nieuw…',
    menuDateAndTime: 'Datum en tijd…',
    menuObject: 'Object…',
    menuInsertFile: 'Bestand…',
    menuField: 'Veld…',
    menuAdvancedSymbol: 'Geavanceerd symbool…',
    menuComment: 'Opmerking',
    menuFormat: 'Opmaak',
    menuBold: 'Vet',
    menuItalic: 'Cursief',
    menuUnderline: 'Onderstrepen',
    menuAlign: 'Uitlijnen',
    menuAlignLeft: 'Links uitlijnen',
    menuAlignCenter: 'Centreren',
    menuAlignRight: 'Rechts uitlijnen',
    menuAlignJustify: 'Uitvullen',
    menuFont: 'Lettertype…',
    menuParagraph: 'Alinea…',
    menuTextEffects: 'Teksteffecten…',
    menuDocument: 'Document…',
    menuBulletsAndNumbering: 'Opsommingstekens en nummering…',
    menuBordersAndShading: 'Randen en arcering…',
    menuColumns: 'Kolommen…',
    menuTabs: 'Tabs…',
    menuDropCap: 'Initiaal…',
    menuTextDirection: 'Tekstrichting…',
    menuChangeCase: 'Hoofdlettergebruik…',
    menuAsianLayout: 'Aziatische indeling',
    menuPhoneticGuide: 'Fonetische gids…',
    menuEncloseCharacters: 'Tekens omsluiten…',
    menuHorizontalInVertical: 'Horizontaal in verticaal…',
    menuCombineCharacters: 'Tekens combineren…',
    menuTwoLinesInOne: 'Twee regels in één…',
    menuFitText: 'Tekst passend maken…',
    menuStyle: 'Stijl…',
    menuFontSubstitution: 'Lettertypevervanging…',
    menuFormatObject: 'Object opmaken…',
    menuEquationOptions: 'Vergelijkingsopties…',
    menuSelectionPane: 'Selectievenster…',
    menuTools: 'Extra',
    menuTable: 'Tabel',
    menuTableInsert: 'Invoegen',
    menuTableInsertTable: 'Tabel…',
    menuTableColsLeft: 'Kolommen links',
    menuTableColsRight: 'Kolommen rechts',
    menuTableRowsAbove: 'Rijen boven',
    menuTableRowsBelow: 'Rijen onder',
    menuTableCells: 'Cellen…',
    menuTableDelete: 'Verwijderen',
    menuTableDeleteTable: 'Tabel',
    menuTableDeleteColumns: 'Kolommen',
    menuTableDeleteRows: 'Rijen',
    menuTableSelect: 'Selecteren',
    menuTableSelectCell: 'Cel',
    menuTableSelectColumn: 'Kolom',
    menuTableSelectRow: 'Rij',
    menuTableSelectTable: 'Tabel',
    menuTableMergeCells: 'Cellen samenvoegen',
    menuTableSplitCells: 'Cellen splitsen…',
    menuTableSplitTable: 'Tabel splitsen',
    menuTableAutoFit: 'AutoAanpassen',
    menuTableAutoFitContents: 'Aanpassen aan inhoud',
    menuTableAutoFitWindow: 'Aanpassen aan venster',
    menuTableFixedWidth: 'Vaste kolombreedte',
    menuTableDistributeRows: 'Rijen gelijkmatig verdelen',
    menuTableDistributeColumns: 'Kolommen gelijkmatig verdelen',
    menuTableRepeatHeader: 'Koprijen herhalen',
    menuTableConvert: 'Converteren',
    menuTableConvertTextToTable: 'Tekst naar tabel converteren…',
    menuTableConvertTableToText: 'Tabel naar tekst converteren…',
    menuTableSort: 'Sorteren…',
    menuTableFormula: 'Formule…',
    menuTableGridlines: 'Rasterlijnen weergeven',
    menuTableProperties: 'Tabeleigenschappen…',
    menuWordCount: 'Woorden tellen…',
    menuAutoCorrect: 'AutoCorrectie-opties…',
    menuSpellingAndGrammar: 'Spelling en grammatica',
    menuSpellingDialog: 'Spelling en grammatica…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Spelfouten verbergen',
    menuHideGrammarErrors: 'Grammaticafouten verbergen',
    menuResetIgnoredWords: 'Genegeerde woorden en grammatica opnieuw instellen',
    menuThesaurus: 'Synoniemenlijst…',
    menuHyphenation: 'Woordafbreking…',
    menuTranslate: 'Vertalen',
    menuTranslateSelection: 'Selectie vertalen',
    menuTranslateDocument: 'Document vertalen',
    menuSetTranslationLanguage: 'Vertaaltaal voor document instellen…',
    menuLanguage: 'Taal…',
    menuTrackChanges: 'Wijzigingen bijhouden',
    menuHighlightChanges: 'Wijzigingen markeren…',
    menuCompareDocuments: 'Documenten vergelijken…',
    menuCombineDocuments: 'Documenten combineren…',
    menuProtectDocument: 'Document beveiligen…',
    menuEnvelopes: 'Enveloppen…',
    menuLabels: 'Etiketten…',
    menuCustomizeKeyboard: 'Toetsenbord aanpassen…',
    menuPreferences: 'Voorkeuren…',
    menuAiProofread: 'AI-proeflezen',
    menuWindow: 'Venster',
    menuHelp: 'Help',
    menuShortcuts: 'Sneltoetsen',
  },
  ms: {
    dlgOpenDoc: 'Buka Dokumen',
    filterWord: 'Dokumen Word',
    odtImportMsg: 'Tukar dokumen .{ext} ini?',
    odtImportDetail:
      'GenOffice menyunting .docx, bukan ODF. Teks kandungan, tajuk, senarai, jadual dan imej ditukar; nota kaki, komen dan kotak teks tidak dikekalkan sepenuhnya. Fail asal tidak pernah diubah; salinan .docx yang ditukar ialah salinan berasingan tempat simpanan seterusnya ditulis.',
    hwpImportDetail:
      'GenOffice menyunting .docx, bukan HWP. Teks kandungan, jadual, imej, pengepala dan pengaki ditukar; persamaan, nota kaki dan objek terbenam tidak ditukar, dan penomboran automatik menjadi teks biasa. Fail asal tidak pernah diubah; salinan .docx yang ditukar ialah salinan berasingan tempat simpanan seterusnya ditulis.',
    odtImportConvertBtn: 'Tukar dan buka',
    importSkipConfirm: 'Lain kali terus tukar tanpa bertanya',
    docImportMsg: 'Tukar dokumen .{ext} ini?',
    docImportDetail:
      'GenOffice menyunting .docx, bukan dokumen Word 97-2003 (.doc). Teks kandungan, tajuk, senarai, jadual dan imej ditukar; nota kaki, komen dan kotak teks tidak dikekalkan sepenuhnya. Fail asal tidak pernah diubah; salinan .docx yang ditukar ialah salinan berasingan tempat simpanan seterusnya ditulis.',
    docImportConvertBtn: 'Tukar dan buka',
    dlgSaveAs: 'Simpan Sebagai',
    closeUnsavedMsg: 'Dokumen ini mempunyai perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Adakah anda mahu menyimpannya sebelum menutup?',
    closeNoReplyMsg: 'Dokumen tidak bertindak balas dan mungkin ada perubahan yang belum disimpan.',
    closeNoReplyDetail: 'Tutup juga? Perubahan yang belum disimpan akan hilang.',
    btnCloseAnyway: 'Tutup Juga',
    autosaveFoundTitle: 'Versi pulihan ditemui',
    autosaveDiskNewer:
      'Nota: fail pada cakera telah diubah suai selepas salinan pemulihan ini disimpan. Memulihkan akan menggantikan perubahan yang lebih baharu itu.',
    autosaveFoundBody:
      'Terdapat perubahan yang belum disimpan daripada sesi terakhir anda. Pulihkan versi yang disimpan secara automatik?',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    extModifiedMsg: 'Fail telah diubah oleh program lain.',
    extModifiedDetail: 'Simpan juga dan tulis ganti perubahan pada cakera?',
    btnOverwrite: 'Tulis Ganti',
    dlgInsertImage: 'Sisipkan Imej',
    filterImages: 'Imej',
    dlgMergeDataSource: 'Pilih senarai penerima',
    filterMergeDataSource: 'Excel atau CSV',
    dlgAddAttachment: 'Tambah Lampiran',
    filterSupported: 'Fail yang Disokong',
    filterAll: 'Semua Fail',
    dlgExportPdf: 'Eksport sebagai PDF',
    dlgExportHtml: 'Eksport sebagai HTML',
    dlgExportMarkdown: 'Eksport sebagai Markdown',
    dlgPickExportDir: 'Pilih Folder Eksport',
    errUnsupportedExt: 'fail .{ext} tidak disokong',
    errNotFile: 'bukan fail',
    errTooLarge: 'melebihi had {mb}MB',
    errTemplateProtected: 'templat ini dilindungi kata laluan; buka melalui Fail > Buka',
    errImageTooLarge: 'imej melebihi had 5MB',
    errUnreadable: 'tidak dapat dibaca',
    errFileTooLarge: 'Fail melebihi had saiz',
    errParseFailed: 'Gagal menghurai fail',
    errImageNoText:
      'Lampiran imej tidak menyediakan teks; imej dihantar bersama mesej pengguna dan boleh dilihat terus',
    errNotImage: 'bukan jenis imej yang disokong',
    errGskNotLoggedIn:
      'Belum log masuk ke Genspark: klik “Log masuk ke Genspark” di bawah, kemudian cuba lagi',
    errNoApiKey: 'Kunci API untuk {provider} belum dikonfigurasikan',
    errAiRegion:
      'Penyedia model ini tidak memberi perkhidmatan di rantau semasa anda — pilih model lain atau tukar rangkaian',
    errAiBusy: 'Perkhidmatan AI sedang sibuk — sila cuba lagi sebentar lagi',
    errNoModel: 'Nama model belum dikonfigurasikan',
    menuFile: 'Fail',
    menuNewDoc: 'Dokumen Baharu',
    menuNewWindow: 'Tetingkap Baharu',
    menuOpen: 'Buka…',
    menuOpenRecent: 'Buka Terkini',
    menuNoRecent: 'Tiada Dokumen Terkini',
    menuRecentMore: 'Lagi…',
    menuClose: 'Tutup',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuPageSetup: 'Persediaan Halaman…',
    menuExportPdf: 'Eksport sebagai PDF…',
    menuExportHtml: 'Eksport sebagai HTML…',
    menuExportMarkdown: 'Eksport sebagai Markdown…',
    menuExportImages: 'Eksport sebagai imej…',
    menuPrint: 'Cetak…',
    menuEdit: 'Edit',
    menuUndo: 'Buat Asal',
    menuRedo: 'Buat Semula',
    menuRepeat: 'Ulang',
    menuCantRepeat: 'Tidak Boleh Ulang',
    menuCut: 'Potong',
    menuCopy: 'Salin',
    menuPaste: 'Tampal',
    menuPasteMatch: 'Tampal dan Padankan Pemformatan',
    menuFindReplace: 'Cari dan Ganti…',
    menuPasteSpecial: 'Tampal Khas…',
    menuClear: 'Kosongkan',
    menuClearFormatting: 'Kosongkan Pemformatan',
    menuClearContents: 'Kandungan',
    menuFindMenu: 'Cari',
    menuFind: 'Cari…',
    menuFindNext: 'Cari Seterusnya',
    menuFindPrevious: 'Cari Sebelumnya',
    menuReplace: 'Ganti…',
    menuAdvancedFind: 'Cari dan Ganti Lanjutan…',
    menuLinks: 'Pautan…',
    menuSelectData: 'Pilih Data…',
    menuToggleDrawing: 'Togol Lukisan',
    menuNewFromTemplate: 'Baharu daripada Templat…',
    menuSaveAsTemplate: 'Simpan sebagai Templat…',
    menuMove: 'Alih…',
    menuRename: 'Namakan Semula…',
    menuAlwaysOpenReadOnly: 'Sentiasa Buka Baca Sahaja',
    menuReduceFileSize: 'Kurangkan Saiz Fail…',
    menuProperties: 'Sifat…',
    menuGoTo: 'Pergi ke…',
    menuSelectAll: 'Pilih Semua',
    menuView: 'Lihat',
    menuZoom: 'Zum',
    menuZoomIn: 'Zum Masuk',
    menuZoomOut: 'Zum Keluar',
    menuZoom100: 'Saiz Sebenar (100%)',
    menuPrintLayout: 'Tataletak Cetakan',
    menuWebLayout: 'Tataletak Web',
    menuOutline: 'Rangka',
    menuDraft: 'Draf',
    menuFocus: 'Fokus',
    menuRibbon: 'Reben',
    menuSidebar: 'Bar Sisi',
    menuThumbnail: 'Imej Kecil',
    menuNavigation: 'Navigasi',
    menuReviewing: 'Semakan',
    menuSidebarSearch: 'Cari',
    menuRuler: 'Pembaris',
    menuHeaderFooter: 'Pengepala dan Pengaki',
    menuFootnotes: 'Nota Kaki',
    menuMarkup: 'Markup',
    menuRevealFormatting: 'Dedahkan Pemformatan',
    menuZoomDialog: 'Zum…',
    menuGridlines: 'Garis Grid',
    menuPageWidth: 'Lebar Halaman',
    menuWholePage: 'Seluruh Halaman',
    menuMultiPage: 'Beberapa halaman',
    menuAiSidebar: 'Bar Sisi AI',
    menuDarkMode: 'Mod Gelap',
    menuFullscreen: 'Masuk Skrin Penuh',
    menuInsert: 'Sisip',
    menuInsertTable: 'Jadual…',
    menuInsertImage: 'Imej…',
    menuInsertPageBreak: 'Pemisah Halaman',
    menuInsertLink: 'Hiperpautan…',
    menuInsertEquation: 'Persamaan…',
    menuPictures: 'Gambar',
    menuPictureFromFile: 'Gambar daripada Fail…',
    menuShape: 'Bentuk',
    menuShapeRectangle: 'Segi Empat Tepat',
    menuShapeRoundedRectangle: 'Segi Empat Tepat Bundar',
    menuShapeTriangle: 'Segi Tiga',
    menuShapeOval: 'Bujur',
    menuShapeLine: 'Garis',
    menuChart: 'Carta',
    menuChartColumn: 'Lajur',
    menuChartBar: 'Bar',
    menuChartLine: 'Garis',
    menuChartArea: 'Kawasan',
    menuChartPie: 'Pai',
    menuChartTreemap: 'Peta Pepohon',
    menuChartSunburst: 'Sunburst',
    menuChartHistogram: 'Histogram',
    menuChartPareto: 'Pareto',
    menuChartBoxWhisker: 'Kotak dan Misai',
    menuChartScatter: 'X Y (Serakan)',
    menuChartWaterfall: 'Air Terjun',
    menuChartFunnel: 'Corong',
    menuChartStock: 'Saham',
    menuChartSurface: 'Permukaan',
    menuChartRadar: 'Radar',
    menuIcons: 'Ikon…',
    menuBookmark: 'Penanda Buku…',
    menuCrossReference: 'Rujukan Silang…',
    menuBreak: 'Pemisah',
    menuColumnBreak: 'Pemisah Lajur',
    menuSectionNextPage: 'Pemisah Seksyen (Halaman Seterusnya)',
    menuSectionContinuous: 'Pemisah Seksyen (Berterusan)',
    menuSectionOddPage: 'Pemisah Seksyen (Halaman Ganjil)',
    menuSectionEvenPage: 'Pemisah Seksyen (Halaman Genap)',
    menuFootnote: 'Nota Kaki…',
    menuCaption: 'Kapsyen…',
    menuIndexAndTables: 'Indeks dan Jadual…',
    menuWatermark: 'Tera Air…',
    menuPageNumbers: 'Nombor Halaman…',
    menuTextBox: 'Kotak Teks',
    menuVerticalTextBox: 'Lukis Kotak Teks Menegak',
    menuAutoText: 'AutoTeks',
    menuAutoTextDialog: 'AutoTeks…',
    menuAutoTextNew: 'Baharu…',
    menuDateAndTime: 'Tarikh dan Masa…',
    menuObject: 'Objek…',
    menuInsertFile: 'Fail…',
    menuField: 'Medan…',
    menuAdvancedSymbol: 'Simbol Lanjutan…',
    menuComment: 'Komen',
    menuFormat: 'Format',
    menuBold: 'Tebal',
    menuItalic: 'Condong',
    menuUnderline: 'Garis Bawah',
    menuAlign: 'Jajarkan',
    menuAlignLeft: 'Jajar Kiri',
    menuAlignCenter: 'Tengah',
    menuAlignRight: 'Jajar Kanan',
    menuAlignJustify: 'Justifikasi',
    menuFont: 'Fon…',
    menuParagraph: 'Perenggan…',
    menuTextEffects: 'Kesan Teks…',
    menuDocument: 'Dokumen…',
    menuBulletsAndNumbering: 'Bulet dan Penomboran…',
    menuBordersAndShading: 'Sempadan dan Lorekan…',
    menuColumns: 'Lajur…',
    menuTabs: 'Tab…',
    menuDropCap: 'Huruf Besar Jatuh…',
    menuTextDirection: 'Arah Teks…',
    menuChangeCase: 'Ubah Huruf Besar/Kecil…',
    menuAsianLayout: 'Tataletak Asia',
    menuPhoneticGuide: 'Panduan Fonetik…',
    menuEncloseCharacters: 'Kurung Aksara…',
    menuHorizontalInVertical: 'Mendatar dalam Menegak…',
    menuCombineCharacters: 'Gabungkan Aksara…',
    menuTwoLinesInOne: 'Dua Baris dalam Satu…',
    menuFitText: 'Muatkan Teks…',
    menuStyle: 'Gaya…',
    menuFontSubstitution: 'Penggantian Fon…',
    menuFormatObject: 'Format Objek…',
    menuEquationOptions: 'Pilihan Persamaan…',
    menuSelectionPane: 'Anak Tetingkap Pilihan…',
    menuTools: 'Alat',
    menuTable: 'Jadual',
    menuTableInsert: 'Sisipkan',
    menuTableInsertTable: 'Jadual…',
    menuTableColsLeft: 'Lajur di Kiri',
    menuTableColsRight: 'Lajur di Kanan',
    menuTableRowsAbove: 'Baris di Atas',
    menuTableRowsBelow: 'Baris di Bawah',
    menuTableCells: 'Sel…',
    menuTableDelete: 'Padam',
    menuTableDeleteTable: 'Jadual',
    menuTableDeleteColumns: 'Lajur',
    menuTableDeleteRows: 'Baris',
    menuTableSelect: 'Pilih',
    menuTableSelectCell: 'Sel',
    menuTableSelectColumn: 'Lajur',
    menuTableSelectRow: 'Baris',
    menuTableSelectTable: 'Jadual',
    menuTableMergeCells: 'Cantum Sel',
    menuTableSplitCells: 'Pisahkan Sel…',
    menuTableSplitTable: 'Pisahkan Jadual',
    menuTableAutoFit: 'Autopadan',
    menuTableAutoFitContents: 'Padan kepada Kandungan',
    menuTableAutoFitWindow: 'Padan kepada Tetingkap',
    menuTableFixedWidth: 'Lebar Lajur Tetap',
    menuTableDistributeRows: 'Agihkan Baris Sama Rata',
    menuTableDistributeColumns: 'Agihkan Lajur Sama Rata',
    menuTableRepeatHeader: 'Ulang Baris Pengepala',
    menuTableConvert: 'Tukar',
    menuTableConvertTextToTable: 'Tukar Teks kepada Jadual…',
    menuTableConvertTableToText: 'Tukar Jadual kepada Teks…',
    menuTableSort: 'Isih…',
    menuTableFormula: 'Formula…',
    menuTableGridlines: 'Lihat Garis Grid',
    menuTableProperties: 'Sifat Jadual…',
    menuWordCount: 'Kiraan Perkataan…',
    menuAutoCorrect: 'Pilihan AutoBetul…',
    menuSpellingAndGrammar: 'Ejaan dan Tatabahasa',
    menuSpellingDialog: 'Ejaan dan Tatabahasa…',
    menuEditor: 'Editor',
    menuHideSpellingErrors: 'Sembunyikan Ralat Ejaan',
    menuHideGrammarErrors: 'Sembunyikan Ralat Tatabahasa',
    menuResetIgnoredWords: 'Set Semula Perkataan dan Tatabahasa yang Diabaikan',
    menuThesaurus: 'Tesaurus…',
    menuHyphenation: 'Penyempangan…',
    menuTranslate: 'Terjemah',
    menuTranslateSelection: 'Terjemah Pilihan',
    menuTranslateDocument: 'Terjemah Dokumen',
    menuSetTranslationLanguage: 'Tetapkan Bahasa Terjemahan Dokumen…',
    menuLanguage: 'Bahasa…',
    menuTrackChanges: 'Jejak Perubahan',
    menuHighlightChanges: 'Serlahkan Perubahan…',
    menuCompareDocuments: 'Bandingkan Dokumen…',
    menuCombineDocuments: 'Gabungkan Dokumen…',
    menuProtectDocument: 'Lindungi Dokumen…',
    menuEnvelopes: 'Sampul Surat…',
    menuLabels: 'Label…',
    menuCustomizeKeyboard: 'Sesuaikan Papan Kekunci…',
    menuPreferences: 'Keutamaan…',
    menuAiProofread: 'Pembacaan Pruf AI',
    menuWindow: 'Tetingkap',
    menuHelp: 'Bantuan',
    menuShortcuts: 'Pintasan Papan Kekunci',
  },
  he: {
    dlgOpenDoc: 'פתיחת מסמך',
    filterWord: 'מסמכי Word',
    odtImportMsg: 'להמיר את מסמך ה‑.{ext} הזה?',
    odtImportDetail:
      'GenOffice עורך קבצי ‎.docx ולא ODF. טקסט, כותרות, רשימות, טבלאות ותמונות מומרים; הערות שוליים, תגובות ותיבות טקסט לא נשמרים במלואם. קובץ המקור לעולם אינו משתנה; קובץ ה‑.docx שהומר הוא עותק נפרד שאליו נכתבות השמירות הבאות.',
    hwpImportDetail:
      'GenOffice עורך קבצי ‎.docx ולא HWP. טקסט, טבלאות, תמונות, כותרות עליונות ותחתונות מומרים; משוואות, הערות שוליים ואובייקטים מוטמעים אינם מומרים, ומספור אוטומטי הופך לטקסט רגיל. קובץ המקור לעולם אינו משתנה; קובץ ה‑.docx שהומר הוא עותק נפרד שאליו נכתבות השמירות הבאות.',
    odtImportConvertBtn: 'המר ופתח',
    importSkipConfirm: 'בפעם הבאה להמיר בלי לשאול',
    docImportMsg: 'להמיר את מסמך ה‑.{ext} הזה?',
    docImportDetail:
      'GenOffice עורך קבצי ‎.docx ולא מסמכי Word 97-2003 (‎.doc). טקסט, כותרות, רשימות, טבלאות ותמונות מומרים; הערות שוליים, תגובות ותיבות טקסט לא נשמרים במלואם. קובץ המקור לעולם אינו משתנה; קובץ ה‑.docx שהומר הוא עותק נפרד שאליו נכתבות השמירות הבאות.',
    docImportConvertBtn: 'המר ופתח',
    dlgSaveAs: 'שמירה בשם',
    closeUnsavedMsg: 'במסמך זה יש שינויים שלא נשמרו.',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    closeNoReplyMsg: 'המסמך אינו מגיב וייתכן שיש בו שינויים שלא נשמרו.',
    closeNoReplyDetail: 'לסגור בכל זאת? שינויים שלא נשמרו יאבדו.',
    btnCloseAnyway: 'סגור בכל זאת',
    autosaveFoundTitle: 'נמצאה גרסה משוחזרת',
    autosaveDiskNewer:
      'שים לב: הקובץ בדיסק שונה לאחר שמירת עותק השחזור הזה. שחזור יחליף את השינויים החדשים יותר.',
    autosaveFoundBody: 'קיימים שינויים שלא נשמרו מהפעלה הקודמת. לשחזר את הגרסה שנשמרה אוטומטית?',
    autosaveRestore: 'שחזר',
    autosaveDiscard: 'התעלם',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
    extModifiedMsg: 'הקובץ שונה על ידי תוכנית אחרת.',
    extModifiedDetail: 'לשמור בכל זאת ולדרוס את השינויים בדיסק?',
    btnOverwrite: 'דרוס',
    dlgInsertImage: 'הוספת תמונה',
    filterImages: 'תמונות',
    dlgMergeDataSource: 'בחירת רשימת נמענים',
    filterMergeDataSource: 'Excel או CSV',
    dlgAddAttachment: 'הוספת קבצים מצורפים',
    filterSupported: 'קבצים נתמכים',
    filterAll: 'כל הקבצים',
    dlgExportPdf: 'ייצוא כ-PDF',
    dlgExportHtml: 'ייצוא כ-HTML',
    dlgExportMarkdown: 'ייצוא כ-Markdown',
    dlgPickExportDir: 'בחירת תיקיית ייצוא',
    errUnsupportedExt: 'קובצי .{ext} אינם נתמכים',
    errNotFile: 'אינו קובץ',
    errTooLarge: 'חורג מהמגבלה של {mb}MB',
    errTemplateProtected: 'תבנית זו מוגנת בסיסמה; פתח אותה דרך קובץ > פתח',
    errImageTooLarge: 'התמונה חורגת מהמגבלה של 5MB',
    errUnreadable: 'לא ניתן לקרוא',
    errFileTooLarge: 'הקובץ חורג ממגבלת הגודל',
    errParseFailed: 'ניתוח הקובץ נכשל',
    errImageNoText:
      'קבצים מצורפים מסוג תמונה אינם מספקים טקסט; התמונה נשלחת יחד עם הודעת המשתמש וניתן לצפות בה ישירות',
    errNotImage: 'סוג תמונה שאינו נתמך',
    errGskNotLoggedIn: 'לא מחובר ל-Genspark: לחץ על "התחבר ל-Genspark" למטה, התחבר ונסה שוב',
    errNoApiKey: 'לא הוגדר מפתח API עבור {provider}',
    errAiRegion: 'ספק המודל הזה אינו פועל באזור הנוכחי שלך — בחרו מודל אחר או החליפו רשת',
    errAiBusy: 'שירות ה-AI עמוס כרגע — נסו שוב בעוד רגע',
    errNoModel: 'לא הוגדר שם מודל',
    menuFile: 'קובץ',
    menuNewDoc: 'מסמך חדש',
    menuNewWindow: 'חלון חדש',
    menuOpen: 'פתיחה…',
    menuOpenRecent: 'פתיחת מסמכים אחרונים',
    menuNoRecent: 'אין מסמכים אחרונים',
    menuRecentMore: 'עוד…',
    menuClose: 'סגירה',
    menuSave: 'שמירה',
    menuSaveAs: 'שמירה בשם…',
    menuPageSetup: 'הגדרת עמוד…',
    menuExportPdf: 'ייצוא כ-PDF…',
    menuExportHtml: 'ייצוא כ-HTML…',
    menuExportMarkdown: 'ייצוא כ-Markdown…',
    menuExportImages: 'ייצוא כתמונות…',
    menuPrint: 'הדפסה…',
    menuEdit: 'עריכה',
    menuUndo: 'בטל',
    menuRedo: 'בצע שוב',
    menuRepeat: 'חזור על הפעולה',
    menuCantRepeat: 'לא ניתן לחזור על הפעולה',
    menuCut: 'גזור',
    menuCopy: 'העתק',
    menuPaste: 'הדבק',
    menuPasteMatch: 'הדבק והתאם עיצוב',
    menuFindReplace: 'חיפוש והחלפה…',
    menuPasteSpecial: 'הדבקה מיוחדת…',
    menuClear: 'נקה',
    menuClearFormatting: 'נקה עיצוב',
    menuClearContents: 'תוכן',
    menuFindMenu: 'חיפוש',
    menuFind: 'חיפוש…',
    menuFindNext: 'חפש את הבא',
    menuFindPrevious: 'חפש את הקודם',
    menuReplace: 'החלפה…',
    menuAdvancedFind: 'חיפוש והחלפה מתקדמים…',
    menuLinks: 'קישורים…',
    menuSelectData: 'בחירת נתונים…',
    menuToggleDrawing: 'החלף מצב ציור',
    menuNewFromTemplate: 'חדש מתבנית…',
    menuSaveAsTemplate: 'שמירה כתבנית…',
    menuMove: 'העבר…',
    menuRename: 'שנה שם…',
    menuAlwaysOpenReadOnly: 'פתח תמיד לקריאה בלבד',
    menuReduceFileSize: 'הקטן גודל קובץ…',
    menuProperties: 'מאפיינים…',
    menuGoTo: 'עבור אל…',
    menuSelectAll: 'בחר הכול',
    menuView: 'תצוגה',
    menuZoom: 'זום',
    menuZoomIn: 'התקרבות',
    menuZoomOut: 'התרחקות',
    menuZoom100: 'גודל אמיתי (100%)',
    menuPrintLayout: 'פריסת הדפסה',
    menuWebLayout: 'פריסת אינטרנט',
    menuOutline: 'חלוקה לרמות',
    menuDraft: 'טיוטה',
    menuFocus: 'מיקוד',
    menuRibbon: 'רצועת הכלים',
    menuSidebar: 'סרגל צד',
    menuThumbnail: 'תמונה ממוזערת',
    menuNavigation: 'ניווט',
    menuReviewing: 'סקירה',
    menuSidebarSearch: 'חיפוש',
    menuRuler: 'סרגל',
    menuHeaderFooter: 'כותרת עליונה ותחתונה',
    menuFootnotes: 'הערות שוליים',
    menuMarkup: 'סימון',
    menuRevealFormatting: 'חשוף עיצוב',
    menuZoomDialog: 'זום…',
    menuGridlines: 'קווי רשת',
    menuPageWidth: 'רוחב עמוד',
    menuWholePage: 'עמוד שלם',
    menuMultiPage: 'עמודים מרובים',
    menuAiSidebar: 'סרגל צד AI',
    menuDarkMode: 'מצב כהה',
    menuFullscreen: 'מעבר למסך מלא',
    menuInsert: 'הוספה',
    menuInsertTable: 'טבלה…',
    menuInsertImage: 'תמונה…',
    menuInsertPageBreak: 'מעבר עמוד',
    menuInsertLink: 'היפר-קישור…',
    menuInsertEquation: 'משוואה…',
    menuPictures: 'תמונות',
    menuPictureFromFile: 'תמונה מקובץ…',
    menuShape: 'צורה',
    menuShapeRectangle: 'מלבן',
    menuShapeRoundedRectangle: 'מלבן מעוגל',
    menuShapeTriangle: 'משולש',
    menuShapeOval: 'אליפסה',
    menuShapeLine: 'קו',
    menuChart: 'תרשים',
    menuChartColumn: 'עמודות',
    menuChartBar: 'עמודות אופקיות',
    menuChartLine: 'קו',
    menuChartArea: 'שטח',
    menuChartPie: 'עוגה',
    menuChartTreemap: 'מפת עץ',
    menuChartSunburst: 'מפת קרני שמש',
    menuChartHistogram: 'היסטוגרמה',
    menuChartPareto: 'פארטו',
    menuChartBoxWhisker: 'קופסה ושפם',
    menuChartScatter: 'פיזור (X Y)',
    menuChartWaterfall: 'מפל',
    menuChartFunnel: 'משפך',
    menuChartStock: 'מניות',
    menuChartSurface: 'משטח',
    menuChartRadar: 'מכ"ם',
    menuIcons: 'סמלים…',
    menuBookmark: 'סימניה…',
    menuCrossReference: 'הפניה מקושרת…',
    menuBreak: 'מעבר',
    menuColumnBreak: 'מעבר עמודה',
    menuSectionNextPage: 'מעבר מקטע (עמוד הבא)',
    menuSectionContinuous: 'מעבר מקטע (רציף)',
    menuSectionOddPage: 'מעבר מקטע (עמוד אי-זוגי)',
    menuSectionEvenPage: 'מעבר מקטע (עמוד זוגי)',
    menuFootnote: 'הערת שוליים…',
    menuCaption: 'כיתוב…',
    menuIndexAndTables: 'אינדקס וטבלאות…',
    menuWatermark: 'סימן מים…',
    menuPageNumbers: 'מספרי עמודים…',
    menuTextBox: 'תיבת טקסט',
    menuVerticalTextBox: 'צייר תיבת טקסט אנכית',
    menuAutoText: 'טקסט אוטומטי',
    menuAutoTextDialog: 'טקסט אוטומטי…',
    menuAutoTextNew: 'חדש…',
    menuDateAndTime: 'תאריך ושעה…',
    menuObject: 'אובייקט…',
    menuInsertFile: 'קובץ…',
    menuField: 'שדה…',
    menuAdvancedSymbol: 'סימן מתקדם…',
    menuComment: 'הערה',
    menuFormat: 'עיצוב',
    menuBold: 'מודגש',
    menuItalic: 'נטוי',
    menuUnderline: 'קו תחתון',
    menuAlign: 'יישור',
    menuAlignLeft: 'יישור לשמאל',
    menuAlignCenter: 'מרכוז',
    menuAlignRight: 'יישור לימין',
    menuAlignJustify: 'יישור לשני הצדדים',
    menuFont: 'גופן…',
    menuParagraph: 'פסקה…',
    menuTextEffects: 'אפקטי טקסט…',
    menuDocument: 'מסמך…',
    menuBulletsAndNumbering: 'תבליטים ומספור…',
    menuBordersAndShading: 'גבולות והצללה…',
    menuColumns: 'עמודות…',
    menuTabs: 'טאבים…',
    menuDropCap: 'אות פתיחה…',
    menuTextDirection: 'כיוון טקסט…',
    menuChangeCase: 'שנה רישיות…',
    menuAsianLayout: 'פריסה אסייתית',
    menuPhoneticGuide: 'מדריך פונטי…',
    menuEncloseCharacters: 'הקף תווים…',
    menuHorizontalInVertical: 'אופקי בתוך אנכי…',
    menuCombineCharacters: 'שלב תווים…',
    menuTwoLinesInOne: 'שתי שורות באחת…',
    menuFitText: 'התאם טקסט…',
    menuStyle: 'סגנון…',
    menuFontSubstitution: 'החלפת גופנים…',
    menuFormatObject: 'עיצוב אובייקט…',
    menuEquationOptions: 'אפשרויות משוואה…',
    menuSelectionPane: 'חלונית בחירה…',
    menuTools: 'כלים',
    menuTable: 'טבלה',
    menuTableInsert: 'הוסף',
    menuTableInsertTable: 'טבלה…',
    menuTableColsLeft: 'עמודות משמאל',
    menuTableColsRight: 'עמודות מימין',
    menuTableRowsAbove: 'שורות מעל',
    menuTableRowsBelow: 'שורות מתחת',
    menuTableCells: 'תאים…',
    menuTableDelete: 'מחק',
    menuTableDeleteTable: 'טבלה',
    menuTableDeleteColumns: 'עמודות',
    menuTableDeleteRows: 'שורות',
    menuTableSelect: 'בחר',
    menuTableSelectCell: 'תא',
    menuTableSelectColumn: 'עמודה',
    menuTableSelectRow: 'שורה',
    menuTableSelectTable: 'טבלה',
    menuTableMergeCells: 'מזג תאים',
    menuTableSplitCells: 'פצל תאים…',
    menuTableSplitTable: 'פצל טבלה',
    menuTableAutoFit: 'התאמה אוטומטית',
    menuTableAutoFitContents: 'התאם לתוכן',
    menuTableAutoFitWindow: 'התאם לחלון',
    menuTableFixedWidth: 'רוחב עמודה קבוע',
    menuTableDistributeRows: 'פזר שורות באופן שווה',
    menuTableDistributeColumns: 'פזר עמודות באופן שווה',
    menuTableRepeatHeader: 'חזור על שורות כותרת',
    menuTableConvert: 'המרה',
    menuTableConvertTextToTable: 'המרת טקסט לטבלה…',
    menuTableConvertTableToText: 'המרת טבלה לטקסט…',
    menuTableSort: 'מיון…',
    menuTableFormula: 'נוסחה…',
    menuTableGridlines: 'הצג קווי רשת',
    menuTableProperties: 'מאפייני טבלה…',
    menuWordCount: 'ספירת מילים…',
    menuAutoCorrect: 'אפשרויות תיקון אוטומטי…',
    menuSpellingAndGrammar: 'איות ודקדוק',
    menuSpellingDialog: 'איות ודקדוק…',
    menuEditor: 'עורך',
    menuHideSpellingErrors: 'הסתר שגיאות איות',
    menuHideGrammarErrors: 'הסתר שגיאות דקדוק',
    menuResetIgnoredWords: 'אפס מילים ודקדוק שהתעלמת מהם',
    menuThesaurus: 'אוצר מילים…',
    menuHyphenation: 'מיקוף…',
    menuTranslate: 'תרגום',
    menuTranslateSelection: 'תרגם בחירה',
    menuTranslateDocument: 'תרגם מסמך',
    menuSetTranslationLanguage: 'הגדר שפת תרגום למסמך…',
    menuLanguage: 'שפה…',
    menuTrackChanges: 'מעקב אחר שינויים',
    menuHighlightChanges: 'סמן שינויים…',
    menuCompareDocuments: 'השווה מסמכים…',
    menuCombineDocuments: 'שלב מסמכים…',
    menuProtectDocument: 'הגן על המסמך…',
    menuEnvelopes: 'מעטפות…',
    menuLabels: 'תוויות…',
    menuCustomizeKeyboard: 'התאמה אישית של המקלדת…',
    menuPreferences: 'העדפות…',
    menuAiProofread: 'הגהת AI',
    menuWindow: 'חלון',
    menuHelp: 'עזרה',
    menuShortcuts: 'קיצורי מקלדת',
  },
  hi: {
    dlgOpenDoc: 'दस्तावेज़ खोलें',
    filterWord: 'Word दस्तावेज़',
    odtImportMsg: 'इस .{ext} दस्तावेज़ को बदलें?',
    odtImportDetail:
      'GenOffice .docx संपादित करता है, ODF नहीं। मुख्य पाठ, शीर्षक, सूचियाँ, तालिकाएँ और चित्र बदल दिए जाते हैं; फुटनोट, टिप्पणियाँ और टेक्स्ट बॉक्स पूरी तरह नहीं रखे जाते। मूल फ़ाइल कभी नहीं बदलती; बदली हुई .docx एक स्वतंत्र प्रति है जिसमें आगे के सेव लिखे जाते हैं।',
    hwpImportDetail:
      'GenOffice .docx संपादित करता है, HWP नहीं। मुख्य पाठ, तालिकाएँ, चित्र, हेडर और फुटर बदल दिए जाते हैं; समीकरण, फुटनोट और एम्बेड किए गए ऑब्जेक्ट नहीं बदले जाते, और स्वचालित क्रमांकन सामान्य पाठ बन जाता है। मूल फ़ाइल कभी नहीं बदलती; बदली हुई .docx एक स्वतंत्र प्रति है जिसमें आगे के सेव लिखे जाते हैं।',
    odtImportConvertBtn: 'बदलकर खोलें',
    importSkipConfirm: 'अगली बार बिना पूछे बदलें',
    docImportMsg: 'इस .{ext} दस्तावेज़ को बदलें?',
    docImportDetail:
      'GenOffice Word 97-2003 (.doc) दस्तावेज़ नहीं, .docx संपादित करता है। मुख्य पाठ, शीर्षक, सूचियाँ, तालिकाएँ और चित्र बदल दिए जाते हैं; फुटनोट, टिप्पणियाँ और टेक्स्ट बॉक्स पूरी तरह नहीं रखे जाते। मूल फ़ाइल कभी नहीं बदलती; बदली हुई .docx एक स्वतंत्र प्रति है जिसमें आगे के सेव लिखे जाते हैं।',
    docImportConvertBtn: 'बदलकर खोलें',
    dlgSaveAs: 'इस रूप में सहेजें',
    closeUnsavedMsg: 'इस दस्तावेज़ में सहेजे नहीं गए परिवर्तन हैं।',
    closeUnsavedDetail: 'क्या आप बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    closeNoReplyMsg: 'दस्तावेज़ प्रतिक्रिया नहीं दे रहा है और उसमें सहेजे न गए बदलाव हो सकते हैं।',
    closeNoReplyDetail: 'फिर भी बंद करें? सहेजे न गए बदलाव खो जाएँगे।',
    btnCloseAnyway: 'फिर भी बंद करें',
    autosaveFoundTitle: 'पुनर्प्राप्त संस्करण मिला',
    autosaveDiskNewer:
      'ध्यान दें: यह पुनर्प्राप्ति प्रति सहेजे जाने के बाद डिस्क पर फ़ाइल में बदलाव किया गया है। पुनर्स्थापित करने पर वे नए बदलाव बदल दिए जाएँगे।',
    autosaveFoundBody:
      'आपके पिछले सत्र से सहेजे नहीं गए परिवर्तन हैं। स्वतः सहेजा गया संस्करण पुनर्स्थापित करें?',
    autosaveRestore: 'पुनर्स्थापित करें',
    autosaveDiscard: 'छोड़ें',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
    extModifiedMsg: 'फ़ाइल को किसी अन्य प्रोग्राम ने बदल दिया है।',
    extModifiedDetail: 'फिर भी सहेजें और डिस्क पर मौजूद बदलावों को अधिलेखित करें?',
    btnOverwrite: 'अधिलेखित करें',
    dlgInsertImage: 'छवि सम्मिलित करें',
    filterImages: 'छवियाँ',
    dlgMergeDataSource: 'प्राप्तकर्ता सूची चुनें',
    filterMergeDataSource: 'Excel या CSV',
    dlgAddAttachment: 'अनुलग्नक जोड़ें',
    filterSupported: 'समर्थित फ़ाइलें',
    filterAll: 'सभी फ़ाइलें',
    dlgExportPdf: 'PDF के रूप में निर्यात करें',
    dlgExportHtml: 'HTML के रूप में निर्यात करें',
    dlgExportMarkdown: 'Markdown के रूप में निर्यात करें',
    dlgPickExportDir: 'निर्यात फ़ोल्डर चुनें',
    errUnsupportedExt: '.{ext} फ़ाइलें समर्थित नहीं हैं',
    errNotFile: 'फ़ाइल नहीं है',
    errTooLarge: '{mb}MB की सीमा से अधिक है',
    errTemplateProtected: 'यह टेम्पलेट पासवर्ड से सुरक्षित है; इसे फ़ाइल > खोलें से खोलें',
    errImageTooLarge: 'छवि 5MB की सीमा से अधिक है',
    errUnreadable: 'पढ़ा नहीं जा सकता',
    errFileTooLarge: 'फ़ाइल आकार सीमा से अधिक है',
    errParseFailed: 'फ़ाइल पार्स करने में विफल',
    errImageNoText:
      'छवि अनुलग्नक टेक्स्ट प्रदान नहीं करते; छवि उपयोगकर्ता संदेश के साथ भेजी जाती है, उसे सीधे देखें',
    errNotImage: 'समर्थित छवि प्रकार नहीं है',
    errGskNotLoggedIn:
      'Genspark में साइन इन नहीं है: नीचे “Genspark में साइन इन करें” पर क्लिक करें, साइन इन करें और फिर से कोशिश करें',
    errNoApiKey: '{provider} के लिए कोई API कुंजी कॉन्फ़िगर नहीं है',
    errAiRegion:
      'इस मॉडल का प्रदाता आपके वर्तमान क्षेत्र में सेवा नहीं देता — कोई दूसरा मॉडल चुनें या नेटवर्क बदलें',
    errAiBusy: 'AI सेवा अभी व्यस्त है — कृपया थोड़ी देर बाद फिर से प्रयास करें',
    errNoModel: 'कोई मॉडल नाम कॉन्फ़िगर नहीं है',
    menuFile: 'फ़ाइल',
    menuNewDoc: 'नया दस्तावेज़',
    menuNewWindow: 'नई विंडो',
    menuOpen: 'खोलें…',
    menuOpenRecent: 'हाल के दस्तावेज़ खोलें',
    menuNoRecent: 'कोई हालिया दस्तावेज़ नहीं',
    menuRecentMore: 'और…',
    menuClose: 'बंद करें',
    menuSave: 'सहेजें',
    menuSaveAs: 'इस रूप में सहेजें…',
    menuPageSetup: 'पृष्ठ सेटअप…',
    menuExportPdf: 'PDF के रूप में निर्यात करें…',
    menuExportHtml: 'HTML के रूप में निर्यात करें…',
    menuExportMarkdown: 'Markdown के रूप में निर्यात करें…',
    menuExportImages: 'छवियों के रूप में निर्यात…',
    menuPrint: 'प्रिंट करें…',
    menuEdit: 'संपादन',
    menuUndo: 'पूर्ववत करें',
    menuRedo: 'फिर से करें',
    menuRepeat: 'दोहराएँ',
    menuCantRepeat: 'दोहराया नहीं जा सकता',
    menuCut: 'काटें',
    menuCopy: 'कॉपी करें',
    menuPaste: 'चिपकाएँ',
    menuPasteMatch: 'चिपकाएँ और स्वरूपण मिलाएँ',
    menuFindReplace: 'ढूँढें और बदलें…',
    menuPasteSpecial: 'विशेष चिपकाएँ…',
    menuClear: 'साफ़ करें',
    menuClearFormatting: 'फ़ॉर्मेटिंग साफ़ करें',
    menuClearContents: 'सामग्री',
    menuFindMenu: 'ढूँढें',
    menuFind: 'ढूँढें…',
    menuFindNext: 'अगला ढूँढें',
    menuFindPrevious: 'पिछला ढूँढें',
    menuReplace: 'बदलें…',
    menuAdvancedFind: 'उन्नत ढूँढें और बदलें…',
    menuLinks: 'लिंक…',
    menuSelectData: 'डेटा चुनें…',
    menuToggleDrawing: 'ड्रॉइंग टॉगल करें',
    menuNewFromTemplate: 'टेम्पलेट से नया…',
    menuSaveAsTemplate: 'टेम्पलेट के रूप में सहेजें…',
    menuMove: 'स्थानांतरित करें…',
    menuRename: 'नाम बदलें…',
    menuAlwaysOpenReadOnly: 'हमेशा केवल-पठन के रूप में खोलें',
    menuReduceFileSize: 'फ़ाइल आकार घटाएँ…',
    menuProperties: 'गुण…',
    menuGoTo: 'यहाँ जाएँ…',
    menuSelectAll: 'सभी चुनें',
    menuView: 'दृश्य',
    menuZoom: 'ज़ूम',
    menuZoomIn: 'ज़ूम इन',
    menuZoomOut: 'ज़ूम आउट',
    menuZoom100: 'वास्तविक आकार (100%)',
    menuPrintLayout: 'प्रिंट लेआउट',
    menuWebLayout: 'वेब लेआउट',
    menuOutline: 'आउटलाइन',
    menuDraft: 'ड्राफ़्ट',
    menuFocus: 'फ़ोकस',
    menuRibbon: 'रिबन',
    menuSidebar: 'साइडबार',
    menuThumbnail: 'थंबनेल',
    menuNavigation: 'नेविगेशन',
    menuReviewing: 'समीक्षा',
    menuSidebarSearch: 'खोजें',
    menuRuler: 'रूलर',
    menuHeaderFooter: 'हेडर और फ़ुटर',
    menuFootnotes: 'फ़ुटनोट',
    menuMarkup: 'मार्कअप',
    menuRevealFormatting: 'फ़ॉर्मेटिंग दिखाएँ',
    menuZoomDialog: 'ज़ूम…',
    menuGridlines: 'ग्रिडलाइन',
    menuPageWidth: 'पृष्ठ चौड़ाई',
    menuWholePage: 'पूरा पृष्ठ',
    menuMultiPage: 'एकाधिक पृष्ठ',
    menuAiSidebar: 'AI साइडबार',
    menuDarkMode: 'डार्क मोड',
    menuFullscreen: 'पूर्ण स्क्रीन में जाएँ',
    menuInsert: 'सम्मिलित करें',
    menuInsertTable: 'तालिका…',
    menuInsertImage: 'छवि…',
    menuInsertPageBreak: 'पृष्ठ विराम',
    menuInsertLink: 'हाइपरलिंक…',
    menuInsertEquation: 'समीकरण…',
    menuPictures: 'चित्र',
    menuPictureFromFile: 'फ़ाइल से चित्र…',
    menuShape: 'आकृति',
    menuShapeRectangle: 'आयत',
    menuShapeRoundedRectangle: 'गोल कोनों वाला आयत',
    menuShapeTriangle: 'त्रिभुज',
    menuShapeOval: 'अंडाकार',
    menuShapeLine: 'रेखा',
    menuChart: 'चार्ट',
    menuChartColumn: 'स्तंभ',
    menuChartBar: 'बार',
    menuChartLine: 'रेखा',
    menuChartArea: 'क्षेत्र',
    menuChartPie: 'पाई',
    menuChartTreemap: 'ट्रीमैप',
    menuChartSunburst: 'सनबर्स्ट',
    menuChartHistogram: 'हिस्टोग्राम',
    menuChartPareto: 'परेटो',
    menuChartBoxWhisker: 'बॉक्स और व्हिस्कर',
    menuChartScatter: 'X Y (स्कैटर)',
    menuChartWaterfall: 'वॉटरफ़ॉल',
    menuChartFunnel: 'फ़नल',
    menuChartStock: 'स्टॉक',
    menuChartSurface: 'सरफ़ेस',
    menuChartRadar: 'रडार',
    menuIcons: 'चिह्न…',
    menuBookmark: 'बुकमार्क…',
    menuCrossReference: 'क्रॉस-संदर्भ…',
    menuBreak: 'विराम',
    menuColumnBreak: 'स्तंभ विराम',
    menuSectionNextPage: 'अनुभाग विराम (अगला पृष्ठ)',
    menuSectionContinuous: 'अनुभाग विराम (निरंतर)',
    menuSectionOddPage: 'अनुभाग विराम (विषम पृष्ठ)',
    menuSectionEvenPage: 'अनुभाग विराम (सम पृष्ठ)',
    menuFootnote: 'फ़ुटनोट…',
    menuCaption: 'कैप्शन…',
    menuIndexAndTables: 'अनुक्रमणिका और तालिकाएँ…',
    menuWatermark: 'वॉटरमार्क…',
    menuPageNumbers: 'पृष्ठ संख्याएँ…',
    menuTextBox: 'पाठ बॉक्स',
    menuVerticalTextBox: 'लंबवत पाठ बॉक्स बनाएँ',
    menuAutoText: 'स्वतः पाठ',
    menuAutoTextDialog: 'स्वतः पाठ…',
    menuAutoTextNew: 'नया…',
    menuDateAndTime: 'दिनांक और समय…',
    menuObject: 'ऑब्जेक्ट…',
    menuInsertFile: 'फ़ाइल…',
    menuField: 'फ़ील्ड…',
    menuAdvancedSymbol: 'उन्नत प्रतीक…',
    menuComment: 'टिप्पणी',
    menuFormat: 'स्वरूप',
    menuBold: 'बोल्ड',
    menuItalic: 'इटैलिक',
    menuUnderline: 'रेखांकित',
    menuAlign: 'संरेखित करें',
    menuAlignLeft: 'बाएँ संरेखित करें',
    menuAlignCenter: 'केंद्र में रखें',
    menuAlignRight: 'दाएँ संरेखित करें',
    menuAlignJustify: 'जस्टिफ़ाई करें',
    menuFont: 'फ़ॉन्ट…',
    menuParagraph: 'अनुच्छेद…',
    menuTextEffects: 'पाठ प्रभाव…',
    menuDocument: 'दस्तावेज़…',
    menuBulletsAndNumbering: 'बुलेट और क्रमांकन…',
    menuBordersAndShading: 'बॉर्डर और छायांकन…',
    menuColumns: 'स्तंभ…',
    menuTabs: 'टैब…',
    menuDropCap: 'ड्रॉप कैप…',
    menuTextDirection: 'पाठ दिशा…',
    menuChangeCase: 'केस बदलें…',
    menuAsianLayout: 'एशियाई लेआउट',
    menuPhoneticGuide: 'ध्वन्यात्मक मार्गदर्शिका…',
    menuEncloseCharacters: 'वर्ण घेरें…',
    menuHorizontalInVertical: 'लंबवत में क्षैतिज…',
    menuCombineCharacters: 'वर्ण संयोजित करें…',
    menuTwoLinesInOne: 'एक में दो पंक्तियाँ…',
    menuFitText: 'पाठ फ़िट करें…',
    menuStyle: 'शैली…',
    menuFontSubstitution: 'फ़ॉन्ट प्रतिस्थापन…',
    menuFormatObject: 'ऑब्जेक्ट फ़ॉर्मेट करें…',
    menuEquationOptions: 'समीकरण विकल्प…',
    menuSelectionPane: 'चयन फलक…',
    menuTools: 'उपकरण',
    menuTable: 'तालिका',
    menuTableInsert: 'सम्मिलित करें',
    menuTableInsertTable: 'तालिका…',
    menuTableColsLeft: 'बाईं ओर स्तंभ',
    menuTableColsRight: 'दाईं ओर स्तंभ',
    menuTableRowsAbove: 'ऊपर पंक्तियाँ',
    menuTableRowsBelow: 'नीचे पंक्तियाँ',
    menuTableCells: 'कक्ष…',
    menuTableDelete: 'हटाएँ',
    menuTableDeleteTable: 'तालिका',
    menuTableDeleteColumns: 'स्तंभ',
    menuTableDeleteRows: 'पंक्तियाँ',
    menuTableSelect: 'चुनें',
    menuTableSelectCell: 'कक्ष',
    menuTableSelectColumn: 'स्तंभ',
    menuTableSelectRow: 'पंक्ति',
    menuTableSelectTable: 'तालिका',
    menuTableMergeCells: 'कक्ष मर्ज करें',
    menuTableSplitCells: 'कक्ष विभाजित करें…',
    menuTableSplitTable: 'तालिका विभाजित करें',
    menuTableAutoFit: 'स्वतः फ़िट',
    menuTableAutoFitContents: 'सामग्री के अनुसार फ़िट',
    menuTableAutoFitWindow: 'विंडो के अनुसार फ़िट',
    menuTableFixedWidth: 'निश्चित स्तंभ चौड़ाई',
    menuTableDistributeRows: 'पंक्तियाँ समान रूप से बाँटें',
    menuTableDistributeColumns: 'स्तंभ समान रूप से बाँटें',
    menuTableRepeatHeader: 'शीर्ष पंक्तियाँ दोहराएँ',
    menuTableConvert: 'कनवर्ट करें',
    menuTableConvertTextToTable: 'टेक्स्ट को तालिका में बदलें…',
    menuTableConvertTableToText: 'तालिका को टेक्स्ट में बदलें…',
    menuTableSort: 'सॉर्ट करें…',
    menuTableFormula: 'सूत्र…',
    menuTableGridlines: 'ग्रिडलाइन देखें',
    menuTableProperties: 'तालिका गुण…',
    menuWordCount: 'शब्द गणना…',
    menuAutoCorrect: 'स्वतः सुधार विकल्प…',
    menuSpellingAndGrammar: 'वर्तनी और व्याकरण',
    menuSpellingDialog: 'वर्तनी और व्याकरण…',
    menuEditor: 'संपादक',
    menuHideSpellingErrors: 'वर्तनी त्रुटियाँ छिपाएँ',
    menuHideGrammarErrors: 'व्याकरण त्रुटियाँ छिपाएँ',
    menuResetIgnoredWords: 'उपेक्षित शब्द और व्याकरण रीसेट करें',
    menuThesaurus: 'समानार्थी शब्दकोश…',
    menuHyphenation: 'हाइफ़नेशन…',
    menuTranslate: 'अनुवाद',
    menuTranslateSelection: 'चयन का अनुवाद करें',
    menuTranslateDocument: 'दस्तावेज़ का अनुवाद करें',
    menuSetTranslationLanguage: 'दस्तावेज़ अनुवाद भाषा सेट करें…',
    menuLanguage: 'भाषा…',
    menuTrackChanges: 'परिवर्तन ट्रैक करें',
    menuHighlightChanges: 'परिवर्तन हाइलाइट करें…',
    menuCompareDocuments: 'दस्तावेज़ों की तुलना करें…',
    menuCombineDocuments: 'दस्तावेज़ संयोजित करें…',
    menuProtectDocument: 'दस्तावेज़ सुरक्षित करें…',
    menuEnvelopes: 'लिफ़ाफ़े…',
    menuLabels: 'लेबल…',
    menuCustomizeKeyboard: 'कीबोर्ड अनुकूलित करें…',
    menuPreferences: 'प्राथमिकताएँ…',
    menuAiProofread: 'AI प्रूफ़रीडिंग',
    menuWindow: 'विंडो',
    menuHelp: 'सहायता',
    menuShortcuts: 'कीबोर्ड शॉर्टकट',
  },
  'zh-TW': {
    dlgOpenDoc: '開啟文件',
    filterWord: 'Word 文件',
    odtImportMsg: '轉換這個 .{ext} 文件？',
    odtImportDetail:
      'GenOffice 只編輯 .docx，不編輯 ODF。內文、標題、清單、表格與圖片會轉換過來；註腳、註解與文字方塊不會完整保留。原始檔案不會被修改，轉換後的 .docx 是獨立副本，之後的儲存都寫入它。', // public-hygiene: allow
    hwpImportDetail:
      'GenOffice 只編輯 .docx，不編輯 HWP。內文、表格、圖片、頁首與頁尾會轉換過來；方程式、註腳與內嵌物件不會轉換，自動編號會變成一般文字。原始檔案不會被修改，轉換後的 .docx 是獨立副本，之後的儲存都寫入它。', // public-hygiene: allow
    odtImportConvertBtn: '轉換並開啟',
    importSkipConfirm: '下次直接轉換，不再提示',
    docImportMsg: '轉換這個 .{ext} 文件？',
    docImportDetail:
      'GenOffice 只編輯 .docx，不編輯 Word 97-2003 舊版 .doc。內文、標題、清單、表格與圖片會轉換過來；註腳、註解與文字方塊不會完整保留。原始檔案不會被修改，轉換後的 .docx 是獨立副本，之後的儲存都寫入它。', // public-hygiene: allow
    docImportConvertBtn: '轉換並開啟',
    dlgSaveAs: '另存新檔',
    closeUnsavedMsg: '此文件有未儲存的變更。',
    closeUnsavedDetail: '關閉前是否要儲存？',
    closeNoReplyMsg: '文件沒有回應,可能有未儲存的變更。',
    closeNoReplyDetail: '仍要關閉嗎?未儲存的變更將遺失。',
    btnCloseAnyway: '仍要關閉',
    autosaveFoundTitle: '發現自動復原版本',
    autosaveDiskNewer: '注意：此復原副本儲存後，磁碟上的檔案又被修改過。復原將取代這些較新的變更。',
    autosaveFoundBody: '上次工作階段有未儲存的變更。要復原自動儲存的版本嗎?',
    autosaveRestore: '復原',
    autosaveDiscard: '放棄',
    btnDontSave: '不儲存',
    btnCancel: '取消',
    extModifiedMsg: '檔案已被其他程式修改。',
    extModifiedDetail: '仍要儲存並覆寫磁碟上的變更嗎?',
    btnOverwrite: '覆寫',
    dlgInsertImage: '插入圖片',
    filterImages: '圖片',
    dlgMergeDataSource: '選擇收件人清單',
    filterMergeDataSource: 'Excel 或 CSV',
    dlgAddAttachment: '新增附件',
    filterSupported: '支援的檔案',
    filterAll: '所有檔案',
    dlgExportPdf: '匯出為 PDF',
    dlgExportHtml: '匯出為 HTML',
    dlgExportMarkdown: '匯出為 Markdown',
    dlgPickExportDir: '選擇匯出目錄',
    errUnsupportedExt: '暫不支援 .{ext} 類型',
    errNotFile: '不是檔案',
    errTooLarge: '超過 {mb}MB 上限',
    errTemplateProtected: '此範本受密碼保護，請透過「檔案 > 開啟」開啟',
    errImageTooLarge: '圖片超過 5MB 上限',
    errUnreadable: '無法讀取',
    errFileTooLarge: '檔案超過大小上限',
    errParseFailed: '檔案解析失敗',
    errImageNoText: '圖片附件不提供文字,已作為影像隨使用者訊息傳送,直接看圖即可',
    errNotImage: '不是支援的圖片類型',
    errGskNotLoggedIn: '未登入 Genspark:請點擊下方「登入 Genspark」完成登入後重試',
    errNoApiKey: '未設定 {provider} 的 API Key',
    errAiRegion: '該模型的廠商不向目前地區提供服務，請換一個模型或切換網路',
    errAiBusy: 'AI 服務目前繁忙，請稍後重試',
    errNoModel: '未設定模型名稱',
    menuFile: '檔案',
    menuNewDoc: '新增文件',
    menuNewWindow: '新增視窗',
    menuOpen: '開啟…',
    menuOpenRecent: '開啟最近的文件',
    menuNoRecent: '沒有最近的文件',
    menuRecentMore: '更多…',
    menuClose: '關閉',
    menuSave: '儲存',
    menuSaveAs: '另存新檔…',
    menuPageSetup: '版面設定…',
    menuExportPdf: '匯出為 PDF…',
    menuExportHtml: '匯出為 HTML…',
    menuExportMarkdown: '匯出為 Markdown…',
    menuExportImages: '匯出為圖片…',
    menuPrint: '列印…',
    menuEdit: '編輯',
    menuUndo: '復原',
    menuRedo: '重做',
    menuRepeat: '重複',
    menuCantRepeat: '無法重複',
    menuCut: '剪下',
    menuCopy: '複製',
    menuPaste: '貼上',
    menuPasteMatch: '貼上並符合格式',
    menuFindReplace: '尋找與取代…',
    menuPasteSpecial: '選擇性貼上…',
    menuClear: '清除',
    menuClearFormatting: '清除格式設定',
    menuClearContents: '內容',
    menuFindMenu: '尋找',
    menuFind: '尋找…',
    menuFindNext: '找下一個',
    menuFindPrevious: '找上一個',
    menuReplace: '取代…',
    menuAdvancedFind: '進階尋找及取代…',
    menuLinks: '連結…',
    menuSelectData: '選取資料…',
    menuToggleDrawing: '切換繪圖',
    menuNewFromTemplate: '從範本新增…',
    menuSaveAsTemplate: '另存為範本…',
    menuMove: '移動…',
    menuRename: '重新命名…',
    menuAlwaysOpenReadOnly: '一律以唯讀方式開啟',
    menuReduceFileSize: '縮減檔案大小…',
    menuProperties: '內容…',
    menuGoTo: '定位…',
    menuSelectAll: '全選',
    menuView: '檢視',
    menuZoom: '縮放',
    menuZoomIn: '放大',
    menuZoomOut: '縮小',
    menuZoom100: '實際大小 (100%)',
    menuPrintLayout: '整頁模式',
    menuWebLayout: 'Web 版面配置',
    menuOutline: '大綱',
    menuDraft: '草稿',
    menuFocus: '專注',
    menuRibbon: '功能區',
    menuSidebar: '側邊欄',
    menuThumbnail: '縮圖',
    menuNavigation: '導覽',
    menuReviewing: '校閱',
    menuSidebarSearch: '搜尋',
    menuRuler: '尺規',
    menuHeaderFooter: '頁首及頁尾',
    menuFootnotes: '註腳',
    menuMarkup: '標記',
    menuRevealFormatting: '顯示格式設定',
    menuZoomDialog: '縮放…',
    menuGridlines: '格線',
    menuPageWidth: '頁面寬度',
    menuWholePage: '整頁',
    menuMultiPage: '多頁',
    menuAiSidebar: 'AI 側邊欄',
    menuDarkMode: '深色模式',
    menuFullscreen: '進入全螢幕',
    menuInsert: '插入',
    menuInsertTable: '表格…',
    menuInsertImage: '圖片…',
    menuInsertPageBreak: '分頁符號',
    menuInsertLink: '超連結…',
    menuInsertEquation: '方程式…',
    menuPictures: '圖片',
    menuPictureFromFile: '從檔案插入圖片…',
    menuShape: '圖案',
    menuShapeRectangle: '矩形',
    menuShapeRoundedRectangle: '圓角矩形',
    menuShapeTriangle: '三角形',
    menuShapeOval: '橢圓',
    menuShapeLine: '線條',
    menuChart: '圖表',
    menuChartColumn: '直條圖',
    menuChartBar: '橫條圖',
    menuChartLine: '折線圖',
    menuChartArea: '區域圖',
    menuChartPie: '圓形圖',
    menuChartTreemap: '矩形式樹狀結構圖',
    menuChartSunburst: '放射環狀圖',
    menuChartHistogram: '長條圖',
    menuChartPareto: '柏拉圖',
    menuChartBoxWhisker: '盒鬚圖',
    menuChartScatter: 'XY 散佈圖',
    menuChartWaterfall: '瀑布圖',
    menuChartFunnel: '漏斗圖',
    menuChartStock: '股票圖',
    menuChartSurface: '曲面圖',
    menuChartRadar: '雷達圖',
    menuIcons: '圖示…',
    menuBookmark: '書籤…',
    menuCrossReference: '交互參照…',
    menuBreak: '分隔符號',
    menuColumnBreak: '分欄符號',
    menuSectionNextPage: '分節符號(下一頁)',
    menuSectionContinuous: '分節符號(接續本頁)',
    menuSectionOddPage: '分節符號(自奇數頁起)',
    menuSectionEvenPage: '分節符號(自偶數頁起)',
    menuFootnote: '註腳…',
    menuCaption: '標號…',
    menuIndexAndTables: '索引及目錄…',
    menuWatermark: '浮水印…',
    menuPageNumbers: '頁碼…',
    menuTextBox: '文字方塊',
    menuVerticalTextBox: '繪製垂直文字方塊',
    menuAutoText: '自動圖文集',
    menuAutoTextDialog: '自動圖文集…',
    menuAutoTextNew: '新增…',
    menuDateAndTime: '日期及時間…',
    menuObject: '物件…',
    menuInsertFile: '檔案…',
    menuField: '功能變數…',
    menuAdvancedSymbol: '進階符號…',
    menuComment: '註解',
    menuFormat: '格式',
    menuBold: '粗體',
    menuItalic: '斜體',
    menuUnderline: '底線',
    menuAlign: '對齊',
    menuAlignLeft: '靠左對齊',
    menuAlignCenter: '置中',
    menuAlignRight: '靠右對齊',
    menuAlignJustify: '左右對齊',
    menuFont: '字型…',
    menuParagraph: '段落…',
    menuTextEffects: '文字效果…',
    menuDocument: '文件…',
    menuBulletsAndNumbering: '項目符號及編號…',
    menuBordersAndShading: '框線及網底…',
    menuColumns: '欄…',
    menuTabs: '定位點…',
    menuDropCap: '首字放大…',
    menuTextDirection: '直書/橫書…',
    menuChangeCase: '大小寫轉換…',
    menuAsianLayout: '亞洲方式配置',
    menuPhoneticGuide: '注音標示…',
    menuEncloseCharacters: '圍繞字元…',
    menuHorizontalInVertical: '橫向文字…',
    menuCombineCharacters: '組排文字…',
    menuTwoLinesInOne: '並列文字…',
    menuFitText: '最適文字大小…',
    menuStyle: '樣式…',
    menuFontSubstitution: '字型替換…',
    menuFormatObject: '設定物件格式…',
    menuEquationOptions: '方程式選項…',
    menuSelectionPane: '選取範圍窗格…',
    menuTools: '工具',
    menuTable: '表格',
    menuTableInsert: '插入',
    menuTableInsertTable: '表格…',
    menuTableColsLeft: '在左側插入欄',
    menuTableColsRight: '在右側插入欄',
    menuTableRowsAbove: '在上方插入列',
    menuTableRowsBelow: '在下方插入列',
    menuTableCells: '儲存格…',
    menuTableDelete: '刪除',
    menuTableDeleteTable: '表格',
    menuTableDeleteColumns: '欄',
    menuTableDeleteRows: '列',
    menuTableSelect: '選取',
    menuTableSelectCell: '儲存格',
    menuTableSelectColumn: '欄',
    menuTableSelectRow: '列',
    menuTableSelectTable: '表格',
    menuTableMergeCells: '合併儲存格',
    menuTableSplitCells: '分割儲存格…',
    menuTableSplitTable: '拆分表格',
    menuTableAutoFit: '自動調整',
    menuTableAutoFitContents: '根據內容自動調整',
    menuTableAutoFitWindow: '根據視窗自動調整',
    menuTableFixedWidth: '固定欄寬',
    menuTableDistributeRows: '平均分佈各列',
    menuTableDistributeColumns: '平均分佈各欄',
    menuTableRepeatHeader: '重複標題列',
    menuTableConvert: '轉換',
    menuTableConvertTextToTable: '文字轉換為表格…',
    menuTableConvertTableToText: '表格轉換為文字…',
    menuTableSort: '排序…',
    menuTableFormula: '公式…',
    menuTableGridlines: '檢視格線',
    menuTableProperties: '表格屬性…',
    menuWordCount: '字數統計…',
    menuAutoCorrect: '自動校正選項…',
    menuSpellingAndGrammar: '拼字及文法檢查',
    menuSpellingDialog: '拼字及文法檢查…',
    menuEditor: '編輯器',
    menuHideSpellingErrors: '隱藏拼字錯誤',
    menuHideGrammarErrors: '隱藏文法錯誤',
    menuResetIgnoredWords: '重設略過的字詞及文法',
    menuThesaurus: '同義字…',
    menuHyphenation: '斷字…',
    menuTranslate: '翻譯',
    menuTranslateSelection: '翻譯選取範圍',
    menuTranslateDocument: '翻譯文件',
    menuSetTranslationLanguage: '設定文件翻譯語言…',
    menuLanguage: '語言…',
    menuTrackChanges: '追蹤修訂',
    menuHighlightChanges: '醒目提示修訂…',
    menuCompareDocuments: '比較文件…',
    menuCombineDocuments: '合併文件…',
    menuProtectDocument: '保護文件…',
    menuEnvelopes: '信封…',
    menuLabels: '標籤…',
    menuCustomizeKeyboard: '自訂鍵盤…',
    menuPreferences: '偏好設定…',
    menuAiProofread: 'AI 校對',
    menuWindow: '視窗',
    menuHelp: '說明',
    menuShortcuts: '鍵盤快速鍵',
  },
})
type MainKey = Parameters<typeof tMain>[1]
// a command's menuKey must be a real entry of the menu dictionary
void (COMMANDS satisfies Record<string, { menuKey?: MainKey; planned?: true }>)
const tm = (key: MainKey, params?: Parameters<typeof tMain>[2]) => tMain(getUiLang(), key, params)

// ---- runtime configuration (paths differ when bundled into the shell) ----

interface DocsRuntimeConfig {
  /** absolute path to the docs preload bundle */
  preloadPath: string
  /** dev-server URL for the docs renderer (wins over rendererFile) */
  rendererUrl?: string | undefined
  /** absolute path to the built docs renderer index.html */
  rendererFile: string
}

let runtime: DocsRuntimeConfig = {
  preloadPath: join(__dirname, '../preload/index.js'),
  rendererUrl: process.env.ELECTRON_RENDERER_URL,
  rendererFile: join(__dirname, '../renderer/index.html'),
}

export function configureDocsRuntime(config: DocsRuntimeConfig): void {
  runtime = config
  // shell mode: the shell queues argv files itself (per-tab pendingWindowOpens);
  // the module-scope fallback would leak the double-clicked file into the next
  // blank tab's consume-pending-open
  pendingOpenPath = null
}

let mainWindow: BrowserWindow | null = null
let rendererReady = false
let pendingOpenPath = findOpenableDocPath(process.argv)
/** documents queued for windows/tabs spawned via New Tab, keyed by webContents id */
const pendingWindowOpens = new Map<number, string>()
/** what a link in a document may hand to the system: web pages and the mail program */
const LINK_PROTOCOLS = ['http:', 'https:', 'mailto:']
/** webContents ids that should open as a new blank doc instead of the start screen */
const pendingNewBlankIds = new Set<number>()

/** mark a docs webContents as "open blank on first consume" (called by the shell for home:new-doc) */
export function markDocsNewBlank(wcId: number): void {
  pendingNewBlankIds.add(wcId)
}

/** AI-authored content waiting for its create_document tab, keyed by webContents id */
const pendingAiDocContents = new Map<number, AiDocContent>()

/** queue AI content for a fresh blank docs tab (called by the shell right after creating the view) */
export function queueDocsAiContent(wcId: number, content: AiDocContent): void {
  pendingAiDocContents.set(wcId, content)
}

/** a document made in another tab (a comparison result) waiting for the tab that shows it */
export interface ResultDoc {
  bytes: Buffer
}
const pendingResultDocs = new Map<number, ResultDoc>()

/** queue a made document for a fresh docs tab, which opens it untitled (called by the shell right after creating the view) */
export function queueDocsResultDoc(wcId: number, doc: ResultDoc): void {
  pendingResultDocs.set(wcId, doc)
}

/** the single real BrowserWindow hosting the tab strip, used as dialog parent in tab mode */
let docsShellWindow: BrowserWindow | null = null
export function setDocsShellWindow(win: BrowserWindow | null): void {
  docsShellWindow = win
}

/** the window hosting a tab's WebContentsView when BrowserWindow.fromWebContents
 *  cannot tell (detached "Open in New Window" editors) */
let hostWindowHook: ((wc: WebContents) => BrowserWindow | undefined) | null = null
export function setDocsHostWindowHook(
  fn: ((wc: WebContents) => BrowserWindow | undefined) | null,
): void {
  hostWindowHook = fn
}

function hostWindowFor(wc: WebContents | null | undefined): BrowserWindow | undefined {
  const own = wc && (hostWindowHook?.(wc) ?? BrowserWindow.fromWebContents(wc))
  if (own && !own.isDestroyed()) return own
  return docsShellWindow && !docsShellWindow.isDestroyed() ? docsShellWindow : undefined
}

/** injected by the shell in tab mode: resolves the webContents of the currently active docs tab,
 * used for menu-command forwarding where there is no IpcMainInvokeEvent to key off of. */
let activeDocsResolver: (() => WebContents | null) | null = null
export function setActiveDocsResolver(fn: (() => WebContents | null) | null): void {
  activeDocsResolver = fn
}

function activeDocsWebContents(): WebContents | null {
  if (activeDocsResolver) return activeDocsResolver()
  return BrowserWindow.getFocusedWindow()?.webContents ?? mainWindow?.webContents ?? null
}

/** dialog parent for the calling tab: the sender's own window when it has one
 *  (standalone mode, detached "Open in New Window" editors), else the shell window */
function dialogParent(event: IpcMainInvokeEvent): BrowserWindow | undefined {
  return hostWindowFor(event.sender)
}

async function openDialog(event: IpcMainInvokeEvent, options: OpenDialogOptions) {
  return showOpenDialogWithMemory(dialog, dialogParent(event), options)
}

async function saveDialog(event: IpcMainInvokeEvent, options: SaveDialogOptions) {
  // before any pick is remembered, bare-name suggestions anchor in the
  // configurable default save folder instead of Electron's Downloads pin
  return showSaveDialogWithMemory(dialog, dialogParent(event), options, defaultSaveDir)
}

/** default folder where new files land on their first (silent) save; shared with the other editors via shell. User-configurable (app-settings.json), falls back to <Documents>/GenOffice. */
export function defaultSaveDir(): string {
  return configuredDefaultSaveDir(app)
}

/**
 * An encrypted file the window holds no password for does not become a
 * document yet: loadDocx hands back a password marker with no side effects and
 * the renderer runs the replace guard once the password decrypted it
 * (submitDocPwd). Guarding here as well would prompt twice, and a Don't Save
 * answer would drop the recovery copy before anything replaced the document.
 */
async function opensAsPasswordPrompt(filePath: string, wcId: number): Promise<boolean> {
  if (docPasswordFor(wcId, filePath)) return false
  try {
    return isEncryptedDocx(await readFile(filePath))
  } catch {
    return false
  }
}

export function openExternalDocx(filePath: string | null): void {
  if (!filePath) return
  // .odt/.ott and legacy Word .doc/.dot join the flow as an import (converted
  // .docx copy), like the shell route
  if (!isOpenableDocPath(filePath)) return
  const win = BrowserWindow.getFocusedWindow() ?? mainWindow
  if (!rendererReady || !win) {
    pendingOpenPath = filePath
    return
  }
  void (async () => {
    const wcId = win.webContents.id
    const prepared = await prepareDocumentForOpen(filePath, win)
    if (!prepared) return null
    if (
      rendererReady &&
      !(await opensAsPasswordPrompt(prepared, wcId)) &&
      !(await requestDocsClose(win.webContents, win))
    ) {
      return
    }
    return loadDocx(prepared, wcId)
  })()
    .then((result) => {
      if (!result || win.isDestroyed()) return
      if (win.isMinimized()) win.restore()
      win.show()
      win.focus()
      win.webContents.send('docs:opened', result)
    })
    .catch((err) => dialog.showErrorBox(tm('dlgOpenDoc'), String(err)))
}

function userDataPath(...parts: string[]): string {
  return join(app.getPath('userData'), ...parts)
}

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* corrupted state file: fall back to defaults */
  }
  return fallback
}

// ---- recent files ----

const RECENT_PATH = () => userDataPath('recent.json')
const APP_SETTINGS_PATH = () => join(app.getPath('userData'), 'app-settings.json')

// the home screen lists all of these; the File menu shows only the first few
const RECENT_LIMIT = 100

function pushRecent(filePath: string): void {
  const recent = readJson<string[]>(RECENT_PATH(), [])
  // Every save lands here (autosave = every 30s): skip the write and the menu
  // rebuild when the file is already at the head of the list
  if (recent[0] === filePath) return
  const next = [filePath, ...recent.filter((p) => p !== filePath)].slice(0, RECENT_LIMIT)
  try {
    writeJsonAtomic(RECENT_PATH(), next)
  } catch (err) {
    // the document itself is already saved; a lost recents entry must not fail the save
    console.warn('[docs] recent.json write failed:', err)
  }
  buildDocsMenu() // keep File > Open Recent in sync
}

/** unified recents for the shell home screen (paths only; type = extension).
 *  No existence filter: a transiently unavailable path (disconnected drive,
 *  pending mount) must stay listed — the stat layer flags it instead (r158) */
export function readRecentFiles(): string[] {
  return readJson<string[]>(RECENT_PATH(), [])
}

export function recordRecentFile(filePath: string): void {
  pushRecent(filePath)
}

export function removeRecentFiles(filePaths: string[]): void {
  const drop = new Set(filePaths)
  const recent = readJson<string[]>(RECENT_PATH(), [])
  writeJsonAtomic(
    RECENT_PATH(),
    recent.filter((p) => !drop.has(p)),
  )
  buildDocsMenu()
}

/** shell notification: the file of an open view was renamed on disk (renamed in
 *  the Home list) — push to the matching renderer so it syncs its save path and
 *  title bar (docs keeps path state on the renderer side). */
export function docsFileRenamed(wc: WebContents, oldPath: string, newPath: string): void {
  // keep the save allowlist in sync so docs:save accepts the renamed path
  docWritablePaths.get(wc.id)?.delete(oldPath)
  allowDocWrite(wc.id, newPath)
  rememberOpenDoc(wc.id, newPath)
  const states = docDiskStates.get(wc.id)
  const recorded = states?.get(oldPath)
  if (states && recorded) {
    states.delete(oldPath)
    states.set(newPath, recorded)
  }
  // an encrypted document's password must follow the path, or the next save
  // finds no password under the new name and silently writes plaintext
  renameDocPassword(wc.id, oldPath, newPath)
  moveLazyMediaSource(oldPath, newPath)
  wc.send('docs:renamed', { oldPath, newPath })
}

/** keep a renamed file at its old position in the recent/starred lists */
export function replaceRecentFile(oldPath: string, newPath: string): void {
  const recent = readJson<string[]>(RECENT_PATH(), [])
  writeJsonAtomic(
    RECENT_PATH(),
    recent.map((p) => (p === oldPath ? newPath : p)),
  )
  const renamed = renameStarredItem(readStarredItems(STARRED_PATH()), oldPath, newPath)
  if (renamed) writeStarredItems(STARRED_PATH(), renamed)
  buildDocsMenu()
}

// ---- starred files (home screen favorites) ----
// the store lives in starred-files.ts (pure, unit-tested); legacy flat
// string[] files migrate to the versioned shape on the first write

const STARRED_PATH = () => userDataPath('starred.json')

/** No existence filter, same rationale as readRecentFiles: a transiently
 *  unavailable starred file must keep its star and its Starred-view row —
 *  filtering here also desynced the star state shown on recents rows (r158) */
export function readStarredFiles(): string[] {
  return readStarredItems(STARRED_PATH()).map((item) => item.path)
}

export function toggleStarredFile(filePath: string): void {
  writeStarredItems(STARRED_PATH(), toggleStarredItem(readStarredItems(STARRED_PATH()), filePath))
}

/** Bulk unstar (in-app delete, bulk unfollow from the Starred view, or
 *  removing an unavailable entry from the recents list): the star must not
 *  outlive the row it pointed at (r158) */
export function removeStarredFiles(filePaths: string[]): void {
  const next = dropStarredItems(readStarredItems(STARRED_PATH()), filePaths)
  if (next) writeStarredItems(STARRED_PATH(), next)
}

/** put each starred path into `group` (null = back to ungrouped). Groups have
 *  no separate list: a group exists while an entry carries its name, so
 *  creating one is assigning its first file and emptying one is implicit. */
export function setStarredGroup(filePaths: string[], group: string | null): void {
  const next = assignStarredGroup(readStarredItems(STARRED_PATH()), filePaths, group)
  if (next) writeStarredItems(STARRED_PATH(), next)
}

/** group names that currently have at least one starred file, first-seen order */
export function readStarredGroups(): string[] {
  return starredGroupNames(readStarredItems(STARRED_PATH()))
}

/** starred path → its group (ungrouped paths absent); feeds the home IPC entries */
export function readStarredGroupMap(): Map<string, string> {
  return starredGroupMap(readStarredItems(STARRED_PATH()))
}

// ---- original archive (pass-through base: original file archived by content hash) ----

async function archiveOriginal(filePath: string, hash: string, size: number): Promise<void> {
  // a copy larger than the whole cap would only evict every other original
  if (size > ORIGINALS_MAX_BYTES) return
  const dir = userDataPath('originals')
  await mkdir(dir, { recursive: true })
  const target = join(dir, `${hash}.docx`)
  if (!existsSync(target)) await copyFile(filePath, target)
  void pruneOriginals(dir)
}

const ORIGINALS_MAX_BYTES = 500 * 1024 * 1024
let originalsPruneRunning = false

/** cap the archive's total size; oldest by mtime go first (never blocks the open path) */
async function pruneOriginals(dir: string): Promise<void> {
  if (originalsPruneRunning) return
  originalsPruneRunning = true
  try {
    const files: Array<{ path: string; size: number; mtimeMs: number }> = []
    for (const name of await readdir(dir)) {
      try {
        const s = await stat(join(dir, name))
        if (s.isFile()) files.push({ path: join(dir, name), size: s.size, mtimeMs: s.mtimeMs })
      } catch {
        /* removed concurrently */
      }
    }
    let total = files.reduce((sum, f) => sum + f.size, 0)
    files.sort((a, b) => a.mtimeMs - b.mtimeMs)
    for (const f of files) {
      if (total <= ORIGINALS_MAX_BYTES) break
      try {
        await unlink(f.path)
        total -= f.size
      } catch {
        /* already gone */
      }
    }
  } catch {
    /* directory unreadable: retry on the next archive */
  } finally {
    originalsPruneRunning = false
  }
}

/** per-renderer paths writable via docs:save — populated by open/save-as flows */
const docWritablePaths = new Map<number, Set<string>>()
/** per-renderer PDF export targets authorized via the export save dialog */
const pdfWritablePaths = new Map<number, Set<string>>()
const tornDownWcIds = new Set<number>()

// per renderer, kept current on open/save/save-as/rename; only the local-media allowlist reads it
const openDocByWc = new Map<number, string>()

// the encrypted document a password prompt is open for, per renderer: the
// decrypt retry's path must be the one main itself flagged as needing a
// password, not any path the renderer names
const pendingDecryptByWc = new Map<number, string>()

/** Attachment reads verify against what this view's user put in front of the
 *  app: dialog picks, OS drops and clipboard file pastes land in the shared
 *  grant ledger (shell records them), and files sitting beside the open
 *  document are its own neighbors. Everything else is refused — without this
 *  the extension whitelist was the only bar between a compromised renderer
 *  and arbitrary file reads. */
function canReadAttachment(wcId: number, filePath: string): boolean {
  return sharedPathGrants().has(wcId, filePath) || isBesideDocument(openDocByWc.get(wcId), filePath)
}

function rememberOpenDoc(wcId: number, filePath: string): void {
  openDocByWc.set(wcId, filePath)
}

function docsMediaRoots(wcId: number): string[] {
  return documentMediaRoots(openDocByWc.get(wcId), join(app.getPath('temp'), 'genoffice-pasted'))
}

function allowDocWrite(wcId: number, filePath: string): void {
  const set = docWritablePaths.get(wcId) ?? new Set<string>()
  set.add(filePath)
  docWritablePaths.set(wcId, set)
}

/** MCP save_session: the shell resolved this path for the tab, so docs:save-to may write it */
export function authorizeMcpDocWrite(wcId: number, filePath: string): void {
  allowDocWrite(wcId, filePath)
}

const organizerPaths = new Map<number, Set<string>>()

function canDocWrite(wcId: number, filePath: string): boolean {
  return docWritablePaths.get(wcId)?.has(filePath) === true
}

function allowPdfWrite(wcId: number, filePath: string): void {
  const set = pdfWritablePaths.get(wcId) ?? new Set<string>()
  set.add(filePath)
  pdfWritablePaths.set(wcId, set)
}

// Fidelity-harness escape hatch: headless runs have no save dialog to authorize
// paths, so an explicitly configured directory (set only by our test scripts)
// is treated as pre-authorized for PDF export.
const testExportDir = process.env.GENOFFICE_TEST_EXPORT_DIR || null

function canPdfWrite(wcId: number, filePath: string): boolean {
  if (testExportDir && filePath.startsWith(testExportDir + '/')) return true
  return pdfWritablePaths.get(wcId)?.has(filePath) === true
}

// Export as images runs the regular PDF export against a temp file: that file must
// not be revealed like a user export, and only the tab that asked may read it back
// (or write PNGs into the folder it picked).
const imageExportTemps = new Map<number, Set<string>>()
const imageExportDirs = new Map<number, Set<string>>()

function isImageExportTemp(wcId: number, filePath: string): boolean {
  return imageExportTemps.get(wcId)?.has(filePath) === true
}

// Word's Ignore All lasts for the document session. Chromium has no
// per-document skip list, so the word sits in the custom dictionary while a
// renderer holds it and leaves when the last holder goes; the journal pulls
// crash-orphaned words back out on the next start.
const SPELL_IGNORED_PATH = () => userDataPath('spell-ignored.json')
const spellIgnored = new Map<string, Set<number>>()
const journalIgnoredWords = () => writeJsonAtomic(SPELL_IGNORED_PATH(), [...spellIgnored.keys()])

function releaseSpellIgnores(wcId: number): void {
  let changed = false
  for (const [word, holders] of spellIgnored) {
    if (!holders.delete(wcId) || holders.size > 0) continue
    spellIgnored.delete(word)
    session.defaultSession.removeWordFromSpellCheckerDictionary(word)
    changed = true
  }
  if (changed) journalIgnoredWords()
}

function dropDocWriter(wcId: number): void {
  releaseSpellIgnores(wcId)
  docWritablePaths.delete(wcId)
  organizerPaths.delete(wcId)
  pdfWritablePaths.delete(wcId)
  openDocByWc.delete(wcId)
  for (const p of imageExportTemps.get(wcId) ?? []) void rm(p, { force: true })
  imageExportTemps.delete(wcId)
  imageExportDirs.delete(wcId)
  docDiskStates.delete(wcId)
  forgetLazyMediaOwner(wcId)
  // Destroyed renderers count as torn down too: window-close paths never run
  // teardownDocsRenderer, but an in-flight save handler resuming after the
  // destruction must still fail its re-check (wcIds are never reused, so the
  // set only accumulates a few integers per session).
  tornDownWcIds.add(wcId)
}

// ── External-modification detection: remember the disk state at every read/write
// so docs:save can refuse to clobber edits made by Word/another window ──
const docDiskStates = new Map<number, Map<string, DiskFileState>>()

const sha256Hex = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex')

async function rememberDiskState(wcId: number, filePath: string, hash: string): Promise<void> {
  try {
    const s = await stat(filePath)
    const states = docDiskStates.get(wcId) ?? new Map<string, DiskFileState>()
    states.set(filePath, { mtimeMs: s.mtimeMs, size: s.size, hash })
    docDiskStates.set(wcId, states)
  } catch {
    /* unstatable target: skip tracking; the next save simply won't flag a conflict */
  }
}

async function diskChangedExternally(wcId: number, filePath: string): Promise<boolean> {
  let current: { mtimeMs: number; size: number } | null
  try {
    current = await stat(filePath)
  } catch {
    current = null
  }
  return isExternallyModified(docDiskStates.get(wcId)?.get(filePath), current, async () => {
    try {
      return sha256Hex(await readFile(filePath))
    } catch {
      return null
    }
  })
}

/** A closed docs tab detaches without destroying its webContents (shell freeze
 * workaround), so the orphan must lose write access and stop its timers — otherwise
 * its 30s recovery loop resurrects content the user already discarded. */
export function teardownDocsRenderer(contents: WebContents): void {
  teardownZoteroIpc(contents)
  tornDownWcIds.add(contents.id)
  releaseSpellIgnores(contents.id)
  forgetLazyMediaOwner(contents.id)
  // Sweep recovery copies for this renderer's documents: every non-crash close
  // either saved (docs:save already cleared it) or explicitly discarded, so a
  // copy still on disk here is a leftover from an in-flight recovery write.
  for (const p of docWritablePaths.get(contents.id) ?? []) clearRecoveryCopy(p)
  // reclaim every per-wcId grant, not just doc saves — the orphaned renderer
  // must also lose its dialog-authorized PDF targets and disk-state cache
  docWritablePaths.delete(contents.id)
  organizerPaths.delete(contents.id)
  pdfWritablePaths.delete(contents.id)
  docDiskStates.delete(contents.id)
  forgetDocPasswords(contents.id)
  if (!contents.isDestroyed()) contents.send('docs:teardown')
}

// ── Crash recovery: dirty renderers push a copy every 30s
// (docs:write-recovery); a normal save cleans it up; open offers Restore/Discard ──
const recoveryDir = () => userDataPath('docs-autosave')
const recoveryPathFor = (filePath: string) =>
  join(recoveryDir(), `${createHash('sha1').update(filePath).digest('hex').slice(0, 16)}.docx`)

/** Bumped by every clear: an in-flight docs:write-recovery that started before
 * the bump must not recreate the file it is about to land (stale-recovery race). */
const recoveryClearEpochs = new Map<string, number>()

function clearRecoveryCopy(filePath: string): void {
  recoveryClearEpochs.set(filePath, (recoveryClearEpochs.get(filePath) ?? 0) + 1)
  try {
    unlinkSync(recoveryPathFor(filePath))
  } catch {
    /* nothing to clean */
  }
}

interface MaybeRecoveredDocBytes {
  bytes: Buffer
  recovered: boolean
}

/** On open, if a recovery copy exists, ask whether to restore (still points at the
 * original path; only save persists it). A newer original does not discard the copy:
 * it may hold the only pre-crash edits, so the prompt just says the disk file is newer. */
async function maybeRecoverDocBytes(
  filePath: string,
  original: Buffer,
): Promise<MaybeRecoveredDocBytes> {
  const asPath = recoveryPathFor(filePath)
  let diskNewer: boolean
  try {
    if (!existsSync(asPath)) return { bytes: original, recovered: false }
    diskNewer = recoveryCopyDiskNewer(statSync(asPath).mtimeMs, statSync(filePath).mtimeMs)
  } catch {
    return { bytes: original, recovered: false }
  }
  const options = {
    type: 'question' as const,
    buttons: [tm('autosaveRestore'), tm('autosaveDiscard')],
    defaultId: 0,
    cancelId: 1,
    message: tm('autosaveFoundTitle'),
    detail: diskNewer
      ? `${tm('autosaveFoundBody')}\n\n${tm('autosaveDiskNewer')}`
      : tm('autosaveFoundBody'),
  }
  const parent = BrowserWindow.getFocusedWindow() ?? mainWindow
  const r =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (r.response === 0) {
    try {
      return { bytes: await readFile(asPath), recovered: true }
    } catch {
      return { bytes: original, recovered: false }
    }
  }
  clearRecoveryCopy(filePath)
  return { bytes: original, recovered: false }
}

// Word's own .docx ceiling
const MAX_OPEN_BYTES = 512 * 1024 * 1024
const MAX_GALLERY_TEMPLATES = 200
const GALLERY_READS_AT_ONCE = 4
const MAX_THUMBNAIL_BYTES = 2 * 1024 * 1024
const THUMBNAIL_WIDTH_PX = 320

/** What the window shows at `rect`, as a small JPEG; null when there is nothing to take. */
async function capturePreview(
  sender: WebContents,
  rect: TemplatePreviewRect | undefined,
): Promise<Uint8Array | null> {
  const sides = rect ? [rect.x, rect.y, rect.width, rect.height].map(Math.round) : []
  if (sides.length !== 4 || sides.some((v) => !Number.isFinite(v) || v < 0)) return null
  const [x, y, width, height] = sides
  if (width < 16 || height < 16) return null
  try {
    const image = await sender.capturePage({ x, y, width, height })
    if (image.isEmpty()) return null
    return image.resize({ width: THUMBNAIL_WIDTH_PX, quality: 'best' }).toJPEG(82)
  } catch {
    return null
  }
}

async function showOpenError(wcId: number, detail: string): Promise<void> {
  const parent = hostWindowFor(webContents.fromId(wcId)) ?? mainWindow
  const options = { type: 'error' as const, message: tm('dlgOpenDoc'), detail }
  if (parent && !parent.isDestroyed()) await dialog.showMessageBox(parent, options)
  else await dialog.showMessageBox(options)
}

/**
 * An .odt/.ott, a legacy Word .doc/.dot or an .hwp/.hwpx opens only as a
 * converted .docx copy: confirm the lossy import, pick the destination (the
 * sibling .docx by default), convert, write the file and open it — the
 * explicit-destination flow the PDF "Export as Word" path already uses, so
 * every later Save lands in the copy the user chose and the original is never
 * written. Anything that is not an import format passes through unchanged;
 * null cancels the open.
 */
async function prepareDocumentForOpen(
  filePath: string,
  parent?: BrowserWindow | undefined,
): Promise<string | null> {
  const importKind = docImportKindForPath(filePath)
  if (!importKind) return filePath
  // .hwp and .hwpx route by content: sites serve each under the other's name
  const isHwp = importKind === 'hwp'
  const word97 = importKind === 'word97'
  if (!(await confirmLossyImport(filePath, importKind, parent))) return null
  const picked = await showSaveDialogWithMemory(
    dialog,
    parent,
    {
      title: tm('dlgSaveAs'),
      defaultPath: join(
        dirname(filePath),
        basename(filePath).replace(/\.(odt|ott|hwpx?|doc|dot)$/i, '.docx'),
      ),
      filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
    },
    defaultSaveDir(),
  )
  if (picked.canceled || !picked.filePath) return null
  try {
    // Imports are read whole and handed to one converter or another — the
    // rhwp WASM path especially is unauditable — so the same cap as the
    // .docx open (MAX_OPEN_BYTES) applies before a byte is read.
    if ((await stat(filePath)).size > MAX_OPEN_BYTES) {
      const mb = MAX_OPEN_BYTES / 1024 / 1024
      if (parent && !parent.isDestroyed()) {
        await dialog.showMessageBox(parent, {
          type: 'error',
          message: tm('dlgOpenDoc'),
          detail: `${basename(filePath)}: ${tm('errTooLarge', { mb })}`,
        })
      }
      return null
    }
    const source = new Uint8Array(await readFile(filePath))
    const result = word97
      ? await convertLegacyDocToDocx(source)
      : isHwp
        ? await convertHwpToDocx(source)
        : await convertOdtToDocx(source)
    await atomicWriteFile(picked.filePath, Buffer.from(result.docx))
  } catch {
    if (parent && !parent.isDestroyed()) {
      await dialog.showMessageBox(parent, {
        type: 'error',
        message: tm('dlgOpenDoc'),
        detail: `${basename(filePath)}: ${tm('errUnreadable')}`,
      })
    }
    return null
  }
  return picked.filePath
}

/** The lossy-import warning, unless the user asked not to see it again for this kind; false cancels. */
async function confirmLossyImport(
  filePath: string,
  importKind: NonNullable<ReturnType<typeof docImportKindForPath>>,
  parent?: BrowserWindow | undefined,
): Promise<boolean> {
  const extension = filePath.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1] ?? ''
  const isHwp = importKind === 'hwp'
  const word97 = importKind === 'word97'
  const skipKind = word97 ? 'word97' : isHwp ? 'hwp' : 'odt'
  if (readLegacyImportSkip(APP_SETTINGS_PATH(), skipKind)) return true
  const options = {
    type: 'warning' as const,
    message: tm(word97 ? 'docImportMsg' : 'odtImportMsg', { ext: extension }),
    detail: tm(word97 ? 'docImportDetail' : isHwp ? 'hwpImportDetail' : 'odtImportDetail'),
    buttons: [tm(word97 ? 'docImportConvertBtn' : 'odtImportConvertBtn'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 1,
    noLink: true,
    checkboxLabel: tm('importSkipConfirm'),
  }
  const { response, checkboxChecked } = parent
    ? await dialog.showMessageBox(parent, options)
    : await dialog.showMessageBox(options)
  if (response !== 0) return false
  if (checkboxChecked) {
    try {
      writeLegacyImportSkip(APP_SETTINGS_PATH(), skipKind, true)
    } catch {
      /* unwritable settings: ask again next time */
    }
  }
  return true
}

async function loadDocx(
  filePath: string,
  wcId: number,
  password?: string,
): Promise<OpenDocxResult> {
  if (typeof filePath !== 'string' || !/\.docx$/i.test(filePath)) return null
  if (!existsSync(filePath)) return null
  const size = (await stat(filePath)).size
  const lazy = await openLazyDocx(filePath, wcId)
  if ((lazy?.bytes.length ?? size) > MAX_OPEN_BYTES) {
    const mb = MAX_OPEN_BYTES / 1024 / 1024
    await showOpenError(wcId, `${basename(filePath)}: ${tm('errTooLarge', { mb })}`)
    return null
  }
  const original = lazy?.bytes ?? (await readFile(filePath))
  // Password-protected docx (ECMA-376 CFB container): without a password, hand
  // back a marker — the renderer prompts and retries via docs:open-decrypt.
  // No side effects (recents/write grant) until the password checks out.
  let plainBytes: Buffer = original
  const encrypted = isEncryptedDocx(original)
  if (encrypted) {
    const pwd = password ?? docPasswordFor(wcId, filePath)
    if (!pwd) {
      pendingDecryptByWc.set(wcId, filePath)
      return { needsPassword: true, path: filePath, name: basename(filePath) }
    }
    plainBytes = await decryptDocx(original, pwd) // throws DocxDecryptError
    rememberDocPassword(wcId, filePath, pwd)
  } else {
    rememberDocPassword(wcId, filePath, null)
  }
  // the archive keeps the on-disk original as-is (encrypted ones included: they
  // reopen with the user's password), so a bad save never loses the source file
  const hash = lazy?.hash ?? sha256Hex(original)
  await archiveOriginal(filePath, hash, size)
  const recovery = await maybeRecoverDocBytes(filePath, plainBytes)
  let bytes = recovery.bytes
  let recovered = recovery.recovered
  // recovery copies of a protected document are themselves encrypted (see
  // docs:write-recovery); an unreadable copy falls back to the original
  if (encrypted && isEncryptedDocx(bytes)) {
    try {
      bytes = await decryptRecoveryCopy(wcId, filePath, bytes)
    } catch {
      bytes = plainBytes
      recovered = false
    }
  }
  if (recovered) await adoptLazyMediaHashes(bytes, filePath, wcId)
  pushRecent(filePath)
  allowDocWrite(wcId, filePath)
  rememberOpenDoc(wcId, filePath)
  if (fileOpenedHook) fileOpenedHook(wcId, filePath)
  markDiskEncrypted(wcId, filePath, encrypted)
  // record the on-disk file, not the recovery copy: what matters is what save would overwrite
  await rememberDiskState(wcId, filePath, hash)
  return {
    path: filePath,
    name: basename(filePath),
    dataUrl: handOffBytes(bytes),
    hash,
    encrypted,
    recovered: recovered || undefined,
  }
}

// ---- IPC ----

const IMAGE_MIME: Record<string, 'image/png' | 'image/jpeg' | 'image/gif'> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
}

// ---- chat attachments: local files parsed for the agent ----

const ATTACHMENT_MAX_BYTES = 50 * 1024 * 1024
/** plain-text extensions read as UTF-8 */
const TEXT_EXTS = new Set([
  'txt',
  'md',
  'markdown',
  'csv',
  'tsv',
  'json',
  'yaml',
  'yml',
  'xml',
  'html',
  'htm',
  'log',
  'js',
  'ts',
  'tsx',
  'jsx',
  'py',
  'java',
  'c',
  'h',
  'cpp',
  'go',
  'rs',
  'rb',
  'sh',
  'sql',
  'css',
])
/** office/pdf formats get text extracted via @genoffice/file-parse; images skip extraction and go multimodal (files:read-image) */
const ATTACHMENT_EXTS = new Set([
  ...TEXT_EXTS,
  'doc',
  'docx',
  'pdf',
  'pptx',
  'ppt',
  'xlsx',
  'xlsm',
  'xls',
  ...ATTACHMENT_IMAGE_EXTS,
])

const ATTACHMENT_IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
}
/** multimodal size cap per image attachment (keeps the context from blowing up) */
const ATTACHMENT_IMAGE_MAX_BYTES = 5 * 1024 * 1024

/** extracted text cache keyed by path; invalidated by mtime+size */
const attachmentTextCache = new Map<string, { stamp: string; text: string }>()

function statAttachment(filePath: string): { meta?: AttachmentMeta; error?: string } {
  const name = basename(filePath)
  const ext = name.split('.').pop()?.toLowerCase() ?? ''
  if (!ATTACHMENT_EXTS.has(ext)) return { error: `${name}: ${tm('errUnsupportedExt', { ext })}` }
  try {
    const stat = statSync(filePath)
    if (!stat.isFile()) return { error: `${name}: ${tm('errNotFile')}` }
    if (stat.size > ATTACHMENT_MAX_BYTES) {
      return {
        error: `${name}: ${tm('errTooLarge', { mb: Math.round(ATTACHMENT_MAX_BYTES / 1024 / 1024) })}`,
      }
    }
    if (ATTACHMENT_IMAGE_EXTS.has(ext) && stat.size > ATTACHMENT_IMAGE_MAX_BYTES) {
      return { error: `${name}: ${tm('errImageTooLarge')}` }
    }
    return { meta: { path: filePath, name, ext, sizeBytes: stat.size } }
  } catch {
    return { error: `${name}: ${tm('errUnreadable')}` }
  }
}

function collectAttachments(paths: string[]): AttachmentAddResult {
  const accepted: AttachmentMeta[] = []
  const rejected: string[] = []
  for (const p of paths) {
    const { meta, error } = statAttachment(p)
    if (meta) accepted.push(meta)
    else if (error) rejected.push(error)
  }
  return { accepted, rejected }
}

/** save clipboard-pasted image bytes to a temp file (screenshots/bitmaps with no local path); returns null for non-images or empty data */
let pastedImageSeq = 0
let pastedDirPruned = false

/** drop pasted-image temp files older than 7 days (once per app run) */
function prunePastedImages(dir: string): void {
  if (pastedDirPruned) return
  pastedDirPruned = true
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000
  try {
    for (const name of readdirSync(dir)) {
      const p = join(dir, name)
      try {
        if (statSync(p).mtimeMs < cutoff) unlinkSync(p)
      } catch {
        // ignore: another tab may have removed it already
      }
    }
  } catch {
    // ignore: directory may not exist yet
  }
}
function savePastedImage(data: unknown, ext: unknown): string | null {
  const cleanExt = typeof ext === 'string' ? ext.toLowerCase() : ''
  if (!ATTACHMENT_IMAGE_EXTS.has(cleanExt)) return null
  const bytes =
    data instanceof ArrayBuffer
      ? Buffer.from(data)
      : ArrayBuffer.isView(data)
        ? Buffer.from(data.buffer, data.byteOffset, data.byteLength)
        : null
  if (!bytes || bytes.byteLength === 0) return null
  const dir = join(app.getPath('temp'), 'genoffice-pasted')
  mkdirSync(dir, { recursive: true })
  prunePastedImages(dir)
  const stamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '-')
  const filePath = join(dir, `pasted-${stamp}-${++pastedImageSeq}.${cleanExt}`)
  writeFileSync(filePath, bytes)
  return filePath
}

/** parse an attachment to text via @genoffice/file-parse (docx/pdf/pptx/xlsx/plain text) */
async function extractAttachmentText(filePath: string): Promise<string> {
  const stat = statSync(filePath)
  const stamp = `${stat.mtimeMs}:${stat.size}`
  const cached = attachmentTextCache.get(filePath)
  if (cached && cached.stamp === stamp) return cached.text
  if (stat.size > ATTACHMENT_MAX_BYTES) throw new Error(tm('errFileTooLarge'))
  const parsed = await parseFileToText(filePath)
  if (!parsed.ok || parsed.kind !== 'text' || parsed.text == null) {
    throw new Error(parsed.error ?? tm('errParseFailed'))
  }
  attachmentTextCache.set(filePath, { stamp, text: parsed.text })
  // keep the cache bounded (a handful of recent files is plenty)
  if (attachmentTextCache.size > 8) {
    const oldest = attachmentTextCache.keys().next().value
    if (oldest) attachmentTextCache.delete(oldest)
  }
  return parsed.text
}

// ---- print / export PDF ----

const TWIPS_PER_INCH = 1440

// ---- AI settings + chat proxy (main process avoids renderer CORS) ----
// provider metadata, settings defaults/migration, and per-provider streaming/chat
// implementations live in @genoffice/ai-provider, shared with apps/sheets.

const SETTINGS_PATH = () => userDataPath('ai-settings.json')

const activeAiStreams = new Map<string, AbortController>()

/** every renderer (tabs, home) re-reads ai-settings.json — the composer chip and
 *  the settings page edit the same file from different windows */
function broadcastAiSettingsChanged(): void {
  for (const wc of webContents.getAllWebContents())
    if (!wc.isDestroyed()) wc.send('ai:settings-changed')
}

/**
 * AI settings + chat/stream proxy handlers. Split out so the shell can
 * register them exactly once for all window types (docs, sheets, home) —
 * sheets' standalone AI handlers use the same channel names.
 */
export function registerAiIpc(): void {
  app.once('before-quit', shutdownCodexAppServers)
  ipcMain.handle('ai:get-settings', async (): Promise<AiSettings> => {
    const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(SETTINGS_PATH(), {})
    // pre-lock legacy file: genspark selected with cloud tools opted out. The
    // settings UI locks the tools switch on with genspark and apps read this
    // file live, so heal the stored flag once. Judged on the *stored* provider
    // — never the activeProvider fallback below, which must not leak into the
    // file and clobber a saved (half-configured) BYOK selection.
    if ((stored.provider ?? 'genspark') === 'genspark' && stored.gskToolsEnabled === false) {
      stored.gskToolsEnabled = true
      writeJsonAtomic(SETTINGS_PATH(), stored, 0o600)
    }
    // heal the file mode of a settings file saved by an older build: BYOK keys
    // are credentials and must not stay readable by other local accounts
    restrictFileToOwner(SETTINGS_PATH())
    const settings = resolveAiSettings(stored, defaultAiSettings())
    // a stored BYOK provider is honored when usable; half-filled configs fall back to genspark
    settings.provider = activeProvider(settings)
    return settings
  })

  // Genspark account (gsk login state): auth source for AI features; the frontend uses it to prompt login when logged out
  ipcMain.handle(
    'ai:gsk-status',
    async (_event, withEmail?: boolean): Promise<GenSparkAccountStatus> => {
      if (!hasGskAuth()) return { loggedIn: false }
      if (!withEmail) return { loggedIn: true }
      const info = await gskLoginInfo()
      return info?.email ? { loggedIn: true, email: info.email } : { loggedIn: true }
    },
  )

  // a 401 on the gsk key opens the device-code login once (genoffice-auth recovery)
  setGenofficeLoginOpener((url) => void shell.openExternal(url))

  ipcMain.handle('ai:gsk-login', () => {
    ensureGenofficeLogin((url) => void shell.openExternal(url))
  })

  ipcMain.handle('ai:set-settings', (_event, settings: AiSettings) => {
    // SECURITY.md: payloads are schema-checked in the main process. The settings
    // file feeds cliPath into spawn() and baseUrl receives the gsk bearer token,
    // so the renderer's copy is sanitized before it touches disk.
    const sanitized = sanitizeAiSettings(settings)
    if (!sanitized) {
      console.warn('[ai] rejected invalid ai:set-settings payload')
      return
    }
    writeJsonAtomic(SETTINGS_PATH(), sanitized, 0o600)
    broadcastAiSettingsChanged()
  })

  ipcMain.handle('ai:codex-models', async (_event, cliPath: unknown) => {
    // the probe spawns the path directly, so it gets the same metacharacter and
    // existence check as the stored setting (anything else: auto-detect)
    return listCodexModels(sanitizeCliPath(cliPath))
  })

  ipcMain.handle('ai:custom-models', (_event, input: unknown) => listCustomModelsForIpc(input))

  ipcMain.handle('ai:stream', async (event, request: AiStreamRequest) => {
    // per-request settings get the same schema check as the persisted ones: a
    // compromised renderer could otherwise hand cliPath/baseUrl straight to
    // the provider layer without ever touching the settings file
    const settings = sanitizeAiSettings(request.settings)
    if (!settings) {
      event.sender.send('ai:stream-chunk', {
        requestId: request.requestId,
        type: 'error',
        error: 'invalid AI settings payload',
      } satisfies AiStreamChunk)
      return
    }
    const { requestId, system, messages } = request
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? maxOutputTokensOf(settings)
    const provider = settings.provider
    let config = settings.providers?.[provider]
    // the genspark key never enters the settings file; requests take it from the gsk login state
    // set only when the gsk key was injected: a 401 then means that key expired
    let gskKey = ''
    if (provider === 'genspark' && config && !config.apiKey) {
      gskKey = gskApiKey()
      config = { ...config, apiKey: gskKey }
    }
    // a custom endpoint may name an environment variable instead of storing a key
    config = config ? resolveProfileCredential(config) : config
    const send = (chunk: AiStreamChunk) => {
      if (!event.sender.isDestroyed()) event.sender.send('ai:stream-chunk', chunk)
    }
    if (!config || (provider !== 'codex' && !config.apiKey)) {
      send({
        requestId,
        type: 'error',
        error: provider === 'genspark' ? tm('errGskNotLoggedIn') : tm('errNoApiKey', { provider }),
      })
      return
    }
    if (provider !== 'codex' && !config.model) {
      send({ requestId, type: 'error', error: tm('errNoModel') })
      return
    }
    const controller = new AbortController()
    activeAiStreams.set(requestId, controller)
    const unwatchSender = abortOnDestroyed(event.sender, controller)
    // wire-activity keepalive: lets the renderer's silence watchdog tell a slow turn from a dead one
    let lastPing = 0
    const ping = () => {
      const now = Date.now()
      if (now - lastPing < 5_000) return
      lastPing = now
      send({ requestId, type: 'ping' })
    }
    const trace = startAiTrace({
      requestId,
      sessionId: request.sessionId,
      source: event.sender.getURL(),
      provider,
      model: config.model,
      maxTokens,
      system,
      messages,
      tools,
    })
    try {
      let stopReason: string | undefined
      await streamForProvider(provider, config, system, messages, tools, maxTokens, {
        ...(request.sessionId ? { sessionId: request.sessionId } : {}),
        signal: controller.signal,
        onDelta: (text) => {
          trace?.onDelta(text)
          send({ requestId, type: 'delta', text })
        },
        onReasoningDelta: (text) => {
          trace?.onReasoningDelta(text)
          send({ requestId, type: 'reasoning', text })
        },
        onToolCall: (toolCall) => {
          trace?.onToolCall(toolCall)
          send({ requestId, type: 'tool-call', toolCall })
        },
        onActivity: ping,
        onStopReason: (reason) => {
          stopReason = reason
          trace?.onStopReason(reason)
        },
        onUsage: (usage) => {
          trace?.onUsage(usage)
          send({ requestId, type: 'usage', usage })
        },
      })
      trace?.finish()
      send({ requestId, type: 'done', stopReason })
    } catch (err) {
      trace?.finish(
        controller.signal.aborted ? 'aborted' : err instanceof Error ? err.message : String(err),
      )
      if (controller.signal.aborted) {
        send({ requestId, type: 'done' })
      } else {
        // an expired/revoked key drops to signed-out so the inline sign-in button shows
        noteGskCallFailure(gskKey, err)
        send({
          requestId,
          type: 'error',
          error: err instanceof Error ? err.message : String(err),
          ...(err instanceof AiTimeoutError
            ? { errorCode: 'timeout' as const }
            : err instanceof AiCreditsError
              ? { errorCode: 'credits' as const }
              : isAiNetworkError(err)
                ? { errorCode: 'network' as const }
                : isAiOverloadedError(err)
                  ? { errorCode: 'overloaded' as const }
                  : isAiRegionBlockedError(err)
                    ? { errorCode: 'region' as const }
                    : {}),
        })
      }
    } finally {
      unwatchSender()
      activeAiStreams.delete(requestId)
    }
  })

  ipcMain.handle('ai:stream-cancel', (_event, requestId: string) => {
    activeAiStreams.get(requestId)?.abort()
  })

  // shared search tools (content + images): Serper with DuckDuckGo fallback (same source as slides/sheets)
  ipcMain.handle('ai:web-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await webSearchTool(
        SETTINGS_PATH(),
        String(query),
        typeof maxResults === 'number' ? maxResults : 6,
      )
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })
  ipcMain.handle('ai:image-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await imageSearchTool(
        SETTINGS_PATH(),
        String(query),
        typeof maxResults === 'number' ? maxResults : 8,
      )
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })

  // media understanding (pictures in the document, attachments, local files): BYOK media
  // provider when one is configured, otherwise the Genspark CLI behind its login gate.
  // docs-prefixed: slides registers its own ai:analyze-media in the same shell process.
  ipcMain.handle(
    'docs:analyze-media',
    async (event, op: { mediaUrls: string[]; requirements: string }) => {
      const mediaUrls = (op.mediaUrls ?? []).map(String).filter(Boolean)
      // a picture opened lazily from a large docx is only addressable by its main-process
      // store; hand its bytes over as a data URL so the loader can read them like any other
      const resolved: string[] = []
      for (const url of mediaUrls) {
        const lazy = await readLazyMedia(url).catch(() => null)
        resolved.push(lazy ? `data:${lazy.mime};base64,${lazy.body.toString('base64')}` : url)
      }
      return analyzeMediaTool(
        SETTINGS_PATH(),
        {
          mediaUrls: resolved,
          requirements: String(op.requirements ?? ''),
        },
        { mediaRoots: docsMediaRoots(event.sender.id) },
      )
    },
  )

  // download image from URL → base64+mime (download in the main process avoids CORS; the renderer builds the image node and measures size itself)
  ipcMain.handle(
    'ai:fetch-image',
    async (_event, url: string): Promise<{ base64: string; mime: string } | null> => {
      try {
        // the URL originates from AI tool calls (prompt-injectable via web search
        // results), so refuse non-http schemes and private/link-local targets;
        // redirects are followed manually so every hop is validated too.
        // fetchRemoteImage adds CDN-friendly headers and transient-error retries.
        const resp = await fetchRemoteImage(String(url))
        if (!resp || !resp.ok) return null
        const buf = Buffer.from(await readBodyCapped(resp, MAX_REMOTE_IMAGE_BYTES))
        const ct = resp.headers.get('content-type') ?? ''
        const mime = ct.includes('png')
          ? 'image/png'
          : ct.includes('gif')
            ? 'image/gif'
            : 'image/jpeg'
        return { base64: buf.toString('base64'), mime }
      } catch {
        return null
      }
    },
  )

  // docs-owned (like pdf:generate-image): slides' ai:generate-image is only
  // registered once a slides view exists, so docs needs its own channel
  ipcMain.handle(
    'docs:ai-generate-image',
    (event, op: { prompt?: unknown; aspectRatio?: unknown }) =>
      generateImageTool(
        SETTINGS_PATH(),
        {
          prompt: String(op?.prompt ?? ''),
          aspectRatio: op?.aspectRatio ? String(op.aspectRatio) : undefined,
        },
        { mediaRoots: docsMediaRoots(event.sender.id) },
      ),
  )

  ipcMain.handle('ai:search-test', (_event, input: unknown) => {
    const { provider, apiKey } = (input ?? {}) as { provider?: AiSearchProviderId; apiKey?: string }
    if (!provider || provider === 'genspark') {
      return hasGskAuth() ? { ok: true } : { ok: false, error: tm('errGskNotLoggedIn') }
    }
    return testSearchProvider(provider, String(apiKey ?? ''))
  })

  // settings-UI connection test for the media provider (genspark = the gsk login state)
  ipcMain.handle('ai:media-test', (_event, input: unknown) => {
    const { provider, config } = (input ?? {}) as {
      provider?: AiMediaProviderId
      config?: AiMediaProviderConfig
    }
    if (!provider || provider === 'genspark') {
      return hasGskAuth() ? { ok: true } : { ok: false, error: tm('errGskNotLoggedIn') }
    }
    if (!config) return { ok: false, error: 'No media provider configuration' }
    return testMediaProvider(provider, config)
  })

  ipcMain.handle('ai:chat', async (_event, request: AiChatRequest) => {
    // same schema check as ai:stream: one-shot requests would otherwise act on
    // the renderer's settings copy verbatim
    const settings = sanitizeAiSettings(request.settings)
    if (!settings) return { ok: false, error: 'invalid AI settings payload' }
    const { system, user } = request
    const provider = settings.provider
    let config = settings.providers?.[provider]
    // set only when the gsk key was injected: a 401 then means that key expired
    let gskKey = ''
    if (provider === 'genspark' && config && !config.apiKey) {
      gskKey = gskApiKey()
      config = { ...config, apiKey: gskKey }
    }
    // a custom endpoint may name an environment variable instead of storing a key
    config = config ? resolveProfileCredential(config) : config
    if (!config || (provider !== 'codex' && !config.apiKey)) {
      return {
        ok: false,
        error: provider === 'genspark' ? tm('errGskNotLoggedIn') : tm('errNoApiKey', { provider }),
      }
    }
    if (provider !== 'codex' && !config.model) return { ok: false, error: tm('errNoModel') }
    try {
      const result = await chatForProvider(provider, config, system, user)
      // the one-shot path reports HTTP failures as ok:false with the raw body —
      // replace capacity/rate-limit dumps with the localized "busy" message
      if (!result.ok && isAiOverloadedError(result.error)) {
        return { ok: false, error: tm('errAiBusy') }
      }
      if (!result.ok && isAiRegionBlockedError(result.error)) {
        return { ok: false, error: tm('errAiRegion') }
      }
      if (!result.ok) noteGskCallFailure(gskKey, result.error)
      return result
    } catch (err) {
      noteGskCallFailure(gskKey, err)
      return {
        ok: false,
        error: isAiOverloadedError(err)
          ? tm('errAiBusy')
          : isAiRegionBlockedError(err)
            ? tm('errAiRegion')
            : String(err),
      }
    }
  })
}

// ── project-store IPC (shared across docs / slides / sheets) ──────────────

let projectStore: ProjectStore | null = null
let projectIpcRegistered = false

function getProjectStore(): ProjectStore {
  if (!projectStore) projectStore = new ProjectStore(app.getPath('userData'))
  return projectStore
}

/**
 * Fired when a save lands on a new path (save-as / first silent save). The shell
 * uses it to sync the tab title/path, record recents and apply a pending project —
 * same contract as the sheets/slides opened hooks. Never called standalone.
 * Returns the final path when the shell filed the new file into a Home folder.
 */
let fileSavedHook: ((wc: WebContents, filePath: string) => string | void) | null = null

export function setDocsFileSavedHook(
  hook: (wc: WebContents, filePath: string) => string | void,
): void {
  fileSavedHook = hook
}

function notifyFileSaved(wc: WebContents, filePath: string): string {
  const moved = fileSavedHook ? fileSavedHook(wc, filePath) : undefined
  return typeof moved === 'string' && moved ? moved : filePath
}

/**
 * Fired when a docx is opened INSIDE an existing tab (File > Open dialog or an
 * explicit path open). Sheets and slides have had this hook from the start;
 * docs only synced the tab on save-as/first-save, so a file opened into an
 * untitled tab kept the "Untitled Document" tab title until a save landed on a
 * NEW path — which a plain Ctrl+S to the original file never does (r115).
 */
let fileOpenedHook: ((wcId: number, filePath: string) => void) | null = null

export function setDocsFileOpenedHook(hook: (wcId: number, filePath: string) => void): void {
  fileOpenedHook = hook
}

/**
 * Reverse lookup from a sheets sessionId to its file path. In shell mode the
 * project:* handlers are registered by this file, but only sheets-main knows the
 * sessionId mapping; the shell injects it at startup (standalone docs doesn't need it).
 */
let sessionPathResolver: ((senderId: number, sessionId: string) => string | null) | null = null

export function setSessionPathResolver(
  fn: (senderId: number, sessionId: string) => string | null,
): void {
  sessionPathResolver = fn
}

/** After a file is renamed/moved on disk, sync project-store (fileMap/chatIdByPath re-key accordingly; history follows the file). */
export function projectFilePaths(): string[] {
  try {
    return getProjectStore().knownFilePaths()
  } catch {
    return []
  }
}

export function projectFileRenamed(oldPath: string, newPath: string): void {
  try {
    getProjectStore().fileRenamed(oldPath, newPath)
  } catch (err) {
    console.warn('[project-store] fileRenamed failed:', err)
  }
}

/**
 * Register the project:* IPC handlers (all three apps share the same channel names).
 * Idempotency guard: registered only once in shell mode.
 */
export function registerProjectIpc(): void {
  if (projectIpcRegistered) return
  projectIpcRegistered = true

  /** Resolve projectId + chatId from a file path (sheets without a path resolves via sessionId) */
  ipcMain.handle(
    'project:resolveChat',
    (event, args: { filePath: string | null; tempChatId?: string; sessionId?: string }) => {
      const store = getProjectStore()
      store.ensureDefaultProject()
      let resolvedPath = args.filePath
      if (!resolvedPath && args.sessionId && sessionPathResolver) {
        resolvedPath = sessionPathResolver(event.sender.id, args.sessionId)
      }
      if (!resolvedPath) {
        return {
          projectId: 'default',
          chatId: args.tempChatId ?? `unsaved-${Date.now()}`,
        }
      }
      return store.resolveChatForFile(resolvedPath)
    },
  )

  /** Append a message */
  ipcMain.handle(
    'project:appendChat',
    (
      _event,
      args: {
        projectId: string
        chatId: string
        role: 'user' | 'assistant'
        text: string
        tools?: Array<{
          name: string
          summary: string
          isError?: boolean
          input?: string
          output?: string
        }>
        attachments?: Array<{ name: string; path?: string; ext?: string; sizeBytes?: number }>
        scope?: { label: string; text?: string }
      },
    ) => {
      if (args.role !== 'user' && args.role !== 'assistant') {
        throw new Error(`Invalid chat role: ${String(args.role)}`)
      }
      if (typeof args.text !== 'string' || args.text.length > 200_000) {
        throw new Error('Invalid chat text: must be a string up to 200000 chars')
      }
      if (args.tools && !Array.isArray(args.tools)) throw new Error('Invalid chat tools')
      if (args.attachments && !Array.isArray(args.attachments)) {
        throw new Error('Invalid chat attachments')
      }
      const store = getProjectStore()
      const msg: Parameters<ProjectStore['appendChatMessage']>[2] = {
        role: args.role,
        text: args.text,
      }
      if (args.tools) msg.tools = args.tools
      if (args.attachments) msg.attachments = args.attachments
      if (args.scope) msg.scope = args.scope

      store.appendChatMessage(args.projectId, args.chatId, msg)
    },
  )

  /** Read history */
  ipcMain.handle(
    'project:loadChat',
    (
      _event,
      args: {
        projectId: string
        chatId: string
        limit?: number
      },
    ) => {
      const store = getProjectStore()
      return store.loadChat(args.projectId, args.chatId, args.limit ?? 200)
    },
  )

  /** Lists the conversations in a project, newest first. */
  ipcMain.handle('project:listChats', (_event, args: { projectId: string }) => {
    const store = getProjectStore()
    return store.listChats(args.projectId)
  })

  /** Deletes one conversation. A file-less chat is the home screen's ask panel. */
  ipcMain.handle('project:deleteChat', (_event, args: { projectId: string; chatId: string }) => {
    const store = getProjectStore()
    store.deleteChat(args.projectId, args.chatId)
  })

  /** rebind chat (called after a file first hits disk): newFilePath/sessionId take priority; the main process computes chatId and records fileMap */
  ipcMain.handle(
    'project:rebindChat',
    (
      event,
      args: {
        projectId: string
        tempChatId: string
        newChatId?: string
        newFilePath?: string
        sessionId?: string
      },
    ) => {
      const store = getProjectStore()
      let path = args.newFilePath ?? null
      if (!path && args.sessionId && sessionPathResolver) {
        path = sessionPathResolver(event.sender.id, args.sessionId)
      }
      if (path) {
        return store.rebindChatToFile(args.projectId, args.tempChatId, path)
      }
      if (args.newChatId) store.rebindChat(args.projectId, args.tempChatId, args.newChatId)
      return { projectId: args.projectId, chatId: args.newChatId ?? args.tempChatId }
    },
  )
}

/** A4 at 96dpi, as the HTML app exports */
const ALT_CHUNK_VIEWPORT = { width: 794, height: 1123, deviceScaleFactor: 2 }
const ALT_CHUNK_HTML_MAX_CHARS = 64 * 1024 * 1024
// An AI-generated page whose scripts never yield must not strand the hidden
// conversion window; the slides export path uses the same shape.
const ALT_CHUNK_TIMEOUT_MS = 120_000

/** an encrypted save leaves no plain file to serve lazy pictures from: the
 *  renderer takes the materialized document back and leaves lazy mode */
const reissuedDoc = (
  encrypted: boolean,
  hashes: Set<string>,
  plain: Buffer,
): { dataUrl?: string } => (encrypted && hashes.size > 0 ? { dataUrl: handOffBytes(plain) } : {})

/** document/attachment/window IPC (everything except the AI proxy above) */
export function registerDocsIpc(): void {
  registerZoteroIpc()
  void app.whenReady().then(registerLazyMediaProtocol)
  // Node fetch (undici) direct connections get reset under VPN/tun setups; retry over Chromium's stack
  setRescueFetch((url, init) => net.fetch(url, init))
  setAiUserAgent(`GenOffice/${app.getVersion()}`)

  // shared with the other editor modules — last (identical) registration wins
  ipcMain.removeHandler('app:get-language')
  ipcMain.handle('app:get-language', () => getUiLang())
  ipcMain.handle('docs:confirm-document-replace', (event) =>
    requestDocsClose(event.sender, dialogParent(event)),
  )
  ipcMain.handle('docs:system-locale', () => app.getSystemLocale())

  configureMetricsCache(userDataPath('font-metrics'))
  ipcMain.handle('docs:font-metrics', (_event, family: string) =>
    typeof family === 'string' ? familyVerticalMetrics(family) : null,
  )
  // the other spellings of the installed faces behind these names: a document
  // that declares a CJK font by its localized name needs the English spelling
  // the browser resolves (genoffice#2449). The renderer is not trusted, but
  // the answer is only font names.
  ipcMain.handle('docs:font-aliases', (_event, names: unknown) => {
    if (!Array.isArray(names)) return {}
    const out: Record<string, readonly string[]> = {}
    for (const raw of names.slice(0, 512)) {
      if (typeof raw !== 'string' || raw.length === 0 || raw.length > 256) continue
      const aliases = familyAliases(raw)
      if (aliases.length) out[raw] = aliases
    }
    return out
  })

  ipcMain.handle('docs:open', async (event) => {
    const result = await openDialog(event, {
      title: tm('dlgOpenDoc'),
      filters: [{ name: tm('filterWord'), extensions: [...OPENABLE_DOC_EXTENSIONS] }],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    const prepared = await prepareDocumentForOpen(result.filePaths[0], dialogParent(event))
    if (!prepared) return null
    return loadDocx(prepared, event.sender.id)
  })

  ipcMain.handle('docs:open-path', async (event, filePath: string) => {
    // renderer-facing open channel: the File > Open Recent menu echoes paths
    // main itself handed out (recent.json), and the shared ledger carries
    // anything this view's user picked or dropped. e2e drives fixture paths.
    if (
      process.env.GENOFFICE_E2E !== '1' &&
      !readRecentFiles().includes(filePath) &&
      !sharedPathGrants().has(event.sender.id, filePath)
    )
      return null
    const prepared = await prepareDocumentForOpen(filePath, dialogParent(event))
    if (!prepared) return null
    return loadDocx(prepared, event.sender.id)
  })

  // w:altChunk HTML: the same html2docx chain as the HTML app's export, in a
  // hidden window; the renderer parses the result and shows its blocks
  ipcMain.handle('docs:altchunk-html-to-docx', async (_event, html: unknown) => {
    if (typeof html !== 'string' || !html.trim() || html.length > ALT_CHUNK_HTML_MAX_CHARS) {
      return null
    }
    const workDir = await mkdtemp(join(tmpdir(), 'genoffice-altchunk-'))
    let driver: ElectronBrowserDriver | null = null
    try {
      const htmlPath = join(workDir, 'chunk.html')
      // the BOM outranks a stale <meta charset> left in the decoded markup
      await writeFile(htmlPath, `\ufeff${html}`, 'utf8')
      driver = await ElectronBrowserDriver.create(ALT_CHUNK_VIEWPORT)
      // The markup is an unsanitised AI artifact: a script that never yields
      // would otherwise keep executeJavaScript pending forever, and the
      // finally below would never run (the hidden window and workDir leak for
      // good). Race a watchdog and destroy the window on timeout, matching
      // the slides export guard.
      const conversion = convertHtmlToDocx({ url: pathToFileURL(htmlPath).href }, driver, {
        naturalTableWidth: true,
      }).then(({ docx }) => docx)
      let watchdog: ReturnType<typeof setTimeout> | undefined
      const docx = await Promise.race([
        conversion,
        new Promise<null>((resolve) => {
          watchdog = setTimeout(() => {
            if (driver && !driver.isWindowDestroyed()) driver.destroyNow()
            driver = null
            console.warn('[docs] altChunk conversion timed out; window destroyed')
            resolve(null)
          }, ALT_CHUNK_TIMEOUT_MS)
        }),
      ]).finally(() => clearTimeout(watchdog))
      return docx
    } catch (err) {
      console.warn('[docs] altChunk conversion failed:', err)
      return null
    } finally {
      await driver?.close()
      await rm(workDir, { recursive: true, force: true }).catch(() => {})
    }
  })

  // Review > Protect > Encrypt with Password: set/clear the open password.
  // Takes effect on the next save (docs:save / save-as / save-new all consult the store).
  ipcMain.handle(
    'docs:set-password',
    (event, filePath: string | null, password: string | null): { ok: boolean } => {
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      if (filePath !== null && typeof filePath !== 'string') return { ok: false }
      if (password !== null && (typeof password !== 'string' || password.length === 0)) {
        return { ok: false }
      }
      // only the document this renderer legitimately has open (same grant as saving)
      if (filePath && !canDocWrite(event.sender.id, filePath)) return { ok: false }
      setDocPassword(event.sender.id, filePath, password)
      return { ok: true }
    },
  )

  ipcMain.handle('docs:password-intent-revision', (event): number => {
    if (tornDownWcIds.has(event.sender.id)) return -1
    return currentDocPasswordIntentRevision()
  })

  ipcMain.handle(
    'docs:discard-password-intents',
    (event, throughRevision: unknown): { ok: boolean } => {
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      if (
        typeof throughRevision !== 'number' ||
        !Number.isSafeInteger(throughRevision) ||
        throughRevision < 0
      ) {
        return { ok: false }
      }
      discardDocPasswordIntents(event.sender.id, throughRevision)
      return { ok: true }
    },
  )

  // decrypt-and-open a password-protected docx; wrong-password keeps the renderer's prompt open
  ipcMain.handle(
    'docs:open-decrypt',
    async (event, filePath: string, password: string): Promise<DecryptOpenResult> => {
      if (typeof filePath !== 'string' || typeof password !== 'string' || password.length === 0) {
        return { ok: false, reason: 'error', error: 'invalid arguments' }
      }
      // the retry's path must be the one main itself flagged as needing a
      // password for this view — not any path the renderer names
      if (pendingDecryptByWc.get(event.sender.id) !== filePath) {
        return { ok: false, reason: 'error', error: 'docs: path not pending decrypt' }
      }
      try {
        const result = await loadDocx(filePath, event.sender.id, password)
        if (!result) return { ok: false, reason: 'error', error: 'file not found' }
        // the file was swapped for a plain docx between prompt and submit — still an open
        if ('needsPassword' in result) return { ok: false, reason: 'error', error: 'not encrypted' }
        return { ok: true, result }
      } catch (err) {
        if (err instanceof DocxDecryptError) {
          return { ok: false, reason: err.reason, error: err.message }
        }
        return { ok: false, reason: 'error', error: String(err) }
      }
    },
  )

  ipcMain.handle('docs:consume-pending-open', async (event) => {
    rendererReady = true
    const made = pendingResultDocs.get(event.sender.id)
    if (made) {
      pendingResultDocs.delete(event.sender.id)
      return {
        path: '',
        name: '',
        dataUrl: handOffBytes(made.bytes),
        hash: '',
        template: true,
        unsaved: true,
      }
    }
    // a tab spawned via New Tab loads the document queued for it specifically
    const queued = pendingWindowOpens.get(event.sender.id)
    if (queued) {
      pendingWindowOpens.delete(event.sender.id)
      const prepared = await prepareDocumentForOpen(queued, dialogParent(event))
      if (!prepared) return null
      return loadDocx(prepared, event.sender.id)
    }
    const filePath = pendingOpenPath
    pendingOpenPath = null
    if (!filePath) return null
    const prepared = await prepareDocumentForOpen(filePath, dialogParent(event))
    if (!prepared) return null
    return loadDocx(prepared, event.sender.id)
  })

  /** returns true when this tab was opened via "New Document" and should start blank */
  ipcMain.handle('docs:consume-new-blank', (event) => {
    rendererReady = true
    if (pendingNewBlankIds.has(event.sender.id)) {
      pendingNewBlankIds.delete(event.sender.id)
      return true
    }
    return false
  })

  /** one-shot AI content queued by create_document for this tab; null when none */
  ipcMain.handle('docs:consume-ai-doc-content', (event): AiDocContent | null => {
    const content = pendingAiDocContents.get(event.sender.id) ?? null
    pendingAiDocContents.delete(event.sender.id)
    return content
  })

  // ---- headless export mode (--headless-export) ----

  ipcMain.handle('docs:consume-headless-export', (event): HeadlessExportTarget | null => {
    const target = headlessExportTargets.get(event.sender.id) ?? null
    headlessExportTargets.delete(event.sender.id)
    return target
  })

  ipcMain.on('docs:headless-export-done', (event, result: unknown) => {
    const settle = headlessExportWaiters.get(event.sender.id)
    if (!settle) return
    headlessExportWaiters.delete(event.sender.id)
    const state = result as { ok?: unknown; error?: unknown } | null
    settle({
      ok: state?.ok === true,
      ...(typeof state?.error === 'string' ? { error: state.error } : {}),
    })
  })

  ipcMain.handle(
    'docs:save',
    async (event, filePath: string, data: ArrayBuffer, auto?: boolean) => {
      try {
        // only paths this renderer opened or chose via save-as may be overwritten
        if (typeof filePath !== 'string' || !canDocWrite(event.sender.id, filePath)) {
          return { ok: false, error: 'save target is not an opened document' }
        }
        if (await diskChangedExternally(event.sender.id, filePath)) {
          // autosave must never clobber another program's edits silently; the
          // renderer stays dirty and the next manual save raises the dialog
          if (auto === true) return { ok: false, reason: 'external-modified' }
          const options = {
            type: 'warning' as const,
            message: tm('extModifiedMsg'),
            detail: tm('extModifiedDetail'),
            buttons: [tm('btnOverwrite'), tm('btnCancel')],
            defaultId: 0,
            cancelId: 1,
            noLink: true,
          }
          const parent = dialogParent(event)
          const { response } =
            parent && !parent.isDestroyed()
              ? await dialog.showMessageBox(parent, options)
              : await dialog.showMessageBox(options)
          if (response !== 0) return { ok: false, reason: 'external-modified' }
        }
        // The tab may have been closed while the external-change check or the
        // overwrite prompt was pending: a Don't Save close revoked this tab's
        // grants, and landing the write now would persist discarded edits.
        if (tornDownWcIds.has(event.sender.id) || !canDocWrite(event.sender.id, filePath)) {
          return { ok: false, error: 'save target is not an opened document' }
        }
        // Snapshot desired state: the disk password remains unchanged until the
        // atomic write succeeds, and a newer ribbon intent survives this save.
        const passwordState = snapshotDocPassword(event.sender.id, filePath)
        const { bytes: plain, hashes } = await materializeLazyDocx(Buffer.from(data))
        const bytes = passwordState.password ? encryptDocx(plain, passwordState.password) : plain
        await atomicWriteFile(filePath, bytes)
        // Teardown may have cleared all in-memory secrets while the atomic
        // write was pending. Never resurrect state for an orphaned renderer.
        if (tornDownWcIds.has(event.sender.id)) {
          return { ok: false, error: 'save target is not an opened document' }
        }
        await rememberDiskState(event.sender.id, filePath, sha256Hex(bytes))
        if (tornDownWcIds.has(event.sender.id)) {
          return { ok: false, error: 'save target is not an opened document' }
        }
        pointLazyMediaAt(
          hashes,
          filePath,
          event.sender.id,
          passwordState.password ? plain : undefined,
        )
        // Commit immediately after the final await: intents received during
        // post-write bookkeeping are included, with no later async race.
        const passwordIntentPending = commitDocPasswordSave(
          event.sender.id,
          passwordState,
          filePath,
        )
        clearRecoveryCopy(filePath)
        pushRecent(filePath)
        return {
          ok: true,
          passwordIntentPending,
          ...reissuedDoc(!!passwordState.password, hashes, plain),
        }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  // crash-recovery copy from a dirty renderer; best-effort, never surfaces
  ipcMain.handle('docs:write-recovery', async (event, filePath: string, data: ArrayBuffer) => {
    try {
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      if (typeof filePath !== 'string' || !canDocWrite(event.sender.id, filePath))
        return { ok: false }
      // snapshot before any await: a save or discard that clears the recovery
      // copy while this write is in flight bumps the epoch and invalidates it
      const epoch = recoveryClearEpochs.get(filePath) ?? 0
      await mkdir(recoveryDir(), { recursive: true })
      // Recovery follows the current disk state, never the desired next-save
      // password. Missing state for an encrypted disk file skips the tick so
      // plaintext can never be written as its recovery copy.
      const bytes = prepareRecoveryDocx(event.sender.id, filePath, Buffer.from(data))
      if (!bytes) return { ok: false }
      await atomicWriteFile(recoveryPathFor(filePath), bytes)
      // The tab may have been closed ("Don't Save" clears the copy, teardown
      // revokes access) or the document saved (docs:save clears the copy) while
      // the write was in flight. A write that lost either race would offer
      // discarded or already-saved content as recovery on the next open — undo it.
      if (
        tornDownWcIds.has(event.sender.id) ||
        !canDocWrite(event.sender.id, filePath) ||
        (recoveryClearEpochs.get(filePath) ?? 0) !== epoch
      ) {
        clearRecoveryCopy(filePath)
        return { ok: false }
      }
      return { ok: true }
    } catch {
      return { ok: false }
    }
  })

  // Blink only respells an editable as a consequence of real (trusted) typing
  // of a word-committing character inside it: attribute flips, focus cycles,
  // script selection moves, execCommand edits, fresh DOM nodes, synthetic
  // clicks/arrow keys — and even a typed zero-width space — all leave existing
  // typos unmarked (each verified pixel-by-pixel).
  // Type one trusted space; the RENDERER removes it again by script (a
  // trusted Backspace would work too, but its deletion re-suppresses the
  // caret paragraph and that line stays unmarked) with ProseMirror's DOM
  // observer paused, so the round trip never becomes a transaction.
  // spell-diag trace (intermittent squiggle loss, platform-bound
  // and unreproducible on demand) — a tiny always-on log support can ask for.
  // Size-capped: over 256KB the file restarts from its last half.
  ipcMain.on('docs:spell-diag', (_event, line: unknown) => {
    if (typeof line !== 'string' || line.length > 500) return
    try {
      const path = userDataPath('spell-diag.log')
      if (existsSync(path) && statSync(path).size > 256 * 1024) {
        const tail = readFileSync(path, 'utf-8').slice(-128 * 1024)
        writeFileSync(path, tail.slice(tail.indexOf('\n') + 1))
      }
      appendFileSync(path, `${new Date().toISOString()} ${line}\n`)
    } catch {
      // diagnostics must never break the app
    }
  })

  // The document body draws its own React context menu, but Chromium's
  // misspelling + suggestions for the clicked word only surface in the main
  // process `context-menu` event. The renderer claims each body right-click
  // synchronously from its DOM handler, i.e. before Blink requests the menu,
  // so claims and events arrive in the same order: a claimed click gets its
  // data forwarded and no native menu, anything else (header/footer surfaces,
  // inputs) pops the native menu as before.
  const ctxMenuClaims = new Map<number, ClickClaims>()
  ipcMain.on('docs:context-menu-claim', (event, seq: unknown) => {
    event.returnValue = true
    if (typeof seq !== 'number') return
    let claims = ctxMenuClaims.get(event.sender.id)
    if (!claims) {
      claims = new ClickClaims()
      ctxMenuClaims.set(event.sender.id, claims)
    }
    claims.claim(seq, Date.now())
  })
  ipcMain.on('docs:context-menu-arm', (event) => {
    const wc = event.sender
    setContextMenuInterceptor(app, wc, (contents, params) => {
      if (contents.isDestroyed() || tornDownWcIds.has(contents.id)) return Promise.resolve(false)
      const seq = ctxMenuClaims.get(contents.id)?.take(Date.now()) ?? null
      if (seq !== null) {
        const request: ContextMenuRequest = {
          seq,
          misspelledWord: params.misspelledWord,
          suggestions: params.dictionarySuggestions,
        }
        contents.send('docs:context-menu', request)
      }
      return Promise.resolve(seq !== null)
    })
    wc.once('destroyed', () => {
      setContextMenuInterceptor(app, wc, null)
      ctxMenuClaims.delete(wc.id)
    })
  })
  void app.whenReady().then(() => {
    const orphans = readJson<string[]>(SPELL_IGNORED_PATH(), [])
    for (const w of orphans) session.defaultSession.removeWordFromSpellCheckerDictionary(w)
    if (orphans.length) journalIgnoredWords()
  })
  ipcMain.handle('docs:spell-ignore-word', (event, word: unknown) => {
    if (typeof word !== 'string' || !word.trim()) return false
    const w = word.trim()
    let holders = spellIgnored.get(w)
    if (!holders) {
      holders = new Set()
      spellIgnored.set(w, holders)
      journalIgnoredWords()
    }
    holders.add(event.sender.id)
    return event.sender.session.addWordToSpellCheckerDictionary(w)
  })
  ipcMain.handle('docs:spell-reset-ignored', (event) => releaseSpellIgnores(event.sender.id))
  ipcMain.handle('docs:spell-add-word', (event, word: unknown) => {
    if (typeof word !== 'string' || !word.trim()) return false
    const w = word.trim()
    if (spellIgnored.delete(w)) journalIgnoredWords()
    return event.sender.session.addWordToSpellCheckerDictionary(w)
  })
  ipcMain.handle('docs:spell-replace', (event, word: unknown) => {
    if (typeof word === 'string' && word) event.sender.replaceMisspelling(word)
  })
  ipcMain.handle('docs:spell-languages', (event): SpellLanguages => {
    const session = event.sender.session
    return {
      active: session.getSpellCheckerLanguages(),
      available: session.availableSpellCheckerLanguages,
    }
  })
  ipcMain.handle('docs:spell-set-languages', (event, langs: unknown): SpellLanguages => {
    const session = event.sender.session
    const available = new Set(session.availableSpellCheckerLanguages)
    const next = Array.isArray(langs)
      ? langs.filter((l): l is string => typeof l === 'string' && available.has(l))
      : []
    if (next.length) session.setSpellCheckerLanguages(next)
    return { active: session.getSpellCheckerLanguages(), available: [...available] }
  })

  ipcMain.handle('docs:respell-kick', async (event) => {
    const wc = event.sender
    if (tornDownWcIds.has(wc.id) || wc.isDestroyed()) return
    wc.focus()
    // Blink only respells after a user activation, and only a keydown grants one
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Space' })
    wc.sendInputEvent({ type: 'char', keyCode: ' ' })
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Space' })
    // resolve only after the input pipeline has delivered the keystroke, so
    // the caller can scrub the space it produced
    await new Promise((r) => setTimeout(r, 120))
  })

  ipcMain.handle(
    'docs:save-as',
    async (event, defaultName: string, data: ArrayBuffer, sourcePath?: string | null) => {
      // an orphaned (closed-tab) renderer must not open dialogs or land new files
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      const result = await saveDialog(event, {
        title: tm('dlgSaveAs'),
        defaultPath: saveAsSuggestion(
          typeof sourcePath === 'string' ? sourcePath : null,
          defaultName,
        ),
        filters: [{ name: tm('filterWord'), extensions: ['docx'] }, ...SAVE_AS_EXPORT_FILTERS],
      })
      if (result.canceled || !result.filePath) return { ok: false }
      // the tab may have been closed while the dialog was open; checked before the
      // write because Save As may overwrite an existing file (no safe rollback)
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      // another format is an exported copy: the renderer runs that export, the document stays where it is
      const exportAs = saveAsExportFormat(result.filePath)
      if (exportAs) {
        allowPdfWrite(event.sender.id, result.filePath)
        return { ok: false, exportAs, path: result.filePath }
      }
      try {
        const passwordState = snapshotDocPassword(
          event.sender.id,
          typeof sourcePath === 'string' && sourcePath ? sourcePath : null,
        )
        const { bytes: plain, hashes } = await materializeLazyDocx(Buffer.from(data))
        const bytes = passwordState.password ? encryptDocx(plain, passwordState.password) : plain
        await atomicWriteFile(result.filePath, bytes)
        if (tornDownWcIds.has(event.sender.id)) return { ok: false }
        allowDocWrite(event.sender.id, result.filePath)
        rememberOpenDoc(event.sender.id, result.filePath)
        await rememberDiskState(event.sender.id, result.filePath, sha256Hex(bytes))
        pointLazyMediaAt(
          hashes,
          result.filePath,
          event.sender.id,
          passwordState.password ? plain : undefined,
        )
        if (tornDownWcIds.has(event.sender.id)) return { ok: false }
        const passwordIntentPending = commitDocPasswordSave(
          event.sender.id,
          passwordState,
          result.filePath,
        )
        pushRecent(result.filePath)
        // the renderer has no path yet to match a rename notification against,
        // so the reply must carry the path it may save to next
        const savedPath = notifyFileSaved(event.sender, result.filePath)
        return {
          ok: true,
          path: savedPath,
          passwordIntentPending,
          ...reissuedDoc(!!passwordState.password, hashes, plain),
        }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  ipcMain.handle('docs:save-new', async (event, defaultName: string, data: ArrayBuffer) => {
    try {
      // a discarded draft in an orphaned renderer must not silently persist
      // itself to the default folder after the user chose Don't Save
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      const filePath = uniquePathIn(defaultSaveDir(), defaultName)
      const passwordState = snapshotDocPassword(event.sender.id, null)
      const { bytes: plain, hashes } = await materializeLazyDocx(Buffer.from(data))
      const bytes = passwordState.password ? encryptDocx(plain, passwordState.password) : plain
      await atomicWriteFile(filePath, bytes)
      // teardown may have happened while the write was in flight — the path is
      // freshly created, so rolling it back is safe (mirrors docs:write-recovery)
      if (tornDownWcIds.has(event.sender.id)) {
        await unlink(filePath).catch(() => {})
        return { ok: false }
      }
      allowDocWrite(event.sender.id, filePath)
      rememberOpenDoc(event.sender.id, filePath)
      await rememberDiskState(event.sender.id, filePath, sha256Hex(bytes))
      pointLazyMediaAt(
        hashes,
        filePath,
        event.sender.id,
        passwordState.password ? plain : undefined,
      )
      if (tornDownWcIds.has(event.sender.id)) {
        await unlink(filePath).catch(() => {})
        return { ok: false }
      }
      const passwordIntentPending = commitDocPasswordSave(event.sender.id, passwordState, filePath)
      pushRecent(filePath)
      const savedPath = notifyFileSaved(event.sender, filePath)
      return {
        ok: true,
        path: savedPath,
        passwordIntentPending,
        ...reissuedDoc(!!passwordState.password, hashes, plain),
      }
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  ipcMain.handle(
    'docs:create-document',
    (_event, request: CreateDocumentRequest): Promise<CreateDocumentResult> =>
      createAiDocument(request),
  )

  // MCP-driven output: write the live document to an explicit absolute path with
  // no dialog. Mirrors docs:save-new's bookkeeping (write allowlist, disk state,
  // recents, tab-title sync) but targets a caller-chosen path and refuses to
  // clobber an existing file unless the caller asked for overwrite.
  ipcMain.handle(
    'docs:save-to',
    async (event, filePath: string, data: ArrayBuffer, overwrite: boolean) => {
      try {
        if (tornDownWcIds.has(event.sender.id)) return { ok: false }
        if (typeof filePath !== 'string' || !isAbsolute(filePath)) {
          return { ok: false, error: 'path must be absolute' }
        }
        if (extname(filePath).toLowerCase() !== '.docx') {
          return { ok: false, error: 'path must point to a .docx file' }
        }
        // only a target the MCP layer resolved for this tab may be written
        if (!canDocWrite(event.sender.id, filePath)) {
          return { ok: false, error: 'save target was not authorized' }
        }
        const existed = existsSync(filePath)
        if (!overwrite && existed) {
          return {
            ok: false,
            error: `file already exists: ${filePath} (pass overwrite:true to replace it)`,
          }
        }
        await mkdir(dirname(filePath), { recursive: true })
        const passwordState = snapshotDocPassword(event.sender.id, null)
        const { bytes: plain, hashes } = await materializeLazyDocx(Buffer.from(data))
        const bytes = passwordState.password ? encryptDocx(plain, passwordState.password) : plain
        await atomicWriteFile(filePath, bytes)
        // teardown may have happened while the write was in flight — only a file
        // this handler created is safe to roll back; an overwritten one stays
        const rollback = async (): Promise<{ ok: false }> => {
          if (!existed) await unlink(filePath).catch(() => {})
          return { ok: false }
        }
        if (tornDownWcIds.has(event.sender.id)) return rollback()
        await rememberDiskState(event.sender.id, filePath, sha256Hex(bytes))
        pointLazyMediaAt(
          hashes,
          filePath,
          event.sender.id,
          passwordState.password ? plain : undefined,
        )
        if (tornDownWcIds.has(event.sender.id)) return rollback()
        const passwordIntentPending = commitDocPasswordSave(
          event.sender.id,
          passwordState,
          filePath,
        )
        pushRecent(filePath)
        rememberOpenDoc(event.sender.id, filePath)
        notifyFileSaved(event.sender, filePath)
        return {
          ok: true,
          path: filePath,
          passwordIntentPending,
          ...reissuedDoc(!!passwordState.password, hashes, plain),
        }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  ipcMain.handle('docs:recent', () => readJson<string[]>(RECENT_PATH(), []))

  ipcMain.handle('docs:pick-image', async (event) => {
    const result = await openDialog(event, {
      title: tm('dlgInsertImage'),
      filters: [{ name: tm('filterImages'), extensions: ['png', 'jpg', 'jpeg', 'gif'] }],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    const filePath = result.filePaths[0]
    const ext = filePath.split('.').pop()?.toLowerCase() ?? ''
    const mime = IMAGE_MIME[ext]
    if (!mime) return null
    return {
      base64: readFileSync(filePath).toString('base64'),
      mime,
      name: basename(filePath),
    }
  })

  // Reads the picked file whole and leaves the tab's document, recents and
  // write grants alone: the caller takes content from it, it is not opened.
  ipcMain.handle('docs:pick-docx', async (event): Promise<PickDocxResult> => {
    const result = await openDialog(event, {
      title: tm('dlgOpenDoc'),
      filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    const filePath = result.filePaths[0]
    if (!/\.docx$/i.test(filePath)) return null
    if ((await stat(filePath)).size > MAX_OPEN_BYTES) {
      const mb = MAX_OPEN_BYTES / 1024 / 1024
      await showOpenError(event.sender.id, `${basename(filePath)}: ${tm('errTooLarge', { mb })}`)
      return null
    }
    const bytes = await readFile(filePath)
    if (isEncryptedDocx(bytes)) return { needsPassword: true }
    return { name: basename(filePath), dataUrl: handOffBytes(bytes) }
  })

  // Styles Organizer: the other document or template, read whole and written
  // back only to the path that was picked here
  ipcMain.handle('docs:organizer-open', async (event): Promise<OrganizerFileResult> => {
    const result = await openDialog(event, {
      title: tm('dlgOpenDoc'),
      filters: [{ name: tm('filterWord'), extensions: ['docx', 'dotx'] }],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    const filePath = result.filePaths[0]
    if (!/\.do[ct]x$/i.test(filePath)) return null
    if ((await stat(filePath)).size > MAX_OPEN_BYTES) {
      const mb = MAX_OPEN_BYTES / 1024 / 1024
      await showOpenError(event.sender.id, `${basename(filePath)}: ${tm('errTooLarge', { mb })}`)
      return null
    }
    const bytes = await readFile(filePath)
    if (isEncryptedDocx(bytes)) return { needsPassword: true }
    const granted = organizerPaths.get(event.sender.id) ?? new Set<string>()
    granted.add(filePath)
    organizerPaths.set(event.sender.id, granted)
    return { name: basename(filePath), path: filePath, dataUrl: handOffBytes(bytes) }
  })

  ipcMain.handle(
    'docs:organizer-save',
    async (
      event,
      filePath: string,
      data: ArrayBuffer,
    ): Promise<{ ok: boolean; error?: string }> => {
      try {
        if (!organizerPaths.get(event.sender.id)?.has(filePath)) {
          return { ok: false, error: 'the file was not opened in the Organizer' }
        }
        const bytes = Buffer.from(data)
        if (!looksLikeZip(bytes)) return { ok: false, error: 'not a document' }
        await atomicWriteFile(filePath, bytes)
        return { ok: true }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  // Mailings tab: pick a local recipient list and read it as a data source.
  // The file is read here, in the main process, where the file system is.
  // File > Rename… / Move…: the document's own file takes a new name or folder
  const moveOpenDoc = async (
    event: IpcMainInvokeEvent,
    target: (from: string) => string | null,
  ): Promise<DocMoveResult> => {
    const from = openDocByWc.get(event.sender.id)
    if (!from || !existsSync(from)) return { ok: false, error: 'missing' }
    const to = target(from)
    if (!to) return { ok: false, error: 'badName' }
    if (to === from) return { ok: true, path: from }
    // a case-only rename meets the file itself on a case-insensitive disk
    if (existsSync(to) && to.toLowerCase() !== from.toLowerCase())
      return { ok: false, error: 'exists' }
    try {
      await moveFile(from, to)
    } catch {
      return { ok: false, error: 'failed' }
    }
    if (shellHooks?.fileMoved) shellHooks.fileMoved(from, to)
    else {
      replaceRecentFile(from, to)
      docsFileRenamed(event.sender, from, to)
    }
    return { ok: true, path: to }
  }
  ipcMain.handle('docs:rename-doc', (event, name: unknown) =>
    moveOpenDoc(event, (from) => (typeof name === 'string' ? renamedDocPath(from, name) : null)),
  )
  ipcMain.handle('docs:move-doc', async (event) => {
    const from = openDocByWc.get(event.sender.id)
    const result = await openDialog(event, {
      title: tm('menuMove').replace(/…$/, ''),
      defaultPath: from ? dirname(from) : undefined,
      properties: ['openDirectory', 'createDirectory'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    return moveOpenDoc(event, (path) => join(result.filePaths[0], basename(path)))
  })
  ipcMain.handle('docs:file-info', (event): DocFileInfo | null => {
    const path = openDocByWc.get(event.sender.id)
    if (!path) return null
    try {
      const stat = statSync(path)
      return {
        path,
        sizeBytes: stat.size,
        createdMs: stat.birthtimeMs,
        modifiedMs: stat.mtimeMs,
      }
    } catch {
      return null
    }
  })
  ipcMain.handle('docs:merge-data-source', async (event) => {
    const result = await openDialog(event, {
      title: tm('dlgMergeDataSource'),
      filters: [
        {
          name: tm('filterMergeDataSource'),
          extensions: [...MERGE_DATA_SOURCE_EXTENSIONS],
        },
      ],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    try {
      return { ok: true, source: await readMergeDataSource(result.filePaths[0]) }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  // Mailings tab: write the merged documents. One file for the whole merge
  // goes where the user says; one file per record goes into a folder they
  // pick, every name kept unique so two records of the same name do not
  // overwrite each other. The bytes carry the document's lazily served
  // pictures, so they are materialized first — the merged file is a document
  // of its own, not a view of the open one.
  ipcMain.handle('docs:merge-output', async (event, request: MergeOutputRequest) => {
    const files = Array.isArray(request?.files) ? request.files : []
    if (files.length === 0) return { ok: false, error: 'nothing to write' }
    const write = async (path: string, data: ArrayBuffer): Promise<void> => {
      const { bytes } = await materializeLazyDocx(Buffer.from(data))
      await writeFile(path, bytes)
    }
    try {
      if (request.mode === 'single' || request.mode === 'labels') {
        const file = files[0]!
        const name = mergeOutputFileName(file.name)
        const result = await saveDialog(event, {
          title: tm('dlgSaveAs'),
          defaultPath: name,
          filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
        })
        if (result.canceled || !result.filePath) return { ok: false, canceled: true }
        await write(result.filePath, file.data)
        return { ok: true, path: result.filePath, count: 1 }
      }
      const result = await openDialog(event, {
        title: tm('dlgPickExportDir'),
        properties: ['openDirectory', 'createDirectory'],
      })
      if (result.canceled || result.filePaths.length === 0) return { ok: false, canceled: true }
      const folder = result.filePaths[0]!
      for (const file of files) {
        await write(mergeOutputPath(folder, file.name), file.data)
      }
      return { ok: true, folder, count: files.length }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  ipcMain.handle('files:pick', async (event): Promise<AttachmentAddResult | null> => {
    const result = await openDialog(event, {
      title: tm('dlgAddAttachment'),
      filters: [
        { name: tm('filterSupported'), extensions: [...ATTACHMENT_EXTS] },
        { name: tm('filterAll'), extensions: ['*'] },
      ],
      properties: ['openFile', 'multiSelections'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    sharedPathGrants().grant(event.sender.id, result.filePaths)
    return collectAttachments(result.filePaths)
  })

  ipcMain.handle('files:add', (_event, paths: string[]) => {
    // panel drops and clipboard file pastes are recorded by the preload
    // bridge before this runs, so the ledger already holds every path a
    // real drop carried; minting a grant for renderer-named paths here
    // would let a compromised renderer read whatever it names
    return collectAttachments(paths)
  })

  ipcMain.handle(
    'files:read',
    async (
      event,
      filePath: string,
      offset: number,
      maxChars: number,
    ): Promise<AttachmentReadResult> => {
      if (!canReadAttachment(event.sender.id, filePath))
        return { ok: false, error: 'docs: path not granted to this view' }
      const name = basename(filePath)
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      if (!ATTACHMENT_EXTS.has(ext)) return { ok: false, error: tm('errUnsupportedExt', { ext }) }
      if (ATTACHMENT_IMAGE_EXTS.has(ext)) {
        return { ok: false, error: tm('errImageNoText') }
      }
      try {
        const text = await extractAttachmentText(filePath)
        const start = Math.max(0, Math.floor(offset) || 0)
        const size = Math.min(Math.max(1, Math.floor(maxChars) || 1), 48_000)
        return {
          ok: true,
          name,
          totalChars: text.length,
          offset: start,
          text: text.slice(start, start + size),
        }
      } catch (e) {
        return { ok: false, error: e instanceof Error ? e.message : String(e) }
      }
    },
  )

  // image attachments read raw bytes → base64; AiPanel puts them into the user message's images for multimodal
  ipcMain.handle('files:read-image', (event, filePath: string): AttachmentImageResult => {
    if (!canReadAttachment(event.sender.id, filePath))
      return { ok: false, error: 'docs: path not granted to this view' }
    const name = basename(filePath)
    const ext = name.split('.').pop()?.toLowerCase() ?? ''
    const mime = ATTACHMENT_IMAGE_MIME[ext]
    if (!mime) return { ok: false, error: `${name}: ${tm('errNotImage')}` }
    try {
      const stat = statSync(filePath)
      if (stat.size > ATTACHMENT_IMAGE_MAX_BYTES) {
        return { ok: false, error: `${name}: ${tm('errImageTooLarge')}` }
      }
      return { ok: true, base64: readFileSync(filePath).toString('base64'), mime }
    } catch {
      return { ok: false, error: `${name}: ${tm('errUnreadable')}` }
    }
  })

  // clipboard-pasted images (screenshots and other bitmaps with no local path): saved to a temp file then use the regular attachment path
  ipcMain.handle(
    'files:add-pasted-image',
    (event, data: unknown, ext: unknown): AttachmentAddResult => {
      const filePath = savePastedImage(data, ext)
      if (filePath) sharedPathGrants().grant(event.sender.id, [filePath])
      return filePath
        ? collectAttachments([filePath])
        : { accepted: [], rejected: [tm('errNotImage')] }
    },
  )

  // r136: copying an embedded picture must yield a real bitmap for external
  // apps (Gmail pasted blank) plus plain <img> html for cross-document paste
  // (the protected wrapper round-tripped as a "protected content" shell).
  ipcMain.handle(
    'docs:copy-image-to-clipboard',
    async (_event, dataUrl: unknown, meta: unknown): Promise<boolean> => {
      // Renderer-supplied bitmap: validate before base64 decode + nativeImage
      // (a huge data URL would OOM the main process). Non-data URLs are
      // lazy-media ids resolved below.
      if (!validCopyImageDataUrl(dataUrl)) return false
      let bytes: Buffer
      let htmlSrc = dataUrl
      if (dataUrl.startsWith('data:image/')) {
        bytes = Buffer.from(dataUrl.slice(dataUrl.indexOf(',') + 1), 'base64')
      } else {
        const media = await readLazyMedia(dataUrl)
        if (!media) return false
        bytes = media.body
        // another document cannot resolve this document's lazy URL; inline the bytes
        htmlSrc = `data:${media.mime};base64,${bytes.toString('base64')}`
      }
      // createFromBuffer, not createFromDataURL — the latter returns an empty
      // image for valid PNGs in this Electron
      const image = nativeImage.createFromBuffer(bytes)
      if (image.isEmpty()) return false
      // the html flavor carries the DISPLAY size + layout meta so an in-app
      // paste keeps size/align/wrap instead of falling back to bitmap pixels
      let width = image.getSize().width
      let height = image.getSize().height
      let metaAttr = ''
      if (typeof meta === 'string' && meta.length <= 2048) {
        try {
          const parsed = copyImageDisplaySize(meta)
          if (parsed.width !== undefined) width = parsed.width
          if (parsed.height !== undefined) height = parsed.height
          const escaped = meta.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
          metaAttr = ` data-image-meta="${escaped}"`
        } catch {
          /* malformed payload: plain img */
        }
      }
      clipboard.write({
        image,
        html: `<img src="${htmlSrc}" width="${width}" height="${height}"${metaAttr}>`,
      })
      return true
    },
  )

  // renderer print scale (inverse of the preview's print zoom, see print-zoom.ts)
  // Infinity passes a `> 0` check, so require finiteness before handing it to Chromium.
  const pdfScale = (scale?: number) => printScaleOption(scale)
  const withInternalLinks = async (data: Buffer, links: unknown): Promise<Buffer> => {
    if (!Array.isArray(links) || links.length === 0) return data
    const { PDFDocument } = await import('pdf-lib')
    const pdf = await PDFDocument.load(data)
    if (addPdfInternalLinks(pdf, links) === 0) return data
    return Buffer.from(await pdf.save())
  }
  const printScale = (scale?: number) =>
    typeof scale === 'number' && Number.isFinite(scale) && scale > 0 && scale !== 1
      ? { scaleFactor: Math.round(scale * 100) }
      : {}

  ipcMain.handle('docs:printers', async (event) => {
    const printers = await event.sender.getPrintersAsync()
    return printers.map((p) => ({ name: p.name, displayName: p.displayName || p.name }))
  })

  ipcMain.handle('docs:print', async (event, scale?: number, job?: unknown, paper?: unknown) => {
    // print the calling tab's own content; zero margins — the docx page padding provides them.
    // Resolves when the system dialog is dismissed (or the silent job is queued); the print
    // dialog stays open on cancel (ok=false without error) and surfaces real failures.
    const printOpts = {
      margins: { marginType: 'none' as const },
      ...printScale(scale),
      ...printJobOptions(job),
      // the document's own paper: Chromium prints on Letter unless the job
      // states one, and an A4 sheet on Letter is taller than the paper — it
      // splits in two, so the printed page loses its bottom (genoffice#2563)
      ...printPaperOption(paper),
    }
    // A named printer prints silently, and Chromium's silent print on macOS
    // can drop the job outright — no queue entry, no callback error
    // (genoffice#2347). Route those jobs through CUPS instead: printToPDF
    // bakes the pages at the document's own paper size, and lp spools a
    // real job whose queue state and failures are diagnosable.
    if (process.platform === 'darwin' && printOpts.silent) {
      const geometry = validPrintPaper(paper)
      const printed = await event.sender.printToPDF({
        printBackground: true,
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        ...(geometry
          ? {
              pageSize: {
                width: geometry.widthTwips / TWIPS_PER_INCH,
                height: geometry.heightTwips / TWIPS_PER_INCH,
              },
            }
          : {}),
        ...printScaleOption(scale),
      })
      return printPdfViaCups(printed, printOpts)
    }
    return new Promise<{ ok: boolean; error?: string }>((resolve) => {
      event.sender.print(printOpts, (success, failureReason) => {
        resolve(printOutcome(success, failureReason, printOpts.silent))
      })
    })
  })

  ipcMain.handle(
    'docs:export-pdf',
    async (
      event,
      defaultName: string,
      pageWidthTwips: number,
      pageHeightTwips: number,
      outPath?: string,
      scale?: number,
      links?: unknown,
    ) => {
      // Renderer-supplied page geometry reaches Chromium printToPDF verbatim:
      // reject non-finite/out-of-range sizes (0.1in..50in) and scales (0.1..5),
      // same guard as docs:print-pdf-buffer (a malformed w:pgSz in a doc would
      // otherwise hand Chromium a page thousands of inches wide).
      if (!validPrintGeometry(pageWidthTwips, pageHeightTwips, scale)) {
        return { ok: false, error: 'invalid page size or scale' }
      }
      // renderer-supplied outPath is only honored when a save dialog authorized it before
      let filePath = outPath ?? null
      if (filePath && !canPdfWrite(event.sender.id, filePath)) {
        return { ok: false, error: 'export target is not an authorized path' }
      }
      if (!filePath) {
        const result = await saveDialog(event, {
          title: tm('dlgExportPdf'),
          defaultPath: defaultName.replace(/\.docx$/i, '') + '.pdf',
          filters: [{ name: 'PDF', extensions: ['pdf'] }],
        })
        if (result.canceled || !result.filePath) return { ok: false }
        filePath = result.filePath
        allowPdfWrite(event.sender.id, filePath)
      }
      try {
        const printed = await event.sender.printToPDF({
          printBackground: true,
          // custom pageSize is in inches; the docx page padding provides the margins
          pageSize: {
            width: pageWidthTwips / TWIPS_PER_INCH,
            height: pageHeightTwips / TWIPS_PER_INCH,
          },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          ...pdfScale(scale),
        })
        const data = await withInternalLinks(printed, links)
        await atomicWriteFile(filePath, data)
        if (!isImageExportTemp(event.sender.id, filePath)) openGeneratedFile(filePath)
        return { ok: true, path: filePath }
      } catch (err) {
        // path is already authorized, so the renderer can retry chunked to the same target
        return { ok: false, error: String(err), path: filePath }
      }
    },
  )

  ipcMain.handle('docs:save-image-as', async (event, src: unknown) => {
    if (tornDownWcIds.has(event.sender.id) || typeof src !== 'string') return { ok: false }
    return saveImageFromUrl(dialogParent(event), src, {
      title: tm('dlgSaveAs'),
      fallbackDir: defaultSaveDir(),
    })
  })

  ipcMain.handle('docs:pick-export-images-target', async (event) => {
    let dir = testExportDir
    if (!dir) {
      const r = await openDialog(event, {
        title: tm('dlgPickExportDir'),
        properties: ['openDirectory', 'createDirectory'],
      })
      dir = r.canceled ? null : (r.filePaths[0] ?? null)
    }
    if (!dir) return null
    const wcId = event.sender.id
    const pdfPath = join(tmpdir(), `genoffice-docs-images-${randomUUID()}.pdf`)
    allowPdfWrite(wcId, pdfPath)
    imageExportTemps.set(wcId, (imageExportTemps.get(wcId) ?? new Set()).add(pdfPath))
    imageExportDirs.set(wcId, (imageExportDirs.get(wcId) ?? new Set()).add(dir))
    return { dir, pdfPath }
  })

  ipcMain.handle('docs:take-export-pdf', async (event, pdfPath: string) => {
    const temps = imageExportTemps.get(event.sender.id)
    if (typeof pdfPath !== 'string' || !temps?.has(pdfPath)) {
      return { ok: false, error: 'not an image-export temp file' }
    }
    temps.delete(pdfPath)
    try {
      const data = await readFile(pdfPath)
      return { ok: true, base64: data.toString('base64') }
    } catch (err) {
      return { ok: false, error: String(err) }
    } finally {
      await rm(pdfPath, { force: true })
    }
  })

  ipcMain.handle(
    'docs:write-export-image',
    async (event, dir: string, fileName: string, pngBase64: string) => {
      if (typeof dir !== 'string' || !imageExportDirs.get(event.sender.id)?.has(dir)) {
        return { ok: false, error: 'export target is not an authorized folder' }
      }
      if (
        typeof fileName !== 'string' ||
        fileName !== basename(fileName) ||
        !/^[^/\\]+\.png$/.test(fileName)
      ) {
        return { ok: false, error: 'invalid image file name' }
      }
      try {
        const filePath = join(dir, fileName)
        await atomicWriteFile(filePath, Buffer.from(String(pngBase64), 'base64'))
        return { ok: true, path: filePath }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  // templates live in one folder of their own, where both dialogs start
  const templatesDir = async (): Promise<string> => {
    const dir = join(app.getPath('userData'), 'Templates')
    await mkdir(dir, { recursive: true })
    return dir
  }

  const TEMPLATE_FILE = /\.(dotx|dotm|dot|docx)$/i

  // `data` is already declared a template by the renderer; nothing about the
  // open document changes: it keeps its path, its password and its unsaved state.
  // `preview` is where the first page is on screen: what is there becomes the
  // picture the template gallery shows.
  ipcMain.handle(
    'docs:save-template',
    async (event, defaultName: string, data: ArrayBuffer, preview?: TemplatePreviewRect) => {
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      const thumbnail = await capturePreview(event.sender, preview)
      const stem =
        basename(String(defaultName || '')).replace(/\.(docx?|dot[xm]?)$/i, '') || 'Template'
      const result = await saveDialog(event, {
        title: tm('dlgSaveAs'),
        defaultPath: join(await templatesDir(), `${stem}.dotx`),
        filters: [{ name: tm('filterWord'), extensions: ['dotx'] }],
      })
      if (result.canceled || !result.filePath) return { ok: false }
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      const filePath = /\.dotx$/i.test(result.filePath)
        ? result.filePath
        : `${result.filePath}.dotx`
      try {
        const { bytes } = await materializeLazyDocx(Buffer.from(data))
        await atomicWriteFile(
          filePath,
          thumbnail ? Buffer.from(await withTemplateThumbnail(bytes, thumbnail)) : bytes,
        )
        return { ok: true, path: filePath }
      } catch (err) {
        return { ok: false, error: String(err), path: filePath }
      }
    },
  )

  // the templates folder as a gallery: every template in it, with the preview picture it carries
  ipcMain.handle('docs:list-templates', async (): Promise<TemplateEntry[]> => {
    const dir = await templatesDir()
    const names = (await readdir(dir))
      .filter((name) => TEMPLATE_FILE.test(name) && !name.startsWith('.') && !name.startsWith('~$'))
      .sort((a, b) => a.localeCompare(b))
      .slice(0, MAX_GALLERY_TEMPLATES)
    // a few at a time: each one is read whole and inflated to find its picture
    return mapBounded(names, GALLERY_READS_AT_ONCE, async (name): Promise<TemplateEntry> => {
      try {
        const filePath = join(dir, name)
        if (/\.dot$/i.test(name) || (await stat(filePath)).size > MAX_OPEN_BYTES) return { name }
        const picture = await templateThumbnail(await readFile(filePath))
        if (!picture || picture.data.length > MAX_THUMBNAIL_BYTES) return { name }
        return {
          name,
          thumbnail: `data:${picture.mime};base64,${Buffer.from(picture.data).toString('base64')}`,
        }
      } catch {
        return { name }
      }
    })
  })

  // a template opens as a document of its own with no path: no recents entry,
  // no write grant, so saving what is made from it never touches the template.
  // `fromGallery` names a file of the templates folder; without it the user picks one.
  ipcMain.handle(
    'docs:open-template',
    async (event, fromGallery?: string): Promise<OpenDocxResult> => {
      let filePath: string
      if (fromGallery !== undefined) {
        const name = String(fromGallery)
        if (name !== basename(name) || !TEMPLATE_FILE.test(name)) return null
        filePath = join(await templatesDir(), name)
        if (!existsSync(filePath)) return null
      } else {
        const result = await openDialog(event, {
          title: tm('dlgOpenDoc'),
          defaultPath: await templatesDir(),
          filters: [{ name: tm('filterWord'), extensions: ['dotx', 'dotm', 'dot', 'docx'] }],
          properties: ['openFile'],
        })
        if (result.canceled || result.filePaths.length === 0) return null
        filePath = result.filePaths[0]
      }
      const tooLarge = async (): Promise<null> => {
        const mb = MAX_OPEN_BYTES / 1024 / 1024
        await showOpenError(event.sender.id, `${basename(filePath)}: ${tm('errTooLarge', { mb })}`)
        return null
      }
      const opened = (bytes: Buffer): OpenDocxResult => ({
        path: '',
        name: basename(filePath),
        dataUrl: handOffBytes(bytes),
        hash: '',
        template: true,
      })
      // a Word 97-2003 template is converted, as File > Open converts a .doc; the
      // document made from it is untitled, so there is no copy to write
      if (/\.dot$/i.test(filePath)) {
        if (!(await confirmLossyImport(filePath, 'word97', dialogParent(event)))) return null
        if ((await stat(filePath)).size > MAX_OPEN_BYTES) return tooLarge()
        try {
          const converted = await convertLegacyDocToDocx(new Uint8Array(await readFile(filePath)))
          return opened(Buffer.from(converted.docx))
        } catch {
          await showOpenError(event.sender.id, `${basename(filePath)}: ${tm('errUnreadable')}`)
          return null
        }
      }
      // large pictures stay in the template file and are served from it, as File > Open does;
      // the limit is on what is left to load, so it is checked after
      const lazy = await openLazyDocx(filePath, event.sender.id)
      if ((lazy?.bytes.length ?? (await stat(filePath)).size) > MAX_OPEN_BYTES) return tooLarge()
      const bytes = lazy?.bytes ?? (await readFile(filePath))
      // a protected file goes through File > Open, which asks for its password
      if (isEncryptedDocx(bytes)) {
        await showOpenError(event.sender.id, `${basename(filePath)}: ${tm('errTemplateProtected')}`)
        return null
      }
      return opened(bytes)
    },
  )

  ipcMain.handle(
    'docs:export-html',
    async (event, defaultName: string, html: string, outPath?: string) => {
      if (typeof html !== 'string' || !html) return { ok: false, error: 'empty document' }
      let filePath = outPath ?? null
      if (filePath && !canPdfWrite(event.sender.id, filePath)) {
        return { ok: false, error: 'export target is not an authorized path' }
      }
      if (!filePath) {
        const result = await saveDialog(event, {
          title: tm('dlgExportHtml'),
          defaultPath: defaultName.replace(/\.docx$/i, '') + '.html',
          filters: [{ name: 'HTML', extensions: ['html'] }],
        })
        if (result.canceled || !result.filePath) return { ok: false }
        filePath = result.filePath
        allowPdfWrite(event.sender.id, filePath)
      }
      try {
        await atomicWriteFile(
          filePath,
          Buffer.from(await inlineLazyMediaInHtml(html, readLazyMedia), 'utf8'),
        )
        openGeneratedFile(filePath)
        return { ok: true, path: filePath }
      } catch (err) {
        return { ok: false, error: String(err), path: filePath }
      }
    },
  )

  ipcMain.handle(
    'docs:export-markdown',
    async (
      event,
      defaultName: string,
      markdown: string,
      images: MarkdownImage[],
      outPath?: string,
    ) => {
      if (typeof markdown !== 'string' || !markdown) return { ok: false, error: 'empty document' }
      if (!Array.isArray(images)) return { ok: false, error: 'bad image list' }
      let filePath = outPath ?? null
      if (filePath && !canPdfWrite(event.sender.id, filePath)) {
        return { ok: false, error: 'export target is not an authorized path' }
      }
      if (!filePath) {
        const result = await saveDialog(event, {
          title: tm('dlgExportMarkdown'),
          defaultPath: defaultName.replace(/\.docx$/i, '') + '.md',
          filters: [{ name: 'Markdown', extensions: ['md'] }],
        })
        if (result.canceled || !result.filePath) return { ok: false }
        filePath = result.filePath
        allowPdfWrite(event.sender.id, filePath)
      }
      try {
        const written = await writeMarkdownMedia(filePath, images)
        // the Markdown destination was escaped before the swap: parentheses must not survive it
        const mediaDir = encodeURIComponent(
          basename(filePath).replace(/\.md$/i, '') + '_files',
        ).replace(/[()]/g, (c) => (c === '(' ? '%28' : '%29'))
        await atomicWriteFile(
          filePath,
          Buffer.from(markdown.split(`${MARKDOWN_MEDIA_DIR_TOKEN}/`).join(`${mediaDir}/`), 'utf8'),
        )
        openGeneratedFile(filePath)
        return { ok: true, path: filePath, imagesDropped: images.length - written }
      } catch (err) {
        return { ok: false, error: String(err), path: filePath }
      }
    },
  )

  // mixed paper-size export: the renderer prints group by group per size (other pages hidden via CSS); this produces one group's bytes
  ipcMain.handle(
    'docs:print-pdf-buffer',
    async (event, pageWidthTwips: number, pageHeightTwips: number, scale?: number) => {
      // Renderer-supplied page geometry reaches Chromium printToPDF verbatim:
      // reject non-finite/out-of-range sizes (0.5in..50in) and scales (0.1..5).
      if (!validPrintGeometry(pageWidthTwips, pageHeightTwips, scale)) {
        return { ok: false, error: 'invalid page size or scale' }
      }
      try {
        const data = await event.sender.printToPDF({
          printBackground: true,
          pageSize: {
            width: pageWidthTwips / TWIPS_PER_INCH,
            height: pageHeightTwips / TWIPS_PER_INCH,
          },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          ...pdfScale(scale),
        })
        return { ok: true, base64: data.toString('base64') }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  // merge grouped PDF fragments into one file in page order (pdf-lib)
  ipcMain.handle(
    'docs:save-merged-pdf',
    async (
      event,
      defaultName: string,
      base64Parts: string[],
      outPath?: string,
      links?: unknown,
    ) => {
      let filePath = outPath ?? null
      if (filePath && !canPdfWrite(event.sender.id, filePath)) {
        return { ok: false, error: 'export target is not an authorized path' }
      }
      if (!filePath) {
        const result = await saveDialog(event, {
          title: tm('dlgExportPdf'),
          defaultPath: defaultName.replace(/\.docx$/i, '') + '.pdf',
          filters: [{ name: 'PDF', extensions: ['pdf'] }],
        })
        if (result.canceled || !result.filePath) return { ok: false }
        filePath = result.filePath
        allowPdfWrite(event.sender.id, filePath)
      }
      try {
        const { PDFDocument } = await import('pdf-lib')
        const merged = await PDFDocument.create()
        for (const b64 of base64Parts) {
          const part = await PDFDocument.load(Buffer.from(b64, 'base64'))
          const pages = await merged.copyPages(part, part.getPageIndices())
          for (const page of pages) merged.addPage(page)
        }
        if (Array.isArray(links) && links.length > 0) addPdfInternalLinks(merged, links)
        await atomicWriteFile(filePath, Buffer.from(await merged.save()))
        if (!isImageExportTemp(event.sender.id, filePath)) openGeneratedFile(filePath)
        return { ok: true, path: filePath }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  ipcMain.handle('win:new', (event, openPath: string | null) => {
    // the only legitimate argument is this view's own open document
    if (
      openPath !== null &&
      openPath !== undefined &&
      openDocByWc.get(event.sender.id) !== openPath
    )
      return
    // A pathless new tab/window starts as a blank document, not the start screen
    const path = openPath ?? undefined
    if (shellHooks) shellHooks.openTab(path, path ? undefined : { newBlank: true })
    else {
      const win = createDocsWindow(path)
      if (!path) markDocsNewBlank(win.webContents.id)
    }
  })

  ipcMain.handle('docs:open-result-tab', (_event, data: Uint8Array) => {
    if (!(data instanceof Uint8Array) || data.byteLength > MAX_OPEN_BYTES) return
    const doc: ResultDoc = { bytes: Buffer.from(data) }
    if (shellHooks?.openResultDocTab) shellHooks.openResultDocTab(doc)
    else queueDocsResultDoc(createDocsWindow().webContents.id, doc)
  })

  ipcMain.handle('win:list', (): DocsTabInfo[] => {
    if (shellHooks) return shellHooks.listTabs()
    return BrowserWindow.getAllWindows().map((w) => ({
      id: String(w.id),
      title: w.getTitle(),
      focused: w.isFocused(),
    }))
  })

  ipcMain.handle('win:arrange-all', (event) => {
    if (shellHooks?.arrangeAll) {
      shellHooks.arrangeAll(event.sender)
      return
    }
    const from = BrowserWindow.fromWebContents(event.sender)
    if (!from) return
    const others = BrowserWindow.getAllWindows().filter((w) => w !== from && w.isVisible())
    tileWindows([from, ...others], screen.getDisplayMatching(from.getBounds()).workArea)
    from.focus()
  })

  // standalone windows have no tab strip to return to
  ipcMain.handle('win:gather', (event) => shellHooks?.gatherWindows?.(event.sender))
  ipcMain.handle('win:detached-count', () => shellHooks?.detachedWindowCount?.() ?? 0)

  ipcMain.handle('win:focus', (_event, id: string) => {
    if (shellHooks) {
      shellHooks.focusTab(id)
      return
    }
    const win = BrowserWindow.fromId(Number(id))
    if (!win) return
    if (win.isMinimized()) win.restore()
    win.show()
    win.focus()
  })
}

/** hooks injected by the shell in tab mode; standalone mode leaves these unset
 * and falls back to real multi-BrowserWindow behavior. */
interface DocsShellHooks {
  openTab(openPath?: string, options?: { newBlank?: boolean }): void
  /** open a blank docs tab that consumes the queued AI content on boot (create_document) */
  openAiDocTab?(content: AiDocContent): void
  /** open a docs tab that shows a made document untitled */
  openResultDocTab?(doc: ResultDoc): void
  listTabs(): DocsTabInfo[]
  focusTab(id: string): void
  arrangeAll?(from: WebContents): void
  /** detached Docs windows become tabs of the shell window again */
  gatherWindows?(from: WebContents): void
  detachedWindowCount?(): number
  /** closes the calling tab instead of the whole shell window (Cmd+W / role:'close') */
  closeActiveTab(): void
  /** Shell router used to open exported PDFs in a new GenOffice tab. */
  openGeneratedPath?(path: string): boolean
  /** the open document's file was renamed or moved: recents, tabs and the views follow it */
  fileMoved?(oldPath: string, newPath: string): void
}
let shellHooks: DocsShellHooks | null = null
export function setDocsShellHooks(hooks: DocsShellHooks | null): void {
  shellHooks = hooks
}

/** After writing an exported/AI-generated file: open it in the right tab
 * (shell) or reveal it in the folder (standalone). Tab-opening failure must
 * not report the write itself as failed — the file is already persisted. */
/**
 * Pictures of a Markdown export land in `<file>_files/` next to it, like a
 * browser's "save page" folder. Returns how many were written; unreadable
 * sources are skipped and the reference in the Markdown dangles.
 */
async function writeMarkdownMedia(mdPath: string, images: MarkdownImage[]): Promise<number> {
  if (images.length === 0) return 0
  const dir = join(dirname(mdPath), basename(mdPath).replace(/\.md$/i, '') + '_files')
  let written = 0
  for (const img of images) {
    if (typeof img?.name !== 'string' || !/^image\d+\.\w+$/.test(img.name)) continue
    if (typeof img.src !== 'string') continue
    const data = /^data:image\/[\w.+-]+;base64,(.*)$/s.exec(img.src)
    const body = data
      ? Buffer.from(data[1], 'base64')
      : ((await readLazyMedia(img.src).catch(() => null))?.body ?? null)
    if (!body || body.length === 0) continue
    if (written === 0) await mkdir(dir, { recursive: true })
    await atomicWriteFile(join(dir, img.name), body)
    written++
  }
  return written
}

function openGeneratedFile(path: string): void {
  // Headless export has no tab strip and no user: revealing the file in Finder
  // would be the only visible effect of a run that must stay silent.
  if (isHeadlessMode()) return
  try {
    if (shellHooks?.openGeneratedPath?.(path)) return
  } catch (err) {
    console.warn('[docs] Failed to open generated file:', err)
  }
  shell.showItemInFolder(path)
}

/** Pick a safe file-name stem for an AI-created document. */
export function sanitizeAiDocFileBase(title: string): string {
  // Control characters are intentionally rejected from generated file names.
  const cleaned = String(title ?? '')
    // eslint-disable-next-line no-control-regex
    .replace(/[/\\:*?"<>|\u0000-\u001f]/g, '_')
    .trim()
    .slice(0, 80)
    .trim()
  return cleaned && cleaned !== '.' && cleaned !== '..' ? cleaned : 'Untitled'
}

/**
 * AI create_document: build a new standalone file in the default folder and
 * open it in a new tab. docx routes through a fresh blank docs tab that
 * inserts the queued content on boot and saves itself (the full-fidelity
 * HTML → docx conversion lives in the docs renderer); pdf and md are written
 * directly here. Also called by other apps' mains via shell-wired hooks.
 */
export async function createAiDocument(
  request: CreateDocumentRequest,
): Promise<CreateDocumentResult> {
  const type = request?.type
  const title = sanitizeAiDocFileBase(request?.title)
  const content = String(request?.content ?? '')
  if (!content.trim()) return { ok: false, error: 'content must not be empty' }
  try {
    if (type === 'docx') {
      const payload: AiDocContent = { title, html: content }
      if (shellHooks?.openAiDocTab) shellHooks.openAiDocTab(payload)
      else {
        const win = createDocsWindow(undefined)
        markDocsNewBlank(win.webContents.id)
        queueDocsAiContent(win.webContents.id, payload)
      }
      return { ok: true }
    }
    if (type === 'pdf') {
      const bytes = await printHtmlToPdf(
        buildPrintableHtml(title, content),
        () =>
          new BrowserWindow({ show: false, webPreferences: { sandbox: true, javascript: false } }),
      )
      const filePath = uniquePathIn(defaultSaveDir(), `${title}.pdf`)
      await writeFile(filePath, bytes)
      openGeneratedFile(filePath)
      return { ok: true, path: filePath }
    }
    if (type === 'md' || type === 'html') {
      const filePath = uniquePathIn(defaultSaveDir(), `${title}.${type}`)
      await writeFile(filePath, content, 'utf8')
      openGeneratedFile(filePath)
      return { ok: true, path: filePath }
    }
    return { ok: false, error: `unsupported document type: ${String(type)}` }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

// ---- application menu ----

function sendCommand(command: MenuCommand, payload?: string): void {
  activeDocsWebContents()?.send('menu:command', command, payload)
}

/**
 * Per-tab View-menu toggle state (AI Sidebar / Dark Mode), reported by each
 * renderer whenever it changes. The template can't hardcode `checked` — the
 * state lives in the renderer and differs per tab — so builds read the active
 * tab's last report, and reports from the active tab patch the built menu in
 * place (buildDocsMenu also re-runs on every tab focus switch).
 * Defaults mirror the renderer's initial state: sidebar shown, light canvas.
 */
const commandStateByWebContents = new Map<number, CommandMenuState>()

function activeCommandState(): CommandMenuState {
  const id = activeDocsWebContents()?.id
  return (
    (id !== undefined ? commandStateByWebContents.get(id) : undefined) ?? {
      checked: ['toggle-ai'],
      disabled: [],
    }
  )
}

const KEYBOARD_PATH = () => userDataPath('docs-keyboard.json')
let keyOverrides: KeyOverrides | null = null

function userKeyOverrides(): KeyOverrides {
  keyOverrides ??= sanitizeKeyOverrides(
    readJson<unknown>(KEYBOARD_PATH(), {}),
    process.platform === 'darwin',
  )
  return keyOverrides
}

/** commands with an item of their own in the menu bar, which is where their chord shows */
const OWN_MENU_ITEM: ReadonlySet<string> = (() => {
  const ids = new Set<string>(['shortcuts'])
  const walk = (nodes: readonly MenuNode[]) => {
    for (const node of nodes) {
      if (typeof node === 'string') ids.add(node)
      else if ('submenu' in node) walk(node.items)
    }
  }
  for (const top of DOCS_MENU) walk(top.items)
  return ids
})()

/** every runnable command of the menu bar, for the Customize Keyboard dialog */
function keyCatalog(): { id: string; label: string; menu: string; path: string }[] {
  const out: { id: string; label: string; menu: string; path: string }[] = []
  const walk = (nodes: readonly MenuNode<MainKey>[], menu: string, trail: string[]) => {
    for (const node of nodes) {
      if (node === '---' || (typeof node === 'object' && ('slot' in node || 'role' in node)))
        continue
      if (typeof node === 'object' && 'submenu' in node) {
        walk(node.items, menu, [...trail, tm(node.submenu)])
        continue
      }
      const cmd = typeof node === 'string' ? node : node.cmd
      const def: CommandDef = COMMANDS[cmd]
      if (def.role || def.planned) continue
      // Repeat's menu key is its greyed-out wording
      const labelKey = cmd === 'repeat' ? 'menuRepeat' : (def.menuKey as MainKey)
      const label = tm(typeof node === 'string' ? labelKey : node.label)
      const id = bindingId(cmd, typeof node === 'string' ? undefined : node.payload)
      out.push({ id, label, menu, path: [...trail, label].join(' \u25B8 ') })
    }
  }
  const layout: readonly MenuTop<MainKey>[] = DOCS_MENU
  for (const top of layout) walk(top.items, tm(top.label), [tm(top.label)])
  walk(['shortcuts'], tm('menuHelp'), [tm('menuHelp')])
  return out
}

const COMMAND_ITEM_ID = 'docs-cmd:'

/** every live menu item of a command (one command can sit in more than one menu) */
function commandMenuItems(
  menu: Menu,
  into = new Map<string, MenuItem[]>(),
): Map<string, MenuItem[]> {
  for (const item of menu.items) {
    if (item.id?.startsWith(COMMAND_ITEM_ID)) {
      const id = item.id.slice(COMMAND_ITEM_ID.length)
      into.set(id, [...(into.get(id) ?? []), item])
    }
    if (item.submenu) commandMenuItems(item.submenu, into)
  }
  return into
}

/** shell-injected items appended to the File menu (e.g. Back to Home); persists
 * across the internal rebuilds pushRecent() triggers */
let extraFileMenuItems: MenuItemConstructorOptions[] = []

export function setDocsExtraFileMenuItems(items: MenuItemConstructorOptions[]): void {
  extraFileMenuItems = items
}

/** Shell-installed gate: inside the shell the docs menu may only take over the
 * application menu while a docs tab is active — internal rebuilds (pushRecent
 * after opening/saving any file) must not clobber another tab's menu.
 * The standalone docs app registers no gate and always installs. */
let docsMenuGate: (() => boolean) | null = null

export function setDocsMenuGate(gate: () => boolean): void {
  docsMenuGate = gate
}

export function buildDocsMenu(): void {
  if (docsMenuGate && !docsMenuGate()) return
  const isMac = process.platform === 'darwin'
  const state = activeCommandState()
  /** ties a menu item to its command, so state reports reach it */
  const stateOf = (id: CommandId): MenuItemConstructorOptions => ({
    id: COMMAND_ITEM_ID + id,
    ...(state.disabled.includes(id) ? { enabled: false } : {}),
  })
  const cmd = (id: CommandId, labelKey?: MainKey, payload?: string): MenuItemConstructorOptions => {
    const def: CommandDef = COMMANDS[id]
    const item: MenuItemConstructorOptions = {
      ...stateOf(id),
      label: tm(labelKey ?? (def.menuKey as MainKey)),
    }
    // a role item is carried out natively by whatever text target has the focus
    if (def.role) return { ...item, role: def.role }
    if (id === 'repeat' && state.repeat)
      item.label = tm(state.repeat === 'redo' ? 'menuRedo' : 'menuRepeat')
    // a command listed a second time under another label leaves the chord to its own item
    const ownAccelerator = labelKey ? undefined : def.accelerator
    const builtIn =
      typeof ownAccelerator === 'object' ? ownAccelerator[isMac ? 'mac' : 'other'] : ownAccelerator
    const showsChord = !labelKey || !!payload || !OWN_MENU_ITEM.has(id)
    const accelerator = showsChord
      ? menuAccelerator(userKeyOverrides(), bindingId(id, payload), builtIn)
      : undefined
    return {
      ...item,
      ...(accelerator ? { accelerator } : {}),
      ...(def.toggle ? { type: 'checkbox' as const, checked: state.checked.includes(id) } : {}),
      click: () => sendCommand(id as MenuCommand, payload),
    }
  }
  // Opening a menu must not trigger an automount; validate when the user opens a file.
  const recent = readJson<string[]>(RECENT_PATH(), []).slice(0, 10)

  const recentSubmenu: MenuItemConstructorOptions[] =
    recent.length > 0
      ? recent.map((p) => ({
          label: basename(p),
          sublabel: isMac ? undefined : p,
          toolTip: p,
          click: () => sendCommand('open-path', p),
        }))
      : [{ label: tm('menuNoRecent'), enabled: false }]

  const labels = appMenuLabels(getUiLang())
  const appMenu = appMenuTemplate(process.platform, labels, {
    appLabel: app.name,
    extra: [settingsMenuItem(labels, () => sendCommand('preferences'))],
  })
  const slots: Record<MenuSlot, MenuItemConstructorOptions[]> = {
    'new-window': [
      {
        label: tm('menuNewWindow'),
        accelerator: 'Shift+CmdOrCtrl+N',
        click: () => {
          if (shellHooks) shellHooks.openTab(undefined, { newBlank: true })
          else markDocsNewBlank(createDocsWindow().webContents.id)
        },
      },
    ],
    'open-recent': [
      {
        label: tm('menuOpenRecent'),
        submenu: [
          ...recentSubmenu,
          { type: 'separator' },
          {
            label: tm('menuRecentMore'),
            // Open's second chord, shown here while Open still has it
            ...(keepsAccelerator(userKeyOverrides(), 'open', 'Shift+CmdOrCtrl+O', isMac)
              ? { accelerator: 'Shift+CmdOrCtrl+O' }
              : {}),
            click: () => sendCommand('open'),
          },
        ],
      },
    ],
    'extra-file-items':
      extraFileMenuItems.length > 0 ? [{ type: 'separator' }, ...extraFileMenuItems] : [],
    close: [
      shellHooks
        ? {
            label: tm('menuClose'),
            accelerator: 'CmdOrCtrl+W',
            click: () => shellHooks?.closeActiveTab(),
          }
        : { role: 'close', label: tm('menuClose') },
    ],
    'name-with-ai': [
      {
        ...stateOf('name-with-ai'),
        label: labels.nameWithAi,
        click: () => sendCommand('name-with-ai'),
      },
    ],
    'zoom-presets': [500, 200, 150, 125, 100, 75, 50, 25, 10].map((pct) => ({
      ...stateOf('zoom-set'),
      label: `${pct}%`,
      click: () => sendCommand('zoom-set', String(pct)),
    })),
    'settings-off-mac': isMac ? [] : [settingsMenuItem(labels, () => sendCommand('preferences'))],
    'dev-tools': isDev ? [toggleDevToolsItem(appMenuLabels(getUiLang()))] : [],
  }
  const build = (nodes: readonly MenuNode<MainKey>[]): MenuItemConstructorOptions[] =>
    nodes.flatMap((node): MenuItemConstructorOptions[] => {
      if (node === '---') return [{ type: 'separator' }]
      if (typeof node === 'string') return [cmd(node)]
      if ('slot' in node) return slots[node.slot]
      if ('cmd' in node) return [cmd(node.cmd, node.label, node.payload)]
      if ('role' in node) return [{ role: node.role, label: tm(node.label) }]
      return [{ label: tm(node.submenu), submenu: build(node.items) }]
    })
  const layout: readonly MenuTop<MainKey>[] = DOCS_MENU
  const template: MenuItemConstructorOptions[] = [
    ...(appMenu ? [appMenu] : []),
    ...layout.map((top) => ({ label: tm(top.label), submenu: build(top.items) })),
    windowMenuTemplate(process.platform, appMenuLabels(getUiLang())),
    {
      label: tm('menuHelp'),
      role: 'help',
      submenu: [
        cmd('shortcuts'),
        { type: 'separator' },
        checkUpdatesMenuItem(appMenuLabels(getUiLang())),
        aboutMenuItem(appMenuLabels(getUiLang())),
      ],
    },
  ]

  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- headless export ----

/** hidden export windows: webContents id -> what the renderer must write */
const headlessExportTargets = new Map<number, HeadlessExportTarget>()
/** settled by 'docs:headless-export-done' (or by the renderer dying) */
const headlessExportWaiters = new Map<number, (result: HeadlessExportReport) => void>()

interface HeadlessExportReport {
  ok: boolean
  error?: string
}

/**
 * Render `input` to `outPath` (as PDF or standalone HTML) with no visible window.
 *
 * The window is wired exactly like createDocsWindow's (same preload, sandbox
 * and `backgroundThrottling: false`) and the document rides the normal
 * pending-open queue, so the renderer runs its usual load -> paginate ->
 * export pipeline; only the save dialog is skipped, by pre-authorizing
 * `outPath` the way a dialog would. Rejects with the renderer's reason when
 * the export fails or the deadline passes.
 */
export async function exportDocsHeadless(
  input: string,
  outPath: string,
  format: HeadlessExportFormat = 'pdf',
  timeoutMs = 300_000,
): Promise<void> {
  const win = new BrowserWindow({
    show: false,
    width: 1360,
    height: 900,
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })
  const wcId = win.webContents.id
  pendingWindowOpens.set(wcId, input)
  headlessExportTargets.set(wcId, { outPath, format })
  allowPdfWrite(wcId, outPath)
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    const report = await new Promise<HeadlessExportReport>((resolve) => {
      headlessExportWaiters.set(wcId, resolve)
      win.webContents.on('render-process-gone', (_event, details) =>
        resolve({ ok: false, error: `docs renderer stopped (${details.reason})` }),
      )
      timer = setTimeout(
        () => resolve({ ok: false, error: `docs export timed out after ${timeoutMs}ms` }),
        timeoutMs,
      )
      void win.webContents.loadURL(rendererUrl(runtime.rendererUrl, 'docs'))
    })
    if (!report.ok) throw new Error(report.error ?? 'docs export failed')
  } finally {
    if (timer) clearTimeout(timer)
    headlessExportWaiters.delete(wcId)
    headlessExportTargets.delete(wcId)
    pendingWindowOpens.delete(wcId)
    if (!win.isDestroyed()) win.destroy()
  }
}

// ---- window ----

export function createDocsWindow(openPath?: string): BrowserWindow {
  const win = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 720,
    minHeight: 550,
    title: 'GenOffice Docs',
    // Word-like custom title bar (document name centered, quick-access buttons)
    ...(process.platform === 'darwin'
      ? { titleBarStyle: 'hiddenInset' as const }
      : {
          titleBarStyle: 'hidden' as const,
          titleBarOverlay: { color: '#ffffff', symbolColor: '#444444', height: 40 },
        }),
    // packaged builds embed the icon (build/icon.icns|ico); dev needs the file path
    ...(isDev && process.platform !== 'darwin'
      ? { icon: join(app.getAppPath(), 'build/icon.png') }
      : {}),
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })

  if (!mainWindow) {
    mainWindow = win
    rendererReady = false
  }
  // captured up front: webContents is already destroyed inside the 'closed' handler
  const webContentsId = win.webContents.id
  if (openPath) pendingWindowOpens.set(webContentsId, openPath)

  win.webContents.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: LINK_PROTOCOLS })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })

  void win.loadURL(rendererUrl(runtime.rendererUrl, 'docs'))
  // close guard for standalone-window mode (tab mode goes through the same flow via the shell's tab-manager/window-close path)
  let closeConfirmed = false
  win.on('close', (event) => {
    if (closeConfirmed) return
    event.preventDefault()
    void requestDocsClose(win.webContents, win).then((proceed) => {
      if (proceed && !win.isDestroyed()) {
        closeConfirmed = true
        win.close()
      }
    })
  })
  win.on('closed', () => {
    pendingWindowOpens.delete(webContentsId)
    dropDocWriter(webContentsId)
    // release any close-guard waiter still keyed on the gone webContents
    closeCheckWaiters.get(webContentsId)?.({ dirty: false, autoSave: false })
    closeCheckWaiters.delete(webContentsId)
    closeSaveWaiters.get(webContentsId)?.(false)
    closeSaveWaiters.delete(webContentsId)
    if (mainWindow === win) {
      mainWindow = BrowserWindow.getAllWindows().find((w) => w !== win) ?? null
      if (!mainWindow) rendererReady = false
    }
  })
  return win
}

/** tab-mode equivalent of createDocsWindow: same runtime/IPC wiring, no BrowserWindow of its own. */
// ── Close guard (aligned with sheets/pdf/slides): dirty documents prompt Save/Don't Save/Cancel before closing a tab/window ──
// docs' dirty state is a composite flag in the renderer; the main process doesn't mirror it and queries once at close time.
interface DocsCloseState {
  dirty: boolean
  autoSave: boolean
  /** Open file path (for recovery-copy cleanup on "Don't Save") */
  filePath?: string | null
  /** The renderer never replied to the close check (busy or wedged) */
  unresponsive?: boolean
}
const closeCheckWaiters = new Map<number, (state: DocsCloseState) => void>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()

ipcMain.on('docs:command-state', (event, value: unknown) => {
  const next = toCommandMenuState(value)
  if (!commandStateByWebContents.has(event.sender.id)) {
    const id = event.sender.id
    event.sender.once('destroyed', () => commandStateByWebContents.delete(id))
  }
  const previous = commandStateByWebContents.get(event.sender.id)
  commandStateByWebContents.set(event.sender.id, next)
  // patch the live menu only for the active tab; an inactive tab's state gets
  // picked up by the buildDocsMenu run its next focus triggers
  if (event.sender.id !== activeDocsWebContents()?.id) return
  // a built item's label cannot change, so Repeat's wording takes a rebuild
  if (previous?.repeat !== next.repeat) {
    buildDocsMenu()
    return
  }
  const menu = Menu.getApplicationMenu()
  if (!menu) return
  for (const [id, items] of commandMenuItems(menu)) {
    for (const item of items) {
      if (item.type === 'checkbox') item.checked = next.checked.includes(id as CommandId)
      item.enabled = !next.disabled.includes(id as CommandId)
    }
  }
})

ipcMain.handle('docs:key-catalog', () => keyCatalog())
ipcMain.handle('docs:key-overrides', () => userKeyOverrides())
ipcMain.handle('docs:set-key-overrides', (_event, value: unknown) => {
  keyOverrides = sanitizeKeyOverrides(value, process.platform === 'darwin')
  writeJsonAtomic(KEYBOARD_PATH(), keyOverrides)
  buildDocsMenu()
  for (const contents of webContents.getAllWebContents()) {
    if (!contents.isDestroyed()) contents.send('docs:key-overrides-changed', keyOverrides)
  }
  return keyOverrides
})

ipcMain.on('docs:close-check-result', (event, state: unknown) => {
  const waiter = closeCheckWaiters.get(event.sender.id)
  if (!waiter) return
  closeCheckWaiters.delete(event.sender.id)
  const s = state as { dirty?: unknown; autoSave?: unknown; filePath?: unknown } | boolean
  waiter(
    typeof s === 'boolean'
      ? { dirty: s, autoSave: false }
      : {
          dirty: s?.dirty === true,
          autoSave: s?.autoSave === true,
          filePath: typeof s?.filePath === 'string' ? s.filePath : null,
        },
  )
})

ipcMain.on('docs:close-save-result', (event, ok: unknown) => {
  const waiter = closeSaveWaiters.get(event.sender.id)
  if (!waiter) return
  closeSaveWaiters.delete(event.sender.id)
  waiter(ok === true)
})

/** Ask the renderer for pre-close state (dirty flag + autosave switch); no reply within 2s
 * fails CLOSED — a busy or wedged renderer may well hold unsaved changes, so the caller
 * prompts instead of closing silently. Concurrent callers share one query: a second
 * request must not overwrite the pending waiter (that stranded the first until timeout). */
const closeStateQueries = new Map<number, Promise<DocsCloseState>>()

function queryCloseState(contents: WebContents): Promise<DocsCloseState> {
  if (contents.isDestroyed()) return Promise.resolve({ dirty: false, autoSave: false })
  const pending = closeStateQueries.get(contents.id)
  if (pending) return pending
  const query = new Promise<DocsCloseState>((resolve) => {
    const timer = setTimeout(() => {
      closeCheckWaiters.delete(contents.id)
      resolve({ dirty: true, autoSave: false, unresponsive: true })
    }, 2000)
    closeCheckWaiters.set(contents.id, (state) => {
      clearTimeout(timer)
      resolve(state)
    })
    contents.send('docs:close-check')
  }).finally(() => closeStateQueries.delete(contents.id))
  closeStateQueries.set(contents.id, query)
  return query
}

export async function docsQueryDirty(contents: WebContents): Promise<boolean> {
  return (await queryCloseState(contents)).dirty
}

/** Ask the renderer to run the full save flow and await the result (failure/timeout = false). */
function requestRendererSave(contents: WebContents): Promise<boolean> {
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send('docs:close-save-request')
  })
}

/**
 * Close guard for the docs renderer: true means proceed with closing.
 * Clean → true; with changes → Save/Don't Save/Cancel. On Save, ask the renderer
 * to run the full save flow (new documents open Save As) and await the result;
 * failure/cancel/timeout keeps the document open.
 */
export function requestDocsClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  // re-entry (close clicked again while the prompt is up) joins the same flow
  // instead of stacking dialogs / stranding the first waiter
  const pending = docsCloseRequests.get(contents.id)
  if (pending) return pending
  const request = performDocsClose(contents, parent).finally(() =>
    docsCloseRequests.delete(contents.id),
  )
  docsCloseRequests.set(contents.id, request)
  return request
}

const docsCloseRequests = new Map<number, Promise<boolean>>()

async function performDocsClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  const state = await queryCloseState(contents)
  if (!state.dirty || contents.isDestroyed()) return true
  if (state.unresponsive) {
    // No reply: saving through the renderer won't work either — offer Close Anyway / Cancel
    const options = {
      type: 'warning' as const,
      message: tm('closeNoReplyMsg'),
      detail: tm('closeNoReplyDetail'),
      buttons: [tm('btnCloseAnyway'), tm('btnCancel')],
      defaultId: 1,
      cancelId: 1,
      noLink: true,
    }
    const { response } =
      parent && !parent.isDestroyed()
        ? await dialog.showMessageBox(parent, options)
        : await dialog.showMessageBox(options)
    return response === 0
  }
  // autosave on (and has a path, already checked when the renderer reported): save silently and proceed; only prompt on failure
  if (state.autoSave && (await requestRendererSave(contents))) return true
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('menuSave'), tm('btnDontSave'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) {
    // Explicitly discarded: also drop the recovery copy so the next open doesn't offer it
    if (state.filePath) clearRecoveryCopy(state.filePath)
    return true
  }
  return requestRendererSave(contents)
}

export function createDocsView(openPath?: string): WebContentsView {
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })

  if (openPath) pendingWindowOpens.set(view.webContents.id, openPath)

  view.webContents.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: LINK_PROTOCOLS })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })

  // mode=tab: the shell's tab strip owns the traffic lights / caption buttons,
  // so the ribbon must not reserve space for them
  void view.webContents.loadURL(rendererUrl(runtime.rendererUrl, 'docs', { mode: 'tab' }))
  // view.webContents becomes undefined after destroy, so grab the id beforehand
  const wcId = view.webContents.id
  view.webContents.once('destroyed', () => {
    pendingWindowOpens.delete(wcId)
    dropDocWriter(wcId)
    closeCheckWaiters.get(wcId)?.({ dirty: false, autoSave: false })
    closeCheckWaiters.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
  })
  return view
}

/** true when at least one docs window is open (shell menu switching) */
export function hasDocsWindow(): boolean {
  return mainWindow !== null
}

// ---- standalone lifecycle (apps/docs running on its own) ----

export function startDocsStandalone(): void {
  registerRendererScheme()
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  // dev runs must not share the packaged app's userData (recent files, AI settings)
  // or its single-instance lock — otherwise `npm run dev` silently quits whenever
  // the installed GenOffice Docs is open and forwards its argv there instead.
  // AI_OFFICE_USER_DATA: E2E/screenshot runs isolate userData (and the
  // single-instance lock) so parallel automation sessions don't evict each other
  if (process.env.AI_OFFICE_USER_DATA) app.setPath('userData', process.env.AI_OFFICE_USER_DATA)
  else if (isDev) {
    // "GenOffice Docs Dev" carried a space, against the XDG one-token
    // convention (genoffice#1817); move the old profile once, as the shell does.
    const appData = app.getPath('appData')
    const devDir = join(appData, 'GenOffice-Docs-Dev')
    const oldDevDir = join(appData, 'GenOffice Docs Dev')
    let userData = devDir
    if (existsSync(oldDevDir) && !existsSync(devDir)) {
      try {
        renameSync(oldDevDir, devDir)
      } catch {
        // another dev instance holds the directory: keep using the old one so
        // the profile is not lost and the move is retried on the next launch
        userData = oldDevDir
      }
    }
    app.setPath('userData', userData)
  }

  const hasSingleInstanceLock = app.requestSingleInstanceLock()
  if (!hasSingleInstanceLock) {
    app.quit()
    return
  }

  app.on('open-file', (event, filePath) => {
    event.preventDefault()
    openExternalDocx(filePath)
  })

  app.on('second-instance', (_event, argv) => {
    openExternalDocx(findOpenableDocPath(argv))
    mainWindow?.show()
    mainWindow?.focus()
  })

  registerAiIpc()
  registerProjectIpc()
  registerDocsIpc()

  app.whenReady().then(() => {
    installRendererProtocol({ docs: join(__dirname, '../renderer') })
    setUiLang(normalizeLang(process.env.GENOFFICE_LANG ?? app.getLocale()))
    // packaged builds get the Dock icon from icon.icns; dev shows Electron's default
    if (isDev && process.platform === 'darwin') {
      app.dock?.setIcon(join(app.getAppPath(), 'build/icon.png'))
    }
    buildDocsMenu()
    createDocsWindow()
    initDocsAutoUpdater(() => mainWindow)
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createDocsWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
}
