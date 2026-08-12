import { execSync, spawn } from 'node:child_process'
import {
  copyFileSync,
  cpSync,
  existsSync,
  readFileSync,
  readdirSync,
  renameSync,
  writeFileSync,
} from 'node:fs'
import { basename, dirname, extname, join } from 'node:path'
import {
  BrowserWindow,
  Menu,
  app,
  dialog,
  ipcMain,
  nativeImage,
  nativeTheme,
  session,
  shell,
  webContents,
} from 'electron'
import type { MenuItemConstructorOptions, NativeImage } from 'electron'
import menuDocxIcon1x from './assets/menu-docx.png?asset'
import menuDocxIcon2x from './assets/menu-docx@2x.png?asset'
import menuXlsxIcon1x from './assets/menu-xlsx.png?asset'
import menuXlsxIcon2x from './assets/menu-xlsx@2x.png?asset'
import menuPptxIcon1x from './assets/menu-pptx.png?asset'
import menuPptxIcon2x from './assets/menu-pptx@2x.png?asset'
import menuPdfIcon1x from './assets/menu-pdf.png?asset'
import menuPdfIcon2x from './assets/menu-pdf@2x.png?asset'
import menuMdIcon1x from './assets/menu-md.png?asset'
import menuMdIcon2x from './assets/menu-md@2x.png?asset'
import menuHomeIcon1x from './assets/menu-home.png?asset'
import menuHomeIcon2x from './assets/menu-home@2x.png?asset'
import { createI18n, isLang, normalizeLang, setUiLang, type Lang } from '@genoffice/i18n'
import {
  DEFAULT_SAVE_DIR_KEY,
  appMenuLabels,
  contextMenuLabels,
  editMenuTemplate,
  installContextMenu,
  installNavigationGuard,
  isUsableSaveDir,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  windowMenuTemplate,
} from '@genoffice/electron-utils'
import { readAppSettings, writeAppSetting } from './app-settings'
import {
  clearCloudProjectsStore,
  cloudProjectExternalUrl,
  readCloudProjectsStore,
  syncCloudProjects,
} from './cloud-projects'
import { ProjectStore } from '@genoffice/project-store'
import {
  ensureGenofficeLogin,
  genofficeLogout,
  gskConvertPdfToDocx,
  gskLoginInfo,
  hasGskAuth,
  loadGenofficeAuth,
  resolveGskEntry,
  setGskProxyUrl,
  startGenofficeLogin,
} from '@genoffice/ai-search'

import {
  buildDocsMenu,
  configureDocsRuntime,
  docsFileRenamed,
  docsQueryDirty,
  requestDocsClose,
  readRecentFiles,
  readStarredFiles,
  recordRecentFile,
  removeRecentFiles,
  replaceRecentFile,
  registerAiIpc,
  registerProjectIpc,
  toggleStarredFile,
  registerDocsIpc,
  setDocsExtraFileMenuItems,
  setDocsMenuGate,
  setDocsShellHooks,
  projectFileRenamed,
  setDocsShellWindow,
  setDocsFileSavedHook,
  setSessionPathResolver,
  defaultSaveDir,
  uniquePathIn,
} from '../../../docs/src/main/docs-main'
import { blankXlsxBuffer } from '../../../sheets/src/gateway/csv-import'
import {
  configureSheetsRuntime,
  hasQueuedWorkbook,
  installSheetsMenu,
  markSheetsShuttingDown,
  requestSheetsClose,
  resolveSheetsSessionPath,
  markSheetsUntitledPath,
  sendSheetsMenuAction,
  sheetsFileRenamed,
  setForcedWorkbookPath,
  setSheetsCloseTabHook,
  setSheetsExtraFileMenuItems,
  setSheetsShellWindow,
  setSheetsWorkbookOpenedHook,
  startSheetsCaptureServer,
  stopSheetsSidecar,
} from '../../../sheets/src/main/sheets-main'
import {
  configureSlidesRuntime,
  installSlidesMenu,
  replaceSlidesRecentFile,
  requestSlidesClose,
  setSlidesCloseTabHook,
  setSlidesExtraFileMenuItems,
  setSlidesOpenedHook,
  setSlidesShellWindow,
  slidesFileRenamed,
} from '../../../slides/src/main/slides-main'
import {
  configurePdfRuntime,
  flushPdfSave,
  pdfIsDirty,
  requestPdfClose,
  requestPdfSaveAs,
  setPdfSaveAsInFlight,
} from '../../../pdf/src/main/pdf-main'
import {
  configureMarkdownRuntime,
  markdownFileRenamed,
  requestMarkdownClose,
  requestMarkdownSave,
  sendMarkdownExportRequest,
  setMarkdownDocxExportedHook,
  setMarkdownFileSavedHook,
} from '../../../markdown/src/main/markdown-main'
import type {
  AccountLoginEvent,
  RecentEntry,
  RecentPage,
  RenameResult,
  UiTheme,
} from '../shared/home-api'
import { HOME_CHANNELS } from '../shared/home-api'
import type { TabKind } from '../shared/tabs-api'
import { TABS_CHANNELS } from '../shared/tabs-api'
import { showErrorDialog } from './error-dialog'
import { normalizeRecentQuery, pageRecentPaths, statExistingPaths } from './recent-files'
import { TabManager } from './tab-manager'
import { applyUpdateChannel, initAutoUpdater } from './updater'
import { isUpdateChannel, type UpdateChannel } from '../shared/update-api'

/**
 * GenOffice unified shell: ONE Electron app, ONE BrowserWindow, hosting the
 * docs and sheets modules as WebContentsView tabs behind a WPS-style tab
 * strip. The shell owns the lifecycle — single-instance lock, file-
 * association routing by extension, and per-active-tab menu switching.
 * Renderers load from each module's build output (apps/docs/out,
 * apps/sheets/out), so build those before running the shell.
 */

// ANY unpacked run (`npm run shell`, `npm run dev`, `npx electron .`) must not
// share the installed app's userData or single-instance lock — otherwise a dev
// run silently quits and forwards its argv to the running installed GenOffice.
// GENOFFICE_USER_DATA: test drivers point this at a scratch dir so an
// automated instance can run alongside the dev instance (separate lock).
if (!app.isPackaged)
  app.setPath(
    'userData',
    process.env.GENOFFICE_USER_DATA ?? join(app.getPath('appData'), 'GenOffice Dev'),
  )

// The product rename from "AI Office" to GenOffice changed the userData path; migrate old user data once
if (app.isPackaged) {
  const oldDir = join(app.getPath('appData'), 'AI Office')
  const newDir = app.getPath('userData')
  const newEmpty = !existsSync(newDir) || readdirSync(newDir).length === 0
  if (newEmpty && existsSync(oldDir)) cpSync(oldDir, newDir, { recursive: true })
}

// module build outputs: packaged builds carry them as extraResources
// (resources/modules/*, resources/native/*); dev/unpacked resolves them
// relative to apps/shell in the monorepo layout.
const SIDECAR_EXE = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
const APPS_ROOT = join(app.getAppPath(), '..')
const DOCS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'docs')
  : join(APPS_ROOT, 'docs', 'out')
const SHEETS_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'sheets')
  : join(APPS_ROOT, 'sheets', 'out')
const SLIDES_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'slides')
  : join(APPS_ROOT, 'slides', 'out')
const PDF_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'pdf')
  : join(APPS_ROOT, 'pdf', 'out')
const MARKDOWN_OUT = app.isPackaged
  ? join(process.resourcesPath, 'modules', 'markdown')
  : join(APPS_ROOT, 'markdown', 'out')
const SIDECAR_BIN = app.isPackaged
  ? join(process.resourcesPath, 'native', SIDECAR_EXE)
  : join(APPS_ROOT, 'sheets', 'native', 'xlsx-engine', 'target', 'release', SIDECAR_EXE)

configureDocsRuntime({
  preloadPath: join(DOCS_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.DOCS_RENDERER_URL,
  rendererFile: join(DOCS_OUT, 'renderer', 'index.html'),
})
configureSheetsRuntime({
  preloadPath: join(SHEETS_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.SHEETS_RENDERER_URL,
  rendererFile: join(SHEETS_OUT, 'renderer', 'index.html'),
  sidecarPath: SIDECAR_BIN,
})
configureSlidesRuntime({
  preloadPath: join(SLIDES_OUT, 'preload', 'index.js'),
  rendererDevUrl: process.env.SLIDES_RENDERER_URL,
  rendererFilePath: join(SLIDES_OUT, 'renderer', 'index.html'),
})
configurePdfRuntime({
  preloadPath: join(PDF_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.PDF_RENDERER_URL,
  rendererFile: join(PDF_OUT, 'renderer', 'index.html'),
})
configureMarkdownRuntime({
  preloadPath: join(MARKDOWN_OUT, 'preload', 'index.js'),
  rendererUrl: process.env.MARKDOWN_RENDERER_URL,
  rendererFile: join(MARKDOWN_OUT, 'renderer', 'index.html'),
})

// ---- UI language ----
// Persisted in userData/app-settings.json so the editor modules can read the
// same file when they pick up i18n later. GENOFFICE_LANG overrides for tests.

const APP_SETTINGS_PATH = () => join(app.getPath('userData'), 'app-settings.json')

let uiLang: Lang | null = null

function currentLang(): Lang {
  if (uiLang) return uiLang
  if (process.env.GENOFFICE_LANG) {
    uiLang = normalizeLang(process.env.GENOFFICE_LANG)
    setUiLang(uiLang)
    return uiLang
  }
  const saved = readAppSettings(APP_SETTINGS_PATH()).language
  if (isLang(saved)) uiLang = saved
  uiLang ??= normalizeLang(app.getLocale())
  setUiLang(uiLang)
  return uiLang
}

function persistLang(lang: Lang): void {
  uiLang = lang
  setUiLang(lang)
  writeAppSetting(APP_SETTINGS_PATH(), 'language', lang)
}

let cachedUpdateChannel: UpdateChannel | null = null

function currentUpdateChannel(): UpdateChannel {
  if (cachedUpdateChannel) return cachedUpdateChannel
  const saved = readAppSettings(APP_SETTINGS_PATH()).updateChannel
  cachedUpdateChannel = isUpdateChannel(saved) ? saved : 'stable'
  return cachedUpdateChannel
}

let cachedTheme: UiTheme | null = null

function currentTheme(): UiTheme {
  if (cachedTheme) return cachedTheme
  const saved = readAppSettings(APP_SETTINGS_PATH()).theme
  cachedTheme = saved === 'light' || saved === 'dark' ? saved : 'system'
  return cachedTheme
}

// ---- first-run onboarding ----
// The GenTeam community page opened from the onboarding's second slide.
// Stable short link served by the genoffice.ai site; it 302s to the tokened
// invite link, which stays out of this repo and rotates server-side.
const GENTEAM_URL = 'https://genoffice.ai/join'

// Genspark credit-usage page opened from the account menu's credits row.
// Kept main-side so the renderer never supplies the URL.
const CREDIT_USAGE_URL = 'https://www.genspark.ai/credit-usage'

const tMain = createI18n({
  zh: {
    menuFile: '文件',
    menuSectionNew: '新建',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: '未命名表格',
    untitledDoc: '未命名文档',
    untitledDeck: '未命名演示文稿',
    untitledMarkdown: '未命名 Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: '导出为 PDF…',
    menuOpenInDocs: '转换为 Docs 文档并打开',
    menuOpen: '打开…',
    menuSave: '保存',
    menuSaveAs: '另存为…',
    menuClose: '关闭',
    menuEdit: '编辑',
    menuWindow: '窗口',
    menuHome: '首页',
    backToHome: '返回首页',
    dlgOpenTitle: '打开文件',
    filterSupported: '支持的文件',
    filterWord: 'Word 文档',
    filterExcel: 'Excel 工作簿',
    filterPpt: 'PowerPoint 演示文稿',
    filterMarkdown: 'Markdown 文档',
    filterPdf: 'PDF 文档',
    errBadArgs: '参数无效',
    errBadName: '文件名不合法',
    errMissing: '文件不存在',
    errExists: '同名文件已存在',
    errRenameFailed: '重命名失败',
    errNewTabFailed: '新建文档失败',
    errUnsupportedExt: '暂不支持 .{ext} 类型',
    copySuffix: '副本',
    menuHelp: '帮助',
    thirdPartyNotices: '第三方软件声明',
    menuExportDocx: '导出为 Word…',
    pdfDocxLoginMsg: '导出为 Word 需要登录 Genspark 账号。',
    pdfDocxLoginDetail: '点击“登录”将打开浏览器完成授权，完成后请重新点击导出。',
    pdfDocxBtnLogin: '登录',
    pdfDocxConfirmMsg: '将此 PDF 上传到 Genspark 云端转换为 Word？',
    pdfDocxConfirmDetail: '本次转换将消耗 5 credits，文件将上传至云端处理。',
    pdfDocxConfirmBalance: '当前余额 {balance} credits。',
    pdfDocxBtnConvert: '继续',
    btnCancel: '取消',
    pdfDocxFailedMsg: '导出为 Word 失败',
    pdfDocxNoCliMsg: '无法登录 Genspark：缺少必需组件（gsk），请重新安装应用。',
    pdfDocxBusyMsg: '正在转换中，请等待当前导出完成。',
    dlgPickSaveDir: '选择默认保存位置',
    errSaveDirUnusable: '所选文件夹不可写，无法用作默认保存位置',
  },
  en: {
    menuFile: 'File',
    menuSectionNew: 'New',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Untitled Spreadsheet',
    untitledDoc: 'Untitled Document',
    untitledDeck: 'Untitled Presentation',
    untitledMarkdown: 'Untitled Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Export as PDF…',
    menuOpenInDocs: 'Convert and Open in Docs',
    menuOpen: 'Open…',
    menuSave: 'Save',
    menuSaveAs: 'Save As…',
    menuClose: 'Close',
    menuEdit: 'Edit',
    menuWindow: 'Window',
    menuHome: 'Home',
    backToHome: 'Back to Home',
    dlgOpenTitle: 'Open File',
    filterSupported: 'Supported Files',
    filterWord: 'Word Documents',
    filterExcel: 'Excel Workbooks',
    filterPpt: 'PowerPoint Presentations',
    filterMarkdown: 'Markdown Documents',
    filterPdf: 'PDF Documents',
    errBadArgs: 'Invalid arguments',
    errBadName: 'Invalid file name',
    errMissing: 'File not found',
    errExists: 'A file with that name already exists',
    errRenameFailed: 'Rename failed',
    errNewTabFailed: 'Could not create the new document',
    errUnsupportedExt: '.{ext} files are not supported',
    copySuffix: 'copy',
    menuHelp: 'Help',
    thirdPartyNotices: 'Third-Party Notices',
    menuExportDocx: 'Export as Word…',
    pdfDocxLoginMsg: 'Exporting as Word requires signing in to Genspark.',
    pdfDocxLoginDetail:
      'Clicking “Sign In” opens your browser to authorize; once done, click Export again.',
    pdfDocxBtnLogin: 'Sign In',
    pdfDocxConfirmMsg: 'Upload this PDF to Genspark cloud and convert it to Word?',
    pdfDocxConfirmDetail:
      'The conversion costs 5 credits. The file will be uploaded for cloud processing.',
    pdfDocxConfirmBalance: 'Current balance: {balance} credits.',
    pdfDocxBtnConvert: 'Continue',
    btnCancel: 'Cancel',
    pdfDocxFailedMsg: 'Export as Word failed',
    pdfDocxNoCliMsg:
      'Cannot sign in to Genspark: a required component (gsk) is missing. Please reinstall the app.',
    pdfDocxBusyMsg: 'A Word export is already in progress. Please wait for it to finish.',
    dlgPickSaveDir: 'Choose Default Save Location',
    errSaveDirUnusable:
      'The selected folder is not writable and cannot be used as the default save location',
  },
  ja: {
    menuFile: 'ファイル',
    menuSectionNew: '新規作成',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: '無題のスプレッドシート',
    untitledDoc: '無題のドキュメント',
    untitledDeck: '無題のプレゼンテーション',
    untitledMarkdown: '無題の Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'PDF として書き出す…',
    menuOpenInDocs: 'Docs 文書に変換して開く',
    menuOpen: '開く…',
    menuSave: '保存',
    menuSaveAs: '名前を付けて保存…',
    menuClose: '閉じる',
    menuEdit: '編集',
    menuWindow: 'ウィンドウ',
    menuHome: 'ホーム',
    backToHome: 'ホームに戻る',
    dlgOpenTitle: 'ファイルを開く',
    filterSupported: '対応ファイル',
    filterWord: 'Word 文書',
    filterExcel: 'Excel ブック',
    filterPpt: 'PowerPoint プレゼンテーション',
    filterMarkdown: 'Markdown ドキュメント',
    filterPdf: 'PDF ドキュメント',
    errBadArgs: '引数が無効です',
    errBadName: 'ファイル名が無効です',
    errMissing: 'ファイルが見つかりません',
    errExists: '同名のファイルが既に存在します',
    errRenameFailed: '名前の変更に失敗しました',
    errNewTabFailed: '新規ドキュメントを作成できませんでした',
    errUnsupportedExt: '.{ext} 形式には対応していません',
    copySuffix: 'コピー',
    menuHelp: 'ヘルプ',
    thirdPartyNotices: 'サードパーティソフトウェアに関する通知',
    menuExportDocx: 'Word として書き出す…',
    pdfDocxLoginMsg: 'Word への書き出しには Genspark へのログインが必要です。',
    pdfDocxLoginDetail:
      '「ログイン」をクリックするとブラウザで認証します。完了後、もう一度書き出しを実行してください。',
    pdfDocxBtnLogin: 'ログイン',
    pdfDocxConfirmMsg: 'この PDF を Genspark クラウドにアップロードして Word に変換しますか？',
    pdfDocxConfirmDetail:
      '変換には 5 クレジットを消費します。ファイルはクラウドにアップロードされ処理されます。',
    pdfDocxConfirmBalance: '現在の残高：{balance} クレジット。',
    pdfDocxBtnConvert: '続行',
    btnCancel: 'キャンセル',
    pdfDocxFailedMsg: 'Word への書き出しに失敗しました',
    pdfDocxNoCliMsg:
      'Genspark にサインインできません：必要なコンポーネント（gsk）が見つかりません。アプリを再インストールしてください。',
    pdfDocxBusyMsg: 'Word への書き出しが進行中です。完了までお待ちください。',
    dlgPickSaveDir: '既定の保存先を選択',
    errSaveDirUnusable:
      '選択したフォルダーは書き込みできないため、既定の保存先として使用できません',
  },
  ko: {
    menuFile: '파일',
    menuSectionNew: '새로 만들기',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: '제목 없는 스프레드시트',
    untitledDoc: '제목 없는 문서',
    untitledDeck: '제목 없는 프레젠테이션',
    untitledMarkdown: '제목 없는 Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'PDF로 내보내기…',
    menuOpenInDocs: 'Docs 문서로 변환하여 열기',
    menuOpen: '열기…',
    menuSave: '저장',
    menuSaveAs: '다른 이름으로 저장…',
    menuClose: '닫기',
    menuEdit: '편집',
    menuWindow: '창',
    menuHome: '홈',
    backToHome: '홈으로 돌아가기',
    dlgOpenTitle: '파일 열기',
    filterSupported: '지원되는 파일',
    filterWord: 'Word 문서',
    filterExcel: 'Excel 통합 문서',
    filterPpt: 'PowerPoint 프레젠테이션',
    filterMarkdown: 'Markdown 문서',
    filterPdf: 'PDF 문서',
    errBadArgs: '잘못된 인수입니다',
    errBadName: '파일 이름이 잘못되었습니다',
    errMissing: '파일을 찾을 수 없습니다',
    errExists: '같은 이름의 파일이 이미 있습니다',
    errRenameFailed: '이름 바꾸기에 실패했습니다',
    errNewTabFailed: '새 문서를 만들지 못했습니다',
    errUnsupportedExt: '.{ext} 형식은 지원되지 않습니다',
    copySuffix: '복사본',
    menuHelp: '도움말',
    thirdPartyNotices: '타사 소프트웨어 고지',
    menuExportDocx: 'Word로 내보내기…',
    pdfDocxLoginMsg: 'Word로 내보내려면 Genspark 로그인이 필요합니다.',
    pdfDocxLoginDetail:
      '“로그인”을 클릭하면 브라우저에서 인증합니다. 완료 후 내보내기를 다시 클릭하세요.',
    pdfDocxBtnLogin: '로그인',
    pdfDocxConfirmMsg: '이 PDF를 Genspark 클라우드에 업로드하여 Word로 변환할까요?',
    pdfDocxConfirmDetail:
      '변환에는 5 크레딧이 소모됩니다. 파일은 클라우드로 업로드되어 처리됩니다.',
    pdfDocxConfirmBalance: '현재 잔액: {balance} 크레딧.',
    pdfDocxBtnConvert: '계속',
    btnCancel: '취소',
    pdfDocxFailedMsg: 'Word로 내보내기 실패',
    pdfDocxNoCliMsg:
      'Genspark에 로그인할 수 없습니다. 필수 구성 요소(gsk)가 없습니다. 앱을 다시 설치해 주세요.',
    pdfDocxBusyMsg: 'Word 내보내기가 이미 진행 중입니다. 완료될 때까지 기다려 주세요.',
    dlgPickSaveDir: '기본 저장 위치 선택',
    errSaveDirUnusable: '선택한 폴더에 쓸 수 없어 기본 저장 위치로 사용할 수 없습니다',
  },
  fr: {
    menuFile: 'Fichier',
    menuSectionNew: 'Nouveau',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Feuille de calcul sans titre',
    untitledDoc: 'Document sans titre',
    untitledDeck: 'Présentation sans titre',
    untitledMarkdown: 'Markdown sans titre',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Exporter en PDF…',
    menuOpenInDocs: 'Convertir et ouvrir dans Docs',
    menuOpen: 'Ouvrir…',
    menuSave: 'Enregistrer',
    menuSaveAs: 'Enregistrer sous…',
    menuClose: 'Fermer',
    menuEdit: 'Édition',
    menuWindow: 'Fenêtre',
    menuHome: 'Accueil',
    backToHome: "Retour à l'accueil",
    dlgOpenTitle: 'Ouvrir un fichier',
    filterSupported: 'Fichiers pris en charge',
    filterWord: 'Documents Word',
    filterExcel: 'Classeurs Excel',
    filterPpt: 'Présentations PowerPoint',
    filterMarkdown: 'Documents Markdown',
    filterPdf: 'Documents PDF',
    errBadArgs: 'Arguments non valides',
    errBadName: 'Nom de fichier non valide',
    errMissing: 'Fichier introuvable',
    errExists: 'Un fichier du même nom existe déjà',
    errRenameFailed: 'Échec du renommage',
    errNewTabFailed: 'Impossible de créer le nouveau document',
    errUnsupportedExt: 'les fichiers .{ext} ne sont pas pris en charge',
    copySuffix: 'copie',
    menuHelp: 'Aide',
    thirdPartyNotices: 'Mentions relatives aux logiciels tiers',
    menuExportDocx: 'Exporter en Word…',
    pdfDocxLoginMsg: "L'export en Word nécessite une connexion à Genspark.",
    pdfDocxLoginDetail:
      "Cliquez sur « Se connecter » pour autoriser dans le navigateur, puis relancez l'export.",
    pdfDocxBtnLogin: 'Se connecter',
    pdfDocxConfirmMsg: 'Téléverser ce PDF vers le cloud Genspark pour le convertir en Word ?',
    pdfDocxConfirmDetail:
      'La conversion coûte 5 crédits. Le fichier sera téléversé pour traitement dans le cloud.',
    pdfDocxConfirmBalance: 'Solde actuel : {balance} crédits.',
    pdfDocxBtnConvert: 'Continuer',
    btnCancel: 'Annuler',
    pdfDocxFailedMsg: "Échec de l'export en Word",
    pdfDocxNoCliMsg:
      "Connexion à Genspark impossible : un composant requis (gsk) est manquant. Veuillez réinstaller l'application.",
    pdfDocxBusyMsg: "Un export en Word est déjà en cours. Veuillez attendre qu'il se termine.",
    dlgPickSaveDir: "Choisir l'emplacement d'enregistrement par défaut",
    errSaveDirUnusable:
      "Le dossier sélectionné n'est pas accessible en écriture et ne peut pas servir d'emplacement d'enregistrement par défaut",
  },
  de: {
    menuFile: 'Datei',
    menuSectionNew: 'Neu',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Unbenannte Tabelle',
    untitledDoc: 'Unbenanntes Dokument',
    untitledDeck: 'Unbenannte Präsentation',
    untitledMarkdown: 'Unbenanntes Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Als PDF exportieren…',
    menuOpenInDocs: 'In Docs umwandeln und öffnen',
    menuOpen: 'Öffnen…',
    menuSave: 'Speichern',
    menuSaveAs: 'Speichern unter…',
    menuClose: 'Schließen',
    menuEdit: 'Bearbeiten',
    menuWindow: 'Fenster',
    menuHome: 'Startseite',
    backToHome: 'Zurück zur Startseite',
    dlgOpenTitle: 'Datei öffnen',
    filterSupported: 'Unterstützte Dateien',
    filterWord: 'Word-Dokumente',
    filterExcel: 'Excel-Arbeitsmappen',
    filterPpt: 'PowerPoint-Präsentationen',
    filterMarkdown: 'Markdown-Dokumente',
    filterPdf: 'PDF-Dokumente',
    errBadArgs: 'Ungültige Argumente',
    errBadName: 'Ungültiger Dateiname',
    errMissing: 'Datei nicht gefunden',
    errExists: 'Eine Datei mit diesem Namen existiert bereits',
    errRenameFailed: 'Umbenennen fehlgeschlagen',
    errNewTabFailed: 'Neues Dokument konnte nicht erstellt werden',
    errUnsupportedExt: '.{ext}-Dateien werden nicht unterstützt',
    copySuffix: 'Kopie',
    menuHelp: 'Hilfe',
    thirdPartyNotices: 'Hinweise zu Drittanbietersoftware',
    menuExportDocx: 'Als Word exportieren…',
    pdfDocxLoginMsg: 'Für den Word-Export ist eine Anmeldung bei Genspark erforderlich.',
    pdfDocxLoginDetail:
      'Klicken Sie auf „Anmelden“, um die Autorisierung im Browser abzuschließen, und starten Sie den Export danach erneut.',
    pdfDocxBtnLogin: 'Anmelden',
    pdfDocxConfirmMsg: 'Dieses PDF in die Genspark-Cloud hochladen und in Word konvertieren?',
    pdfDocxConfirmDetail:
      'Die Konvertierung kostet 5 Credits. Die Datei wird zur Verarbeitung in die Cloud hochgeladen.',
    pdfDocxConfirmBalance: 'Aktuelles Guthaben: {balance} Credits.',
    pdfDocxBtnConvert: 'Fortfahren',
    btnCancel: 'Abbrechen',
    pdfDocxFailedMsg: 'Word-Export fehlgeschlagen',
    pdfDocxNoCliMsg:
      'Anmeldung bei Genspark nicht möglich: Eine erforderliche Komponente (gsk) fehlt. Bitte installieren Sie die App neu.',
    pdfDocxBusyMsg: 'Ein Word-Export läuft bereits. Bitte warten Sie, bis er abgeschlossen ist.',
    dlgPickSaveDir: 'Standard-Speicherort auswählen',
    errSaveDirUnusable:
      'Der ausgewählte Ordner ist nicht beschreibbar und kann nicht als Standard-Speicherort verwendet werden',
  },
  es: {
    menuFile: 'Archivo',
    menuSectionNew: 'Nuevo',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Hoja de cálculo sin título',
    untitledDoc: 'Documento sin título',
    untitledDeck: 'Presentación sin título',
    untitledMarkdown: 'Markdown sin título',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Exportar como PDF…',
    menuOpenInDocs: 'Convertir y abrir en Docs',
    menuOpen: 'Abrir…',
    menuSave: 'Guardar',
    menuSaveAs: 'Guardar como…',
    menuClose: 'Cerrar',
    menuEdit: 'Edición',
    menuWindow: 'Ventana',
    menuHome: 'Inicio',
    backToHome: 'Volver al inicio',
    dlgOpenTitle: 'Abrir archivo',
    filterSupported: 'Archivos compatibles',
    filterWord: 'Documentos de Word',
    filterExcel: 'Libros de Excel',
    filterPpt: 'Presentaciones de PowerPoint',
    filterMarkdown: 'Documentos Markdown',
    filterPdf: 'Documentos PDF',
    errBadArgs: 'Argumentos no válidos',
    errBadName: 'Nombre de archivo no válido',
    errMissing: 'Archivo no encontrado',
    errExists: 'Ya existe un archivo con ese nombre',
    errRenameFailed: 'No se pudo cambiar el nombre',
    errNewTabFailed: 'No se pudo crear el nuevo documento',
    errUnsupportedExt: 'los archivos .{ext} no son compatibles',
    copySuffix: 'copia',
    menuHelp: 'Ayuda',
    thirdPartyNotices: 'Avisos de software de terceros',
    menuExportDocx: 'Exportar como Word…',
    pdfDocxLoginMsg: 'Para exportar como Word es necesario iniciar sesión en Genspark.',
    pdfDocxLoginDetail:
      'Al hacer clic en «Iniciar sesión» se abrirá el navegador para autorizar; después, vuelve a hacer clic en Exportar.',
    pdfDocxBtnLogin: 'Iniciar sesión',
    pdfDocxConfirmMsg: '¿Subir este PDF a la nube de Genspark para convertirlo a Word?',
    pdfDocxConfirmDetail:
      'La conversión cuesta 5 créditos. El archivo se subirá para procesarse en la nube.',
    pdfDocxConfirmBalance: 'Saldo actual: {balance} créditos.',
    pdfDocxBtnConvert: 'Continuar',
    btnCancel: 'Cancelar',
    pdfDocxFailedMsg: 'Error al exportar como Word',
    pdfDocxNoCliMsg:
      'No se puede iniciar sesión en Genspark: falta un componente necesario (gsk). Reinstale la aplicación.',
    pdfDocxBusyMsg: 'Ya hay una exportación a Word en curso. Espera a que termine.',
    dlgPickSaveDir: 'Elegir ubicación de guardado predeterminada',
    errSaveDirUnusable:
      'La carpeta seleccionada no admite escritura y no puede usarse como ubicación de guardado predeterminada',
  },
  th: {
    menuFile: 'ไฟล์',
    menuSectionNew: 'สร้างใหม่',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'สเปรดชีตไม่มีชื่อ',
    untitledDoc: 'เอกสารไม่มีชื่อ',
    untitledDeck: 'งานนำเสนอไม่มีชื่อ',
    untitledMarkdown: 'Markdown ไม่มีชื่อ',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'ส่งออกเป็น PDF…',
    menuOpenInDocs: 'แปลงและเปิดใน Docs',
    menuOpen: 'เปิด…',
    menuSave: 'บันทึก',
    menuSaveAs: 'บันทึกเป็น…',
    menuClose: 'ปิด',
    menuEdit: 'แก้ไข',
    menuWindow: 'หน้าต่าง',
    menuHome: 'หน้าแรก',
    backToHome: 'กลับไปหน้าแรก',
    dlgOpenTitle: 'เปิดไฟล์',
    filterSupported: 'ไฟล์ที่รองรับ',
    filterWord: 'เอกสาร Word',
    filterExcel: 'เวิร์กบุ๊ก Excel',
    filterPpt: 'งานนำเสนอ PowerPoint',
    filterMarkdown: 'เอกสาร Markdown',
    filterPdf: 'เอกสาร PDF',
    errBadArgs: 'อาร์กิวเมนต์ไม่ถูกต้อง',
    errBadName: 'ชื่อไฟล์ไม่ถูกต้อง',
    errMissing: 'ไม่พบไฟล์',
    errExists: 'มีไฟล์ชื่อเดียวกันอยู่แล้ว',
    errRenameFailed: 'เปลี่ยนชื่อไม่สำเร็จ',
    errNewTabFailed: 'สร้างเอกสารใหม่ไม่สำเร็จ',
    errUnsupportedExt: 'ไม่รองรับไฟล์ .{ext}',
    copySuffix: 'สำเนา',
    menuHelp: 'วิธีใช้',
    thirdPartyNotices: 'ประกาศเกี่ยวกับซอฟต์แวร์ของบุคคลที่สาม',
    menuExportDocx: 'ส่งออกเป็น Word…',
    pdfDocxLoginMsg: 'การส่งออกเป็น Word ต้องเข้าสู่ระบบ Genspark',
    pdfDocxLoginDetail:
      'คลิก “เข้าสู่ระบบ” เพื่อเปิดเบราว์เซอร์ยืนยันตัวตน เสร็จแล้วให้คลิกส่งออกอีกครั้ง',
    pdfDocxBtnLogin: 'เข้าสู่ระบบ',
    pdfDocxConfirmMsg: 'อัปโหลด PDF นี้ไปยังคลาวด์ Genspark เพื่อแปลงเป็น Word หรือไม่?',
    pdfDocxConfirmDetail: 'การแปลงใช้ 5 เครดิต ไฟล์จะถูกอัปโหลดเพื่อประมวลผลบนคลาวด์',
    pdfDocxConfirmBalance: 'ยอดคงเหลือปัจจุบัน: {balance} เครดิต',
    pdfDocxBtnConvert: 'ดำเนินการต่อ',
    btnCancel: 'ยกเลิก',
    pdfDocxFailedMsg: 'ส่งออกเป็น Word ไม่สำเร็จ',
    pdfDocxNoCliMsg:
      'ไม่สามารถลงชื่อเข้าใช้ Genspark ได้: ไม่พบคอมโพเนนต์ที่จำเป็น (gsk) โปรดติดตั้งแอปใหม่',
    pdfDocxBusyMsg: 'กำลังส่งออกเป็น Word อยู่ โปรดรอให้เสร็จสิ้นก่อน',
    dlgPickSaveDir: 'เลือกตำแหน่งบันทึกเริ่มต้น',
    errSaveDirUnusable: 'โฟลเดอร์ที่เลือกไม่สามารถเขียนได้ จึงใช้เป็นตำแหน่งบันทึกเริ่มต้นไม่ได้',
  },
  id: {
    menuFile: 'File',
    menuSectionNew: 'Baru',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Spreadsheet tanpa judul',
    untitledDoc: 'Dokumen tanpa judul',
    untitledDeck: 'Presentasi tanpa judul',
    untitledMarkdown: 'Markdown tanpa judul',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Ekspor sebagai PDF…',
    menuOpenInDocs: 'Konversi dan buka di Docs',
    menuOpen: 'Buka…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuClose: 'Tutup',
    menuEdit: 'Edit',
    menuWindow: 'Jendela',
    menuHome: 'Beranda',
    backToHome: 'Kembali ke Beranda',
    dlgOpenTitle: 'Buka File',
    filterSupported: 'File yang Didukung',
    filterWord: 'Dokumen Word',
    filterExcel: 'Buku Kerja Excel',
    filterPpt: 'Presentasi PowerPoint',
    filterMarkdown: 'Dokumen Markdown',
    filterPdf: 'Dokumen PDF',
    errBadArgs: 'Argumen tidak valid',
    errBadName: 'Nama file tidak valid',
    errMissing: 'File tidak ditemukan',
    errExists: 'File dengan nama tersebut sudah ada',
    errRenameFailed: 'Gagal mengganti nama',
    errNewTabFailed: 'Gagal membuat dokumen baru',
    errUnsupportedExt: 'file .{ext} tidak didukung',
    copySuffix: 'salinan',
    menuHelp: 'Bantuan',
    thirdPartyNotices: 'Pemberitahuan Perangkat Lunak Pihak Ketiga',
    menuExportDocx: 'Ekspor sebagai Word…',
    pdfDocxLoginMsg: 'Ekspor sebagai Word memerlukan login ke Genspark.',
    pdfDocxLoginDetail:
      'Klik “Masuk” untuk membuka browser dan memberi otorisasi; setelah selesai, klik Ekspor lagi.',
    pdfDocxBtnLogin: 'Masuk',
    pdfDocxConfirmMsg: 'Unggah PDF ini ke cloud Genspark untuk dikonversi ke Word?',
    pdfDocxConfirmDetail:
      'Konversi ini menggunakan 5 kredit. File akan diunggah untuk diproses di cloud.',
    pdfDocxConfirmBalance: 'Saldo saat ini: {balance} kredit.',
    pdfDocxBtnConvert: 'Lanjutkan',
    btnCancel: 'Batal',
    pdfDocxFailedMsg: 'Gagal mengekspor sebagai Word',
    pdfDocxNoCliMsg:
      'Tidak dapat masuk ke Genspark: komponen yang diperlukan (gsk) tidak ditemukan. Silakan instal ulang aplikasi.',
    pdfDocxBusyMsg: 'Ekspor ke Word sedang berlangsung. Harap tunggu hingga selesai.',
    dlgPickSaveDir: 'Pilih Lokasi Penyimpanan Default',
    errSaveDirUnusable:
      'Folder yang dipilih tidak dapat ditulis dan tidak bisa digunakan sebagai lokasi penyimpanan default',
  },
  ru: {
    menuFile: 'Файл',
    menuSectionNew: 'Создать',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Таблица без названия',
    untitledDoc: 'Документ без названия',
    untitledDeck: 'Презентация без названия',
    untitledMarkdown: 'Markdown без названия',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Экспортировать в PDF…',
    menuOpenInDocs: 'Преобразовать и открыть в Docs',
    menuOpen: 'Открыть…',
    menuSave: 'Сохранить',
    menuSaveAs: 'Сохранить как…',
    menuClose: 'Закрыть',
    menuEdit: 'Правка',
    menuWindow: 'Окно',
    menuHome: 'Главная',
    backToHome: 'Вернуться на главную',
    dlgOpenTitle: 'Открытие файла',
    filterSupported: 'Поддерживаемые файлы',
    filterWord: 'Документы Word',
    filterExcel: 'Книги Excel',
    filterPpt: 'Презентации PowerPoint',
    filterMarkdown: 'Документы Markdown',
    filterPdf: 'Документы PDF',
    errBadArgs: 'Недопустимые аргументы',
    errBadName: 'Недопустимое имя файла',
    errMissing: 'Файл не найден',
    errExists: 'Файл с таким именем уже существует',
    errRenameFailed: 'Не удалось переименовать',
    errNewTabFailed: 'Не удалось создать новый документ',
    errUnsupportedExt: 'файлы .{ext} не поддерживаются',
    copySuffix: 'копия',
    menuHelp: 'Справка',
    thirdPartyNotices: 'Уведомления о стороннем ПО',
    menuExportDocx: 'Экспортировать в Word…',
    pdfDocxLoginMsg: 'Для экспорта в Word требуется вход в Genspark.',
    pdfDocxLoginDetail:
      'Нажмите «Войти», чтобы авторизоваться в браузере, затем снова запустите экспорт.',
    pdfDocxBtnLogin: 'Войти',
    pdfDocxConfirmMsg: 'Загрузить этот PDF в облако Genspark и конвертировать в Word?',
    pdfDocxConfirmDetail:
      'Конвертация стоит 5 кредитов. Файл будет загружен для обработки в облаке.',
    pdfDocxConfirmBalance: 'Текущий баланс: {balance} кредитов.',
    pdfDocxBtnConvert: 'Продолжить',
    btnCancel: 'Отмена',
    pdfDocxFailedMsg: 'Не удалось экспортировать в Word',
    pdfDocxNoCliMsg:
      'Не удаётся войти в Genspark: отсутствует необходимый компонент (gsk). Переустановите приложение.',
    pdfDocxBusyMsg: 'Экспорт в Word уже выполняется. Дождитесь его завершения.',
    dlgPickSaveDir: 'Выбрать папку сохранения по умолчанию',
    errSaveDirUnusable:
      'Выбранная папка недоступна для записи и не может использоваться как папка сохранения по умолчанию',
  },
  ar: {
    menuFile: 'ملف',
    menuSectionNew: 'جديد',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'جدول بيانات بلا عنوان',
    untitledDoc: 'مستند بدون عنوان',
    untitledDeck: 'عرض تقديمي بدون عنوان',
    untitledMarkdown: 'Markdown بدون عنوان',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'تصدير بتنسيق PDF…',
    menuOpenInDocs: 'التحويل والفتح في Docs',
    menuOpen: 'فتح…',
    menuSave: 'حفظ',
    menuSaveAs: 'حفظ باسم…',
    menuClose: 'إغلاق',
    menuEdit: 'تحرير',
    menuWindow: 'نافذة',
    menuHome: 'الصفحة الرئيسية',
    backToHome: 'العودة إلى الصفحة الرئيسية',
    dlgOpenTitle: 'فتح ملف',
    filterSupported: 'الملفات المدعومة',
    filterWord: 'مستندات Word',
    filterExcel: 'مصنفات Excel',
    filterPpt: 'عروض PowerPoint التقديمية',
    filterMarkdown: 'مستندات Markdown',
    filterPdf: 'مستندات PDF',
    errBadArgs: 'وسيطات غير صالحة',
    errBadName: 'اسم ملف غير صالح',
    errMissing: 'الملف غير موجود',
    errExists: 'يوجد ملف بالاسم نفسه بالفعل',
    errRenameFailed: 'فشلت إعادة التسمية',
    errNewTabFailed: 'تعذّر إنشاء المستند الجديد',
    errUnsupportedExt: 'ملفات .{ext} غير مدعومة',
    copySuffix: 'نسخة',
    menuHelp: 'تعليمات',
    thirdPartyNotices: 'إشعارات برامج الجهات الخارجية',
    menuExportDocx: 'تصدير كملف Word…',
    pdfDocxLoginMsg: 'يتطلب التصدير كملف Word تسجيل الدخول إلى Genspark.',
    pdfDocxLoginDetail:
      'انقر على «تسجيل الدخول» لفتح المتصفح وإتمام التفويض، ثم انقر على التصدير مرة أخرى.',
    pdfDocxBtnLogin: 'تسجيل الدخول',
    pdfDocxConfirmMsg: 'رفع هذا الـ PDF إلى سحابة Genspark وتحويله إلى Word؟',
    pdfDocxConfirmDetail: 'يكلف التحويل 5 أرصدة. سيتم رفع الملف للمعالجة في السحابة.',
    pdfDocxConfirmBalance: 'الرصيد الحالي: {balance} من الأرصدة.',
    pdfDocxBtnConvert: 'متابعة',
    btnCancel: 'إلغاء',
    pdfDocxFailedMsg: 'فشل التصدير كملف Word',
    pdfDocxNoCliMsg:
      'تعذّر تسجيل الدخول إلى Genspark: المكوّن المطلوب (gsk) مفقود. يُرجى إعادة تثبيت التطبيق.',
    pdfDocxBusyMsg: 'يجري حاليًا تصدير إلى Word. يُرجى الانتظار حتى يكتمل.',
    dlgPickSaveDir: 'اختيار موقع الحفظ الافتراضي',
    errSaveDirUnusable: 'المجلد المحدد غير قابل للكتابة ولا يمكن استخدامه كموقع حفظ افتراضي',
  },
  pt: {
    menuFile: 'Arquivo',
    menuSectionNew: 'Novo',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Planilha sem título',
    untitledDoc: 'Documento sem título',
    untitledDeck: 'Apresentação sem título',
    untitledMarkdown: 'Markdown sem título',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Exportar como PDF…',
    menuOpenInDocs: 'Converter e abrir no Docs',
    menuOpen: 'Abrir…',
    menuSave: 'Salvar',
    menuSaveAs: 'Salvar Como…',
    menuClose: 'Fechar',
    menuEdit: 'Editar',
    menuWindow: 'Janela',
    menuHome: 'Início',
    backToHome: 'Voltar ao início',
    dlgOpenTitle: 'Abrir arquivo',
    filterSupported: 'Arquivos compatíveis',
    filterWord: 'Documentos do Word',
    filterExcel: 'Pastas de trabalho do Excel',
    filterPpt: 'Apresentações do PowerPoint',
    filterMarkdown: 'Documentos Markdown',
    filterPdf: 'Documentos PDF',
    errBadArgs: 'Argumentos inválidos',
    errBadName: 'Nome de arquivo inválido',
    errMissing: 'Arquivo não encontrado',
    errExists: 'Já existe um arquivo com esse nome',
    errRenameFailed: 'Falha ao renomear',
    errNewTabFailed: 'Falha ao criar o novo documento',
    errUnsupportedExt: 'arquivos .{ext} não são suportados',
    copySuffix: 'cópia',
    menuHelp: 'Ajuda',
    thirdPartyNotices: 'Avisos de software de terceiros',
    menuExportDocx: 'Exportar como Word…',
    pdfDocxLoginMsg: 'Exportar como Word requer login no Genspark.',
    pdfDocxLoginDetail:
      'Clique em “Entrar” para autorizar no navegador; depois, clique em Exportar novamente.',
    pdfDocxBtnLogin: 'Entrar',
    pdfDocxConfirmMsg: 'Enviar este PDF para a nuvem do Genspark e convertê-lo em Word?',
    pdfDocxConfirmDetail:
      'A conversão custa 5 créditos. O arquivo será enviado para processamento na nuvem.',
    pdfDocxConfirmBalance: 'Saldo atual: {balance} créditos.',
    pdfDocxBtnConvert: 'Continuar',
    btnCancel: 'Cancelar',
    pdfDocxFailedMsg: 'Falha ao exportar como Word',
    pdfDocxNoCliMsg:
      'Não é possível iniciar sessão no Genspark: falta um componente necessário (gsk). Reinstale o aplicativo.',
    pdfDocxBusyMsg: 'Já há uma exportação para Word em andamento. Aguarde a conclusão.',
    dlgPickSaveDir: 'Escolher local de salvamento padrão',
    errSaveDirUnusable:
      'A pasta selecionada não permite gravação e não pode ser usada como local de salvamento padrão',
  },
  it: {
    menuFile: 'File',
    menuSectionNew: 'Nuovo',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Foglio di calcolo senza titolo',
    untitledDoc: 'Documento senza titolo',
    untitledDeck: 'Presentazione senza titolo',
    untitledMarkdown: 'Markdown senza titolo',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Esporta come PDF…',
    menuOpenInDocs: 'Converti e apri in Docs',
    menuOpen: 'Apri…',
    menuSave: 'Salva',
    menuSaveAs: 'Salva con nome…',
    menuClose: 'Chiudi',
    menuEdit: 'Modifica',
    menuWindow: 'Finestra',
    menuHome: 'Home',
    backToHome: 'Torna alla Home',
    dlgOpenTitle: 'Apri file',
    filterSupported: 'File supportati',
    filterWord: 'Documenti Word',
    filterExcel: 'Cartelle di lavoro Excel',
    filterPpt: 'Presentazioni PowerPoint',
    filterMarkdown: 'Documenti Markdown',
    filterPdf: 'Documenti PDF',
    errBadArgs: 'Argomenti non validi',
    errBadName: 'Nome file non valido',
    errMissing: 'File non trovato',
    errExists: 'Esiste già un file con questo nome',
    errRenameFailed: 'Impossibile rinominare',
    errNewTabFailed: 'Impossibile creare il nuovo documento',
    errUnsupportedExt: 'i file .{ext} non sono supportati',
    copySuffix: 'copia',
    menuHelp: 'Aiuto',
    thirdPartyNotices: 'Note sul software di terze parti',
    menuExportDocx: 'Esporta come Word…',
    pdfDocxLoginMsg: 'Per esportare come Word è necessario accedere a Genspark.',
    pdfDocxLoginDetail:
      'Fai clic su “Accedi” per autorizzare nel browser; al termine, fai di nuovo clic su Esporta.',
    pdfDocxBtnLogin: 'Accedi',
    pdfDocxConfirmMsg: 'Caricare questo PDF sul cloud Genspark e convertirlo in Word?',
    pdfDocxConfirmDetail:
      "La conversione costa 5 crediti. Il file verrà caricato per l'elaborazione nel cloud.",
    pdfDocxConfirmBalance: 'Saldo attuale: {balance} crediti.',
    pdfDocxBtnConvert: 'Continua',
    btnCancel: 'Annulla',
    pdfDocxFailedMsg: 'Esportazione in Word non riuscita',
    pdfDocxNoCliMsg:
      "Impossibile accedere a Genspark: manca un componente necessario (gsk). Reinstallare l'app.",
    pdfDocxBusyMsg: "Un'esportazione in Word è già in corso. Attendi il completamento.",
    dlgPickSaveDir: 'Scegli la posizione di salvataggio predefinita',
    errSaveDirUnusable:
      'La cartella selezionata non è scrivibile e non può essere usata come posizione di salvataggio predefinita',
  },
  pl: {
    menuFile: 'Plik',
    menuSectionNew: 'Nowy',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Arkusz bez tytułu',
    untitledDoc: 'Dokument bez tytułu',
    untitledDeck: 'Prezentacja bez tytułu',
    untitledMarkdown: 'Markdown bez tytułu',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Eksportuj jako PDF…',
    menuOpenInDocs: 'Konwertuj i otwórz w Docs',
    menuOpen: 'Otwórz…',
    menuSave: 'Zapisz',
    menuSaveAs: 'Zapisz jako…',
    menuClose: 'Zamknij',
    menuEdit: 'Edycja',
    menuWindow: 'Okno',
    menuHome: 'Strona główna',
    backToHome: 'Wróć do strony głównej',
    dlgOpenTitle: 'Otwieranie pliku',
    filterSupported: 'Obsługiwane pliki',
    filterWord: 'Dokumenty programu Word',
    filterExcel: 'Skoroszyty programu Excel',
    filterPpt: 'Prezentacje programu PowerPoint',
    filterMarkdown: 'Dokumenty Markdown',
    filterPdf: 'Dokumenty PDF',
    errBadArgs: 'Nieprawidłowe argumenty',
    errBadName: 'Nieprawidłowa nazwa pliku',
    errMissing: 'Nie znaleziono pliku',
    errExists: 'Plik o tej nazwie już istnieje',
    errRenameFailed: 'Nie udało się zmienić nazwy',
    errNewTabFailed: 'Nie udało się utworzyć nowego dokumentu',
    errUnsupportedExt: 'pliki .{ext} nie są obsługiwane',
    copySuffix: 'kopia',
    menuHelp: 'Pomoc',
    thirdPartyNotices: 'Informacje o oprogramowaniu innych firm',
    menuExportDocx: 'Eksportuj jako Word…',
    pdfDocxLoginMsg: 'Eksport do formatu Word wymaga zalogowania do Genspark.',
    pdfDocxLoginDetail:
      'Kliknij „Zaloguj się”, aby autoryzować w przeglądarce; po zakończeniu kliknij Eksportuj ponownie.',
    pdfDocxBtnLogin: 'Zaloguj się',
    pdfDocxConfirmMsg: 'Przesłać ten PDF do chmury Genspark i przekonwertować na Word?',
    pdfDocxConfirmDetail:
      'Konwersja kosztuje 5 kredytów. Plik zostanie przesłany do przetworzenia w chmurze.',
    pdfDocxConfirmBalance: 'Aktualne saldo: {balance} kredytów.',
    pdfDocxBtnConvert: 'Kontynuuj',
    btnCancel: 'Anuluj',
    pdfDocxFailedMsg: 'Eksport do formatu Word nie powiódł się',
    pdfDocxNoCliMsg:
      'Nie można zalogować się do Genspark: brakuje wymaganego komponentu (gsk). Zainstaluj aplikację ponownie.',
    pdfDocxBusyMsg: 'Eksport do formatu Word już trwa. Poczekaj na jego zakończenie.',
    dlgPickSaveDir: 'Wybierz domyślną lokalizację zapisu',
    errSaveDirUnusable:
      'Wybrany folder nie pozwala na zapis i nie może być domyślną lokalizacją zapisu',
  },
  nl: {
    menuFile: 'Bestand',
    menuSectionNew: 'Nieuw',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Naamloze spreadsheet',
    untitledDoc: 'Naamloos document',
    untitledDeck: 'Naamloze presentatie',
    untitledMarkdown: 'Naamloos Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Exporteren als PDF…',
    menuOpenInDocs: 'Converteren en openen in Docs',
    menuOpen: 'Openen…',
    menuSave: 'Opslaan',
    menuSaveAs: 'Opslaan als…',
    menuClose: 'Sluiten',
    menuEdit: 'Bewerken',
    menuWindow: 'Venster',
    menuHome: 'Start',
    backToHome: 'Terug naar start',
    dlgOpenTitle: 'Bestand openen',
    filterSupported: 'Ondersteunde bestanden',
    filterWord: 'Word-documenten',
    filterExcel: 'Excel-werkmappen',
    filterPpt: 'PowerPoint-presentaties',
    filterMarkdown: 'Markdown-documenten',
    filterPdf: 'PDF-documenten',
    errBadArgs: 'Ongeldige argumenten',
    errBadName: 'Ongeldige bestandsnaam',
    errMissing: 'Bestand niet gevonden',
    errExists: 'Er bestaat al een bestand met die naam',
    errRenameFailed: 'Naam wijzigen mislukt',
    errNewTabFailed: 'Kan het nieuwe document niet maken',
    errUnsupportedExt: '.{ext}-bestanden worden niet ondersteund',
    copySuffix: 'kopie',
    menuHelp: 'Help',
    thirdPartyNotices: 'Kennisgevingen over software van derden',
    menuExportDocx: 'Exporteren als Word…',
    pdfDocxLoginMsg: 'Exporteren als Word vereist inloggen bij Genspark.',
    pdfDocxLoginDetail:
      'Klik op “Inloggen” om in de browser te autoriseren; klik daarna opnieuw op Exporteren.',
    pdfDocxBtnLogin: 'Inloggen',
    pdfDocxConfirmMsg: 'Deze PDF uploaden naar de Genspark-cloud en converteren naar Word?',
    pdfDocxConfirmDetail:
      'De conversie kost 5 credits. Het bestand wordt geüpload voor verwerking in de cloud.',
    pdfDocxConfirmBalance: 'Huidig saldo: {balance} credits.',
    pdfDocxBtnConvert: 'Doorgaan',
    btnCancel: 'Annuleren',
    pdfDocxFailedMsg: 'Exporteren als Word mislukt',
    pdfDocxNoCliMsg:
      'Kan niet inloggen bij Genspark: een vereist onderdeel (gsk) ontbreekt. Installeer de app opnieuw.',
    pdfDocxBusyMsg: 'Er is al een Word-export bezig. Wacht tot deze is voltooid.',
    dlgPickSaveDir: 'Standaard opslaglocatie kiezen',
    errSaveDirUnusable:
      'De geselecteerde map is niet beschrijfbaar en kan niet als standaard opslaglocatie worden gebruikt',
  },
  ms: {
    menuFile: 'Fail',
    menuSectionNew: 'Baharu',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'Hamparan tanpa tajuk',
    untitledDoc: 'Dokumen tanpa tajuk',
    untitledDeck: 'Persembahan tanpa tajuk',
    untitledMarkdown: 'Markdown tanpa tajuk',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'Eksport sebagai PDF…',
    menuOpenInDocs: 'Tukar dan buka dalam Docs',
    menuOpen: 'Buka…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuClose: 'Tutup',
    menuEdit: 'Edit',
    menuWindow: 'Tetingkap',
    menuHome: 'Laman Utama',
    backToHome: 'Kembali ke Laman Utama',
    dlgOpenTitle: 'Buka Fail',
    filterSupported: 'Fail yang Disokong',
    filterWord: 'Dokumen Word',
    filterExcel: 'Buku Kerja Excel',
    filterPpt: 'Persembahan PowerPoint',
    filterMarkdown: 'Dokumen Markdown',
    filterPdf: 'Dokumen PDF',
    errBadArgs: 'Argumen tidak sah',
    errBadName: 'Nama fail tidak sah',
    errMissing: 'Fail tidak ditemui',
    errExists: 'Fail dengan nama yang sama sudah wujud',
    errRenameFailed: 'Gagal menamakan semula',
    errNewTabFailed: 'Gagal mencipta dokumen baharu',
    errUnsupportedExt: 'fail .{ext} tidak disokong',
    copySuffix: 'salinan',
    menuHelp: 'Bantuan',
    thirdPartyNotices: 'Notis Perisian Pihak Ketiga',
    menuExportDocx: 'Eksport sebagai Word…',
    pdfDocxLoginMsg: 'Eksport sebagai Word memerlukan log masuk ke Genspark.',
    pdfDocxLoginDetail:
      'Klik “Log Masuk” untuk membuka pelayar dan memberi kebenaran; selepas selesai, klik Eksport sekali lagi.',
    pdfDocxBtnLogin: 'Log Masuk',
    pdfDocxConfirmMsg: 'Muat naik PDF ini ke awan Genspark untuk ditukar kepada Word?',
    pdfDocxConfirmDetail:
      'Penukaran ini menggunakan 5 kredit. Fail akan dimuat naik untuk diproses di awan.',
    pdfDocxConfirmBalance: 'Baki semasa: {balance} kredit.',
    pdfDocxBtnConvert: 'Teruskan',
    btnCancel: 'Batal',
    pdfDocxFailedMsg: 'Gagal mengeksport sebagai Word',
    pdfDocxNoCliMsg:
      'Tidak dapat log masuk ke Genspark: komponen yang diperlukan (gsk) tiada. Sila pasang semula aplikasi.',
    pdfDocxBusyMsg: 'Eksport ke Word sedang dijalankan. Sila tunggu sehingga selesai.',
    dlgPickSaveDir: 'Pilih Lokasi Simpanan Lalai',
    errSaveDirUnusable:
      'Folder yang dipilih tidak boleh ditulis dan tidak dapat digunakan sebagai lokasi simpanan lalai',
  },
  he: {
    menuFile: 'קובץ',
    menuSectionNew: 'חדש',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'גיליון אלקטרוני ללא שם',
    untitledDoc: 'מסמך ללא שם',
    untitledDeck: 'מצגת ללא שם',
    untitledMarkdown: 'Markdown ללא שם',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'ייצוא כ-PDF…',
    menuOpenInDocs: 'המרה ופתיחה ב-Docs',
    menuOpen: 'פתיחה…',
    menuSave: 'שמירה',
    menuSaveAs: 'שמירה בשם…',
    menuClose: 'סגירה',
    menuEdit: 'עריכה',
    menuWindow: 'חלון',
    menuHome: 'דף הבית',
    backToHome: 'חזרה לדף הבית',
    dlgOpenTitle: 'פתיחת קובץ',
    filterSupported: 'קבצים נתמכים',
    filterWord: 'מסמכי Word',
    filterExcel: 'חוברות עבודה של Excel',
    filterPpt: 'מצגות PowerPoint',
    filterMarkdown: 'מסמכי Markdown',
    filterPdf: 'מסמכי PDF',
    errBadArgs: 'ארגומנטים לא חוקיים',
    errBadName: 'שם קובץ לא חוקי',
    errMissing: 'הקובץ לא נמצא',
    errExists: 'כבר קיים קובץ באותו שם',
    errRenameFailed: 'שינוי השם נכשל',
    errNewTabFailed: 'יצירת המסמך החדש נכשלה',
    errUnsupportedExt: 'קובצי .{ext} אינם נתמכים',
    copySuffix: 'עותק',
    menuHelp: 'עזרה',
    thirdPartyNotices: 'הודעות על תוכנות צד שלישי',
    menuExportDocx: 'ייצוא כ-Word…',
    pdfDocxLoginMsg: 'ייצוא כ-Word דורש התחברות ל-Genspark.',
    pdfDocxLoginDetail: 'לחיצה על ”התחברות” תפתח את הדפדפן לאישור; בסיום, לחצו שוב על ייצוא.',
    pdfDocxBtnLogin: 'התחברות',
    pdfDocxConfirmMsg: 'להעלות את ה-PDF לענן של Genspark ולהמיר אותו ל-Word?',
    pdfDocxConfirmDetail: 'ההמרה עולה 5 קרדיטים. הקובץ יועלה לעיבוד בענן.',
    pdfDocxConfirmBalance: 'יתרה נוכחית: {balance} קרדיטים.',
    pdfDocxBtnConvert: 'המשך',
    btnCancel: 'ביטול',
    pdfDocxFailedMsg: 'הייצוא כ-Word נכשל',
    pdfDocxNoCliMsg: 'לא ניתן להתחבר ל-Genspark: רכיב נדרש (gsk) חסר. נא להתקין מחדש את האפליקציה.',
    pdfDocxBusyMsg: 'ייצוא ל-Word כבר מתבצע. נא להמתין לסיומו.',
    dlgPickSaveDir: 'בחירת מיקום שמירה כברירת מחדל',
    errSaveDirUnusable:
      'התיקייה שנבחרה אינה ניתנת לכתיבה ולא ניתן להשתמש בה כמיקום שמירה כברירת מחדל',
  },
  hi: {
    menuFile: 'फ़ाइल',
    menuSectionNew: 'नया',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: 'शीर्षकहीन स्प्रेडशीट',
    untitledDoc: 'बिना शीर्षक दस्तावेज़',
    untitledDeck: 'बिना शीर्षक प्रस्तुति',
    untitledMarkdown: 'अनाम Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: 'PDF के रूप में निर्यात…',
    menuOpenInDocs: 'Docs में बदलें और खोलें',
    menuOpen: 'खोलें…',
    menuSave: 'सहेजें',
    menuSaveAs: 'इस रूप में सहेजें…',
    menuClose: 'बंद करें',
    menuEdit: 'संपादन',
    menuWindow: 'विंडो',
    menuHome: 'होम',
    backToHome: 'होम पर वापस जाएँ',
    dlgOpenTitle: 'फ़ाइल खोलें',
    filterSupported: 'समर्थित फ़ाइलें',
    filterWord: 'Word दस्तावेज़',
    filterExcel: 'Excel वर्कबुक',
    filterPpt: 'PowerPoint प्रस्तुतियाँ',
    filterMarkdown: 'Markdown दस्तावेज़',
    filterPdf: 'PDF दस्तावेज़',
    errBadArgs: 'अमान्य आर्ग्युमेंट',
    errBadName: 'अमान्य फ़ाइल नाम',
    errMissing: 'फ़ाइल नहीं मिली',
    errExists: 'इस नाम की फ़ाइल पहले से मौजूद है',
    errRenameFailed: 'नाम बदलने में विफल',
    errNewTabFailed: 'नया दस्तावेज़ बनाने में विफल',
    errUnsupportedExt: '.{ext} फ़ाइलें समर्थित नहीं हैं',
    copySuffix: 'प्रतिलिपि',
    menuHelp: 'सहायता',
    thirdPartyNotices: 'तृतीय-पक्ष सॉफ़्टवेयर सूचनाएँ',
    menuExportDocx: 'Word के रूप में निर्यात करें…',
    pdfDocxLoginMsg: 'Word के रूप में निर्यात करने के लिए Genspark में लॉगिन आवश्यक है।',
    pdfDocxLoginDetail:
      '“लॉगिन” पर क्लिक करने से ब्राउज़र में प्राधिकरण खुलेगा; पूरा होने पर फिर से निर्यात पर क्लिक करें।',
    pdfDocxBtnLogin: 'लॉगिन',
    pdfDocxConfirmMsg: 'इस PDF को Genspark क्लाउड पर अपलोड करके Word में बदलें?',
    pdfDocxConfirmDetail:
      'रूपांतरण में 5 क्रेडिट लगते हैं। फ़ाइल क्लाउड में प्रोसेसिंग के लिए अपलोड की जाएगी।',
    pdfDocxConfirmBalance: 'वर्तमान शेष: {balance} क्रेडिट।',
    pdfDocxBtnConvert: 'जारी रखें',
    btnCancel: 'रद्द करें',
    pdfDocxFailedMsg: 'Word के रूप में निर्यात विफल रहा',
    pdfDocxNoCliMsg:
      'Genspark में साइन इन नहीं किया जा सकता: आवश्यक घटक (gsk) मौजूद नहीं है। कृपया ऐप को फिर से इंस्टॉल करें।',
    pdfDocxBusyMsg: 'Word के रूप में निर्यात पहले से चल रहा है। कृपया पूरा होने तक प्रतीक्षा करें।',
    dlgPickSaveDir: 'डिफ़ॉल्ट सहेजने का स्थान चुनें',
    errSaveDirUnusable:
      'चयनित फ़ोल्डर में लिखा नहीं जा सकता, इसलिए इसे डिफ़ॉल्ट सहेजने के स्थान के रूप में उपयोग नहीं किया जा सकता',
  },
  'zh-TW': {
    menuFile: '檔案',
    menuSectionNew: '新增',
    menuNewDoc: 'AI Docs',
    menuNewSheet: 'AI Sheets',
    untitledSheet: '未命名試算表',
    untitledDoc: '未命名文件',
    untitledDeck: '未命名簡報',
    untitledMarkdown: '未命名 Markdown',
    menuNewSlide: 'AI Slides',
    menuNewMarkdown: 'AI Markdown',
    menuExportPdf: '匯出為 PDF…',
    menuOpenInDocs: '轉換為 Docs 文件並開啟',
    menuOpen: '開啟…',
    menuSave: '儲存',
    menuSaveAs: '另存新檔…',
    menuClose: '關閉',
    menuEdit: '編輯',
    menuWindow: '視窗',
    menuHome: '首頁',
    backToHome: '返回首頁',
    dlgOpenTitle: '開啟檔案',
    filterSupported: '支援的檔案',
    filterWord: 'Word 文件',
    filterExcel: 'Excel 活頁簿',
    filterPpt: 'PowerPoint 簡報',
    filterMarkdown: 'Markdown 文件',
    filterPdf: 'PDF 文件',
    errBadArgs: '參數無效',
    errBadName: '檔案名稱不合法',
    errMissing: '檔案不存在',
    errExists: '同名檔案已存在',
    errRenameFailed: '重新命名失敗',
    errNewTabFailed: '新建文件失敗',
    errUnsupportedExt: '暫不支援 .{ext} 類型',
    copySuffix: '副本',
    menuHelp: '說明',
    thirdPartyNotices: '第三方軟體聲明',
    menuExportDocx: '匯出為 Word…',
    pdfDocxLoginMsg: '匯出為 Word 需要登入 Genspark 帳號。',
    pdfDocxLoginDetail: '點擊「登入」將開啟瀏覽器完成授權，完成後請重新點擊匯出。',
    pdfDocxBtnLogin: '登入',
    pdfDocxConfirmMsg: '將此 PDF 上傳到 Genspark 雲端轉換為 Word？',
    pdfDocxConfirmDetail: '本次轉換將消耗 5 credits，檔案將上傳至雲端處理。',
    pdfDocxConfirmBalance: '目前餘額 {balance} credits。',
    pdfDocxBtnConvert: '繼續',
    btnCancel: '取消',
    pdfDocxFailedMsg: '匯出為 Word 失敗',
    pdfDocxNoCliMsg: '無法登入 Genspark：缺少必要元件（gsk），請重新安裝應用程式。',
    pdfDocxBusyMsg: '正在轉換中，請等待目前的匯出完成。',
    dlgPickSaveDir: '選擇預設儲存位置',
    errSaveDirUnusable: '所選資料夾無法寫入，無法作為預設儲存位置',
  },
})

const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(currentLang(), key, params)

// ---- the shell window + its tab manager (recreated if the user closes it on macOS) ----

let shellWindow: BrowserWindow | null = null
let tabManager: TabManager | null = null

/**
 * When the user creates a file from a specific project view, remember which
 * project the next save should belong to. key: 'doc' | 'sheet' | 'slide', value: projectId.
 * Consumed by each app's saveHook once the file first hits disk (P1 item 3).
 */
const pendingNewFileProject = new Map<string, string>()

/**
 * P1: after a file first hits disk, if a pending project was set earlier via
 * "create from project view", move the new file into that project automatically.
 * Called from createShellWindow's opened/saved hooks.
 */
function applyPendingProject(filePath: string): void {
  const ext = extname(filePath).slice(1).toLowerCase()
  let key: string | undefined
  if (ext === 'docx') key = 'doc'
  else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') key = 'sheet'
  else if (ext === 'pptx') key = 'slide'
  else if (ext === 'md' || ext === 'markdown') key = 'markdown'
  if (!key) return
  const projectId = pendingNewFileProject.get(key)
  if (!projectId) return
  pendingNewFileProject.delete(key)
  try {
    const store = new ProjectStore(app.getPath('userData'))
    store.ensureDefaultProject()
    store.resolveProjectForFile(filePath) // assign to default first (idempotent)
    store.moveFileToProject(filePath, projectId)
  } catch (err) {
    console.warn('[shell] applyPendingProject failed:', err)
  }
}

function applyMenuFor(kind: TabKind): void {
  switch (kind) {
    case 'docs':
      buildDocsMenu()
      break
    case 'sheets':
      installSheetsMenu()
      break
    case 'slides':
      installSlidesMenu()
      break
    case 'pdf':
      buildPdfMenu()
      break
    case 'markdown':
      buildMarkdownMenu()
      break
    default:
      buildHomeMenu()
  }
}

function createShellWindow(): void {
  const win = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 980,
    minHeight: 600,
    title: 'GenOffice',
    // vibrancy: editor modules punch translucent regions (e.g. the slides
    // thumbnail pane) through to the desktop
    ...(process.platform === 'darwin'
      ? { titleBarStyle: 'hiddenInset' as const, vibrancy: 'sidebar' as const }
      : {}),
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  shellWindow = win
  // dragging the window by the tab strip's blank (draggable) area produces no
  // DOM event anywhere — will-move is the only signal to dismiss popovers
  win.on('will-move', broadcastChromePressed)

  const manager = new TabManager(
    win,
    () => win.webContents.send(TABS_CHANNELS.changed, manager.list()),
    applyMenuFor,
    // no extension: these tabs have no file on disk yet; the title becomes the
    // real filename (the localized untitled default + .docx etc.) once the first save lands
    (kind) =>
      kind === 'docs'
        ? tm('untitledDoc')
        : kind === 'slides'
          ? tm('untitledDeck')
          : kind === 'markdown'
            ? tm('untitledMarkdown')
            : tm('untitledSheet'),
  )
  tabManager = manager

  // pushRecent-triggered docs menu rebuilds must not clobber the active tab's menu
  setDocsMenuGate(() => manager.list().some((t) => t.active && t.kind === 'docs'))

  setDocsShellWindow(win)
  setSheetsShellWindow(win)
  setSlidesShellWindow(win)
  setDocsShellHooks({
    openTab: (openPath, options) => manager.openDocsTab(openPath, options),
    listTabs: () =>
      manager
        .list()
        .filter((t) => t.kind === 'docs')
        .map((t) => ({ id: t.id, title: t.title, focused: t.active })),
    focusTab: (id) => manager.activateTab(id),
    closeActiveTab: () => manager.closeActiveTab(),
  })
  setSheetsCloseTabHook(() => manager.closeActiveTab())
  setSlidesCloseTabHook(() => manager.closeActiveTab())
  // When ⌘O opens a file inside a tab, sync the tab title/path (used for de-dup by path) and record it as recent.
  // The first save / save-as fires this too, so applyPendingProject also runs here.
  setSheetsWorkbookOpenedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  setSlidesOpenedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // docs' save-as / silent first save lands on a new path → sync the tab title too
  setDocsFileSavedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // markdown untitled first save / Save As lands on a new path
  setMarkdownFileSavedHook((wc, path) => {
    manager.setTabFileFor(wc.id, path)
    recordRecentFile(path)
    applyPendingProject(path)
  })
  // markdown "convert & open in Docs" → route the fresh .docx to a docs tab
  setMarkdownDocxExportedHook((path) => {
    openDocumentPath(path)
  })

  // Closing the whole window walks every dirty sheets/pdf/slides/docs tab through
  // the same save/don't-save/cancel prompt; any cancel aborts the close.
  // docs dirtiness lives renderer-side, so any live docs tab forces the async path
  // and gets queried there (clean tabs pass through without activation).
  let closeConfirmed = false
  win.on('close', (event) => {
    if (closeConfirmed) return
    const dirtySheets = manager.dirtySheetsTabs()
    const dirtyPdf = manager.dirtyPdfTabs()
    const dirtyMarkdown = manager.dirtyMarkdownTabs()
    const dirtySlides = manager.dirtySlidesTabs()
    const docsTabs = manager.docsTabs()
    if (
      dirtySheets.length === 0 &&
      dirtyPdf.length === 0 &&
      dirtyMarkdown.length === 0 &&
      dirtySlides.length === 0 &&
      docsTabs.length === 0
    )
      return
    event.preventDefault()
    void (async () => {
      for (const tab of dirtySheets) {
        manager.activateTab(tab.id)
        if (!(await requestSheetsClose(tab.webContents, win))) return
      }
      for (const tab of dirtyPdf) {
        manager.activateTab(tab.id)
        if (!(await requestPdfClose(tab.webContents, win))) return
      }
      for (const tab of dirtyMarkdown) {
        manager.activateTab(tab.id)
        if (!(await requestMarkdownClose(tab.webContents, win))) return
      }
      for (const tab of dirtySlides) {
        manager.activateTab(tab.id)
        if (!(await requestSlidesClose(tab.webContents, win))) return
      }
      for (const tab of docsTabs) {
        if (!(await docsQueryDirty(tab.webContents))) continue
        manager.activateTab(tab.id)
        if (!(await requestDocsClose(tab.webContents, win))) return
      }
      closeConfirmed = true
      if (!win.isDestroyed()) win.close()
    })()
  })

  win.on('closed', () => {
    if (shellWindow === win) shellWindow = null
    if (tabManager === manager) tabManager = null
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    void win.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    void win.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

// ---- routing: one dispatch function for every open path ----

const DOCX_RE = /\.docx$/i
const XLSX_RE = /\.(xlsx|xls|csv)$/i
const PPTX_RE = /\.pptx$/i
const PDF_RE = /\.pdf$/i
const MD_RE = /\.(md|markdown)$/i

/** document formats we recognize but don't open — surfaced as a dialog, not silently dropped */
const UNSUPPORTED_DOC_RE = /\.(doc|rtf|odt|ppt|pps|odp|ods|xlsm|xlsb|pages|key|numbers)$/i

/**
 * Single source of truth for the open-dialog filter. Includes the
 * legacy .doc/.ppt binaries so they are selectable and surface the explicit
 * "not supported" dialog via openDocumentPath instead of being grayed out.
 */
const OPEN_DIALOG_EXTENSIONS = [
  'docx',
  'doc',
  'xlsx',
  'xls',
  'csv',
  'pptx',
  'ppt',
  'pdf',
  'md',
  'markdown',
]

function supportedFileIn(argv: string[]): string | null {
  return (
    argv.find(
      (arg) =>
        (DOCX_RE.test(arg) ||
          XLSX_RE.test(arg) ||
          PPTX_RE.test(arg) ||
          PDF_RE.test(arg) ||
          MD_RE.test(arg)) &&
        existsSync(arg),
    ) ?? null
  )
}

function unsupportedFileIn(argv: string[]): string | null {
  return argv.find((arg) => UNSUPPORTED_DOC_RE.test(arg) && existsSync(arg)) ?? null
}

function notifyUnsupportedFile(filePath: string): void {
  const ext = extname(filePath).slice(1).toLowerCase() || basename(filePath)
  const options = { type: 'warning' as const, message: tm('errUnsupportedExt', { ext }) }
  if (shellWindow) {
    shellWindow.show()
    shellWindow.focus()
    void dialog.showMessageBox(shellWindow, options)
  } else {
    void dialog.showMessageBox(options)
  }
}

/** the single router: extension decides which module owns the file; false = nothing opened */
function openDocumentPath(filePath: string): boolean {
  if (!existsSync(filePath) || !tabManager) return false
  if (DOCX_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findDocsTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openDocsTab(filePath)
    return true
  }
  if (XLSX_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findSheetsTabByPath(filePath)
    if (existing) {
      tabManager.activateTab(existing)
    } else {
      setForcedWorkbookPath(filePath)
      tabManager.openSheetsTab(filePath)
      startQueuedWorkbookNudge()
    }
    return true
  }
  if (PPTX_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findSlidesTabByPath(filePath)
    if (existing) {
      tabManager.activateTab(existing)
    } else {
      // For a new tab the path goes through the pending queue; the renderer consumes it after mounting
      tabManager.openSlidesTab(filePath)
    }
    return true
  }
  if (PDF_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findPdfTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openPdfTab(filePath)
    return true
  }
  if (MD_RE.test(filePath)) {
    recordRecentFile(filePath)
    const existing = tabManager.findMarkdownTabByPath(filePath)
    if (existing) tabManager.activateTab(existing)
    else tabManager.openMarkdownTab(filePath)
    return true
  }
  notifyUnsupportedFile(filePath)
  return false
}

/**
 * "New spreadsheet" creates the backing .xlsx in the default folder up front and
 * opens it as a regular file tab — the blank in-memory demo mode has no save
 * pipeline, so the file must exist before edits. Falls back to the old blank
 * tab if the write fails.
 */
async function newSheetTab(): Promise<void> {
  try {
    const filePath = uniquePathIn(defaultSaveDir(), `${tm('untitledSheet')}.xlsx`)
    writeFileSync(filePath, await blankXlsxBuffer())
    // eligible for content-derived auto-rename after the first AI generation
    markSheetsUntitledPath(filePath)
    openDocumentPath(filePath)
  } catch (err) {
    console.warn('[shell] blank workbook create failed, opening in-memory blank tab:', err)
    try {
      tabManager?.openSheetsTab(undefined, { newBlank: true })
    } catch (fallbackErr) {
      surfaceNewTabError(fallbackErr)
    }
  }
}

/**
 * A throw anywhere in the create-tab path (view creation, sidecar resolution,
 * renderer load) used to be swallowed by `void`-ed promises and ipc-invoke
 * rejections, so the click looked like a pure no-op — the exact "AI Sheets /
 * AI Slides do nothing" alpha report. Surface the failure instead.
 */
function surfaceNewTabError(err: unknown): void {
  console.error('[shell] new tab failed:', err)
  showErrorDialog(shellWindow, tm('errNewTabFailed'), err)
}

function newDocTab(): void {
  try {
    tabManager?.openDocsTab(undefined, { newBlank: true })
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newSlideTab(): void {
  try {
    tabManager?.openSlidesTab()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

function newMarkdownTab(): void {
  try {
    tabManager?.openMarkdownTab()
  } catch (err) {
    surfaceNewTabError(err)
  }
}

/**
 * The sheets renderer subscribes to menu actions only after Univer finishes
 * mounting (seconds on cold start), so a single 'open' can fire into the
 * void. Re-send until the queued workbook is consumed; consumption clears the
 * queue flag main-side (sheets-main), which stops the loop.
 */
let workbookNudgeTimer: ReturnType<typeof setInterval> | null = null

function startQueuedWorkbookNudge(): void {
  if (workbookNudgeTimer) clearInterval(workbookNudgeTimer)
  const startedAt = Date.now()
  sendSheetsMenuAction('open')
  workbookNudgeTimer = setInterval(() => {
    if (!hasQueuedWorkbook() || Date.now() - startedAt > 30_000 || !tabManager?.findSheetsTab()) {
      if (workbookNudgeTimer) clearInterval(workbookNudgeTimer)
      workbookNudgeTimer = null
      return
    }
    sendSheetsMenuAction('open')
  }, 700)
}

// ---- home IPC ----

function statEntries(paths: string[]): RecentEntry[] {
  return statExistingPaths(paths, new Set(readStarredFiles()))
}

function registerHomeIpc(): void {
  // signed-in means GenOffice's own device-code login; the shared gsk CLI key
  // is only a silent fallback, deliberately not shown here to nudge users onto our key
  ipcMain.handle(HOME_CHANNELS.accountStatus, async () => {
    if (!loadGenofficeAuth()) return { loggedIn: false }
    await proxyBootstrap
    const info = await gskLoginInfo()
    return info
      ? { loggedIn: true, email: info.email, creditBalance: info.creditBalance }
      : { loggedIn: true }
  })

  // login progress is streamed to the requesting renderer; the auth URL is
  // kept main-side so the "open manually" rescue never opens a renderer-supplied URL
  let pendingLoginUrl = ''
  ipcMain.handle(HOME_CHANNELS.accountLogin, async (event) => {
    const sender = event.sender
    pendingLoginUrl = ''
    await proxyBootstrap
    const send = (payload: AccountLoginEvent) => {
      if (!sender.isDestroyed()) sender.send(HOME_CHANNELS.accountLoginEvent, payload)
    }
    // open the browser on the first url event only; later events refresh the rescue URL
    let opened = false
    const launched = startGenofficeLogin((progress) => {
      if (progress.url) {
        pendingLoginUrl = progress.url
        if (!opened) {
          opened = true
          void shell.openExternal(progress.url)
        }
      }
      send(progress)
    })
    if (launched) send({ phase: 'launched' })
    return launched
  })

  ipcMain.handle(HOME_CHANNELS.accountLoginOpenUrl, () => {
    if (pendingLoginUrl) void shell.openExternal(pendingLoginUrl)
  })

  ipcMain.handle(HOME_CHANNELS.accountLogout, async () => {
    await genofficeLogout()
    // the cloud projects cache belongs to the account that just signed out
    clearCloudProjectsStore(cloudProjectsStorePath())
  })

  ipcMain.handle(HOME_CHANNELS.getAppVersion, (): string => app.getVersion())

  ipcMain.handle(HOME_CHANNELS.recents, (_event, query: unknown): RecentPage =>
    pageRecentPaths(readRecentFiles(), query, new Set(readStarredFiles())),
  )

  // Starred files sort by mtime, which requires stat-ing them all first; they are hand-picked and few, so this is fine
  ipcMain.handle(HOME_CHANNELS.starred, (_event, query: unknown): RecentPage => {
    const { offset, limit, ext } = normalizeRecentQuery(query)
    const all = statEntries(readStarredFiles()).sort((a, b) => b.mtimeMs - a.mtimeMs)
    const filtered = ext ? all.filter((entry) => entry.ext === ext) : all
    return {
      entries: limit === 0 ? [] : filtered.slice(offset, offset + limit),
      total: filtered.length,
      totalAll: all.length,
    }
  })

  ipcMain.handle(HOME_CHANNELS.statPaths, (_event, paths: unknown): RecentEntry[] =>
    statEntries(stringPaths(paths)),
  )

  ipcMain.handle(HOME_CHANNELS.toggleStar, (_event, path: unknown) => {
    if (typeof path === 'string') toggleStarredFile(path)
  })

  ipcMain.handle(HOME_CHANNELS.openPath, (_event, path: unknown) => {
    if (typeof path === 'string') openDocumentPath(path)
  })

  ipcMain.handle(HOME_CHANNELS.browse, async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender) ?? shellWindow
    if (!win) return
    const result = await showOpenDialogWithMemory(dialog, win, {
      title: tm('dlgOpenTitle'),
      filters: [
        { name: tm('filterSupported'), extensions: OPEN_DIALOG_EXTENSIONS },
        { name: tm('filterWord'), extensions: ['docx', 'doc'] },
        { name: tm('filterExcel'), extensions: ['xlsx', 'xls', 'csv'] },
        { name: tm('filterPpt'), extensions: ['pptx', 'ppt'] },
        { name: tm('filterPdf'), extensions: ['pdf'] },
        { name: tm('filterMarkdown'), extensions: ['md', 'markdown'] },
      ],
      properties: ['openFile'],
    })
    if (!result.canceled && result.filePaths[0]) openDocumentPath(result.filePaths[0])
  })

  ipcMain.handle(HOME_CHANNELS.newDoc, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('doc', opts.projectId)
    }
    newDocTab()
  })

  ipcMain.handle(HOME_CHANNELS.newSheet, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('sheet', opts.projectId)
    }
    void newSheetTab()
  })

  ipcMain.handle(HOME_CHANNELS.newSlide, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('slide', opts.projectId)
    }
    newSlideTab()
  })

  ipcMain.handle(HOME_CHANNELS.newMarkdown, (_event, opts?: { projectId?: string }) => {
    if (opts?.projectId && opts.projectId !== 'default') {
      pendingNewFileProject.set('markdown', opts.projectId)
    }
    newMarkdownTab()
  })

  ipcMain.handle(HOME_CHANNELS.removeRecent, (_event, paths: unknown) => {
    removeRecentFiles(stringPaths(paths))
  })

  ipcMain.handle(HOME_CHANNELS.revealPath, (_event, path: unknown) => {
    if (typeof path === 'string' && existsSync(path)) shell.showItemInFolder(path)
  })

  ipcMain.handle(
    HOME_CHANNELS.renameFile,
    (_event, path: unknown, newName: unknown): RenameResult => {
      if (typeof path !== 'string' || typeof newName !== 'string')
        return { ok: false, error: tm('errBadArgs') }
      const name = newName.trim()
      if (!name || /[\\/:]/.test(name)) return { ok: false, error: tm('errBadName') }
      if (!existsSync(path)) return { ok: false, error: tm('errMissing') }
      const target = join(dirname(path), name)
      if (target === path) return { ok: true, path }
      if (existsSync(target)) return { ok: false, error: tm('errExists') }
      try {
        renameSync(path, target)
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : tm('errRenameFailed') }
      }
      replaceRecentFile(path, target)
      // project-store's fileMap/chatIdByPath re-key too, so AI chat history follows the file
      projectFileRenamed(path, target)
      // the slides module's own recent list switches to the new path as well (used by the start screen)
      if (/\.pptx$/i.test(target)) void replaceSlidesRecentFile(path, target)
      // open tabs sync their title/path; each editor then syncs its internal save path and title bar
      const affected = tabManager?.renameTabFile(path, target) ?? []
      for (const t of affected) {
        if (t.kind === 'slides') slidesFileRenamed(t.webContents, path, target)
        else if (t.kind === 'docs') docsFileRenamed(t.webContents, path, target)
        else if (t.kind === 'sheets') sheetsFileRenamed(t.webContents, path, target)
        else if (t.kind === 'markdown') markdownFileRenamed(t.webContents, path, target)
      }
      return { ok: true, path: target }
    },
  )

  ipcMain.handle(HOME_CHANNELS.duplicateFile, (_event, path: unknown) => {
    if (typeof path !== 'string' || !existsSync(path)) return
    const ext = extname(path)
    const base = basename(path, ext)
    const dir = dirname(path)
    for (let i = 1; ; i++) {
      const target = join(dir, `${base} ${tm('copySuffix')}${i === 1 ? '' : ` ${i}`}${ext}`)
      if (existsSync(target)) continue
      copyFileSync(path, target)
      recordRecentFile(target)
      return
    }
  })

  ipcMain.handle(HOME_CHANNELS.deleteFiles, async (_event, paths: unknown) => {
    const list = stringPaths(paths)
    for (const p of list) {
      try {
        await shell.trashItem(p)
      } catch {
        // file already gone or trash unavailable; still drop it from the list
      }
    }
    removeRecentFiles(list)
  })

  ipcMain.handle(HOME_CHANNELS.openTrash, () => {
    if (process.platform === 'darwin') {
      void shell.openPath(join(app.getPath('home'), '.Trash'))
    } else if (process.platform === 'win32') {
      spawn('explorer.exe', ['shell:RecycleBin'], { detached: true }).unref()
    } else {
      void shell.openPath(join(app.getPath('home'), '.local', 'share', 'Trash', 'files'))
    }
  })

  ipcMain.handle(HOME_CHANNELS.getLanguage, (): Lang => currentLang())

  ipcMain.handle(HOME_CHANNELS.setLanguage, (_event, lang: unknown) => {
    if (!isLang(lang) || lang === currentLang()) return
    persistLang(lang)
    // the switcher lives on the home page, so the home menu is the active one
    buildHomeMenu()
    installDockMenu()
    installBackToHomeItems()
    for (const wc of webContents.getAllWebContents()) wc.send('app:language-changed', lang)
  })

  ipcMain.handle(HOME_CHANNELS.getUpdateChannel, (): UpdateChannel => currentUpdateChannel())

  ipcMain.handle(HOME_CHANNELS.setUpdateChannel, (_event, channel: unknown) => {
    if (!isUpdateChannel(channel) || channel === currentUpdateChannel()) return
    cachedUpdateChannel = channel
    writeAppSetting(APP_SETTINGS_PATH(), 'updateChannel', channel)
    applyUpdateChannel(channel)
  })

  ipcMain.handle(
    HOME_CHANNELS.onboardingSeen,
    (): boolean => readAppSettings(APP_SETTINGS_PATH()).onboardingSeen === true,
  )

  ipcMain.handle(HOME_CHANNELS.setOnboardingSeen, () => {
    writeAppSetting(APP_SETTINGS_PATH(), 'onboardingSeen', true)
  })

  ipcMain.handle(HOME_CHANNELS.getTheme, (): UiTheme => currentTheme())
  // editor tabs ask via the app-wide channel (symmetric with app:get-language)
  ipcMain.handle('app:get-theme', (): UiTheme => currentTheme())

  ipcMain.handle(HOME_CHANNELS.setTheme, (_event, theme: unknown) => {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'system') return
    if (theme === currentTheme()) return
    cachedTheme = theme
    writeAppSetting(APP_SETTINGS_PATH(), 'theme', theme)
    nativeTheme.themeSource = theme
    for (const wc of webContents.getAllWebContents()) wc.send('app:theme-changed', theme)
  })

  // effective folder where new/untitled files land; the editor mains resolve
  // the same setting themselves (configuredDefaultSaveDir via docs' defaultSaveDir)
  ipcMain.handle(HOME_CHANNELS.getDefaultSaveDir, (): string => defaultSaveDir())

  ipcMain.handle(HOME_CHANNELS.pickDefaultSaveDir, async (): Promise<string | null> => {
    const result = await showOpenDialogWithMemory(dialog, shellWindow, {
      title: tm('dlgPickSaveDir'),
      defaultPath: defaultSaveDir(),
      properties: ['openDirectory', 'createDirectory'],
    })
    const picked = result.filePaths[0]
    if (result.canceled || !picked) return null
    if (!isUsableSaveDir(picked)) {
      showErrorDialog(shellWindow, tm('errSaveDirUnusable'), picked)
      return null
    }
    writeAppSetting(APP_SETTINGS_PATH(), DEFAULT_SAVE_DIR_KEY, picked)
    return picked
  })

  ipcMain.handle(HOME_CHANNELS.openGenTeam, () => {
    shell.openExternal(GENTEAM_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  ipcMain.handle(HOME_CHANNELS.openCreditUsage, () => {
    shell.openExternal(CREDIT_USAGE_URL).catch(() => {
      // no browser handler available; nothing actionable for the user here
    })
  })

  const cloudProjectsStorePath = () => join(app.getPath('userData'), 'cloud-projects.json')

  ipcMain.handle(HOME_CHANNELS.cloudProjectsCached, () =>
    readCloudProjectsStore(cloudProjectsStorePath()),
  )

  ipcMain.handle(HOME_CHANNELS.cloudProjects, () => syncCloudProjects(cloudProjectsStorePath()))

  ipcMain.handle(HOME_CHANNELS.openCloudProject, (_event, projectUrl: unknown) => {
    const url = cloudProjectExternalUrl(projectUrl)
    if (url) void shell.openExternal(url)
  })
}

function stringPaths(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((p): p is string => typeof p === 'string') : []
}

// electron-vite emits ?asset files under hashed names, which breaks nativeImage's
// automatic `@2x` sibling lookup — attach the retina representation by hand
function loadMenuIcon(path1x: string, path2x: string): NativeImage {
  const icon = nativeImage.createFromPath(path1x)
  icon.addRepresentation({ scaleFactor: 2, buffer: readFileSync(path2x) })
  return icon
}

// loaded once, not on every menu open
interface MenuIconSet {
  docx: NativeImage
  xlsx: NativeImage
  pptx: NativeImage
  pdf: NativeImage
  md: NativeImage
  home: NativeImage
}
let menuIconCache: MenuIconSet | null = null
function menuIcons(): MenuIconSet {
  menuIconCache ??= {
    docx: loadMenuIcon(menuDocxIcon1x, menuDocxIcon2x),
    xlsx: loadMenuIcon(menuXlsxIcon1x, menuXlsxIcon2x),
    pptx: loadMenuIcon(menuPptxIcon1x, menuPptxIcon2x),
    pdf: loadMenuIcon(menuPdfIcon1x, menuPdfIcon2x),
    md: loadMenuIcon(menuMdIcon1x, menuMdIcon2x),
    home: loadMenuIcon(menuHomeIcon1x, menuHomeIcon2x),
  }
  return menuIconCache
}

const TAB_MENU_ICON: Record<TabKind, keyof MenuIconSet> = {
  home: 'home',
  docs: 'docx',
  sheets: 'xlsx',
  slides: 'pptx',
  pdf: 'pdf',
  markdown: 'md',
}

// tab views see neither DOM events nor a focus change when the user clicks the
// shell chrome — relay the press so open popovers in documents can dismiss
function broadcastChromePressed(): void {
  for (const wc of webContents.getAllWebContents()) wc.send('app:chrome-pressed')
}

function registerTabsIpc(): void {
  ipcMain.on(TABS_CHANNELS.chromePressed, broadcastChromePressed)
  ipcMain.handle(TABS_CHANNELS.list, () => tabManager?.list() ?? [])
  ipcMain.handle(TABS_CHANNELS.activate, (_event, id: string) => tabManager?.activateTab(id))
  ipcMain.handle(TABS_CHANNELS.close, (_event, id: string) => tabManager?.closeTab(id))
  ipcMain.handle(TABS_CHANNELS.reorder, (_event, id: string, toIndex: number) => {
    if (typeof id === 'string' && Number.isInteger(toIndex)) tabManager?.reorderTab(id, toIndex)
  })
  // "all tabs" overflow menu — native popup because the editors' WebContentsView
  // would cover any DOM dropdown the shell renderer draws below the tab strip
  ipcMain.handle(TABS_CHANNELS.showMenu, (_event, x: unknown, y: unknown) => {
    if (!tabManager || !shellWindow) return
    const menu = Menu.buildFromTemplate(
      tabManager.list().map((tab) => ({
        label: tab.title,
        type: 'checkbox' as const,
        checked: tab.active,
        icon: menuIcons()[TAB_MENU_ICON[tab.kind]],
        click: () => tabManager?.activateTab(tab.id),
      })),
    )
    menu.popup({
      window: shellWindow,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
  // "+" new-file menu — native for the same reason as the tab list above
  ipcMain.handle(TABS_CHANNELS.showNewMenu, (_event, x: unknown, y: unknown) => {
    if (!tabManager || !shellWindow) return
    const menu = Menu.buildFromTemplate([
      // enabled:false so pre-Sonoma macOS / Windows (no 'header' support) degrade
      // to an inert label instead of a clickable no-op item
      { label: tm('menuSectionNew'), type: 'header', enabled: false },
      {
        label: tm('menuNewDoc'),
        icon: menuIcons().docx,
        click: () => newDocTab(),
      },
      {
        label: tm('menuNewSheet'),
        icon: menuIcons().xlsx,
        click: () => void newSheetTab(),
      },
      {
        label: tm('menuNewSlide'),
        icon: menuIcons().pptx,
        click: () => newSlideTab(),
      },
      {
        label: tm('menuNewMarkdown'),
        icon: menuIcons().md,
        click: () => newMarkdownTab(),
      },
      { type: 'separator' },
      { label: tm('menuOpen'), click: () => void openFileViaDialog() },
    ])
    menu.popup({
      window: shellWindow,
      ...(typeof x === 'number' && typeof y === 'number'
        ? { x: Math.round(x), y: Math.round(y) }
        : {}),
    })
  })
}

// ---- home menu ----

async function openFileViaDialog(): Promise<void> {
  const win = shellWindow ?? BrowserWindow.getFocusedWindow()
  if (!win) return
  const result = await showOpenDialogWithMemory(dialog, win, {
    filters: [{ name: tm('filterSupported'), extensions: OPEN_DIALOG_EXTENSIONS }],
    properties: ['openFile'],
  })
  if (!result.canceled && result.filePaths[0]) openDocumentPath(result.filePaths[0])
}

function buildHomeMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        { label: tm('menuSectionNew'), type: 'header', enabled: false },
        {
          label: tm('menuNewDoc'),
          accelerator: 'CmdOrCtrl+N',
          click: () => newDocTab(),
        },
        {
          label: tm('menuNewSheet'),
          click: () => void newSheetTab(),
        },
        { label: tm('menuNewSlide'), click: () => newSlideTab() },
        { label: tm('menuNewMarkdown'), click: () => newMarkdownTab() },
        { type: 'separator' },
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        { role: 'close', label: tm('menuClose') },
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    windowMenuTemplate(process.platform, appMenuLabels(currentLang())),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [{ label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() }],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- pdf menu (pdf-main has no menu of its own; the shell owns pdf tabs, so it builds one) ----

function buildPdfMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        {
          label: tm('backToHome'),
          accelerator: 'Shift+CmdOrCtrl+H',
          click: () => tabManager?.openHomeTab(),
        },
        { type: 'separator' },
        {
          label: tm('menuSave'),
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const tab = tabManager?.activePdfTab()
            if (tab) void flushPdfSave(tab.webContents)
          },
        },
        {
          label: tm('menuSaveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => void savePdfAs(),
        },
        { type: 'separator' },
        {
          label: tm('menuExportDocx'),
          click: () => void exportPdfAsDocx(),
        },
        { type: 'separator' },
        {
          label: tm('menuClose'),
          accelerator: 'CmdOrCtrl+W',
          click: () => tabManager?.closeActiveTab(),
        },
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    windowMenuTemplate(process.platform, appMenuLabels(currentLang())),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [{ label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() }],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- markdown menu (markdown-main has no menu of its own; the shell owns markdown tabs) ----

function buildMarkdownMenu(): void {
  const isMac = process.platform === 'darwin'
  const template: MenuItemConstructorOptions[] = [
    ...(isMac ? [{ role: 'appMenu' as const }] : []),
    {
      label: tm('menuFile'),
      submenu: [
        {
          label: tm('menuOpen'),
          accelerator: 'CmdOrCtrl+O',
          click: () => void openFileViaDialog(),
        },
        { type: 'separator' },
        {
          label: tm('backToHome'),
          accelerator: 'Shift+CmdOrCtrl+H',
          click: () => tabManager?.openHomeTab(),
        },
        { type: 'separator' },
        {
          label: tm('menuSave'),
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) void requestMarkdownSave(tab.webContents, 'save')
          },
        },
        {
          label: tm('menuSaveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) void requestMarkdownSave(tab.webContents, 'saveAs')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuExportDocx'),
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'docx')
          },
        },
        {
          label: tm('menuExportPdf'),
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'pdf')
          },
        },
        {
          label: tm('menuOpenInDocs'),
          click: () => {
            const tab = tabManager?.activeMarkdownTab()
            if (tab) sendMarkdownExportRequest(tab.webContents, 'docs')
          },
        },
        { type: 'separator' },
        {
          label: tm('menuClose'),
          accelerator: 'CmdOrCtrl+W',
          click: () => tabManager?.closeActiveTab(),
        },
      ],
    },
    editMenuTemplate(process.platform, appMenuLabels(currentLang())),
    windowMenuTemplate(process.platform, appMenuLabels(currentLang())),
    {
      role: 'help',
      label: tm('menuHelp'),
      submenu: [{ label: tm('thirdPartyNotices'), click: () => void openThirdPartyNotices() }],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

/**
 * Save As for pdf tabs: write pending edits to the picked path only, then open the copy.
 * Non-destructive: the original file is never written, and a cancelled dialog changes
 * nothing on disk (dialog first, no flush into the source).
 */
/** In-flight guard (same pattern as exportPdfAsDocx): a re-trigger while the dialog
    or write is active must not start a second flow that overwrites the first one's
    waiter/target grant or clears its autosave pause early */
let savingPdfAs = false

async function savePdfAs(): Promise<void> {
  const tab = tabManager?.activePdfTab()
  if (!tab?.filePath || !shellWindow || savingPdfAs) return
  savingPdfAs = true
  // Pause renderer autosave for the whole flow: the dialog blurs the window, and a
  // blur-triggered autosave would write the pending edits into the original file
  setPdfSaveAsInFlight(tab.webContents, true)
  try {
    const picked = await showSaveDialogWithMemory(dialog, shellWindow, {
      defaultPath: tab.filePath,
      filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
    })
    if (picked.canceled || !picked.filePath || picked.filePath === tab.filePath) return
    if (pdfIsDirty(tab.webContents.id)) {
      // Renderer applies its pending edits onto the source bytes; the pdf main
      // process writes the result to the picked path only
      if (!(await requestPdfSaveAs(tab.webContents, picked.filePath))) return
    } else {
      // No pending edits → a byte-identical copy
      copyFileSync(tab.filePath, picked.filePath)
    }
    openDocumentPath(picked.filePath)
  } finally {
    savingPdfAs = false
    setPdfSaveAsInFlight(tab.webContents, false)
  }
}

/**
 * In-flight guard: covers the whole flow (dialogs included, conversion takes
 * ~10s+) so re-triggering from the menu can never start a second paid conversion
 */
let exportingPdfDocx = false

/**
 * Export as Word for pdf tabs: flush pending edits, confirm the 5-credit cost,
 * pick the destination, then upload + cloud-convert via gsk file_convert. Not
 * logged in → offer browser login and let the user re-trigger the export
 * afterwards. The destination is picked before converting so cancelling the
 * save dialog never wastes a paid conversion.
 */
async function exportPdfAsDocx(): Promise<void> {
  const tab = tabManager?.activePdfTab()
  if (!tab?.filePath || !shellWindow) return
  if (exportingPdfDocx) {
    // Re-triggered while a previous export (dialogs or cloud conversion) is
    // still in flight: tell the user instead of silently ignoring the click.
    void dialog.showMessageBox(shellWindow, {
      type: 'info',
      message: tm('pdfDocxBusyMsg'),
    })
    return
  }
  exportingPdfDocx = true
  try {
    if (!(await flushPdfSave(tab.webContents))) return
    if (!hasGskAuth()) {
      // hasGskAuth() is also false when the gsk CLI itself cannot be resolved
      // (broken install); Sign In could not launch in that case, so surface
      // the real problem instead of a login dialog that cannot succeed.
      if (!resolveGskEntry()) {
        void dialog.showMessageBox(shellWindow, {
          type: 'error',
          message: tm('pdfDocxNoCliMsg'),
        })
        return
      }
      const { response } = await dialog.showMessageBox(shellWindow, {
        type: 'info',
        message: tm('pdfDocxLoginMsg'),
        detail: tm('pdfDocxLoginDetail'),
        buttons: [tm('pdfDocxBtnLogin'), tm('btnCancel')],
        defaultId: 0,
        cancelId: 1,
        noLink: true,
      })
      if (response === 0) ensureGenofficeLogin((url) => void shell.openExternal(url))
      return
    }
    const balance = (await gskLoginInfo())?.creditBalance
    const balanceLine =
      balance === undefined
        ? ''
        : ` ${tm('pdfDocxConfirmBalance', { balance: Math.floor(balance).toLocaleString('en-US') })}`
    const confirm = await dialog.showMessageBox(shellWindow, {
      type: 'question',
      message: tm('pdfDocxConfirmMsg'),
      detail: `${tm('pdfDocxConfirmDetail')}${balanceLine}`,
      buttons: [tm('pdfDocxBtnConvert'), tm('btnCancel')],
      defaultId: 0,
      cancelId: 1,
      noLink: true,
    })
    if (confirm.response !== 0) return
    const picked = await showSaveDialogWithMemory(dialog, shellWindow, {
      defaultPath: tab.filePath.replace(/\.pdf$/i, '.docx'),
      filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
    })
    if (picked.canceled || !picked.filePath) return
    // If the destination is already open in a docs tab, close it first (its
    // normal unsaved-changes guard applies) so the converted file opens fresh
    // instead of leaving a stale tab whose next save would clobber the result.
    // Cancelling the close aborts the export before any credits are spent.
    const staleTabId = tabManager?.findDocsTabByPath(picked.filePath)
    if (staleTabId) {
      await tabManager?.closeTab(staleTabId)
      // closeTab activates the docs tab for its unsaved-changes prompt (and a
      // fallback tab after a successful close), so bring the pdf tab back
      // either way — especially when the user cancels and the export aborts.
      tabManager?.activateTab(tab.id)
      if (tabManager?.findDocsTabByPath(picked.filePath)) return
    }
    shellWindow.setProgressBar(2)
    const bytes = await gskConvertPdfToDocx(tab.filePath)
    writeFileSync(picked.filePath, bytes)
    openDocumentPath(picked.filePath)
  } catch (err) {
    if (shellWindow && !shellWindow.isDestroyed()) {
      void dialog.showMessageBox(shellWindow, {
        type: 'error',
        message: tm('pdfDocxFailedMsg'),
        detail: err instanceof Error ? err.message : String(err),
      })
    }
  } finally {
    exportingPdfDocx = false
    if (shellWindow && !shellWindow.isDestroyed()) shellWindow.setProgressBar(-1)
  }
}

function openThirdPartyNotices(): Promise<string> {
  const path = app.isPackaged
    ? join(process.resourcesPath, 'THIRD-PARTY-NOTICES.txt')
    : join(app.getAppPath(), 'build', 'THIRD-PARTY-NOTICES.txt')
  return shell.openPath(path)
}

/** every module's File menu gets a way back to the launcher */
function installBackToHomeItems(): void {
  const backToHomeItem: MenuItemConstructorOptions = {
    label: tm('backToHome'),
    accelerator: 'Shift+CmdOrCtrl+H',
    click: () => tabManager?.openHomeTab(),
  }
  setDocsExtraFileMenuItems([backToHomeItem])
  setSheetsExtraFileMenuItems([backToHomeItem])
  setSlidesExtraFileMenuItems([backToHomeItem])
}

function installDockMenu(): void {
  if (process.platform !== 'darwin') return
  app.dock?.setMenu(
    Menu.buildFromTemplate([
      { label: tm('menuHome'), click: () => tabManager?.openHomeTab() },
      {
        label: tm('menuNewDoc'),
        click: () => newDocTab(),
      },
      {
        label: tm('menuNewSheet'),
        click: () => void newSheetTab(),
      },
      { label: tm('menuNewSlide'), click: () => newSlideTab() },
      { label: tm('menuNewMarkdown'), click: () => newMarkdownTab() },
    ]),
  )
}

// On mainland-China networks the main process's Node fetch (undici) bypasses the system proxy,
// so direct calls to overseas LLM/image-search APIs time out or get region-blocked (403).
// Prefer proxy env vars (terminal launch); a packaged app launched from Finder inherits no shell
// env vars, so fall back to the system HTTP proxy. The renderer uses Chromium's system proxy and
// is unaffected. Same bootstrap as slides-main startSlidesStandalone.
// awaited by login IPC so the first status probe / login click cannot race the proxy resolution
let proxyBootstrap: Promise<void> = Promise.resolve()

async function installMainProcessProxy(): Promise<void> {
  let proxyUrl = [
    process.env.HTTPS_PROXY,
    process.env.https_proxy,
    process.env.HTTP_PROXY,
    process.env.http_proxy,
    process.env.ALL_PROXY,
    process.env.all_proxy,
  ].find((v) => v && /^https?:\/\//.test(v))
  if (!proxyUrl) {
    try {
      // PAC/rule proxies answer per-host: probe the host the login flow, the
      // Genspark LLM proxy and the gsk CLI actually target
      const resolved = await session.defaultSession.resolveProxy('https://www.genspark.ai/')
      const m = /PROXY\s+([^;\s]+)/.exec(resolved)
      if (m) proxyUrl = `http://${m[1]}`
    } catch {
      /* no system proxy */
    }
  }
  if (!proxyUrl) return
  // spawned gsk CLI children (login/search/…) do their own fetch and never see
  // the dispatcher below — forward the proxy to them via env
  setGskProxyUrl(proxyUrl)
  try {
    const { ProxyAgent, setGlobalDispatcher } = await import('undici')
    setGlobalDispatcher(new ProxyAgent(proxyUrl))
    // strip user:pass credentials before logging
    console.log('[proxy] main-process fetch via', proxyUrl.replace(/\/\/[^@/]*@/, '//***@'))
  } catch (e) {
    console.warn('[proxy] failed to set ProxyAgent:', e)
  }
}

// ---- lifecycle (the shell is the only owner) ----

let pendingLaunchPath = supportedFileIn(process.argv) ?? unsupportedFileIn(process.argv)

// show() does not un-minimize, and on macOS ⌘W destroys the shell window while the
// app keeps running — either way a file opened from Finder would land out of sight.
function revealShellWindow(): void {
  if (!shellWindow) createShellWindow()
  if (shellWindow?.isMinimized()) shellWindow.restore()
  shellWindow?.show()
  shellWindow?.focus()
}

// On macOS a file opened from Finder is not in argv; it arrives via the open-file event (before ready).
// If another instance already holds the lock, this process exits, and the path must ride along in
// the lock request's additionalData to the surviving instance — so the lock request is deferred
// until ready, after the path is known.
app.on('open-file', (event, filePath) => {
  event.preventDefault()
  if (!app.isReady()) {
    pendingLaunchPath = filePath
    return
  }
  revealShellWindow()
  if (!openDocumentPath(filePath)) tabManager?.openHomeTab()
})

app.on('second-instance', (_event, argv, _cwd, additionalData) => {
  const file =
    supportedFileIn(argv) ??
    unsupportedFileIn(argv) ??
    (additionalData as { launchPath?: string } | null)?.launchPath
  revealShellWindow()
  if (!file || !openDocumentPath(file)) tabManager?.openHomeTab()
})

installNavigationGuard(app)
installContextMenu(app, () => contextMenuLabels(currentLang()))
registerAiIpc()
registerProjectIpc()
registerDocsIpc()
registerHomeIpc()
registerTabsIpc()

// sheets' project:resolveChat goes through the handler registered by docs-main; the sessionId reverse lookup hooks in here
setSessionPathResolver(resolveSheetsSessionPath)

/** Dev-only pid marker for the takeover below; scoped to userData like the lock itself. */
const devPidFile = () => join(app.getPath('userData'), 'dev-instance.pid')

app.whenReady().then(async () => {
  const lockData = () => (pendingLaunchPath ? { launchPath: pendingLaunchPath } : {})
  let hasLock = app.requestSingleInstanceLock(lockData())
  if (!hasLock && !app.isPackaged) {
    // Dev watch restart: electron-vite SIGTERMs the previous instance and spawns this
    // one immediately. Chromium turns that SIGTERM into a graceful quit (Node's
    // process.on('SIGTERM') never fires in the main process), and the quit can wedge
    // in the close-confirmation flow — the zombie then keeps the single-instance lock,
    // this instance quits, and electron-vite's on-close handler exits with it, killing
    // the renderer dev server (blank shell window until a manual dev restart).
    // The previous instance is doomed either way: kill it and take over the lock.
    try {
      const oldPid = Number(readFileSync(devPidFile(), 'utf-8').trim())
      if (Number.isFinite(oldPid) && oldPid > 0 && oldPid !== process.pid) {
        // pid-recycling guard: only kill if that pid is still an Electron process
        const cmd = execSync(`ps -o command= -p ${oldPid}`).toString()
        if (cmd.includes('Electron')) process.kill(oldPid, 'SIGKILL')
      }
    } catch {
      // no previous instance recorded / already gone (ps exits non-zero)
    }
    for (let i = 0; i < 20 && !hasLock; i++) {
      await new Promise((r) => setTimeout(r, 150))
      hasLock = app.requestSingleInstanceLock(lockData())
    }
  }
  if (!hasLock) {
    app.quit()
    return
  }
  if (!app.isPackaged) {
    try {
      writeFileSync(devPidFile(), String(process.pid))
    } catch {
      // best-effort: without the marker the next restart just retries the lock
    }
  }

  proxyBootstrap = installMainProcessProxy()
  app.setAccessibilitySupportEnabled(true)
  // Settle the shared uiLang from saved settings BEFORE any tab renderer can
  // ask 'app:get-language': the editor handlers return the i18n module's
  // mutable lang, whose 'zh' default otherwise wins the race for whichever
  // tab loads first (e.g. sheets booting in Chinese while docs shows English).
  currentLang()
  // native menus/dialogs/scrollbars follow the persisted theme from first paint
  nativeTheme.themeSource = currentTheme()
  startSheetsCaptureServer()
  createShellWindow()
  // deferred to ready: labels need currentLang(), which reads app.getLocale()
  installBackToHomeItems()
  installDockMenu()
  initAutoUpdater(() => shellWindow, currentUpdateChannel())

  if (!pendingLaunchPath || !openDocumentPath(pendingLaunchPath)) tabManager?.openHomeTab()
  pendingLaunchPath = null

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createShellWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', () => {
  // No close prompt may fall through to "Save" during shutdown
  markSheetsShuttingDown()
  stopSheetsSidecar()
})
