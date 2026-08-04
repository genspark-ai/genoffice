import { createHash } from 'node:crypto'
import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { copyFile, mkdir, readFile, readdir, stat, unlink } from 'node:fs/promises'
import { basename, join } from 'node:path'
import { BrowserWindow, Menu, WebContentsView, app, dialog, ipcMain, shell } from 'electron'
import {
  appMenuLabels,
  contextMenuLabels,
  fetchWithSsrfGuard,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  windowMenuTemplate,
} from '@genoffice/electron-utils'
import { createI18n, getUiLang, normalizeLang, setUiLang } from '@genoffice/i18n'
import { ProjectStore } from '@genoffice/project-store'
import type {
  IpcMainInvokeEvent,
  MenuItemConstructorOptions,
  OpenDialogOptions,
  SaveDialogOptions,
  WebContents,
} from 'electron'
import { parseFileToText } from '@genoffice/file-parse'
import {
  AiCreditsError,
  AiTimeoutError,
  chatForProvider,
  defaultAiSettings,
  resolveAiSettings,
  streamForProvider,
  type AiChatRequest,
  type AiSettings,
  type AiStreamChunk,
  type AiStreamRequest,
  type GenSparkAccountStatus,
  type LegacyAiSettings,
} from '@genoffice/ai-provider'
import {
  gskApiKey,
  gskLogin,
  gskLoginInfo,
  hasGskAuth,
  webSearch,
  imageSearch,
} from '@genoffice/ai-search'
import type {
  AttachmentAddResult,
  AttachmentImageResult,
  AttachmentMeta,
  AttachmentReadResult,
  DocsTabInfo,
  MenuCommand,
  OpenFileResult,
} from '../shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../shared/ipc'
import { findDocxPath } from '../shared/open-file'
import { atomicWriteFile, looksLikeZip } from './atomic-write'
import { isExternallyModified, type DiskFileState } from './external-change'
import { initDocsAutoUpdater } from './updater'

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
    dlgSaveAs: '另存为',
    closeUnsavedMsg: '此文档有未保存的更改。',
    closeUnsavedDetail: '关闭前是否保存？',
    closeNoReplyMsg: '文档没有响应,可能有未保存的更改。',
    closeNoReplyDetail: '仍要关闭吗?未保存的更改将丢失。',
    btnCloseAnyway: '仍要关闭',
    autosaveFoundTitle: '发现自动恢复版本',
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
    dlgAddAttachment: '添加附件',
    filterSupported: '支持的文件',
    filterAll: '所有文件',
    dlgExportPdf: '导出为 PDF',
    errUnsupportedExt: '暂不支持 .{ext} 类型',
    errNotFile: '不是文件',
    errTooLarge: '超过 {mb}MB 上限',
    errImageTooLarge: '图片超过 5MB 上限',
    errUnreadable: '无法读取',
    errFileTooLarge: '文件超过大小上限',
    errParseFailed: '文件解析失败',
    errImageNoText: '图片附件不提供文本,已作为图像随用户消息发送,直接看图即可',
    errNotImage: '不是支持的图片类型',
    errGskNotLoggedIn: '未登录 Genspark:请点击下方「登录 Genspark」完成登录后重试',
    errNoApiKey: '未配置 {provider} 的 API Key',
    errNoModel: '未配置模型名称',
    menuFile: '文件',
    menuNewDoc: '新建文档',
    menuNewWindow: '新建窗口',
    menuOpen: '打开…',
    menuOpenRecent: '打开最近的文档',
    menuNoRecent: '无最近文档',
    menuClose: '关闭',
    menuSave: '保存',
    menuSaveAs: '另存为…',
    menuPageSetup: '页面设置…',
    menuExportPdf: '导出为 PDF…',
    menuPrint: '打印…',
    menuEdit: '编辑',
    menuUndo: '撤销',
    menuRedo: '重做',
    menuCut: '剪切',
    menuCopy: '复制',
    menuPaste: '粘贴',
    menuPasteMatch: '粘贴并匹配格式',
    menuFindReplace: '查找和替换…',
    menuSelectAll: '全选',
    menuView: '视图',
    menuZoomIn: '放大',
    menuZoomOut: '缩小',
    menuZoom100: '实际大小 (100%)',
    menuPageWidth: '页宽',
    menuWholePage: '整页',
    menuAiSidebar: 'AI 侧栏',
    menuDarkMode: '深色模式',
    menuFullscreen: '进入全屏',
    menuInsert: '插入',
    menuInsertTable: '表格(3×3)',
    menuInsertImage: '图片…',
    menuInsertPageBreak: '分页符',
    menuInsertLink: '超链接…',
    menuInsertEquation: '公式…',
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
    menuTools: '工具',
    menuWordCount: '字数统计…',
    menuSpelling: '拼写和语法检查',
    menuMacros: '宏',
    menuWindow: '窗口',
    menuHelp: '帮助',
    menuDocsHelp: 'GenOffice Docs 帮助',
  },
  en: {
    dlgOpenDoc: 'Open Document',
    filterWord: 'Word Documents',
    dlgSaveAs: 'Save As',
    closeUnsavedMsg: 'This document has unsaved changes.',
    closeUnsavedDetail: 'Do you want to save them before closing?',
    closeNoReplyMsg: 'The document is not responding and may have unsaved changes.',
    closeNoReplyDetail: 'Close anyway? Unsaved changes will be lost.',
    btnCloseAnyway: 'Close Anyway',
    autosaveFoundTitle: 'Recovered version found',
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
    dlgAddAttachment: 'Add Attachments',
    filterSupported: 'Supported Files',
    filterAll: 'All Files',
    dlgExportPdf: 'Export as PDF',
    errUnsupportedExt: '.{ext} files are not supported',
    errNotFile: 'not a file',
    errTooLarge: 'exceeds the {mb}MB limit',
    errImageTooLarge: 'image exceeds the 5MB limit',
    errUnreadable: 'cannot be read',
    errFileTooLarge: 'File exceeds the size limit',
    errParseFailed: 'Failed to parse file',
    errImageNoText: 'Image attachments have no text; the image is sent along with the user message',
    errNotImage: 'not a supported image type',
    errGskNotLoggedIn:
      'Not signed in to Genspark: click “Sign in to Genspark” below, sign in, then retry',
    errNoApiKey: 'No API key configured for {provider}',
    errNoModel: 'No model name configured',
    menuFile: 'File',
    menuNewDoc: 'New Document',
    menuNewWindow: 'New Window',
    menuOpen: 'Open…',
    menuOpenRecent: 'Open Recent',
    menuNoRecent: 'No Recent Documents',
    menuClose: 'Close',
    menuSave: 'Save',
    menuSaveAs: 'Save As…',
    menuPageSetup: 'Page Setup…',
    menuExportPdf: 'Export as PDF…',
    menuPrint: 'Print…',
    menuEdit: 'Edit',
    menuUndo: 'Undo',
    menuRedo: 'Redo',
    menuCut: 'Cut',
    menuCopy: 'Copy',
    menuPaste: 'Paste',
    menuPasteMatch: 'Paste and Match Style',
    menuFindReplace: 'Find and Replace…',
    menuSelectAll: 'Select All',
    menuView: 'View',
    menuZoomIn: 'Zoom In',
    menuZoomOut: 'Zoom Out',
    menuZoom100: 'Actual Size (100%)',
    menuPageWidth: 'Page Width',
    menuWholePage: 'Whole Page',
    menuAiSidebar: 'AI Sidebar',
    menuDarkMode: 'Dark Mode',
    menuFullscreen: 'Enter Full Screen',
    menuInsert: 'Insert',
    menuInsertTable: 'Table (3×3)',
    menuInsertImage: 'Image…',
    menuInsertPageBreak: 'Page Break',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Equation…',
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
    menuTools: 'Tools',
    menuWordCount: 'Word Count…',
    menuSpelling: 'Spelling and Grammar',
    menuMacros: 'Macros',
    menuWindow: 'Window',
    menuHelp: 'Help',
    menuDocsHelp: 'GenOffice Docs Help',
  },
  ja: {
    dlgOpenDoc: '文書を開く',
    filterWord: 'Word 文書',
    dlgSaveAs: '名前を付けて保存',
    closeUnsavedMsg: 'このドキュメントに未保存の変更があります。',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    closeNoReplyMsg: 'ドキュメントが応答していません。未保存の変更がある可能性があります。',
    closeNoReplyDetail: 'それでも閉じますか?未保存の変更は失われます。',
    btnCloseAnyway: '閉じる',
    autosaveFoundTitle: '自動回復バージョンがあります',
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
    dlgAddAttachment: '添付ファイルの追加',
    filterSupported: 'サポートされているファイル',
    filterAll: 'すべてのファイル',
    dlgExportPdf: 'PDF としてエクスポート',
    errUnsupportedExt: '.{ext} 形式には対応していません',
    errNotFile: 'ファイルではありません',
    errTooLarge: '{mb}MB の上限を超えています',
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
    errNoModel: 'モデル名が設定されていません',
    menuFile: 'ファイル',
    menuNewDoc: '新規文書',
    menuNewWindow: '新規ウィンドウ',
    menuOpen: '開く…',
    menuOpenRecent: '最近使った文書を開く',
    menuNoRecent: '最近使った文書はありません',
    menuClose: '閉じる',
    menuSave: '保存',
    menuSaveAs: '名前を付けて保存…',
    menuPageSetup: 'ページ設定…',
    menuExportPdf: 'PDF としてエクスポート…',
    menuPrint: '印刷…',
    menuEdit: '編集',
    menuUndo: '元に戻す',
    menuRedo: 'やり直す',
    menuCut: '切り取り',
    menuCopy: 'コピー',
    menuPaste: '貼り付け',
    menuPasteMatch: '貼り付けて書式を合わせる',
    menuFindReplace: '検索と置換…',
    menuSelectAll: 'すべて選択',
    menuView: '表示',
    menuZoomIn: '拡大',
    menuZoomOut: '縮小',
    menuZoom100: '実際のサイズ (100%)',
    menuPageWidth: 'ページ幅',
    menuWholePage: 'ページ全体',
    menuAiSidebar: 'AI サイドバー',
    menuDarkMode: 'ダークモード',
    menuFullscreen: 'フルスクリーンにする',
    menuInsert: '挿入',
    menuInsertTable: '表 (3×3)',
    menuInsertImage: '画像…',
    menuInsertPageBreak: '改ページ',
    menuInsertLink: 'ハイパーリンク…',
    menuInsertEquation: '数式…',
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
    menuTools: 'ツール',
    menuWordCount: '文字カウント…',
    menuSpelling: 'スペルチェックと文章校正',
    menuMacros: 'マクロ',
    menuWindow: 'ウィンドウ',
    menuHelp: 'ヘルプ',
    menuDocsHelp: 'GenOffice Docs ヘルプ',
  },
  ko: {
    dlgOpenDoc: '문서 열기',
    filterWord: 'Word 문서',
    dlgSaveAs: '다른 이름으로 저장',
    closeUnsavedMsg: '이 문서에 저장하지 않은 변경 사항이 있습니다.',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    closeNoReplyMsg: '문서가 응답하지 않으며 저장하지 않은 변경 사항이 있을 수 있습니다.',
    closeNoReplyDetail: '그래도 닫을까요? 저장하지 않은 변경 사항은 사라집니다.',
    btnCloseAnyway: '닫기',
    autosaveFoundTitle: '자동 복구 버전 발견',
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
    dlgAddAttachment: '첨부 파일 추가',
    filterSupported: '지원되는 파일',
    filterAll: '모든 파일',
    dlgExportPdf: 'PDF로 내보내기',
    errUnsupportedExt: '.{ext} 형식은 지원되지 않습니다',
    errNotFile: '파일이 아닙니다',
    errTooLarge: '{mb}MB 제한을 초과했습니다',
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
    errNoModel: '모델 이름이 설정되지 않았습니다',
    menuFile: '파일',
    menuNewDoc: '새 문서',
    menuNewWindow: '새 창',
    menuOpen: '열기…',
    menuOpenRecent: '최근 문서 열기',
    menuNoRecent: '최근 문서 없음',
    menuClose: '닫기',
    menuSave: '저장',
    menuSaveAs: '다른 이름으로 저장…',
    menuPageSetup: '페이지 설정…',
    menuExportPdf: 'PDF로 내보내기…',
    menuPrint: '인쇄…',
    menuEdit: '편집',
    menuUndo: '실행 취소',
    menuRedo: '다시 실행',
    menuCut: '잘라내기',
    menuCopy: '복사',
    menuPaste: '붙여넣기',
    menuPasteMatch: '서식 맞춰 붙여넣기',
    menuFindReplace: '찾기 및 바꾸기…',
    menuSelectAll: '모두 선택',
    menuView: '보기',
    menuZoomIn: '확대',
    menuZoomOut: '축소',
    menuZoom100: '실제 크기(100%)',
    menuPageWidth: '페이지 너비',
    menuWholePage: '전체 페이지',
    menuAiSidebar: 'AI 사이드바',
    menuDarkMode: '다크 모드',
    menuFullscreen: '전체 화면 시작',
    menuInsert: '삽입',
    menuInsertTable: '표(3×3)',
    menuInsertImage: '그림…',
    menuInsertPageBreak: '페이지 나누기',
    menuInsertLink: '하이퍼링크…',
    menuInsertEquation: '수식…',
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
    menuTools: '도구',
    menuWordCount: '단어 개수…',
    menuSpelling: '맞춤법 및 문법 검사',
    menuMacros: '매크로',
    menuWindow: '창',
    menuHelp: '도움말',
    menuDocsHelp: 'GenOffice Docs 도움말',
  },
  fr: {
    dlgOpenDoc: 'Ouvrir un document',
    filterWord: 'Documents Word',
    dlgSaveAs: 'Enregistrer sous',
    closeUnsavedMsg: 'Ce document contient des modifications non enregistrées.',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    closeNoReplyMsg:
      'Le document ne répond pas et peut contenir des modifications non enregistrées.',
    closeNoReplyDetail: 'Fermer quand même ? Les modifications non enregistrées seront perdues.',
    btnCloseAnyway: 'Fermer quand même',
    autosaveFoundTitle: 'Version récupérée trouvée',
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
    dlgAddAttachment: 'Ajouter des pièces jointes',
    filterSupported: 'Fichiers pris en charge',
    filterAll: 'Tous les fichiers',
    dlgExportPdf: 'Exporter au format PDF',
    errUnsupportedExt: 'les fichiers .{ext} ne sont pas pris en charge',
    errNotFile: "n'est pas un fichier",
    errTooLarge: 'dépasse la limite de {mb} Mo',
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
    errNoModel: 'Aucun nom de modèle configuré',
    menuFile: 'Fichier',
    menuNewDoc: 'Nouveau document',
    menuNewWindow: 'Nouvelle fenêtre',
    menuOpen: 'Ouvrir…',
    menuOpenRecent: 'Ouvrir un document récent',
    menuNoRecent: 'Aucun document récent',
    menuClose: 'Fermer',
    menuSave: 'Enregistrer',
    menuSaveAs: 'Enregistrer sous…',
    menuPageSetup: 'Mise en page…',
    menuExportPdf: 'Exporter au format PDF…',
    menuPrint: 'Imprimer…',
    menuEdit: 'Édition',
    menuUndo: 'Annuler',
    menuRedo: 'Rétablir',
    menuCut: 'Couper',
    menuCopy: 'Copier',
    menuPaste: 'Coller',
    menuPasteMatch: 'Coller et adapter le style',
    menuFindReplace: 'Rechercher et remplacer…',
    menuSelectAll: 'Tout sélectionner',
    menuView: 'Affichage',
    menuZoomIn: 'Zoom avant',
    menuZoomOut: 'Zoom arrière',
    menuZoom100: 'Taille réelle (100 %)',
    menuPageWidth: 'Largeur de page',
    menuWholePage: 'Page entière',
    menuAiSidebar: 'Volet IA',
    menuDarkMode: 'Mode sombre',
    menuFullscreen: 'Activer le mode plein écran',
    menuInsert: 'Insertion',
    menuInsertTable: 'Tableau (3×3)',
    menuInsertImage: 'Image…',
    menuInsertPageBreak: 'Saut de page',
    menuInsertLink: 'Lien hypertexte…',
    menuInsertEquation: 'Équation…',
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
    menuTools: 'Outils',
    menuWordCount: 'Statistiques…',
    menuSpelling: 'Grammaire et orthographe',
    menuMacros: 'Macros',
    menuWindow: 'Fenêtre',
    menuHelp: 'Aide',
    menuDocsHelp: 'Aide GenOffice Docs',
  },
  de: {
    dlgOpenDoc: 'Dokument öffnen',
    filterWord: 'Word-Dokumente',
    dlgSaveAs: 'Speichern unter',
    closeUnsavedMsg: 'Dieses Dokument enthält ungespeicherte Änderungen.',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    closeNoReplyMsg:
      'Das Dokument reagiert nicht und enthält möglicherweise ungespeicherte Änderungen.',
    closeNoReplyDetail: 'Trotzdem schließen? Ungespeicherte Änderungen gehen verloren.',
    btnCloseAnyway: 'Trotzdem schließen',
    autosaveFoundTitle: 'Wiederhergestellte Version gefunden',
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
    dlgAddAttachment: 'Anlagen hinzufügen',
    filterSupported: 'Unterstützte Dateien',
    filterAll: 'Alle Dateien',
    dlgExportPdf: 'Als PDF exportieren',
    errUnsupportedExt: '.{ext}-Dateien werden nicht unterstützt',
    errNotFile: 'keine Datei',
    errTooLarge: 'überschreitet das Limit von {mb} MB',
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
    errNoModel: 'Kein Modellname konfiguriert',
    menuFile: 'Datei',
    menuNewDoc: 'Neues Dokument',
    menuNewWindow: 'Neues Fenster',
    menuOpen: 'Öffnen…',
    menuOpenRecent: 'Zuletzt verwendete Dokumente',
    menuNoRecent: 'Keine zuletzt verwendeten Dokumente',
    menuClose: 'Schließen',
    menuSave: 'Speichern',
    menuSaveAs: 'Speichern unter…',
    menuPageSetup: 'Seite einrichten…',
    menuExportPdf: 'Als PDF exportieren…',
    menuPrint: 'Drucken…',
    menuEdit: 'Bearbeiten',
    menuUndo: 'Rückgängig',
    menuRedo: 'Wiederholen',
    menuCut: 'Ausschneiden',
    menuCopy: 'Kopieren',
    menuPaste: 'Einfügen',
    menuPasteMatch: 'Einfügen und Stil anpassen',
    menuFindReplace: 'Suchen und Ersetzen…',
    menuSelectAll: 'Alles auswählen',
    menuView: 'Ansicht',
    menuZoomIn: 'Vergrößern',
    menuZoomOut: 'Verkleinern',
    menuZoom100: 'Originalgröße (100 %)',
    menuPageWidth: 'Seitenbreite',
    menuWholePage: 'Ganze Seite',
    menuAiSidebar: 'KI-Seitenleiste',
    menuDarkMode: 'Dunkelmodus',
    menuFullscreen: 'Vollbild ein',
    menuInsert: 'Einfügen',
    menuInsertTable: 'Tabelle (3×3)',
    menuInsertImage: 'Bild…',
    menuInsertPageBreak: 'Seitenumbruch',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Formel…',
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
    menuTools: 'Extras',
    menuWordCount: 'Wörter zählen…',
    menuSpelling: 'Rechtschreibung und Grammatik',
    menuMacros: 'Makros',
    menuWindow: 'Fenster',
    menuHelp: 'Hilfe',
    menuDocsHelp: 'GenOffice Docs-Hilfe',
  },
  es: {
    dlgOpenDoc: 'Abrir documento',
    filterWord: 'Documentos de Word',
    dlgSaveAs: 'Guardar como',
    closeUnsavedMsg: 'Este documento tiene cambios sin guardar.',
    closeUnsavedDetail: '¿Desea guardarlos antes de cerrar?',
    closeNoReplyMsg: 'El documento no responde y puede tener cambios sin guardar.',
    closeNoReplyDetail: '¿Cerrar de todos modos? Los cambios sin guardar se perderán.',
    btnCloseAnyway: 'Cerrar de todos modos',
    autosaveFoundTitle: 'Se encontró una versión recuperada',
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
    dlgAddAttachment: 'Agregar datos adjuntos',
    filterSupported: 'Archivos compatibles',
    filterAll: 'Todos los archivos',
    dlgExportPdf: 'Exportar como PDF',
    errUnsupportedExt: 'los archivos .{ext} no son compatibles',
    errNotFile: 'no es un archivo',
    errTooLarge: 'supera el límite de {mb} MB',
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
    errNoModel: 'No se ha configurado el nombre del modelo',
    menuFile: 'Archivo',
    menuNewDoc: 'Nuevo documento',
    menuNewWindow: 'Nueva ventana',
    menuOpen: 'Abrir…',
    menuOpenRecent: 'Abrir recientes',
    menuNoRecent: 'No hay documentos recientes',
    menuClose: 'Cerrar',
    menuSave: 'Guardar',
    menuSaveAs: 'Guardar como…',
    menuPageSetup: 'Configurar página…',
    menuExportPdf: 'Exportar como PDF…',
    menuPrint: 'Imprimir…',
    menuEdit: 'Edición',
    menuUndo: 'Deshacer',
    menuRedo: 'Rehacer',
    menuCut: 'Cortar',
    menuCopy: 'Copiar',
    menuPaste: 'Pegar',
    menuPasteMatch: 'Pegar con el mismo estilo',
    menuFindReplace: 'Buscar y reemplazar…',
    menuSelectAll: 'Seleccionar todo',
    menuView: 'Ver',
    menuZoomIn: 'Acercar',
    menuZoomOut: 'Alejar',
    menuZoom100: 'Tamaño real (100 %)',
    menuPageWidth: 'Ancho de página',
    menuWholePage: 'Página completa',
    menuAiSidebar: 'Barra lateral de IA',
    menuDarkMode: 'Modo oscuro',
    menuFullscreen: 'Usar pantalla completa',
    menuInsert: 'Insertar',
    menuInsertTable: 'Tabla (3×3)',
    menuInsertImage: 'Imagen…',
    menuInsertPageBreak: 'Salto de página',
    menuInsertLink: 'Hipervínculo…',
    menuInsertEquation: 'Ecuación…',
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
    menuTools: 'Herramientas',
    menuWordCount: 'Contar palabras…',
    menuSpelling: 'Ortografía y gramática',
    menuMacros: 'Macros',
    menuWindow: 'Ventana',
    menuHelp: 'Ayuda',
    menuDocsHelp: 'Ayuda de GenOffice Docs',
  },
  th: {
    dlgOpenDoc: 'เปิดเอกสาร',
    filterWord: 'เอกสาร Word',
    dlgSaveAs: 'บันทึกเป็น',
    closeUnsavedMsg: 'เอกสารนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    closeNoReplyMsg: 'เอกสารไม่ตอบสนองและอาจมีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeNoReplyDetail: 'ปิดต่อไปหรือไม่ การเปลี่ยนแปลงที่ยังไม่ได้บันทึกจะหายไป',
    btnCloseAnyway: 'ปิดต่อไป',
    autosaveFoundTitle: 'พบเวอร์ชันกู้คืนอัตโนมัติ',
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
    dlgAddAttachment: 'เพิ่มสิ่งที่แนบ',
    filterSupported: 'ไฟล์ที่รองรับ',
    filterAll: 'ไฟล์ทั้งหมด',
    dlgExportPdf: 'ส่งออกเป็น PDF',
    errUnsupportedExt: 'ไม่รองรับไฟล์ .{ext}',
    errNotFile: 'ไม่ใช่ไฟล์',
    errTooLarge: 'เกินขีดจำกัด {mb}MB',
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
    errNoModel: 'ยังไม่ได้ตั้งค่าชื่อโมเดล',
    menuFile: 'ไฟล์',
    menuNewDoc: 'เอกสารใหม่',
    menuNewWindow: 'หน้าต่างใหม่',
    menuOpen: 'เปิด…',
    menuOpenRecent: 'เปิดเอกสารล่าสุด',
    menuNoRecent: 'ไม่มีเอกสารล่าสุด',
    menuClose: 'ปิด',
    menuSave: 'บันทึก',
    menuSaveAs: 'บันทึกเป็น…',
    menuPageSetup: 'ตั้งค่าหน้ากระดาษ…',
    menuExportPdf: 'ส่งออกเป็น PDF…',
    menuPrint: 'พิมพ์…',
    menuEdit: 'แก้ไข',
    menuUndo: 'เลิกทำ',
    menuRedo: 'ทำซ้ำ',
    menuCut: 'ตัด',
    menuCopy: 'คัดลอก',
    menuPaste: 'วาง',
    menuPasteMatch: 'วางแบบจับคู่ลักษณะ',
    menuFindReplace: 'ค้นหาและแทนที่…',
    menuSelectAll: 'เลือกทั้งหมด',
    menuView: 'มุมมอง',
    menuZoomIn: 'ขยาย',
    menuZoomOut: 'ย่อ',
    menuZoom100: 'ขนาดจริง (100%)',
    menuPageWidth: 'ความกว้างของหน้า',
    menuWholePage: 'ทั้งหน้า',
    menuAiSidebar: 'แถบข้าง AI',
    menuDarkMode: 'โหมดมืด',
    menuFullscreen: 'เข้าสู่โหมดเต็มหน้าจอ',
    menuInsert: 'แทรก',
    menuInsertTable: 'ตาราง (3×3)',
    menuInsertImage: 'รูปภาพ…',
    menuInsertPageBreak: 'ตัวแบ่งหน้า',
    menuInsertLink: 'ไฮเปอร์ลิงก์…',
    menuInsertEquation: 'สมการ…',
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
    menuTools: 'เครื่องมือ',
    menuWordCount: 'นับจำนวนคำ…',
    menuSpelling: 'การสะกดและไวยากรณ์',
    menuMacros: 'แมโคร',
    menuWindow: 'หน้าต่าง',
    menuHelp: 'วิธีใช้',
    menuDocsHelp: 'วิธีใช้ GenOffice Docs',
  },
  id: {
    dlgOpenDoc: 'Buka Dokumen',
    filterWord: 'Dokumen Word',
    dlgSaveAs: 'Simpan Sebagai',
    closeUnsavedMsg: 'Dokumen ini memiliki perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    closeNoReplyMsg: 'Dokumen tidak merespons dan mungkin memiliki perubahan yang belum disimpan.',
    closeNoReplyDetail: 'Tetap tutup? Perubahan yang belum disimpan akan hilang.',
    btnCloseAnyway: 'Tetap Tutup',
    autosaveFoundTitle: 'Versi pemulihan ditemukan',
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
    dlgAddAttachment: 'Tambahkan Lampiran',
    filterSupported: 'File yang Didukung',
    filterAll: 'Semua File',
    dlgExportPdf: 'Ekspor sebagai PDF',
    errUnsupportedExt: 'file .{ext} tidak didukung',
    errNotFile: 'bukan file',
    errTooLarge: 'melebihi batas {mb}MB',
    errImageTooLarge: 'gambar melebihi batas 5MB',
    errUnreadable: 'tidak dapat dibaca',
    errFileTooLarge: 'File melebihi batas ukuran',
    errParseFailed: 'Gagal mengurai file',
    errImageNoText:
      'Lampiran gambar tidak menyediakan teks; gambar dikirim bersama pesan pengguna dan dapat dilihat langsung',
    errNotImage: 'bukan jenis gambar yang didukung',
    errGskNotLoggedIn: 'Belum masuk ke Genspark: klik “Masuk ke Genspark” di bawah, lalu coba lagi',
    errNoApiKey: 'API Key untuk {provider} belum dikonfigurasi',
    errNoModel: 'Nama model belum dikonfigurasi',
    menuFile: 'File',
    menuNewDoc: 'Dokumen Baru',
    menuNewWindow: 'Jendela Baru',
    menuOpen: 'Buka…',
    menuOpenRecent: 'Buka Terbaru',
    menuNoRecent: 'Tidak Ada Dokumen Terbaru',
    menuClose: 'Tutup',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuPageSetup: 'Penyetelan Halaman…',
    menuExportPdf: 'Ekspor sebagai PDF…',
    menuPrint: 'Cetak…',
    menuEdit: 'Edit',
    menuUndo: 'Urungkan',
    menuRedo: 'Ulangi',
    menuCut: 'Potong',
    menuCopy: 'Salin',
    menuPaste: 'Tempel',
    menuPasteMatch: 'Tempel dan Samakan Gaya',
    menuFindReplace: 'Temukan dan Ganti…',
    menuSelectAll: 'Pilih Semua',
    menuView: 'Tampilan',
    menuZoomIn: 'Perbesar',
    menuZoomOut: 'Perkecil',
    menuZoom100: 'Ukuran Sebenarnya (100%)',
    menuPageWidth: 'Lebar Halaman',
    menuWholePage: 'Seluruh Halaman',
    menuAiSidebar: 'Bilah Samping AI',
    menuDarkMode: 'Mode Gelap',
    menuFullscreen: 'Masuk Layar Penuh',
    menuInsert: 'Sisipkan',
    menuInsertTable: 'Tabel (3×3)',
    menuInsertImage: 'Gambar…',
    menuInsertPageBreak: 'Pemisah Halaman',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Persamaan…',
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
    menuTools: 'Alat',
    menuWordCount: 'Hitungan Kata…',
    menuSpelling: 'Ejaan dan Tata Bahasa',
    menuMacros: 'Makro',
    menuWindow: 'Jendela',
    menuHelp: 'Bantuan',
    menuDocsHelp: 'Bantuan GenOffice Docs',
  },
  ru: {
    dlgOpenDoc: 'Открыть документ',
    filterWord: 'Документы Word',
    dlgSaveAs: 'Сохранить как',
    closeUnsavedMsg: 'В этом документе есть несохранённые изменения.',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    closeNoReplyMsg: 'Документ не отвечает; возможно, есть несохранённые изменения.',
    closeNoReplyDetail: 'Всё равно закрыть? Несохранённые изменения будут потеряны.',
    btnCloseAnyway: 'Закрыть всё равно',
    autosaveFoundTitle: 'Найдена восстановленная версия',
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
    dlgAddAttachment: 'Добавить вложения',
    filterSupported: 'Поддерживаемые файлы',
    filterAll: 'Все файлы',
    dlgExportPdf: 'Экспорт в PDF',
    errUnsupportedExt: 'файлы .{ext} не поддерживаются',
    errNotFile: 'не является файлом',
    errTooLarge: 'превышает лимит {mb} МБ',
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
    errNoModel: 'Не указано имя модели',
    menuFile: 'Файл',
    menuNewDoc: 'Создать документ',
    menuNewWindow: 'Новое окно',
    menuOpen: 'Открыть…',
    menuOpenRecent: 'Открыть последние',
    menuNoRecent: 'Нет последних документов',
    menuClose: 'Закрыть',
    menuSave: 'Сохранить',
    menuSaveAs: 'Сохранить как…',
    menuPageSetup: 'Параметры страницы…',
    menuExportPdf: 'Экспорт в PDF…',
    menuPrint: 'Печать…',
    menuEdit: 'Правка',
    menuUndo: 'Отменить',
    menuRedo: 'Повторить',
    menuCut: 'Вырезать',
    menuCopy: 'Копировать',
    menuPaste: 'Вставить',
    menuPasteMatch: 'Вставить и согласовать стиль',
    menuFindReplace: 'Найти и заменить…',
    menuSelectAll: 'Выделить все',
    menuView: 'Вид',
    menuZoomIn: 'Увеличить',
    menuZoomOut: 'Уменьшить',
    menuZoom100: 'Фактический размер (100%)',
    menuPageWidth: 'По ширине страницы',
    menuWholePage: 'Страница целиком',
    menuAiSidebar: 'Боковая панель ИИ',
    menuDarkMode: 'Темный режим',
    menuFullscreen: 'Перейти в полноэкранный режим',
    menuInsert: 'Вставка',
    menuInsertTable: 'Таблица (3×3)',
    menuInsertImage: 'Рисунок…',
    menuInsertPageBreak: 'Разрыв страницы',
    menuInsertLink: 'Гиперссылка…',
    menuInsertEquation: 'Уравнение…',
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
    menuTools: 'Сервис',
    menuWordCount: 'Статистика…',
    menuSpelling: 'Правописание',
    menuMacros: 'Макросы',
    menuWindow: 'Окно',
    menuHelp: 'Справка',
    menuDocsHelp: 'Справка GenOffice Docs',
  },
  ar: {
    dlgOpenDoc: 'فتح مستند',
    filterWord: 'مستندات Word',
    dlgSaveAs: 'حفظ باسم',
    closeUnsavedMsg: 'يحتوي هذا المستند على تغييرات غير محفوظة.',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    closeNoReplyMsg: 'المستند لا يستجيب وقد يحتوي على تغييرات غير محفوظة.',
    closeNoReplyDetail: 'هل تريد الإغلاق رغم ذلك؟ ستفقد التغييرات غير المحفوظة.',
    btnCloseAnyway: 'إغلاق على أي حال',
    autosaveFoundTitle: 'تم العثور على نسخة مستردة',
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
    dlgAddAttachment: 'إضافة مرفقات',
    filterSupported: 'الملفات المدعومة',
    filterAll: 'كل الملفات',
    dlgExportPdf: 'تصدير بتنسيق PDF',
    errUnsupportedExt: 'ملفات .{ext} غير مدعومة',
    errNotFile: 'ليس ملفًا',
    errTooLarge: 'يتجاوز الحد {mb}MB',
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
    errNoModel: 'لم يتم تكوين اسم النموذج',
    menuFile: 'ملف',
    menuNewDoc: 'مستند جديد',
    menuNewWindow: 'نافذة جديدة',
    menuOpen: 'فتح…',
    menuOpenRecent: 'فتح الأخيرة',
    menuNoRecent: 'لا توجد مستندات أخيرة',
    menuClose: 'إغلاق',
    menuSave: 'حفظ',
    menuSaveAs: 'حفظ باسم…',
    menuPageSetup: 'إعداد الصفحة…',
    menuExportPdf: 'تصدير بتنسيق PDF…',
    menuPrint: 'طباعة…',
    menuEdit: 'تحرير',
    menuUndo: 'تراجع',
    menuRedo: 'إعادة',
    menuCut: 'قص',
    menuCopy: 'نسخ',
    menuPaste: 'لصق',
    menuPasteMatch: 'لصق مع مطابقة النمط',
    menuFindReplace: 'بحث واستبدال…',
    menuSelectAll: 'تحديد الكل',
    menuView: 'عرض',
    menuZoomIn: 'تكبير',
    menuZoomOut: 'تصغير',
    menuZoom100: 'الحجم الفعلي (100%)',
    menuPageWidth: 'عرض الصفحة',
    menuWholePage: 'صفحة كاملة',
    menuAiSidebar: 'الشريط الجانبي للذكاء الاصطناعي',
    menuDarkMode: 'الوضع الداكن',
    menuFullscreen: 'الدخول إلى ملء الشاشة',
    menuInsert: 'إدراج',
    menuInsertTable: 'جدول (3×3)',
    menuInsertImage: 'صورة…',
    menuInsertPageBreak: 'فاصل صفحات',
    menuInsertLink: 'ارتباط تشعبي…',
    menuInsertEquation: 'معادلة…',
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
    menuTools: 'أدوات',
    menuWordCount: 'عدد الكلمات…',
    menuSpelling: 'تدقيق إملائي ونحوي',
    menuMacros: 'وحدات الماكرو',
    menuWindow: 'نافذة',
    menuHelp: 'تعليمات',
    menuDocsHelp: 'تعليمات GenOffice Docs',
  },
  pt: {
    dlgOpenDoc: 'Abrir Documento',
    filterWord: 'Documentos do Word',
    dlgSaveAs: 'Salvar Como',
    closeUnsavedMsg: 'Este documento tem alterações não salvas.',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    closeNoReplyMsg: 'O documento não está respondendo e pode ter alterações não salvas.',
    closeNoReplyDetail: 'Fechar mesmo assim? As alterações não salvas serão perdidas.',
    btnCloseAnyway: 'Fechar mesmo assim',
    autosaveFoundTitle: 'Versão recuperada encontrada',
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
    dlgAddAttachment: 'Adicionar Anexos',
    filterSupported: 'Arquivos Compatíveis',
    filterAll: 'Todos os Arquivos',
    dlgExportPdf: 'Exportar como PDF',
    errUnsupportedExt: 'arquivos .{ext} não são suportados',
    errNotFile: 'não é um arquivo',
    errTooLarge: 'excede o limite de {mb}MB',
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
    errNoModel: 'Nenhum nome de modelo configurado',
    menuFile: 'Arquivo',
    menuNewDoc: 'Novo Documento',
    menuNewWindow: 'Nova Janela',
    menuOpen: 'Abrir…',
    menuOpenRecent: 'Abrir Recente',
    menuNoRecent: 'Nenhum Documento Recente',
    menuClose: 'Fechar',
    menuSave: 'Salvar',
    menuSaveAs: 'Salvar Como…',
    menuPageSetup: 'Configurar Página…',
    menuExportPdf: 'Exportar como PDF…',
    menuPrint: 'Imprimir…',
    menuEdit: 'Editar',
    menuUndo: 'Desfazer',
    menuRedo: 'Refazer',
    menuCut: 'Recortar',
    menuCopy: 'Copiar',
    menuPaste: 'Colar',
    menuPasteMatch: 'Colar com a Mesma Formatação',
    menuFindReplace: 'Localizar e Substituir…',
    menuSelectAll: 'Selecionar Tudo',
    menuView: 'Exibir',
    menuZoomIn: 'Ampliar',
    menuZoomOut: 'Reduzir',
    menuZoom100: 'Tamanho Real (100%)',
    menuPageWidth: 'Largura da Página',
    menuWholePage: 'Página Inteira',
    menuAiSidebar: 'Barra Lateral de IA',
    menuDarkMode: 'Modo Escuro',
    menuFullscreen: 'Entrar em Tela Cheia',
    menuInsert: 'Inserir',
    menuInsertTable: 'Tabela (3×3)',
    menuInsertImage: 'Imagem…',
    menuInsertPageBreak: 'Quebra de Página',
    menuInsertLink: 'Hiperlink…',
    menuInsertEquation: 'Equação…',
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
    menuTools: 'Ferramentas',
    menuWordCount: 'Contagem de Palavras…',
    menuSpelling: 'Ortografia e Gramática',
    menuMacros: 'Macros',
    menuWindow: 'Janela',
    menuHelp: 'Ajuda',
    menuDocsHelp: 'Ajuda do GenOffice Docs',
  },
  it: {
    dlgOpenDoc: 'Apri documento',
    filterWord: 'Documenti Word',
    dlgSaveAs: 'Salva con nome',
    closeUnsavedMsg: 'Questo documento contiene modifiche non salvate.',
    closeUnsavedDetail: 'Salvarle prima di chiudere?',
    closeNoReplyMsg: 'Il documento non risponde e potrebbe avere modifiche non salvate.',
    closeNoReplyDetail: 'Chiudere comunque? Le modifiche non salvate andranno perse.',
    btnCloseAnyway: 'Chiudi comunque',
    autosaveFoundTitle: 'Trovata versione recuperata',
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
    dlgAddAttachment: 'Aggiungi allegati',
    filterSupported: 'File supportati',
    filterAll: 'Tutti i file',
    dlgExportPdf: 'Esporta come PDF',
    errUnsupportedExt: 'i file .{ext} non sono supportati',
    errNotFile: 'non è un file',
    errTooLarge: 'supera il limite di {mb} MB',
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
    errNoModel: 'Nessun nome di modello configurato',
    menuFile: 'File',
    menuNewDoc: 'Nuovo documento',
    menuNewWindow: 'Nuova finestra',
    menuOpen: 'Apri…',
    menuOpenRecent: 'Apri recenti',
    menuNoRecent: 'Nessun documento recente',
    menuClose: 'Chiudi',
    menuSave: 'Salva',
    menuSaveAs: 'Salva con nome…',
    menuPageSetup: 'Imposta pagina…',
    menuExportPdf: 'Esporta come PDF…',
    menuPrint: 'Stampa…',
    menuEdit: 'Modifica',
    menuUndo: 'Annulla',
    menuRedo: 'Ripeti',
    menuCut: 'Taglia',
    menuCopy: 'Copia',
    menuPaste: 'Incolla',
    menuPasteMatch: 'Incolla e adatta lo stile',
    menuFindReplace: 'Trova e sostituisci…',
    menuSelectAll: 'Seleziona tutto',
    menuView: 'Visualizza',
    menuZoomIn: 'Ingrandisci',
    menuZoomOut: 'Riduci',
    menuZoom100: 'Dimensioni effettive (100%)',
    menuPageWidth: 'Larghezza pagina',
    menuWholePage: 'Pagina intera',
    menuAiSidebar: 'Barra laterale IA',
    menuDarkMode: 'Modalità scura',
    menuFullscreen: 'Attiva schermo intero',
    menuInsert: 'Inserisci',
    menuInsertTable: 'Tabella (3×3)',
    menuInsertImage: 'Immagine…',
    menuInsertPageBreak: 'Interruzione di pagina',
    menuInsertLink: 'Collegamento ipertestuale…',
    menuInsertEquation: 'Equazione…',
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
    menuTools: 'Strumenti',
    menuWordCount: 'Conteggio parole…',
    menuSpelling: 'Ortografia e grammatica',
    menuMacros: 'Macro',
    menuWindow: 'Finestra',
    menuHelp: 'Aiuto',
    menuDocsHelp: 'Guida di GenOffice Docs',
  },
  pl: {
    dlgOpenDoc: 'Otwórz dokument',
    filterWord: 'Dokumenty programu Word',
    dlgSaveAs: 'Zapisz jako',
    closeUnsavedMsg: 'Ten dokument zawiera niezapisane zmiany.',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    closeNoReplyMsg: 'Dokument nie odpowiada i może zawierać niezapisane zmiany.',
    closeNoReplyDetail: 'Zamknąć mimo to? Niezapisane zmiany zostaną utracone.',
    btnCloseAnyway: 'Zamknij mimo to',
    autosaveFoundTitle: 'Znaleziono odzyskaną wersję',
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
    dlgAddAttachment: 'Dodaj załączniki',
    filterSupported: 'Obsługiwane pliki',
    filterAll: 'Wszystkie pliki',
    dlgExportPdf: 'Eksportuj jako PDF',
    errUnsupportedExt: 'pliki .{ext} nie są obsługiwane',
    errNotFile: 'to nie jest plik',
    errTooLarge: 'przekracza limit {mb} MB',
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
    errNoModel: 'Nie skonfigurowano nazwy modelu',
    menuFile: 'Plik',
    menuNewDoc: 'Nowy dokument',
    menuNewWindow: 'Nowe okno',
    menuOpen: 'Otwórz…',
    menuOpenRecent: 'Otwórz ostatnie',
    menuNoRecent: 'Brak ostatnich dokumentów',
    menuClose: 'Zamknij',
    menuSave: 'Zapisz',
    menuSaveAs: 'Zapisz jako…',
    menuPageSetup: 'Ustawienia strony…',
    menuExportPdf: 'Eksportuj jako PDF…',
    menuPrint: 'Drukuj…',
    menuEdit: 'Edycja',
    menuUndo: 'Cofnij',
    menuRedo: 'Ponów',
    menuCut: 'Wytnij',
    menuCopy: 'Kopiuj',
    menuPaste: 'Wklej',
    menuPasteMatch: 'Wklej i dopasuj styl',
    menuFindReplace: 'Znajdź i zamień…',
    menuSelectAll: 'Zaznacz wszystko',
    menuView: 'Widok',
    menuZoomIn: 'Powiększ',
    menuZoomOut: 'Pomniejsz',
    menuZoom100: 'Rzeczywisty rozmiar (100%)',
    menuPageWidth: 'Szerokość strony',
    menuWholePage: 'Cała strona',
    menuAiSidebar: 'Pasek boczny AI',
    menuDarkMode: 'Tryb ciemny',
    menuFullscreen: 'Przejdź do pełnego ekranu',
    menuInsert: 'Wstaw',
    menuInsertTable: 'Tabela (3×3)',
    menuInsertImage: 'Obraz…',
    menuInsertPageBreak: 'Podział strony',
    menuInsertLink: 'Hiperłącze…',
    menuInsertEquation: 'Równanie…',
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
    menuTools: 'Narzędzia',
    menuWordCount: 'Statystyka wyrazów…',
    menuSpelling: 'Pisownia i gramatyka',
    menuMacros: 'Makra',
    menuWindow: 'Okno',
    menuHelp: 'Pomoc',
    menuDocsHelp: 'Pomoc GenOffice Docs',
  },
  nl: {
    dlgOpenDoc: 'Document openen',
    filterWord: 'Word-documenten',
    dlgSaveAs: 'Opslaan als',
    closeUnsavedMsg: 'Dit document bevat niet-opgeslagen wijzigingen.',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    closeNoReplyMsg: 'Het document reageert niet en heeft mogelijk niet-opgeslagen wijzigingen.',
    closeNoReplyDetail: 'Toch sluiten? Niet-opgeslagen wijzigingen gaan verloren.',
    btnCloseAnyway: 'Toch sluiten',
    autosaveFoundTitle: 'Herstelde versie gevonden',
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
    dlgAddAttachment: 'Bijlagen toevoegen',
    filterSupported: 'Ondersteunde bestanden',
    filterAll: 'Alle bestanden',
    dlgExportPdf: 'Exporteren als PDF',
    errUnsupportedExt: '.{ext}-bestanden worden niet ondersteund',
    errNotFile: 'geen bestand',
    errTooLarge: 'overschrijdt de limiet van {mb} MB',
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
    errNoModel: 'Geen modelnaam geconfigureerd',
    menuFile: 'Bestand',
    menuNewDoc: 'Nieuw document',
    menuNewWindow: 'Nieuw venster',
    menuOpen: 'Openen…',
    menuOpenRecent: 'Recent geopend',
    menuNoRecent: 'Geen recente documenten',
    menuClose: 'Sluiten',
    menuSave: 'Opslaan',
    menuSaveAs: 'Opslaan als…',
    menuPageSetup: 'Pagina-instelling…',
    menuExportPdf: 'Exporteren als PDF…',
    menuPrint: 'Afdrukken…',
    menuEdit: 'Bewerken',
    menuUndo: 'Ongedaan maken',
    menuRedo: 'Opnieuw',
    menuCut: 'Knippen',
    menuCopy: 'Kopiëren',
    menuPaste: 'Plakken',
    menuPasteMatch: 'Plakken met dezelfde stijl',
    menuFindReplace: 'Zoeken en vervangen…',
    menuSelectAll: 'Alles selecteren',
    menuView: 'Beeld',
    menuZoomIn: 'Inzoomen',
    menuZoomOut: 'Uitzoomen',
    menuZoom100: 'Ware grootte (100%)',
    menuPageWidth: 'Paginabreedte',
    menuWholePage: 'Hele pagina',
    menuAiSidebar: 'AI-zijbalk',
    menuDarkMode: 'Donkere modus',
    menuFullscreen: 'Schermvullende weergave',
    menuInsert: 'Invoegen',
    menuInsertTable: 'Tabel (3×3)',
    menuInsertImage: 'Afbeelding…',
    menuInsertPageBreak: 'Pagina-einde',
    menuInsertLink: 'Hyperlink…',
    menuInsertEquation: 'Vergelijking…',
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
    menuTools: 'Extra',
    menuWordCount: 'Woorden tellen…',
    menuSpelling: 'Spelling en grammatica',
    menuMacros: "Macro's",
    menuWindow: 'Venster',
    menuHelp: 'Help',
    menuDocsHelp: 'GenOffice Docs Help',
  },
  ms: {
    dlgOpenDoc: 'Buka Dokumen',
    filterWord: 'Dokumen Word',
    dlgSaveAs: 'Simpan Sebagai',
    closeUnsavedMsg: 'Dokumen ini mempunyai perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Adakah anda mahu menyimpannya sebelum menutup?',
    closeNoReplyMsg: 'Dokumen tidak bertindak balas dan mungkin ada perubahan yang belum disimpan.',
    closeNoReplyDetail: 'Tutup juga? Perubahan yang belum disimpan akan hilang.',
    btnCloseAnyway: 'Tutup Juga',
    autosaveFoundTitle: 'Versi pulihan ditemui',
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
    dlgAddAttachment: 'Tambah Lampiran',
    filterSupported: 'Fail yang Disokong',
    filterAll: 'Semua Fail',
    dlgExportPdf: 'Eksport sebagai PDF',
    errUnsupportedExt: 'fail .{ext} tidak disokong',
    errNotFile: 'bukan fail',
    errTooLarge: 'melebihi had {mb}MB',
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
    errNoModel: 'Nama model belum dikonfigurasikan',
    menuFile: 'Fail',
    menuNewDoc: 'Dokumen Baharu',
    menuNewWindow: 'Tetingkap Baharu',
    menuOpen: 'Buka…',
    menuOpenRecent: 'Buka Terkini',
    menuNoRecent: 'Tiada Dokumen Terkini',
    menuClose: 'Tutup',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuPageSetup: 'Persediaan Halaman…',
    menuExportPdf: 'Eksport sebagai PDF…',
    menuPrint: 'Cetak…',
    menuEdit: 'Edit',
    menuUndo: 'Buat Asal',
    menuRedo: 'Buat Semula',
    menuCut: 'Potong',
    menuCopy: 'Salin',
    menuPaste: 'Tampal',
    menuPasteMatch: 'Tampal dan Padankan Gaya',
    menuFindReplace: 'Cari dan Ganti…',
    menuSelectAll: 'Pilih Semua',
    menuView: 'Lihat',
    menuZoomIn: 'Zum Masuk',
    menuZoomOut: 'Zum Keluar',
    menuZoom100: 'Saiz Sebenar (100%)',
    menuPageWidth: 'Lebar Halaman',
    menuWholePage: 'Seluruh Halaman',
    menuAiSidebar: 'Bar Sisi AI',
    menuDarkMode: 'Mod Gelap',
    menuFullscreen: 'Masuk Skrin Penuh',
    menuInsert: 'Sisip',
    menuInsertTable: 'Jadual (3×3)',
    menuInsertImage: 'Imej…',
    menuInsertPageBreak: 'Pemisah Halaman',
    menuInsertLink: 'Hiperpautan…',
    menuInsertEquation: 'Persamaan…',
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
    menuTools: 'Alat',
    menuWordCount: 'Kiraan Perkataan…',
    menuSpelling: 'Ejaan dan Tatabahasa',
    menuMacros: 'Makro',
    menuWindow: 'Tetingkap',
    menuHelp: 'Bantuan',
    menuDocsHelp: 'Bantuan GenOffice Docs',
  },
  he: {
    dlgOpenDoc: 'פתיחת מסמך',
    filterWord: 'מסמכי Word',
    dlgSaveAs: 'שמירה בשם',
    closeUnsavedMsg: 'במסמך זה יש שינויים שלא נשמרו.',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    closeNoReplyMsg: 'המסמך אינו מגיב וייתכן שיש בו שינויים שלא נשמרו.',
    closeNoReplyDetail: 'לסגור בכל זאת? שינויים שלא נשמרו יאבדו.',
    btnCloseAnyway: 'סגור בכל זאת',
    autosaveFoundTitle: 'נמצאה גרסה משוחזרת',
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
    dlgAddAttachment: 'הוספת קבצים מצורפים',
    filterSupported: 'קבצים נתמכים',
    filterAll: 'כל הקבצים',
    dlgExportPdf: 'ייצוא כ-PDF',
    errUnsupportedExt: 'קובצי .{ext} אינם נתמכים',
    errNotFile: 'אינו קובץ',
    errTooLarge: 'חורג מהמגבלה של {mb}MB',
    errImageTooLarge: 'התמונה חורגת מהמגבלה של 5MB',
    errUnreadable: 'לא ניתן לקרוא',
    errFileTooLarge: 'הקובץ חורג ממגבלת הגודל',
    errParseFailed: 'ניתוח הקובץ נכשל',
    errImageNoText:
      'קבצים מצורפים מסוג תמונה אינם מספקים טקסט; התמונה נשלחת יחד עם הודעת המשתמש וניתן לצפות בה ישירות',
    errNotImage: 'סוג תמונה שאינו נתמך',
    errGskNotLoggedIn: 'לא מחובר ל-Genspark: לחץ על "התחבר ל-Genspark" למטה, התחבר ונסה שוב',
    errNoApiKey: 'לא הוגדר מפתח API עבור {provider}',
    errNoModel: 'לא הוגדר שם מודל',
    menuFile: 'קובץ',
    menuNewDoc: 'מסמך חדש',
    menuNewWindow: 'חלון חדש',
    menuOpen: 'פתיחה…',
    menuOpenRecent: 'פתיחת מסמכים אחרונים',
    menuNoRecent: 'אין מסמכים אחרונים',
    menuClose: 'סגירה',
    menuSave: 'שמירה',
    menuSaveAs: 'שמירה בשם…',
    menuPageSetup: 'הגדרת עמוד…',
    menuExportPdf: 'ייצוא כ-PDF…',
    menuPrint: 'הדפסה…',
    menuEdit: 'עריכה',
    menuUndo: 'בטל',
    menuRedo: 'בצע שוב',
    menuCut: 'גזור',
    menuCopy: 'העתק',
    menuPaste: 'הדבק',
    menuPasteMatch: 'הדבק והתאם סגנון',
    menuFindReplace: 'חיפוש והחלפה…',
    menuSelectAll: 'בחר הכול',
    menuView: 'תצוגה',
    menuZoomIn: 'התקרבות',
    menuZoomOut: 'התרחקות',
    menuZoom100: 'גודל אמיתי (100%)',
    menuPageWidth: 'רוחב עמוד',
    menuWholePage: 'עמוד שלם',
    menuAiSidebar: 'סרגל צד AI',
    menuDarkMode: 'מצב כהה',
    menuFullscreen: 'מעבר למסך מלא',
    menuInsert: 'הוספה',
    menuInsertTable: 'טבלה (3×3)',
    menuInsertImage: 'תמונה…',
    menuInsertPageBreak: 'מעבר עמוד',
    menuInsertLink: 'היפר-קישור…',
    menuInsertEquation: 'משוואה…',
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
    menuTools: 'כלים',
    menuWordCount: 'ספירת מילים…',
    menuSpelling: 'איות ודקדוק',
    menuMacros: 'פקודות מאקרו',
    menuWindow: 'חלון',
    menuHelp: 'עזרה',
    menuDocsHelp: 'עזרה של GenOffice Docs',
  },
  hi: {
    dlgOpenDoc: 'दस्तावेज़ खोलें',
    filterWord: 'Word दस्तावेज़',
    dlgSaveAs: 'इस रूप में सहेजें',
    closeUnsavedMsg: 'इस दस्तावेज़ में सहेजे नहीं गए परिवर्तन हैं।',
    closeUnsavedDetail: 'क्या आप बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    closeNoReplyMsg: 'दस्तावेज़ प्रतिक्रिया नहीं दे रहा है और उसमें सहेजे न गए बदलाव हो सकते हैं।',
    closeNoReplyDetail: 'फिर भी बंद करें? सहेजे न गए बदलाव खो जाएँगे।',
    btnCloseAnyway: 'फिर भी बंद करें',
    autosaveFoundTitle: 'पुनर्प्राप्त संस्करण मिला',
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
    dlgAddAttachment: 'अनुलग्नक जोड़ें',
    filterSupported: 'समर्थित फ़ाइलें',
    filterAll: 'सभी फ़ाइलें',
    dlgExportPdf: 'PDF के रूप में निर्यात करें',
    errUnsupportedExt: '.{ext} फ़ाइलें समर्थित नहीं हैं',
    errNotFile: 'फ़ाइल नहीं है',
    errTooLarge: '{mb}MB की सीमा से अधिक है',
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
    errNoModel: 'कोई मॉडल नाम कॉन्फ़िगर नहीं है',
    menuFile: 'फ़ाइल',
    menuNewDoc: 'नया दस्तावेज़',
    menuNewWindow: 'नई विंडो',
    menuOpen: 'खोलें…',
    menuOpenRecent: 'हाल के दस्तावेज़ खोलें',
    menuNoRecent: 'कोई हालिया दस्तावेज़ नहीं',
    menuClose: 'बंद करें',
    menuSave: 'सहेजें',
    menuSaveAs: 'इस रूप में सहेजें…',
    menuPageSetup: 'पृष्ठ सेटअप…',
    menuExportPdf: 'PDF के रूप में निर्यात करें…',
    menuPrint: 'प्रिंट करें…',
    menuEdit: 'संपादन',
    menuUndo: 'पूर्ववत करें',
    menuRedo: 'फिर से करें',
    menuCut: 'काटें',
    menuCopy: 'कॉपी करें',
    menuPaste: 'चिपकाएँ',
    menuPasteMatch: 'चिपकाएँ और शैली मिलाएँ',
    menuFindReplace: 'ढूँढें और बदलें…',
    menuSelectAll: 'सभी चुनें',
    menuView: 'दृश्य',
    menuZoomIn: 'ज़ूम इन',
    menuZoomOut: 'ज़ूम आउट',
    menuZoom100: 'वास्तविक आकार (100%)',
    menuPageWidth: 'पृष्ठ चौड़ाई',
    menuWholePage: 'पूरा पृष्ठ',
    menuAiSidebar: 'AI साइडबार',
    menuDarkMode: 'डार्क मोड',
    menuFullscreen: 'पूर्ण स्क्रीन में जाएँ',
    menuInsert: 'सम्मिलित करें',
    menuInsertTable: 'तालिका (3×3)',
    menuInsertImage: 'छवि…',
    menuInsertPageBreak: 'पृष्ठ विराम',
    menuInsertLink: 'हाइपरलिंक…',
    menuInsertEquation: 'समीकरण…',
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
    menuTools: 'उपकरण',
    menuWordCount: 'शब्द गणना…',
    menuSpelling: 'वर्तनी और व्याकरण',
    menuMacros: 'मैक्रो',
    menuWindow: 'विंडो',
    menuHelp: 'सहायता',
    menuDocsHelp: 'GenOffice Docs सहायता',
  },
  'zh-TW': {
    dlgOpenDoc: '開啟文件',
    filterWord: 'Word 文件',
    dlgSaveAs: '另存新檔',
    closeUnsavedMsg: '此文件有未儲存的變更。',
    closeUnsavedDetail: '關閉前是否要儲存？',
    closeNoReplyMsg: '文件沒有回應,可能有未儲存的變更。',
    closeNoReplyDetail: '仍要關閉嗎?未儲存的變更將遺失。',
    btnCloseAnyway: '仍要關閉',
    autosaveFoundTitle: '發現自動復原版本',
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
    dlgAddAttachment: '新增附件',
    filterSupported: '支援的檔案',
    filterAll: '所有檔案',
    dlgExportPdf: '匯出為 PDF',
    errUnsupportedExt: '暫不支援 .{ext} 類型',
    errNotFile: '不是檔案',
    errTooLarge: '超過 {mb}MB 上限',
    errImageTooLarge: '圖片超過 5MB 上限',
    errUnreadable: '無法讀取',
    errFileTooLarge: '檔案超過大小上限',
    errParseFailed: '檔案解析失敗',
    errImageNoText: '圖片附件不提供文字,已作為影像隨使用者訊息傳送,直接看圖即可',
    errNotImage: '不是支援的圖片類型',
    errGskNotLoggedIn: '未登入 Genspark:請點擊下方「登入 Genspark」完成登入後重試',
    errNoApiKey: '未設定 {provider} 的 API Key',
    errNoModel: '未設定模型名稱',
    menuFile: '檔案',
    menuNewDoc: '新增文件',
    menuNewWindow: '新增視窗',
    menuOpen: '開啟…',
    menuOpenRecent: '開啟最近的文件',
    menuNoRecent: '沒有最近的文件',
    menuClose: '關閉',
    menuSave: '儲存',
    menuSaveAs: '另存新檔…',
    menuPageSetup: '版面設定…',
    menuExportPdf: '匯出為 PDF…',
    menuPrint: '列印…',
    menuEdit: '編輯',
    menuUndo: '復原',
    menuRedo: '重做',
    menuCut: '剪下',
    menuCopy: '複製',
    menuPaste: '貼上',
    menuPasteMatch: '貼上並符合格式',
    menuFindReplace: '尋找與取代…',
    menuSelectAll: '全選',
    menuView: '檢視',
    menuZoomIn: '放大',
    menuZoomOut: '縮小',
    menuZoom100: '實際大小 (100%)',
    menuPageWidth: '頁面寬度',
    menuWholePage: '整頁',
    menuAiSidebar: 'AI 側邊欄',
    menuDarkMode: '深色模式',
    menuFullscreen: '進入全螢幕',
    menuInsert: '插入',
    menuInsertTable: '表格(3×3)',
    menuInsertImage: '圖片…',
    menuInsertPageBreak: '分頁符號',
    menuInsertLink: '超連結…',
    menuInsertEquation: '方程式…',
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
    menuTools: '工具',
    menuWordCount: '字數統計…',
    menuSpelling: '拼字及文法檢查',
    menuMacros: '巨集',
    menuWindow: '視窗',
    menuHelp: '說明',
    menuDocsHelp: 'GenOffice Docs 說明',
  },
})
const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(getUiLang(), key, params)

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
let pendingOpenPath = findDocxPath(process.argv)
/** documents queued for windows/tabs spawned via New Tab, keyed by webContents id */
const pendingWindowOpens = new Map<number, string>()
/** webContents ids that should open as a new blank doc instead of the start screen */
const pendingNewBlankIds = new Set<number>()

/** mark a docs webContents as "open blank on first consume" (called by the shell for home:new-doc) */
export function markDocsNewBlank(wcId: number): void {
  pendingNewBlankIds.add(wcId)
}

/** the single real BrowserWindow hosting the tab strip, used as dialog parent in tab mode */
let docsShellWindow: BrowserWindow | null = null
export function setDocsShellWindow(win: BrowserWindow | null): void {
  docsShellWindow = win
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

/** dialog parent for the calling tab (standalone mode falls back to its own BrowserWindow) */
function dialogParent(event: IpcMainInvokeEvent): BrowserWindow | undefined {
  return docsShellWindow ?? BrowserWindow.fromWebContents(event.sender) ?? undefined
}

async function openDialog(event: IpcMainInvokeEvent, options: OpenDialogOptions) {
  const parent = dialogParent(event)
  return parent ? dialog.showOpenDialog(parent, options) : dialog.showOpenDialog(options)
}

async function saveDialog(event: IpcMainInvokeEvent, options: SaveDialogOptions) {
  const parent = dialogParent(event)
  return parent ? dialog.showSaveDialog(parent, options) : dialog.showSaveDialog(options)
}

/** default folder where new files land on their first (silent) save; shared with the other editors via shell */
export function defaultSaveDir(): string {
  const dir = join(app.getPath('documents'), 'GenOffice')
  mkdirSync(dir, { recursive: true })
  return dir
}

/** first free path for fileName inside dir: name.ext, name-2.ext, name-3.ext… */
export function uniquePathIn(dir: string, fileName: string): string {
  const dot = fileName.lastIndexOf('.')
  const base = dot > 0 ? fileName.slice(0, dot) : fileName
  const ext = dot > 0 ? fileName.slice(dot) : ''
  let candidate = join(dir, fileName)
  for (let i = 2; existsSync(candidate); i++) candidate = join(dir, `${base}-${i}${ext}`)
  return candidate
}

export function openExternalDocx(filePath: string | null): void {
  if (!filePath || !/\.docx$/i.test(filePath)) return
  const win = BrowserWindow.getFocusedWindow() ?? mainWindow
  if (!rendererReady || !win) {
    pendingOpenPath = filePath
    return
  }
  void loadDocx(filePath, win.webContents.id)
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

function writeJson(path: string, value: unknown): void {
  mkdirSync(join(path, '..'), { recursive: true })
  writeFileSync(path, JSON.stringify(value, null, 2))
}

// ---- recent files ----

const RECENT_PATH = () => userDataPath('recent.json')

// the home screen lists all of these; the File menu shows only the first few
const RECENT_LIMIT = 100

function pushRecent(filePath: string): void {
  const recent = readJson<string[]>(RECENT_PATH(), [])
  // Every save lands here (autosave = every 30s): skip the write and the menu
  // rebuild when the file is already at the head of the list
  if (recent[0] === filePath) return
  const next = [filePath, ...recent.filter((p) => p !== filePath)].slice(0, RECENT_LIMIT)
  writeJson(RECENT_PATH(), next)
  buildDocsMenu() // keep File > Open Recent in sync
}

/** unified recents for the shell home screen (paths only; type = extension) */
export function readRecentFiles(): string[] {
  return readJson<string[]>(RECENT_PATH(), []).filter((p) => existsSync(p))
}

export function recordRecentFile(filePath: string): void {
  pushRecent(filePath)
}

export function removeRecentFiles(filePaths: string[]): void {
  const drop = new Set(filePaths)
  const recent = readJson<string[]>(RECENT_PATH(), [])
  writeJson(
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
  const states = docDiskStates.get(wc.id)
  const recorded = states?.get(oldPath)
  if (states && recorded) {
    states.delete(oldPath)
    states.set(newPath, recorded)
  }
  wc.send('docs:renamed', { oldPath, newPath })
}

/** keep a renamed file at its old position in the recent/starred lists */
export function replaceRecentFile(oldPath: string, newPath: string): void {
  const recent = readJson<string[]>(RECENT_PATH(), [])
  writeJson(
    RECENT_PATH(),
    recent.map((p) => (p === oldPath ? newPath : p)),
  )
  const starred = readJson<string[]>(STARRED_PATH(), [])
  if (starred.includes(oldPath)) {
    writeJson(
      STARRED_PATH(),
      starred.map((p) => (p === oldPath ? newPath : p)),
    )
  }
  buildDocsMenu()
}

// ---- starred files (home screen favorites) ----

const STARRED_PATH = () => userDataPath('starred.json')

export function readStarredFiles(): string[] {
  return readJson<string[]>(STARRED_PATH(), []).filter((p) => existsSync(p))
}

export function toggleStarredFile(filePath: string): void {
  const starred = readJson<string[]>(STARRED_PATH(), [])
  const next = starred.includes(filePath)
    ? starred.filter((p) => p !== filePath)
    : [...starred, filePath]
  writeJson(STARRED_PATH(), next)
}

// ---- original archive (pass-through base: original file archived by content hash) ----

async function archiveOriginal(filePath: string, bytes: Buffer): Promise<string> {
  const hash = sha256Hex(bytes)
  const dir = userDataPath('originals')
  await mkdir(dir, { recursive: true })
  const target = join(dir, `${hash}.docx`)
  if (!existsSync(target)) await copyFile(filePath, target)
  void pruneOriginals(dir)
  return hash
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

function allowDocWrite(wcId: number, filePath: string): void {
  const set = docWritablePaths.get(wcId) ?? new Set<string>()
  set.add(filePath)
  docWritablePaths.set(wcId, set)
}

function canDocWrite(wcId: number, filePath: string): boolean {
  return docWritablePaths.get(wcId)?.has(filePath) === true
}

function allowPdfWrite(wcId: number, filePath: string): void {
  const set = pdfWritablePaths.get(wcId) ?? new Set<string>()
  set.add(filePath)
  pdfWritablePaths.set(wcId, set)
}

function canPdfWrite(wcId: number, filePath: string): boolean {
  return pdfWritablePaths.get(wcId)?.has(filePath) === true
}

function dropDocWriter(wcId: number): void {
  docWritablePaths.delete(wcId)
  pdfWritablePaths.delete(wcId)
  docDiskStates.delete(wcId)
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

async function rememberDiskState(wcId: number, filePath: string, bytes: Buffer): Promise<void> {
  try {
    const s = await stat(filePath)
    const states = docDiskStates.get(wcId) ?? new Map<string, DiskFileState>()
    states.set(filePath, { mtimeMs: s.mtimeMs, size: s.size, hash: sha256Hex(bytes) })
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
  tornDownWcIds.add(contents.id)
  // Sweep recovery copies for this renderer's documents: every non-crash close
  // either saved (docs:save already cleared it) or explicitly discarded, so a
  // copy still on disk here is a leftover from an in-flight recovery write.
  for (const p of docWritablePaths.get(contents.id) ?? []) clearRecoveryCopy(p)
  // reclaim every per-wcId grant, not just doc saves — the orphaned renderer
  // must also lose its dialog-authorized PDF targets and disk-state cache
  docWritablePaths.delete(contents.id)
  pdfWritablePaths.delete(contents.id)
  docDiskStates.delete(contents.id)
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

/** On open, if a recovery copy newer than the original exists, ask whether to restore
 * (still points at the original path; only save persists it). */
async function maybeRecoverDocBytes(filePath: string, original: Buffer): Promise<Buffer> {
  const asPath = recoveryPathFor(filePath)
  try {
    if (!existsSync(asPath)) return original
    if (statSync(asPath).mtimeMs <= statSync(filePath).mtimeMs) {
      // a crashed partial write bumps mtime yet corrupts the file — keep the copy then
      if (looksLikeZip(original)) {
        unlinkSync(asPath)
        return original
      }
    }
  } catch {
    return original
  }
  const options = {
    type: 'question' as const,
    buttons: [tm('autosaveRestore'), tm('autosaveDiscard')],
    defaultId: 0,
    cancelId: 1,
    message: tm('autosaveFoundTitle'),
    detail: tm('autosaveFoundBody'),
  }
  const parent = BrowserWindow.getFocusedWindow() ?? mainWindow
  const r =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (r.response === 0) {
    try {
      return await readFile(asPath)
    } catch {
      return original
    }
  }
  clearRecoveryCopy(filePath)
  return original
}

async function loadDocx(filePath: string, wcId: number): Promise<OpenFileResult | null> {
  if (typeof filePath !== 'string' || !/\.docx$/i.test(filePath)) return null
  if (!existsSync(filePath)) return null
  const original = await readFile(filePath)
  const hash = await archiveOriginal(filePath, original)
  const bytes = await maybeRecoverDocBytes(filePath, original)
  pushRecent(filePath)
  allowDocWrite(wcId, filePath)
  // record the on-disk file, not the recovery copy: what matters is what save would overwrite
  await rememberDiskState(wcId, filePath, original)
  return {
    path: filePath,
    name: basename(filePath),
    data: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
    hash,
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
  'docx',
  'pdf',
  'pptx',
  'ppt',
  'xlsx',
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

/**
 * AI settings + chat/stream proxy handlers. Split out so the shell can
 * register them exactly once for all window types (docs, sheets, home) —
 * sheets' standalone AI handlers use the same channel names.
 */
export function registerAiIpc(): void {
  ipcMain.handle('ai:get-settings', (): AiSettings => {
    const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(SETTINGS_PATH(), {})
    const settings = resolveAiSettings(stored, defaultAiSettings())
    // AI features all go through Genspark (gsk login); legacy settings with another provider are reset
    settings.provider = 'genspark'
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

  ipcMain.handle('ai:gsk-login', () => {
    gskLogin()
  })

  ipcMain.handle('ai:set-settings', (_event, settings: AiSettings) => {
    writeJson(SETTINGS_PATH(), settings)
  })

  ipcMain.handle('ai:stream', async (event, request: AiStreamRequest) => {
    const { requestId, settings, system, messages } = request
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? 8192
    const provider = settings.provider
    let config = settings.providers?.[provider]
    // the genspark key never enters the settings file; requests take it from the gsk login state
    if (provider === 'genspark' && config && !config.apiKey) {
      config = { ...config, apiKey: gskApiKey() }
    }
    const send = (chunk: AiStreamChunk) => {
      if (!event.sender.isDestroyed()) event.sender.send('ai:stream-chunk', chunk)
    }
    if (!config?.apiKey) {
      send({
        requestId,
        type: 'error',
        error: provider === 'genspark' ? tm('errGskNotLoggedIn') : tm('errNoApiKey', { provider }),
      })
      return
    }
    if (!config.model) {
      send({ requestId, type: 'error', error: tm('errNoModel') })
      return
    }
    const controller = new AbortController()
    activeAiStreams.set(requestId, controller)
    // wire-activity keepalive: lets the renderer's silence watchdog tell a slow turn from a dead one
    let lastPing = 0
    const ping = () => {
      const now = Date.now()
      if (now - lastPing < 5_000) return
      lastPing = now
      send({ requestId, type: 'ping' })
    }
    try {
      let stopReason: string | undefined
      await streamForProvider(provider, config, system, messages, tools, maxTokens, {
        signal: controller.signal,
        onDelta: (text) => send({ requestId, type: 'delta', text }),
        onToolCall: (toolCall) => send({ requestId, type: 'tool-call', toolCall }),
        onActivity: ping,
        onStopReason: (reason) => {
          stopReason = reason
        },
      })
      send({ requestId, type: 'done', stopReason })
    } catch (err) {
      if (controller.signal.aborted) {
        send({ requestId, type: 'done' })
      } else {
        send({
          requestId,
          type: 'error',
          error: err instanceof Error ? err.message : String(err),
          ...(err instanceof AiTimeoutError
            ? { errorCode: 'timeout' as const }
            : err instanceof AiCreditsError
              ? { errorCode: 'credits' as const }
              : {}),
        })
      }
    } finally {
      activeAiStreams.delete(requestId)
    }
  })

  ipcMain.handle('ai:stream-cancel', (_event, requestId: string) => {
    activeAiStreams.get(requestId)?.abort()
  })

  // shared search tools (content + images): Serper with DuckDuckGo fallback (same source as slides/sheets)
  ipcMain.handle('ai:web-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await webSearch(String(query), typeof maxResults === 'number' ? maxResults : 6)
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })
  ipcMain.handle('ai:image-search', async (_event, query: string, maxResults?: number) => {
    try {
      return await imageSearch(String(query), typeof maxResults === 'number' ? maxResults : 8)
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })

  // download image from URL → base64+mime (download in the main process avoids CORS; the renderer builds the image node and measures size itself)
  ipcMain.handle(
    'ai:fetch-image',
    async (_event, url: string): Promise<{ base64: string; mime: string } | null> => {
      try {
        // the URL originates from AI tool calls (prompt-injectable via web search
        // results), so refuse non-http schemes and private/link-local targets;
        // redirects are followed manually so every hop is validated too
        const resp = await fetchWithSsrfGuard(String(url), {
          headers: { 'User-Agent': 'Mozilla/5.0' },
        })
        if (!resp || !resp.ok) return null
        const buf = Buffer.from(await resp.arrayBuffer())
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

  ipcMain.handle('ai:chat', async (_event, request: AiChatRequest) => {
    const { settings, system, user } = request
    const provider = settings.provider
    let config = settings.providers?.[provider]
    if (provider === 'genspark' && config && !config.apiKey) {
      config = { ...config, apiKey: gskApiKey() }
    }
    if (!config?.apiKey) {
      return {
        ok: false,
        error: provider === 'genspark' ? tm('errGskNotLoggedIn') : tm('errNoApiKey', { provider }),
      }
    }
    if (!config.model) return { ok: false, error: tm('errNoModel') }
    try {
      return await chatForProvider(provider, config, system, user)
    } catch (err) {
      return { ok: false, error: String(err) }
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
 */
let fileSavedHook: ((wc: WebContents, filePath: string) => void) | null = null

export function setDocsFileSavedHook(hook: (wc: WebContents, filePath: string) => void): void {
  fileSavedHook = hook
}

function notifyFileSaved(wc: WebContents, filePath: string): void {
  if (fileSavedHook) fileSavedHook(wc, filePath)
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
      },
    ) => {
      const store = getProjectStore()
      const msg: Parameters<ProjectStore['appendChatMessage']>[2] = {
        role: args.role,
        text: args.text,
      }
      if (args.tools) msg.tools = args.tools
      if (args.attachments) msg.attachments = args.attachments
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

  // ── P1 extension IPC ─────────────────────────────────────

  /** List all projects (with file count + last-active time) */
  ipcMain.handle('project:list', () => {
    return getProjectStore().listProjectsSummary()
  })

  /** List existing files belonging to one project */
  ipcMain.handle('project:files', (_event, args: { projectId: string }) => {
    return getProjectStore().listProjectFiles(args.projectId)
  })

  /** Create a project */
  ipcMain.handle('project:create', (_event, args: { name: string }) => {
    const store = getProjectStore()
    const data = store.createProject(args.name)
    // returns ProjectSummary shape
    return store.listProjectsSummary().find((s) => s.id === data.id) ?? data
  })

  /** Rename a project */
  ipcMain.handle('project:rename', (_event, args: { id: string; name: string }) => {
    getProjectStore().renameProject(args.id, args.name)
  })

  /** Soft-delete a project */
  ipcMain.handle('project:delete', (_event, args: { id: string }) => {
    getProjectStore().deleteProject(args.id)
  })

  /** Move a file into the given project */
  ipcMain.handle('project:moveFile', (_event, args: { filePath: string; projectId: string }) => {
    getProjectStore().moveFileToProject(args.filePath, args.projectId)
  })

  /** Get the project timeline */
  ipcMain.handle('project:timeline', (_event, args: { projectId: string; limit?: number }) => {
    return getProjectStore().getProjectTimeline(args.projectId, args.limit ?? 20)
  })
}

/** document/attachment/window IPC (everything except the AI proxy above) */
export function registerDocsIpc(): void {
  // shared with the other editor modules — last (identical) registration wins
  ipcMain.removeHandler('app:get-language')
  ipcMain.handle('app:get-language', () => getUiLang())

  ipcMain.handle('docs:open', async (event) => {
    const result = await openDialog(event, {
      title: tm('dlgOpenDoc'),
      filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    return loadDocx(result.filePaths[0], event.sender.id)
  })

  ipcMain.handle('docs:open-path', (event, filePath: string) => loadDocx(filePath, event.sender.id))

  ipcMain.handle('docs:consume-pending-open', (event) => {
    rendererReady = true
    // a tab spawned via New Tab loads the document queued for it specifically
    const queued = pendingWindowOpens.get(event.sender.id)
    if (queued) {
      pendingWindowOpens.delete(event.sender.id)
      return loadDocx(queued, event.sender.id)
    }
    const filePath = pendingOpenPath
    pendingOpenPath = null
    return filePath ? loadDocx(filePath, event.sender.id) : null
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
        const bytes = Buffer.from(data)
        await atomicWriteFile(filePath, bytes)
        await rememberDiskState(event.sender.id, filePath, bytes)
        clearRecoveryCopy(filePath)
        pushRecent(filePath)
        return { ok: true }
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
      await atomicWriteFile(recoveryPathFor(filePath), Buffer.from(data))
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

  ipcMain.handle('docs:save-as', async (event, defaultName: string, data: ArrayBuffer) => {
    // an orphaned (closed-tab) renderer must not open dialogs or land new files
    if (tornDownWcIds.has(event.sender.id)) return { ok: false }
    const result = await saveDialog(event, {
      title: tm('dlgSaveAs'),
      defaultPath: defaultName,
      filters: [{ name: tm('filterWord'), extensions: ['docx'] }],
    })
    if (result.canceled || !result.filePath) return { ok: false }
    // the tab may have been closed while the dialog was open; checked before the
    // write because Save As may overwrite an existing file (no safe rollback)
    if (tornDownWcIds.has(event.sender.id)) return { ok: false }
    try {
      const bytes = Buffer.from(data)
      await atomicWriteFile(result.filePath, bytes)
      allowDocWrite(event.sender.id, result.filePath)
      await rememberDiskState(event.sender.id, result.filePath, bytes)
      pushRecent(result.filePath)
      notifyFileSaved(event.sender, result.filePath)
      return { ok: true, path: result.filePath }
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  ipcMain.handle('docs:save-new', async (event, defaultName: string, data: ArrayBuffer) => {
    try {
      // a discarded draft in an orphaned renderer must not silently persist
      // itself to the default folder after the user chose Don't Save
      if (tornDownWcIds.has(event.sender.id)) return { ok: false }
      const filePath = uniquePathIn(defaultSaveDir(), defaultName)
      const bytes = Buffer.from(data)
      await atomicWriteFile(filePath, bytes)
      // teardown may have happened while the write was in flight — the path is
      // freshly created, so rolling it back is safe (mirrors docs:write-recovery)
      if (tornDownWcIds.has(event.sender.id)) {
        await unlink(filePath).catch(() => {})
        return { ok: false }
      }
      allowDocWrite(event.sender.id, filePath)
      await rememberDiskState(event.sender.id, filePath, bytes)
      pushRecent(filePath)
      notifyFileSaved(event.sender, filePath)
      return { ok: true, path: filePath }
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  ipcMain.handle('docs:recent', () =>
    readJson<string[]>(RECENT_PATH(), []).filter((p) => existsSync(p)),
  )

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
    return collectAttachments(result.filePaths)
  })

  ipcMain.handle('files:add', (_event, paths: string[]) => collectAttachments(paths))

  ipcMain.handle(
    'files:read',
    async (
      _event,
      filePath: string,
      offset: number,
      maxChars: number,
    ): Promise<AttachmentReadResult> => {
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
  ipcMain.handle('files:read-image', (_event, filePath: string): AttachmentImageResult => {
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
    (_event, data: unknown, ext: unknown): AttachmentAddResult => {
      const filePath = savePastedImage(data, ext)
      return filePath
        ? collectAttachments([filePath])
        : { accepted: [], rejected: [tm('errNotImage')] }
    },
  )

  ipcMain.handle('docs:print', (event) => {
    // print the calling tab's own content; zero margins — the docx page padding provides them
    event.sender.print({ margins: { marginType: 'none' } })
  })

  ipcMain.handle(
    'docs:export-pdf',
    async (
      event,
      defaultName: string,
      pageWidthTwips: number,
      pageHeightTwips: number,
      outPath?: string,
    ) => {
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
        const data = await event.sender.printToPDF({
          printBackground: true,
          // custom pageSize is in inches; the docx page padding provides the margins
          pageSize: {
            width: pageWidthTwips / TWIPS_PER_INCH,
            height: pageHeightTwips / TWIPS_PER_INCH,
          },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
        })
        writeFileSync(filePath, data)
        return { ok: true, path: filePath }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  // mixed paper-size export: the renderer prints group by group per size (other pages hidden via CSS); this produces one group's bytes
  ipcMain.handle(
    'docs:print-pdf-buffer',
    async (event, pageWidthTwips: number, pageHeightTwips: number) => {
      try {
        const data = await event.sender.printToPDF({
          printBackground: true,
          pageSize: {
            width: pageWidthTwips / TWIPS_PER_INCH,
            height: pageHeightTwips / TWIPS_PER_INCH,
          },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
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
    async (event, defaultName: string, base64Parts: string[], outPath?: string) => {
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
        writeFileSync(filePath, Buffer.from(await merged.save()))
        return { ok: true, path: filePath }
      } catch (err) {
        return { ok: false, error: String(err) }
      }
    },
  )

  ipcMain.handle('win:new', (_event, openPath: string | null) => {
    // A pathless new tab/window starts as a blank document, not the start screen
    const path = openPath ?? undefined
    if (shellHooks) shellHooks.openTab(path, path ? undefined : { newBlank: true })
    else {
      const win = createDocsWindow(path)
      if (!path) markDocsNewBlank(win.webContents.id)
    }
  })

  ipcMain.handle('win:list', (): DocsTabInfo[] => {
    if (shellHooks) return shellHooks.listTabs()
    return BrowserWindow.getAllWindows().map((w) => ({
      id: String(w.id),
      title: w.getTitle(),
      focused: w.isFocused(),
    }))
  })

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
  listTabs(): DocsTabInfo[]
  focusTab(id: string): void
  /** closes the calling tab instead of the whole shell window (Cmd+W / role:'close') */
  closeActiveTab(): void
}
let shellHooks: DocsShellHooks | null = null
export function setDocsShellHooks(hooks: DocsShellHooks | null): void {
  shellHooks = hooks
}

// ---- application menu ----

function sendCommand(command: MenuCommand, payload?: string): void {
  activeDocsWebContents()?.send('menu:command', command, payload)
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
  const recent = readJson<string[]>(RECENT_PATH(), [])
    .filter((p) => existsSync(p))
    .slice(0, 10)

  const recentSubmenu: MenuItemConstructorOptions[] =
    recent.length > 0
      ? recent.map((p) => ({
          label: basename(p),
          sublabel: isMac ? undefined : p,
          toolTip: p,
          click: () => sendCommand('open-path', p),
        }))
      : [{ label: tm('menuNoRecent'), enabled: false }]

  const template: MenuItemConstructorOptions[] = [
    ...(isMac
      ? [
          {
            label: app.name,
            submenu: [
              { role: 'about' as const },
              { type: 'separator' as const },
              { role: 'services' as const },
              { type: 'separator' as const },
              { role: 'hide' as const },
              { role: 'hideOthers' as const },
              { role: 'unhide' as const },
              { type: 'separator' as const },
              { role: 'quit' as const },
            ],
          },
        ]
      : []),
    {
      label: tm('menuFile'),
      submenu: [
        { label: tm('menuNewDoc'), accelerator: 'CmdOrCtrl+N', click: () => sendCommand('new') },
        {
          label: tm('menuNewWindow'),
          accelerator: 'Shift+CmdOrCtrl+N',
          click: () => {
            if (shellHooks) shellHooks.openTab(undefined, { newBlank: true })
            else markDocsNewBlank(createDocsWindow().webContents.id)
          },
        },
        { label: tm('menuOpen'), accelerator: 'CmdOrCtrl+O', click: () => sendCommand('open') },
        { label: tm('menuOpenRecent'), submenu: recentSubmenu },
        ...(extraFileMenuItems.length > 0
          ? [{ type: 'separator' as const }, ...extraFileMenuItems]
          : []),
        { type: 'separator' },
        shellHooks
          ? {
              label: tm('menuClose'),
              accelerator: 'CmdOrCtrl+W',
              click: () => shellHooks?.closeActiveTab(),
            }
          : { role: 'close' as const, label: tm('menuClose') },
        { label: tm('menuSave'), accelerator: 'CmdOrCtrl+S', click: () => sendCommand('save') },
        {
          label: tm('menuSaveAs'),
          accelerator: 'Shift+CmdOrCtrl+S',
          click: () => sendCommand('save-as'),
        },
        { type: 'separator' },
        { label: tm('menuPageSetup'), click: () => sendCommand('page-setup') },
        { label: tm('menuExportPdf'), click: () => sendCommand('export-pdf') },
        {
          label: tm('menuPrint'),
          accelerator: 'CmdOrCtrl+P',
          click: () => activeDocsWebContents()?.print({}),
        },
      ],
    },
    {
      label: tm('menuEdit'),
      submenu: [
        { label: tm('menuUndo'), accelerator: 'CmdOrCtrl+Z', click: () => sendCommand('undo') },
        {
          label: tm('menuRedo'),
          accelerator: 'Shift+CmdOrCtrl+Z',
          click: () => sendCommand('redo'),
        },
        { type: 'separator' },
        { role: 'cut', label: tm('menuCut') },
        { role: 'copy', label: tm('menuCopy') },
        { role: 'paste', label: tm('menuPaste') },
        { role: 'pasteAndMatchStyle', label: tm('menuPasteMatch') },
        { type: 'separator' },
        {
          label: tm('menuFindReplace'),
          accelerator: 'CmdOrCtrl+F',
          click: () => sendCommand('find'),
        },
        { type: 'separator' },
        { role: 'selectAll', label: tm('menuSelectAll') },
      ],
    },
    {
      label: tm('menuView'),
      submenu: [
        {
          label: tm('menuZoomIn'),
          accelerator: 'CmdOrCtrl+=',
          click: () => sendCommand('zoom-in'),
        },
        {
          label: tm('menuZoomOut'),
          accelerator: 'CmdOrCtrl+-',
          click: () => sendCommand('zoom-out'),
        },
        {
          label: tm('menuZoom100'),
          accelerator: 'CmdOrCtrl+0',
          click: () => sendCommand('zoom-100'),
        },
        { label: tm('menuPageWidth'), click: () => sendCommand('zoom-page-width') },
        { label: tm('menuWholePage'), click: () => sendCommand('zoom-whole-page') },
        { type: 'separator' },
        { label: tm('menuAiSidebar'), click: () => sendCommand('toggle-ai') },
        { label: tm('menuDarkMode'), click: () => sendCommand('toggle-dark') },
        { type: 'separator' },
        { role: 'togglefullscreen', label: tm('menuFullscreen') },
        ...(isDev ? [{ role: 'toggleDevTools' as const }] : []),
      ],
    },
    {
      label: tm('menuInsert'),
      submenu: [
        { label: tm('menuInsertTable'), click: () => sendCommand('insert-table') },
        { label: tm('menuInsertImage'), click: () => sendCommand('insert-image') },
        { label: tm('menuInsertPageBreak'), click: () => sendCommand('insert-page-break') },
        {
          label: tm('menuInsertLink'),
          accelerator: 'CmdOrCtrl+K',
          click: () => sendCommand('insert-link'),
        },
        { label: tm('menuInsertEquation'), click: () => sendCommand('insert-equation') },
        { type: 'separator' },
        { label: tm('menuComment'), click: () => sendCommand('insert-comment') },
      ],
    },
    {
      label: tm('menuFormat'),
      submenu: [
        { label: tm('menuBold'), accelerator: 'CmdOrCtrl+B', click: () => sendCommand('bold') },
        { label: tm('menuItalic'), accelerator: 'CmdOrCtrl+I', click: () => sendCommand('italic') },
        {
          label: tm('menuUnderline'),
          accelerator: 'CmdOrCtrl+U',
          click: () => sendCommand('underline'),
        },
        { type: 'separator' },
        {
          label: tm('menuAlign'),
          submenu: [
            { label: tm('menuAlignLeft'), click: () => sendCommand('align-left') },
            { label: tm('menuAlignCenter'), click: () => sendCommand('align-center') },
            { label: tm('menuAlignRight'), click: () => sendCommand('align-right') },
            { label: tm('menuAlignJustify'), click: () => sendCommand('align-justify') },
          ],
        },
        { type: 'separator' },
        {
          label: tm('menuFont'),
          accelerator: 'CmdOrCtrl+D',
          click: () => sendCommand('font-dialog'),
        },
        {
          label: tm('menuParagraph'),
          accelerator: 'Alt+CmdOrCtrl+M',
          click: () => sendCommand('paragraph-dialog'),
        },
      ],
    },
    {
      // Word for Mac keeps Word Count in the Tools menu, not on the ribbon
      label: tm('menuTools'),
      submenu: [
        { label: tm('menuWordCount'), click: () => sendCommand('word-count') },
        { type: 'separator' },
        { label: tm('menuSpelling'), enabled: false },
        { label: tm('menuMacros'), enabled: false },
      ],
    },
    windowMenuTemplate(process.platform, appMenuLabels(getUiLang())),
    {
      label: tm('menuHelp'),
      role: 'help',
      submenu: [{ label: tm('menuDocsHelp'), enabled: false }],
    },
  ]

  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// ---- window ----

export function createDocsWindow(openPath?: string): BrowserWindow {
  const win = new BrowserWindow({
    width: 1360,
    height: 900,
    minWidth: 980,
    minHeight: 600,
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
    const target = safeExternalUrl(url)
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })

  if (runtime.rendererUrl) {
    void win.loadURL(runtime.rendererUrl)
  } else {
    void win.loadFile(runtime.rendererFile)
  }
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
    const target = safeExternalUrl(url)
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })

  // mode=tab: the shell's tab strip owns the traffic lights / caption buttons,
  // so the ribbon must not reserve space for them
  if (runtime.rendererUrl) {
    // append via URL so a dev URL that already carries query params stays valid
    const devUrl = new URL(runtime.rendererUrl)
    devUrl.searchParams.set('mode', 'tab')
    void view.webContents.loadURL(devUrl.toString())
  } else {
    void view.webContents.loadFile(runtime.rendererFile, { query: { mode: 'tab' } })
  }
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
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  // dev runs must not share the packaged app's userData (recent files, AI settings)
  // or its single-instance lock — otherwise `npm run dev` silently quits whenever
  // the installed GenOffice Docs is open and forwards its argv there instead.
  if (isDev) app.setPath('userData', join(app.getPath('appData'), 'GenOffice Docs Dev'))

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
    openExternalDocx(findDocxPath(argv))
    mainWindow?.show()
    mainWindow?.focus()
  })

  registerAiIpc()
  registerProjectIpc()
  registerDocsIpc()

  app.whenReady().then(() => {
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
