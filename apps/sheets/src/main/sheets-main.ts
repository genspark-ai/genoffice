import { createHash, randomUUID } from 'node:crypto'
import {
  createReadStream,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises'
import { createServer } from 'node:http'
import { basename, dirname, isAbsolute, join } from 'node:path'

import {
  app,
  BrowserWindow,
  desktopCapturer,
  dialog,
  ipcMain,
  Menu,
  screen,
  session as electronSession,
  shell,
  systemPreferences,
  WebContentsView,
} from 'electron'
import type {
  IpcMainInvokeEvent,
  MenuItemConstructorOptions,
  OpenDialogOptions,
  SaveDialogOptions,
  WebContents,
} from 'electron'
import { z } from 'zod'
import {
  appMenuLabels,
  configuredDefaultSaveDir,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  viewMenuTemplate,
  windowMenuTemplate,
} from '@genoffice/electron-utils'
import { createI18n, getUiLang, type Lang, normalizeLang, setUiLang } from '@genoffice/i18n'
import { ProjectStore } from '@genoffice/project-store'

import {
  AiCreditsError,
  AiTimeoutError,
  chatForProvider,
  defaultAiSettings,
  resolveAiSettings,
  streamForProvider,
  type AiProviderId,
  type AiSettings,
  type AiStreamChunk,
  type GenSparkAccountStatus,
  type LegacyAiSettings,
} from '@genoffice/ai-provider'
import { csvToXlsxBuffer, decodeCsvBuffer } from '../gateway/csv-import'
import {
  ensureGenofficeLogin,
  gskApiKey,
  gskLoginInfo,
  hasGskAuth,
  setGskProxyUrl,
  webSearch,
  imageSearch,
} from '@genoffice/ai-search'
import { parseFileToText } from '@genoffice/file-parse'
import type { CellEdit, SheetStructuralOps } from '../gateway/xlsx-gateway'
import { readArchiveEntryText, saveWorkbookViaSidecar } from '../gateway/xlsx-package-io'
import { parsePivotDefinition } from '../gateway/xlsx-pivot'
import type { SheetEditPlan } from '../gateway/xlsx-sheets'
import type {
  AttachmentAddResult,
  AttachmentImageResult,
  AttachmentMeta,
  AttachmentReadResult,
  WorkbookFile,
} from '../shared/desktop-api'
import {
  ATTACHMENT_IMAGE_EXTS,
  aiChatRequestSchema,
  aiSettingsInputSchema,
  aiStreamRequestSchema,
  workbookFileSchema,
  workbookFormulaCellsRequestSchema,
  workbookFormulaCellsResultSchema,
  workbookRecalcRequestSchema,
  workbookRecalcResultSchema,
  workbookMediaRequestSchema,
  workbookMediaResultSchema,
  workbookPivotRequestSchema,
  localImageRequestSchema,
  localImageResultSchema,
  screenCaptureRequestSchema,
  screenCaptureResultSchema,
  screenSourcesResultSchema,
  workbookPivotDefinitionSchema,
  workbookExportPdfRequestSchema,
  workbookRangeRequestSchema,
  workbookRangeResultSchema,
  workbookSaveRequestSchema,
  type WorkbookSaveRequest,
} from '../shared/desktop-api'
import { IPC_CHANNELS } from '../shared/ipc-channels'
import { closeGuardDecision } from './close-guard'
import { exportPdf } from './pdf-export'
import { XlsxSidecarClient } from './xlsx-sidecar-client'

/**
 * Sheets main-process logic as an embeddable module: no top-level lifecycle.
 * Standalone mode (apps/sheets entry) calls startSheetsStandalone(); the
 * unified shell calls configureSheetsRuntime() + createSheetsWindow() and
 * owns the app lifecycle. AI IPC is registered separately so the shell can
 * substitute its single unified handler set (same channel names as docs).
 */

const tMain = createI18n({
  zh: {
    filterSpreadsheets: '电子表格',
    filterXlsx: 'Excel 工作簿',
    dlgAddAttachment: '添加附件',
    filterSupported: '支持的文件',
    filterAll: '所有文件',
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
    errImgAbsPath: '图片路径必须是绝对路径。',
    errImgNotFound: '找不到图片文件: {path}',
    errImgTooLarge20: '图片超过 20MB,不支持插入。',
    errImgBadType: '该文件不是 PNG/JPEG/GIF 图片。',
    errDiskChanged: '工作簿在打开后被磁盘上的改动覆盖——请改用另存为。',
    autosaveFoundTitle: '发现自动恢复版本',
    autosaveFoundBody: '上次会话有未保存的更改。要恢复自动保存的版本吗?',
    autosaveRestore: '恢复',
    autosaveDiscard: '放弃',
    menuFile: '文件',
    menuOpenWorkbook: '打开工作簿…',
    menuSave: '保存',
    menuSaveAs: '另存为…',
    menuExportPdf: '导出 PDF…',
    menuClose: '关闭',
    menuQuit: '退出',
    menuEdit: '编辑',
    menuUndo: '撤销',
    menuRedo: '重做',
    closeUnsavedMsg: '有 {count} 处未保存的修改',
    closeUnsavedDetail: '不保存直接关闭,这些修改将丢失。',
    btnDontSave: '不保存',
    btnCancel: '取消',
    csvSaveAsNotice: 'CSV 格式不保留样式等格式修改——另存为 .xlsx 可保留全部内容。',
  },
  en: {
    filterSpreadsheets: 'Spreadsheets',
    filterXlsx: 'Excel Workbooks',
    dlgAddAttachment: 'Add Attachments',
    filterSupported: 'Supported Files',
    filterAll: 'All Files',
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
    errImgAbsPath: 'Image path must be absolute.',
    errImgNotFound: 'Image file not found: {path}',
    errImgTooLarge20: 'Image exceeds 20MB and cannot be inserted.',
    errImgBadType: 'The file is not a PNG/JPEG/GIF image.',
    errDiskChanged: 'The workbook changed on disk after it was opened — use Save As instead.',
    autosaveFoundTitle: 'Recovered version found',
    autosaveFoundBody:
      'There are unsaved changes from your last session. Restore the autosaved version?',
    autosaveRestore: 'Restore',
    autosaveDiscard: 'Discard',
    menuFile: 'File',
    menuOpenWorkbook: 'Open Workbook…',
    menuSave: 'Save',
    menuSaveAs: 'Save As…',
    menuExportPdf: 'Export PDF…',
    menuClose: 'Close',
    menuQuit: 'Quit',
    menuEdit: 'Edit',
    menuUndo: 'Undo',
    menuRedo: 'Redo',
    closeUnsavedMsg: '{count} unsaved change(s)',
    closeUnsavedDetail: 'Your changes will be lost if you close without saving.',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
    csvSaveAsNotice: "CSV files can't keep formatting — saving as .xlsx keeps all your changes.",
  },
  ja: {
    filterSpreadsheets: 'スプレッドシート',
    filterXlsx: 'Excel ブック',
    dlgAddAttachment: '添付ファイルを追加',
    filterSupported: 'サポートされているファイル',
    filterAll: 'すべてのファイル',
    errUnsupportedExt: '.{ext} 形式には対応していません',
    errNotFile: 'ファイルではありません',
    errTooLarge: '{mb}MB の上限を超えています',
    errImageTooLarge: '画像が 5MB の上限を超えています',
    errUnreadable: '読み取れません',
    errFileTooLarge: 'ファイルがサイズ上限を超えています',
    errParseFailed: 'ファイルの解析に失敗しました',
    errImageNoText:
      '画像添付にはテキストがありません。画像はユーザー メッセージと一緒に送信されるため、そのまま画像をご確認ください',
    errNotImage: 'サポートされていない画像形式です',
    errGskNotLoggedIn:
      'Genspark にサインインしていません。下の「Genspark にサインイン」からサインインして再試行してください',
    errNoApiKey: '{provider} の API キーが設定されていません',
    errNoModel: 'モデル名が設定されていません',
    errImgAbsPath: '画像パスは絶対パスで指定してください。',
    errImgNotFound: '画像ファイルが見つかりません: {path}',
    errImgTooLarge20: '画像が 20MB を超えているため挿入できません。',
    errImgBadType: 'このファイルは PNG/JPEG/GIF 画像ではありません。',
    errDiskChanged:
      'ブックを開いた後にディスク上で変更されています — 名前を付けて保存を使用してください。',
    autosaveFoundTitle: '自動回復バージョンがあります',
    autosaveFoundBody: '前回のセッションに未保存の変更があります。自動保存版を復元しますか?',
    autosaveRestore: '復元',
    autosaveDiscard: '破棄',
    menuFile: 'ファイル',
    menuOpenWorkbook: 'ブックを開く…',
    menuSave: '保存',
    menuSaveAs: '名前を付けて保存…',
    menuExportPdf: 'PDF をエクスポート…',
    menuClose: '閉じる',
    menuQuit: '終了',
    menuEdit: '編集',
    menuUndo: '元に戻す',
    menuRedo: 'やり直し',
    closeUnsavedMsg: '未保存の変更が {count} 件あります',
    closeUnsavedDetail: '保存せずに閉じると、これらの変更は失われます。',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
    csvSaveAsNotice:
      'CSV 形式は書式を保持できません。.xlsx として保存すると変更をすべて保持できます。',
  },
  ko: {
    filterSpreadsheets: '스프레드시트',
    filterXlsx: 'Excel 통합 문서',
    dlgAddAttachment: '첨부 파일 추가',
    filterSupported: '지원되는 파일',
    filterAll: '모든 파일',
    errUnsupportedExt: '.{ext} 형식은 지원되지 않습니다',
    errNotFile: '파일이 아닙니다',
    errTooLarge: '{mb}MB 제한을 초과했습니다',
    errImageTooLarge: '이미지가 5MB 제한을 초과했습니다',
    errUnreadable: '읽을 수 없습니다',
    errFileTooLarge: '파일이 크기 제한을 초과했습니다',
    errParseFailed: '파일을 구문 분석하지 못했습니다',
    errImageNoText:
      '이미지 첨부에는 텍스트가 없습니다. 이미지는 사용자 메시지와 함께 전송되므로 이미지를 직접 확인하세요',
    errNotImage: '지원되는 이미지 형식이 아닙니다',
    errGskNotLoggedIn:
      'Genspark에 로그인되어 있지 않습니다. 아래 "Genspark 로그인"을 눌러 로그인한 뒤 다시 시도하세요',
    errNoApiKey: '{provider}의 API 키가 설정되지 않았습니다',
    errNoModel: '모델 이름이 설정되지 않았습니다',
    errImgAbsPath: '이미지 경로는 절대 경로여야 합니다.',
    errImgNotFound: '이미지 파일을 찾을 수 없습니다: {path}',
    errImgTooLarge20: '이미지가 20MB를 초과하여 삽입할 수 없습니다.',
    errImgBadType: '이 파일은 PNG/JPEG/GIF 이미지가 아닙니다.',
    errDiskChanged:
      '통합 문서가 열린 후 디스크에서 변경되었습니다. 다른 이름으로 저장을 사용하세요.',
    autosaveFoundTitle: '자동 복구 버전 발견',
    autosaveFoundBody:
      '마지막 세션에 저장되지 않은 변경 내용이 있습니다. 자동 저장 버전을 복원할까요?',
    autosaveRestore: '복원',
    autosaveDiscard: '취소',
    menuFile: '파일',
    menuOpenWorkbook: '통합 문서 열기…',
    menuSave: '저장',
    menuSaveAs: '다른 이름으로 저장…',
    menuExportPdf: 'PDF 내보내기…',
    menuClose: '닫기',
    menuQuit: '끝내기',
    menuEdit: '편집',
    menuUndo: '실행 취소',
    menuRedo: '다시 실행',
    closeUnsavedMsg: '저장하지 않은 변경이 {count}건 있습니다',
    closeUnsavedDetail: '저장하지 않고 닫으면 변경 내용이 손실됩니다.',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
    csvSaveAsNotice:
      'CSV 형식은 서식을 저장할 수 없습니다. .xlsx로 저장하면 모든 변경 내용이 유지됩니다.',
  },
  fr: {
    filterSpreadsheets: 'Feuilles de calcul',
    filterXlsx: 'Classeurs Excel',
    dlgAddAttachment: 'Ajouter des pièces jointes',
    filterSupported: 'Fichiers pris en charge',
    filterAll: 'Tous les fichiers',
    errUnsupportedExt: 'Les fichiers .{ext} ne sont pas pris en charge',
    errNotFile: "n'est pas un fichier",
    errTooLarge: 'dépasse la limite de {mb} Mo',
    errImageTooLarge: "l'image dépasse la limite de 5 Mo",
    errUnreadable: 'illisible',
    errFileTooLarge: 'Le fichier dépasse la taille limite',
    errParseFailed: "Échec de l'analyse du fichier",
    errImageNoText:
      "Les images jointes n'ont pas de texte ; l'image est envoyée avec le message de l'utilisateur",
    errNotImage: "type d'image non pris en charge",
    errGskNotLoggedIn:
      'Non connecté à Genspark : cliquez sur « Se connecter à Genspark » ci-dessous, connectez-vous puis réessayez',
    errNoApiKey: 'Aucune clé API configurée pour {provider}',
    errNoModel: 'Aucun nom de modèle configuré',
    errImgAbsPath: "Le chemin de l'image doit être absolu.",
    errImgNotFound: 'Fichier image introuvable : {path}',
    errImgTooLarge20: "L'image dépasse 20 Mo et ne peut pas être insérée.",
    errImgBadType: "Ce fichier n'est pas une image PNG/JPEG/GIF.",
    errDiskChanged:
      'Le classeur a été modifié sur le disque après son ouverture — utilisez Enregistrer sous.',
    autosaveFoundTitle: 'Version récupérée trouvée',
    autosaveFoundBody:
      'Des modifications non enregistrées existent. Restaurer la version auto-enregistrée ?',
    autosaveRestore: 'Restaurer',
    autosaveDiscard: 'Ignorer',
    menuFile: 'Fichier',
    menuOpenWorkbook: 'Ouvrir un classeur…',
    menuSave: 'Enregistrer',
    menuSaveAs: 'Enregistrer sous…',
    menuExportPdf: 'Exporter en PDF…',
    menuClose: 'Fermer',
    menuQuit: 'Quitter',
    menuEdit: 'Édition',
    menuUndo: 'Annuler',
    menuRedo: 'Rétablir',
    closeUnsavedMsg: '{count} modification(s) non enregistrée(s)',
    closeUnsavedDetail: 'Vos modifications seront perdues si vous fermez sans enregistrer.',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
    csvSaveAsNotice:
      'Le format CSV ne conserve pas la mise en forme — enregistrez en .xlsx pour conserver toutes vos modifications.',
  },
  de: {
    filterSpreadsheets: 'Tabellenkalkulationen',
    filterXlsx: 'Excel-Arbeitsmappen',
    dlgAddAttachment: 'Anlagen hinzufügen',
    filterSupported: 'Unterstützte Dateien',
    filterAll: 'Alle Dateien',
    errUnsupportedExt: '.{ext}-Dateien werden nicht unterstützt',
    errNotFile: 'keine Datei',
    errTooLarge: 'überschreitet das Limit von {mb} MB',
    errImageTooLarge: 'Bild überschreitet das Limit von 5 MB',
    errUnreadable: 'kann nicht gelesen werden',
    errFileTooLarge: 'Datei überschreitet die Größenbeschränkung',
    errParseFailed: 'Datei konnte nicht analysiert werden',
    errImageNoText:
      'Bildanlagen enthalten keinen Text; das Bild wird zusammen mit der Benutzernachricht gesendet',
    errNotImage: 'kein unterstützter Bildtyp',
    errGskNotLoggedIn:
      'Nicht bei Genspark angemeldet: Klicken Sie unten auf „Bei Genspark anmelden“, melden Sie sich an und versuchen Sie es erneut',
    errNoApiKey: 'Kein API-Schlüssel für {provider} konfiguriert',
    errNoModel: 'Kein Modellname konfiguriert',
    errImgAbsPath: 'Der Bildpfad muss absolut sein.',
    errImgNotFound: 'Bilddatei nicht gefunden: {path}',
    errImgTooLarge20: 'Das Bild überschreitet 20 MB und kann nicht eingefügt werden.',
    errImgBadType: 'Die Datei ist kein PNG/JPEG/GIF-Bild.',
    errDiskChanged:
      'Die Arbeitsmappe wurde nach dem Öffnen auf dem Datenträger geändert — verwenden Sie stattdessen „Speichern unter“.',
    autosaveFoundTitle: 'Wiederhergestellte Version gefunden',
    autosaveFoundBody:
      'Es gibt ungespeicherte Änderungen. Automatisch gespeicherte Version wiederherstellen?',
    autosaveRestore: 'Wiederherstellen',
    autosaveDiscard: 'Verwerfen',
    menuFile: 'Datei',
    menuOpenWorkbook: 'Arbeitsmappe öffnen…',
    menuSave: 'Speichern',
    menuSaveAs: 'Speichern unter…',
    menuExportPdf: 'PDF exportieren…',
    menuClose: 'Schließen',
    menuQuit: 'Beenden',
    menuEdit: 'Bearbeiten',
    menuUndo: 'Rückgängig',
    menuRedo: 'Wiederholen',
    closeUnsavedMsg: '{count} nicht gespeicherte Änderung(en)',
    closeUnsavedDetail: 'Ihre Änderungen gehen verloren, wenn Sie ohne Speichern schließen.',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
    csvSaveAsNotice:
      'CSV-Dateien können keine Formatierung speichern – als .xlsx speichern, um alle Änderungen zu behalten.',
  },
  es: {
    filterSpreadsheets: 'Hojas de cálculo',
    filterXlsx: 'Libros de Excel',
    dlgAddAttachment: 'Agregar datos adjuntos',
    filterSupported: 'Archivos compatibles',
    filterAll: 'Todos los archivos',
    errUnsupportedExt: 'Los archivos .{ext} no son compatibles',
    errNotFile: 'no es un archivo',
    errTooLarge: 'supera el límite de {mb} MB',
    errImageTooLarge: 'la imagen supera el límite de 5 MB',
    errUnreadable: 'no se puede leer',
    errFileTooLarge: 'El archivo supera el límite de tamaño',
    errParseFailed: 'No se pudo analizar el archivo',
    errImageNoText:
      'Las imágenes adjuntas no tienen texto; la imagen se envía junto con el mensaje del usuario',
    errNotImage: 'no es un tipo de imagen compatible',
    errGskNotLoggedIn:
      'No has iniciado sesión en Genspark: pulsa «Iniciar sesión en Genspark» abajo, inicia sesión y vuelve a intentarlo',
    errNoApiKey: 'No hay clave de API configurada para {provider}',
    errNoModel: 'No hay nombre de modelo configurado',
    errImgAbsPath: 'La ruta de la imagen debe ser absoluta.',
    errImgNotFound: 'No se encontró el archivo de imagen: {path}',
    errImgTooLarge20: 'La imagen supera los 20 MB y no se puede insertar.',
    errImgBadType: 'El archivo no es una imagen PNG/JPEG/GIF.',
    errDiskChanged: 'El libro cambió en el disco después de abrirse; usa Guardar como en su lugar.',
    autosaveFoundTitle: 'Se encontró una versión recuperada',
    autosaveFoundBody:
      'Hay cambios sin guardar de la última sesión. ¿Restaurar la versión autoguardada?',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    menuFile: 'Archivo',
    menuOpenWorkbook: 'Abrir libro…',
    menuSave: 'Guardar',
    menuSaveAs: 'Guardar como…',
    menuExportPdf: 'Exportar a PDF…',
    menuClose: 'Cerrar',
    menuQuit: 'Salir',
    menuEdit: 'Edición',
    menuUndo: 'Deshacer',
    menuRedo: 'Rehacer',
    closeUnsavedMsg: '{count} cambio(s) sin guardar',
    closeUnsavedDetail: 'Los cambios se perderán si cierras sin guardar.',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
    csvSaveAsNotice:
      'El formato CSV no conserva el formato: guarda como .xlsx para conservar todos tus cambios.',
  },
  th: {
    filterSpreadsheets: 'สเปรดชีต',
    filterXlsx: 'เวิร์กบุ๊ก Excel',
    dlgAddAttachment: 'เพิ่มสิ่งที่แนบ',
    filterSupported: 'ไฟล์ที่รองรับ',
    filterAll: 'ไฟล์ทั้งหมด',
    errUnsupportedExt: 'ไม่รองรับไฟล์ชนิด .{ext}',
    errNotFile: 'ไม่ใช่ไฟล์',
    errTooLarge: 'เกินขีดจำกัด {mb}MB',
    errImageTooLarge: 'รูปภาพเกินขีดจำกัด 5MB',
    errUnreadable: 'อ่านไม่ได้',
    errFileTooLarge: 'ไฟล์มีขนาดเกินขีดจำกัด',
    errParseFailed: 'แยกวิเคราะห์ไฟล์ไม่สำเร็จ',
    errImageNoText:
      'รูปภาพแนบไม่มีข้อความ รูปภาพจะถูกส่งไปพร้อมข้อความของผู้ใช้ ให้ดูที่รูปภาพโดยตรง',
    errNotImage: 'ไม่ใช่ชนิดรูปภาพที่รองรับ',
    errGskNotLoggedIn:
      'ยังไม่ได้ลงชื่อเข้าใช้ Genspark: แตะ “ลงชื่อเข้าใช้ Genspark” ด้านล่าง แล้วลองอีกครั้ง',
    errNoApiKey: 'ยังไม่ได้ตั้งค่า API Key ของ {provider}',
    errNoModel: 'ยังไม่ได้กำหนดชื่อโมเดล',
    errImgAbsPath: 'เส้นทางรูปภาพต้องเป็นเส้นทางแบบสัมบูรณ์',
    errImgNotFound: 'ไม่พบไฟล์รูปภาพ: {path}',
    errImgTooLarge20: 'รูปภาพเกิน 20MB ไม่สามารถแทรกได้',
    errImgBadType: 'ไฟล์นี้ไม่ใช่รูปภาพ PNG/JPEG/GIF',
    errDiskChanged: 'เวิร์กบุ๊กถูกเปลี่ยนแปลงบนดิสก์หลังจากเปิด — โปรดใช้บันทึกเป็นแทน',
    autosaveFoundTitle: 'พบเวอร์ชันกู้คืนอัตโนมัติ',
    autosaveFoundBody: 'มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึกจากครั้งก่อน ต้องการกู้คืนหรือไม่?',
    autosaveRestore: 'กู้คืน',
    autosaveDiscard: 'ละทิ้ง',
    menuFile: 'ไฟล์',
    menuOpenWorkbook: 'เปิดเวิร์กบุ๊ก…',
    menuSave: 'บันทึก',
    menuSaveAs: 'บันทึกเป็น…',
    menuExportPdf: 'ส่งออก PDF…',
    menuClose: 'ปิด',
    menuQuit: 'ออก',
    menuEdit: 'แก้ไข',
    menuUndo: 'เลิกทำ',
    menuRedo: 'ทำซ้ำ',
    closeUnsavedMsg: 'มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก {count} รายการ',
    closeUnsavedDetail: 'หากปิดโดยไม่บันทึก การเปลี่ยนแปลงเหล่านี้จะหายไป',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
    csvSaveAsNotice:
      'ไฟล์ CSV ไม่สามารถเก็บการจัดรูปแบบได้ — บันทึกเป็น .xlsx เพื่อเก็บการเปลี่ยนแปลงทั้งหมดของคุณ',
  },
  id: {
    filterSpreadsheets: 'Lembar bentang',
    filterXlsx: 'Buku kerja Excel',
    dlgAddAttachment: 'Tambahkan lampiran',
    filterSupported: 'File yang didukung',
    filterAll: 'Semua file',
    errUnsupportedExt: 'File .{ext} tidak didukung',
    errNotFile: 'bukan file',
    errTooLarge: 'melebihi batas {mb}MB',
    errImageTooLarge: 'gambar melebihi batas 5MB',
    errUnreadable: 'tidak dapat dibaca',
    errFileTooLarge: 'File melebihi batas ukuran',
    errParseFailed: 'Gagal mengurai file',
    errImageNoText: 'Lampiran gambar tidak memiliki teks; gambar dikirim bersama pesan pengguna',
    errNotImage: 'bukan jenis gambar yang didukung',
    errGskNotLoggedIn: 'Belum masuk ke Genspark: klik “Masuk ke Genspark” di bawah, lalu coba lagi',
    errNoApiKey: 'API Key untuk {provider} belum dikonfigurasi',
    errNoModel: 'Nama model belum dikonfigurasi',
    errImgAbsPath: 'Jalur gambar harus berupa jalur absolut.',
    errImgNotFound: 'File gambar tidak ditemukan: {path}',
    errImgTooLarge20: 'Gambar melebihi 20MB dan tidak dapat disisipkan.',
    errImgBadType: 'File ini bukan gambar PNG/JPEG/GIF.',
    errDiskChanged: 'Buku kerja berubah di disk setelah dibuka — gunakan Simpan Sebagai.',
    autosaveFoundTitle: 'Versi pemulihan ditemukan',
    autosaveFoundBody:
      'Ada perubahan yang belum disimpan dari sesi terakhir. Pulihkan versi tersimpan otomatis?',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    menuFile: 'File',
    menuOpenWorkbook: 'Buka Buku Kerja…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuExportPdf: 'Ekspor PDF…',
    menuClose: 'Tutup',
    menuQuit: 'Keluar',
    menuEdit: 'Edit',
    menuUndo: 'Urungkan',
    menuRedo: 'Ulangi',
    closeUnsavedMsg: '{count} perubahan belum disimpan',
    closeUnsavedDetail: 'Perubahan Anda akan hilang jika menutup tanpa menyimpan.',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    csvSaveAsNotice:
      'File CSV tidak dapat menyimpan pemformatan — simpan sebagai .xlsx untuk mempertahankan semua perubahan Anda.',
  },
  ru: {
    filterSpreadsheets: 'Электронные таблицы',
    filterXlsx: 'Книги Excel',
    dlgAddAttachment: 'Добавить вложения',
    filterSupported: 'Поддерживаемые файлы',
    filterAll: 'Все файлы',
    errUnsupportedExt: 'Файлы .{ext} не поддерживаются',
    errNotFile: 'не является файлом',
    errTooLarge: 'превышает лимит {mb} МБ',
    errImageTooLarge: 'изображение превышает лимит 5 МБ',
    errUnreadable: 'не удаётся прочитать',
    errFileTooLarge: 'Файл превышает предельный размер',
    errParseFailed: 'Не удалось разобрать файл',
    errImageNoText:
      'Вложенные изображения не содержат текста; изображение отправляется вместе с сообщением пользователя',
    errNotImage: 'неподдерживаемый тип изображения',
    errGskNotLoggedIn:
      'Вы не вошли в Genspark: нажмите «Войти в Genspark» ниже, войдите и повторите попытку',
    errNoApiKey: 'API-ключ для {provider} не настроен',
    errNoModel: 'Имя модели не настроено',
    errImgAbsPath: 'Путь к изображению должен быть абсолютным.',
    errImgNotFound: 'Файл изображения не найден: {path}',
    errImgTooLarge20: 'Изображение превышает 20 МБ и не может быть вставлено.',
    errImgBadType: 'Этот файл не является изображением PNG/JPEG/GIF.',
    errDiskChanged: 'Книга была изменена на диске после открытия — используйте «Сохранить как».',
    autosaveFoundTitle: 'Найдена восстановленная версия',
    autosaveFoundBody:
      'Есть несохранённые изменения из прошлого сеанса. Восстановить автосохранённую версию?',
    autosaveRestore: 'Восстановить',
    autosaveDiscard: 'Отклонить',
    menuFile: 'Файл',
    menuOpenWorkbook: 'Открыть книгу…',
    menuSave: 'Сохранить',
    menuSaveAs: 'Сохранить как…',
    menuExportPdf: 'Экспорт в PDF…',
    menuClose: 'Закрыть',
    menuQuit: 'Выход',
    menuEdit: 'Правка',
    menuUndo: 'Отменить',
    menuRedo: 'Повторить',
    closeUnsavedMsg: 'Несохранённых изменений: {count}',
    closeUnsavedDetail: 'Если закрыть без сохранения, эти изменения будут потеряны.',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
    csvSaveAsNotice:
      'Формат CSV не сохраняет форматирование — сохраните в .xlsx, чтобы не потерять изменения.',
  },
  ar: {
    filterSpreadsheets: 'جداول البيانات',
    filterXlsx: 'مصنفات Excel',
    dlgAddAttachment: 'إضافة مرفقات',
    filterSupported: 'الملفات المدعومة',
    filterAll: 'كل الملفات',
    errUnsupportedExt: 'ملفات .{ext} غير مدعومة',
    errNotFile: 'ليس ملفًا',
    errTooLarge: 'يتجاوز الحد البالغ {mb} ميغابايت',
    errImageTooLarge: 'الصورة تتجاوز الحد البالغ 5 ميغابايت',
    errUnreadable: 'تعذّرت قراءته',
    errFileTooLarge: 'الملف يتجاوز حد الحجم',
    errParseFailed: 'فشل تحليل الملف',
    errImageNoText: 'مرفقات الصور لا تحتوي على نص؛ تُرسل الصورة مع رسالة المستخدم',
    errNotImage: 'نوع صورة غير مدعوم',
    errGskNotLoggedIn:
      'لم تسجّل الدخول إلى Genspark: انقر على «تسجيل الدخول إلى Genspark» أدناه ثم أعد المحاولة',
    errNoApiKey: 'لم يتم تكوين مفتاح API لـ {provider}',
    errNoModel: 'لم يتم تكوين اسم النموذج',
    errImgAbsPath: 'يجب أن يكون مسار الصورة مسارًا مطلقًا.',
    errImgNotFound: 'لم يتم العثور على ملف الصورة: {path}',
    errImgTooLarge20: 'الصورة تتجاوز 20 ميغابايت ولا يمكن إدراجها.',
    errImgBadType: 'هذا الملف ليس صورة PNG/JPEG/GIF.',
    errDiskChanged: 'تم تغيير المصنف على القرص بعد فتحه — استخدم «حفظ باسم» بدلاً من ذلك.',
    autosaveFoundTitle: 'تم العثور على نسخة مستردة',
    autosaveFoundBody:
      'توجد تغييرات غير محفوظة من الجلسة الأخيرة. هل تريد استعادة النسخة المحفوظة تلقائيًا؟',
    autosaveRestore: 'استعادة',
    autosaveDiscard: 'تجاهل',
    menuFile: 'ملف',
    menuOpenWorkbook: 'فتح مصنف…',
    menuSave: 'حفظ',
    menuSaveAs: 'حفظ باسم…',
    menuExportPdf: 'تصدير PDF…',
    menuClose: 'إغلاق',
    menuQuit: 'إنهاء',
    menuEdit: 'تحرير',
    menuUndo: 'تراجع',
    menuRedo: 'إعادة',
    closeUnsavedMsg: 'يوجد {count} من التغييرات غير المحفوظة',
    closeUnsavedDetail: 'ستفقد هذه التغييرات إذا أغلقت دون حفظ.',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
    csvSaveAsNotice: 'ملفات CSV لا تحتفظ بالتنسيق — احفظ بصيغة ‎.xlsx للاحتفاظ بجميع تغييراتك.',
  },
  pt: {
    filterSpreadsheets: 'Planilhas',
    filterXlsx: 'Pastas de Trabalho do Excel',
    dlgAddAttachment: 'Adicionar Anexos',
    filterSupported: 'Arquivos Compatíveis',
    filterAll: 'Todos os Arquivos',
    errUnsupportedExt: 'arquivos .{ext} não são suportados',
    errNotFile: 'não é um arquivo',
    errTooLarge: 'excede o limite de {mb}MB',
    errImageTooLarge: 'a imagem excede o limite de 5MB',
    errUnreadable: 'não é possível ler',
    errFileTooLarge: 'O arquivo excede o limite de tamanho',
    errParseFailed: 'Falha ao analisar o arquivo',
    errImageNoText:
      'Anexos de imagem não têm texto; a imagem é enviada junto com a mensagem do usuário',
    errNotImage: 'não é um tipo de imagem suportado',
    errGskNotLoggedIn:
      'Não conectado ao Genspark: clique em “Entrar no Genspark” abaixo, entre e tente novamente',
    errNoApiKey: 'Nenhuma chave de API configurada para {provider}',
    errNoModel: 'Nenhum nome de modelo configurado',
    errImgAbsPath: 'O caminho da imagem deve ser absoluto.',
    errImgNotFound: 'Arquivo de imagem não encontrado: {path}',
    errImgTooLarge20: 'A imagem excede 20MB e não pode ser inserida.',
    errImgBadType: 'O arquivo não é uma imagem PNG/JPEG/GIF.',
    errDiskChanged: 'A pasta de trabalho foi alterada no disco após ser aberta — use Salvar Como.',
    autosaveFoundTitle: 'Versão recuperada encontrada',
    autosaveFoundBody:
      'Há alterações não salvas da sua última sessão. Restaurar a versão salva automaticamente?',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    menuFile: 'Arquivo',
    menuOpenWorkbook: 'Abrir Pasta de Trabalho…',
    menuSave: 'Salvar',
    menuSaveAs: 'Salvar Como…',
    menuExportPdf: 'Exportar PDF…',
    menuClose: 'Fechar',
    menuQuit: 'Sair',
    menuEdit: 'Editar',
    menuUndo: 'Desfazer',
    menuRedo: 'Refazer',
    closeUnsavedMsg: '{count} alteração(ões) não salva(s)',
    closeUnsavedDetail: 'Suas alterações serão perdidas se você fechar sem salvar.',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
    csvSaveAsNotice:
      'Arquivos CSV não mantêm a formatação — salve como .xlsx para manter todas as suas alterações.',
  },
  it: {
    filterSpreadsheets: 'Fogli di calcolo',
    filterXlsx: 'Cartelle di lavoro di Excel',
    dlgAddAttachment: 'Aggiungi allegati',
    filterSupported: 'File supportati',
    filterAll: 'Tutti i file',
    errUnsupportedExt: 'i file .{ext} non sono supportati',
    errNotFile: 'non è un file',
    errTooLarge: 'supera il limite di {mb} MB',
    errImageTooLarge: "l'immagine supera il limite di 5 MB",
    errUnreadable: 'impossibile leggere',
    errFileTooLarge: 'Il file supera il limite di dimensione',
    errParseFailed: 'Impossibile analizzare il file',
    errImageNoText:
      "Gli allegati immagine non hanno testo; l'immagine viene inviata insieme al messaggio dell'utente",
    errNotImage: 'tipo di immagine non supportato',
    errGskNotLoggedIn:
      'Accesso a Genspark non effettuato: fai clic su “Accedi a Genspark” qui sotto, accedi e riprova',
    errNoApiKey: 'Nessuna chiave API configurata per {provider}',
    errNoModel: 'Nessun nome di modello configurato',
    errImgAbsPath: "Il percorso dell'immagine deve essere assoluto.",
    errImgNotFound: 'File immagine non trovato: {path}',
    errImgTooLarge20: "L'immagine supera i 20 MB e non può essere inserita.",
    errImgBadType: "Il file non è un'immagine PNG/JPEG/GIF.",
    errDiskChanged:
      "La cartella di lavoro è stata modificata sul disco dopo l'apertura — usa Salva con nome.",
    autosaveFoundTitle: 'Trovata versione recuperata',
    autosaveFoundBody:
      "Ci sono modifiche non salvate dall'ultima sessione. Ripristinare la versione salvata automaticamente?",
    autosaveRestore: 'Ripristina',
    autosaveDiscard: 'Ignora',
    menuFile: 'File',
    menuOpenWorkbook: 'Apri cartella di lavoro…',
    menuSave: 'Salva',
    menuSaveAs: 'Salva con nome…',
    menuExportPdf: 'Esporta PDF…',
    menuClose: 'Chiudi',
    menuQuit: 'Esci',
    menuEdit: 'Modifica',
    menuUndo: 'Annulla',
    menuRedo: 'Ripeti',
    closeUnsavedMsg: '{count} modifica/e non salvata/e',
    closeUnsavedDetail: 'Le modifiche andranno perse se chiudi senza salvare.',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
    csvSaveAsNotice:
      'I file CSV non conservano la formattazione: salva come .xlsx per mantenere tutte le modifiche.',
  },
  pl: {
    filterSpreadsheets: 'Arkusze kalkulacyjne',
    filterXlsx: 'Skoroszyty programu Excel',
    dlgAddAttachment: 'Dodaj załączniki',
    filterSupported: 'Obsługiwane pliki',
    filterAll: 'Wszystkie pliki',
    errUnsupportedExt: 'pliki .{ext} nie są obsługiwane',
    errNotFile: 'to nie jest plik',
    errTooLarge: 'przekracza limit {mb} MB',
    errImageTooLarge: 'obraz przekracza limit 5 MB',
    errUnreadable: 'nie można odczytać',
    errFileTooLarge: 'Plik przekracza limit rozmiaru',
    errParseFailed: 'Nie udało się przeanalizować pliku',
    errImageNoText:
      'Załączniki graficzne nie zawierają tekstu; obraz jest wysyłany razem z wiadomością użytkownika',
    errNotImage: 'nieobsługiwany typ obrazu',
    errGskNotLoggedIn:
      'Nie zalogowano do Genspark: kliknij „Zaloguj się do Genspark” poniżej, zaloguj się i spróbuj ponownie',
    errNoApiKey: 'Nie skonfigurowano klucza API dla {provider}',
    errNoModel: 'Nie skonfigurowano nazwy modelu',
    errImgAbsPath: 'Ścieżka obrazu musi być bezwzględna.',
    errImgNotFound: 'Nie znaleziono pliku obrazu: {path}',
    errImgTooLarge20: 'Obraz przekracza 20 MB i nie może zostać wstawiony.',
    errImgBadType: 'Plik nie jest obrazem PNG/JPEG/GIF.',
    errDiskChanged: 'Skoroszyt został zmieniony na dysku po otwarciu — użyj polecenia Zapisz jako.',
    autosaveFoundTitle: 'Znaleziono odzyskaną wersję',
    autosaveFoundBody:
      'Istnieją niezapisane zmiany z ostatniej sesji. Przywrócić wersję zapisaną automatycznie?',
    autosaveRestore: 'Przywróć',
    autosaveDiscard: 'Odrzuć',
    menuFile: 'Plik',
    menuOpenWorkbook: 'Otwórz skoroszyt…',
    menuSave: 'Zapisz',
    menuSaveAs: 'Zapisz jako…',
    menuExportPdf: 'Eksportuj PDF…',
    menuClose: 'Zamknij',
    menuQuit: 'Zakończ',
    menuEdit: 'Edycja',
    menuUndo: 'Cofnij',
    menuRedo: 'Ponów',
    closeUnsavedMsg: 'Niezapisane zmiany: {count}',
    closeUnsavedDetail: 'Zmiany zostaną utracone, jeśli zamkniesz bez zapisywania.',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
    csvSaveAsNotice:
      'Pliki CSV nie zachowują formatowania — zapisz jako .xlsx, aby zachować wszystkie zmiany.',
  },
  nl: {
    filterSpreadsheets: 'Spreadsheets',
    filterXlsx: 'Excel-werkmappen',
    dlgAddAttachment: 'Bijlagen toevoegen',
    filterSupported: 'Ondersteunde bestanden',
    filterAll: 'Alle bestanden',
    errUnsupportedExt: '.{ext}-bestanden worden niet ondersteund',
    errNotFile: 'geen bestand',
    errTooLarge: 'overschrijdt de limiet van {mb} MB',
    errImageTooLarge: 'afbeelding overschrijdt de limiet van 5 MB',
    errUnreadable: 'kan niet worden gelezen',
    errFileTooLarge: 'Bestand overschrijdt de maximale grootte',
    errParseFailed: 'Kan bestand niet parseren',
    errImageNoText:
      'Afbeeldingsbijlagen bevatten geen tekst; de afbeelding wordt samen met het gebruikersbericht verzonden',
    errNotImage: 'geen ondersteund afbeeldingstype',
    errGskNotLoggedIn:
      'Niet aangemeld bij Genspark: klik hieronder op “Aanmelden bij Genspark”, meld u aan en probeer het opnieuw',
    errNoApiKey: 'Geen API-sleutel geconfigureerd voor {provider}',
    errNoModel: 'Geen modelnaam geconfigureerd',
    errImgAbsPath: 'Het afbeeldingspad moet absoluut zijn.',
    errImgNotFound: 'Afbeeldingsbestand niet gevonden: {path}',
    errImgTooLarge20: 'De afbeelding is groter dan 20 MB en kan niet worden ingevoegd.',
    errImgBadType: 'Het bestand is geen PNG/JPEG/GIF-afbeelding.',
    errDiskChanged:
      'De werkmap is op de schijf gewijzigd nadat deze was geopend — gebruik Opslaan als.',
    autosaveFoundTitle: 'Herstelde versie gevonden',
    autosaveFoundBody:
      'Er zijn niet-opgeslagen wijzigingen van uw laatste sessie. De automatisch opgeslagen versie herstellen?',
    autosaveRestore: 'Herstellen',
    autosaveDiscard: 'Negeren',
    menuFile: 'Bestand',
    menuOpenWorkbook: 'Werkmap openen…',
    menuSave: 'Opslaan',
    menuSaveAs: 'Opslaan als…',
    menuExportPdf: 'PDF exporteren…',
    menuClose: 'Sluiten',
    menuQuit: 'Stoppen',
    menuEdit: 'Bewerken',
    menuUndo: 'Ongedaan maken',
    menuRedo: 'Opnieuw',
    closeUnsavedMsg: '{count} niet-opgeslagen wijziging(en)',
    closeUnsavedDetail: 'Uw wijzigingen gaan verloren als u sluit zonder op te slaan.',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
    csvSaveAsNotice:
      'CSV-bestanden bewaren geen opmaak — sla op als .xlsx om al uw wijzigingen te behouden.',
  },
  ms: {
    filterSpreadsheets: 'Hamparan',
    filterXlsx: 'Buku Kerja Excel',
    dlgAddAttachment: 'Tambah Lampiran',
    filterSupported: 'Fail yang Disokong',
    filterAll: 'Semua Fail',
    errUnsupportedExt: 'fail .{ext} tidak disokong',
    errNotFile: 'bukan fail',
    errTooLarge: 'melebihi had {mb}MB',
    errImageTooLarge: 'imej melebihi had 5MB',
    errUnreadable: 'tidak dapat dibaca',
    errFileTooLarge: 'Fail melebihi had saiz',
    errParseFailed: 'Gagal menghurai fail',
    errImageNoText: 'Lampiran imej tiada teks; imej dihantar bersama mesej pengguna',
    errNotImage: 'bukan jenis imej yang disokong',
    errGskNotLoggedIn:
      'Belum log masuk ke Genspark: klik “Log masuk ke Genspark” di bawah, kemudian cuba lagi',
    errNoApiKey: 'Kunci API untuk {provider} belum dikonfigurasikan',
    errNoModel: 'Nama model belum dikonfigurasikan',
    errImgAbsPath: 'Laluan imej mestilah laluan mutlak.',
    errImgNotFound: 'Fail imej tidak ditemui: {path}',
    errImgTooLarge20: 'Imej melebihi 20MB dan tidak boleh disisipkan.',
    errImgBadType: 'Fail ini bukan imej PNG/JPEG/GIF.',
    errDiskChanged: 'Buku kerja telah diubah pada cakera selepas dibuka — gunakan Simpan Sebagai.',
    autosaveFoundTitle: 'Versi pulihan ditemui',
    autosaveFoundBody:
      'Terdapat perubahan yang belum disimpan daripada sesi terakhir anda. Pulihkan versi yang disimpan secara automatik?',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    menuFile: 'Fail',
    menuOpenWorkbook: 'Buka Buku Kerja…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuExportPdf: 'Eksport PDF…',
    menuClose: 'Tutup',
    menuQuit: 'Keluar',
    menuEdit: 'Edit',
    menuUndo: 'Buat Asal',
    menuRedo: 'Buat Semula',
    closeUnsavedMsg: '{count} perubahan belum disimpan',
    closeUnsavedDetail: 'Perubahan anda akan hilang jika anda menutup tanpa menyimpan.',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    csvSaveAsNotice:
      'Fail CSV tidak dapat menyimpan pemformatan — simpan sebagai .xlsx untuk mengekalkan semua perubahan anda.',
  },
  he: {
    filterSpreadsheets: 'גיליונות אלקטרוניים',
    filterXlsx: 'חוברות עבודה של Excel',
    dlgAddAttachment: 'הוספת קבצים מצורפים',
    filterSupported: 'קבצים נתמכים',
    filterAll: 'כל הקבצים',
    errUnsupportedExt: 'קובצי .{ext} אינם נתמכים',
    errNotFile: 'אינו קובץ',
    errTooLarge: 'חורג מהמגבלה של {mb}MB',
    errImageTooLarge: 'התמונה חורגת מהמגבלה של 5MB',
    errUnreadable: 'לא ניתן לקרוא',
    errFileTooLarge: 'הקובץ חורג ממגבלת הגודל',
    errParseFailed: 'ניתוח הקובץ נכשל',
    errImageNoText: 'קבצים מצורפים מסוג תמונה אינם מכילים טקסט; התמונה נשלחת יחד עם הודעת המשתמש',
    errNotImage: 'סוג תמונה שאינו נתמך',
    errGskNotLoggedIn: 'לא מחובר ל-Genspark: לחץ על "התחבר ל-Genspark" למטה, התחבר ונסה שוב',
    errNoApiKey: 'לא הוגדר מפתח API עבור {provider}',
    errNoModel: 'לא הוגדר שם מודל',
    errImgAbsPath: 'נתיב התמונה חייב להיות מוחלט.',
    errImgNotFound: 'קובץ התמונה לא נמצא: {path}',
    errImgTooLarge20: 'התמונה חורגת מ-20MB ולא ניתן להוסיף אותה.',
    errImgBadType: 'הקובץ אינו תמונת PNG/JPEG/GIF.',
    errDiskChanged: 'חוברת העבודה השתנתה בדיסק לאחר פתיחתה — השתמש בשמירה בשם.',
    autosaveFoundTitle: 'נמצאה גרסה משוחזרת',
    autosaveFoundBody: 'קיימים שינויים שלא נשמרו מהפעלה הקודמת. לשחזר את הגרסה שנשמרה אוטומטית?',
    autosaveRestore: 'שחזר',
    autosaveDiscard: 'התעלם',
    menuFile: 'קובץ',
    menuOpenWorkbook: 'פתיחת חוברת עבודה…',
    menuSave: 'שמירה',
    menuSaveAs: 'שמירה בשם…',
    menuExportPdf: 'ייצוא PDF…',
    menuClose: 'סגירה',
    menuQuit: 'יציאה',
    menuEdit: 'עריכה',
    menuUndo: 'בטל',
    menuRedo: 'בצע שוב',
    closeUnsavedMsg: '{count} שינויים שלא נשמרו',
    closeUnsavedDetail: 'השינויים שלך יאבדו אם תסגור בלי לשמור.',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
    csvSaveAsNotice: 'קובצי CSV אינם שומרים עיצוב — שמרו כ‑.xlsx כדי לשמור על כל השינויים.',
  },
  hi: {
    filterSpreadsheets: 'स्प्रेडशीट',
    filterXlsx: 'Excel कार्यपुस्तिकाएँ',
    dlgAddAttachment: 'अनुलग्नक जोड़ें',
    filterSupported: 'समर्थित फ़ाइलें',
    filterAll: 'सभी फ़ाइलें',
    errUnsupportedExt: '.{ext} फ़ाइलें समर्थित नहीं हैं',
    errNotFile: 'फ़ाइल नहीं है',
    errTooLarge: '{mb}MB की सीमा से अधिक है',
    errImageTooLarge: 'छवि 5MB की सीमा से अधिक है',
    errUnreadable: 'पढ़ा नहीं जा सकता',
    errFileTooLarge: 'फ़ाइल आकार सीमा से अधिक है',
    errParseFailed: 'फ़ाइल पार्स करने में विफल',
    errImageNoText: 'छवि अनुलग्नक में टेक्स्ट नहीं होता; छवि उपयोगकर्ता संदेश के साथ भेजी जाती है',
    errNotImage: 'समर्थित छवि प्रकार नहीं है',
    errGskNotLoggedIn:
      'Genspark में साइन इन नहीं है: नीचे “Genspark में साइन इन करें” पर क्लिक करें, साइन इन करें और फिर से कोशिश करें',
    errNoApiKey: '{provider} के लिए कोई API कुंजी कॉन्फ़िगर नहीं है',
    errNoModel: 'कोई मॉडल नाम कॉन्फ़िगर नहीं है',
    errImgAbsPath: 'छवि पथ निरपेक्ष होना चाहिए।',
    errImgNotFound: 'छवि फ़ाइल नहीं मिली: {path}',
    errImgTooLarge20: 'छवि 20MB से अधिक है और सम्मिलित नहीं की जा सकती।',
    errImgBadType: 'यह फ़ाइल PNG/JPEG/GIF छवि नहीं है।',
    errDiskChanged:
      'खोले जाने के बाद कार्यपुस्तिका डिस्क पर बदल गई — इसके बजाय इस रूप में सहेजें का उपयोग करें।',
    autosaveFoundTitle: 'पुनर्प्राप्त संस्करण मिला',
    autosaveFoundBody:
      'आपके पिछले सत्र से सहेजे नहीं गए परिवर्तन हैं। स्वतः सहेजा गया संस्करण पुनर्स्थापित करें?',
    autosaveRestore: 'पुनर्स्थापित करें',
    autosaveDiscard: 'छोड़ें',
    menuFile: 'फ़ाइल',
    menuOpenWorkbook: 'कार्यपुस्तिका खोलें…',
    menuSave: 'सहेजें',
    menuSaveAs: 'इस रूप में सहेजें…',
    menuExportPdf: 'PDF निर्यात करें…',
    menuClose: 'बंद करें',
    menuQuit: 'बाहर निकलें',
    menuEdit: 'संपादन',
    menuUndo: 'पूर्ववत करें',
    menuRedo: 'फिर से करें',
    closeUnsavedMsg: '{count} सहेजे नहीं गए परिवर्तन',
    closeUnsavedDetail: 'यदि आप बिना सहेजे बंद करते हैं तो आपके परिवर्तन खो जाएँगे।',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
    csvSaveAsNotice:
      'CSV फ़ाइलें फ़ॉर्मेटिंग सहेज नहीं सकतीं — सभी बदलाव बनाए रखने के लिए .xlsx के रूप में सहेजें।',
  },
  'zh-TW': {
    filterSpreadsheets: '電子試算表',
    filterXlsx: 'Excel 活頁簿',
    dlgAddAttachment: '新增附件',
    filterSupported: '支援的檔案',
    filterAll: '所有檔案',
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
    errImgAbsPath: '圖片路徑必須是絕對路徑。',
    errImgNotFound: '找不到圖片檔案: {path}',
    errImgTooLarge20: '圖片超過 20MB,不支援插入。',
    errImgBadType: '該檔案不是 PNG/JPEG/GIF 圖片。',
    errDiskChanged: '活頁簿在開啟後被磁碟上的變更覆蓋——請改用另存新檔。',
    autosaveFoundTitle: '發現自動復原版本',
    autosaveFoundBody: '上次工作階段有未儲存的變更。要復原自動儲存的版本嗎?',
    autosaveRestore: '復原',
    autosaveDiscard: '放棄',
    menuFile: '檔案',
    menuOpenWorkbook: '開啟活頁簿…',
    menuSave: '儲存',
    menuSaveAs: '另存新檔…',
    menuExportPdf: '匯出 PDF…',
    menuClose: '關閉',
    menuQuit: '結束',
    menuEdit: '編輯',
    menuUndo: '復原',
    menuRedo: '重做',
    closeUnsavedMsg: '有 {count} 處未儲存的修改',
    closeUnsavedDetail: '不儲存直接關閉,這些修改將遺失。',
    btnDontSave: '不儲存',
    btnCancel: '取消',
    csvSaveAsNotice: 'CSV 格式不保留樣式等格式修改——另存為 .xlsx 可保留全部內容。',
  },
})
const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(getUiLang(), key, params)

interface SessionInfo {
  readonly path: string
  readonly sha256: string
  readonly sheetNames: ReadonlyMap<string, string>
  /// Set when the session opened a converted copy (.xls/.csv import): the
  /// first save routes through Save As, defaulting to this .xlsx path.
  readonly suggestSaveAs?: string
  /// The converted copy came from a CSV: the Save As dialog explains that
  /// formatting requires .xlsx (CSV keeps values only).
  readonly csvImport?: boolean
}

// ---- runtime configuration (paths differ when bundled into the shell) ----

interface SheetsRuntimeConfig {
  /** absolute path to the sheets preload bundle */
  preloadPath: string
  /** dev-server URL for the sheets renderer (wins over rendererFile) */
  rendererUrl?: string | undefined
  /** absolute path to the built sheets renderer index.html */
  rendererFile: string
  /** absolute path to the Rust xlsx-sidecar binary */
  sidecarPath?: string | undefined
}

let runtime: SheetsRuntimeConfig = {
  preloadPath: join(__dirname, '../preload/index.js'),
  rendererUrl: process.env.ELECTRON_RENDERER_URL,
  rendererFile: join(__dirname, '../renderer/index.html'),
}

export function configureSheetsRuntime(config: SheetsRuntimeConfig): void {
  runtime = config
}

let mainWindow: BrowserWindow | null = null
let sidecar: XlsxSidecarClient | null = null

/** the single real BrowserWindow hosting the tab strip, used as dialog parent in tab mode */
let sheetsShellWindow: BrowserWindow | null = null
export function setSheetsShellWindow(win: BrowserWindow | null): void {
  sheetsShellWindow = win
}

interface SheetsTabSession {
  readonly webContents: WebContents
  readonly client: XlsxSidecarClient
  readonly sessions: Map<string, SessionInfo>
  readonly aiStreams: Map<string, AbortController>
}

/** per-tab session state, keyed by webContents.id — replaces the old single-window closures
 * that `registerIpcHandlers`/`validateSender` used to capture, which broke as soon as a second
 * tab (or a closed-then-reopened tab) registered and overwrote the previous closure. */
const sheetsTabs = new Map<number, SheetsTabSession>()
let activeSheetsWebContents: WebContents | null = null

function sessionFor(event: IpcMainInvokeEvent): SheetsTabSession {
  const entry = sheetsTabs.get(event.sender.id)
  if (!entry) throw new Error('Untrusted IPC sender.')
  return entry
}

function dialogParent(event: IpcMainInvokeEvent): BrowserWindow | undefined {
  return sheetsShellWindow ?? BrowserWindow.fromWebContents(event.sender) ?? undefined
}

async function openFileDialog(event: IpcMainInvokeEvent, options: OpenDialogOptions) {
  return showOpenDialogWithMemory(dialog, dialogParent(event), options)
}

async function saveFileDialog(event: IpcMainInvokeEvent, options: SaveDialogOptions) {
  // before any pick is remembered, bare-name suggestions anchor in the
  // configurable default save folder instead of Electron's Downloads pin
  return showSaveDialogWithMemory(
    dialog,
    dialogParent(event),
    options,
    configuredDefaultSaveDir(app),
  )
}

/** register a tab's webContents/client pair and wire up cleanup on teardown */
function registerSheetsSession(webContents: WebContents, client: XlsxSidecarClient): void {
  sheetsTabs.set(webContents.id, { webContents, client, sessions: new Map(), aiStreams: new Map() })
  activeSheetsWebContents = webContents
  webContents.once('destroyed', () => {
    const entry = sheetsTabs.get(webContents.id)
    sheetsTabs.delete(webContents.id)
    if (entry) void closeAllSessions(entry)
    if (activeSheetsWebContents === webContents) activeSheetsWebContents = null
  })
}

export function getSheetsWindow(): BrowserWindow | null {
  return mainWindow
}

/** the webContents of whichever sheets tab most recently registered or activated */
export function getActiveSheetsWebContents(): WebContents | null {
  return activeSheetsWebContents
}

/** shell tab switching keeps menu actions routed at the visible sheets tab */
export function setActiveSheetsWebContents(wc: WebContents | null): void {
  activeSheetsWebContents = wc
}

/** Shell notification: an open view's file was renamed on disk (renamed in the
 *  Home list) — sync the matching session's path in that tab (later saves write
 *  the new file) and push the renderer to update the title-bar file name. */
export function sheetsFileRenamed(wc: WebContents, oldPath: string, newPath: string): void {
  // A user-chosen name always wins: the file no longer qualifies for auto-rename
  untitledWorkbookPaths.delete(oldPath)
  const entry = sheetsTabs.get(wc.id)
  if (!entry) return
  let matched = false
  for (const [id, session] of entry.sessions) {
    if (session.path !== oldPath) continue
    entry.sessions.set(id, { ...session, path: newPath })
    matched = true
  }
  if (matched) wc.send(IPC_CHANNELS.workbookRenamed, basename(newPath))
}

/**
 * Workbooks the shell pre-created on disk with the localized untitled name
 * ("New Spreadsheet"). Only these ever qualify for the content-derived
 * auto-rename after an AI run; any manual rename removes the mark.
 */
const untitledWorkbookPaths = new Set<string>()
export function markSheetsUntitledPath(path: string): void {
  untitledWorkbookPaths.add(path)
}

/** Sanitize an AI-provided sheet name into a safe filename base: strip illegal path chars, collapse whitespace, cap length; null if invalid. (Mirrors slides' draft naming.) */
function sanitizeAutoRenameBase(raw: string): string | null {
  const cleaned = raw
    // eslint-disable-next-line no-control-regex -- stripping control chars is the point here
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .trim()
  if (!cleaned) return null
  return cleaned.length > 40 ? cleaned.slice(0, 40).trim() : cleaned
}

/** shell hook: a tab opened a workbook (dialog or queued path) — used for tab titles/dedupe */
let workbookOpenedHook: ((wc: WebContents, path: string) => void) | null = null
export function setSheetsWorkbookOpenedHook(
  fn: ((wc: WebContents, path: string) => void) | null,
): void {
  workbookOpenedHook = fn
}

/** forward an application-menu File command into the sheets renderer */
export function sendSheetsMenuAction(
  action: 'open' | 'save' | 'save-as' | 'export-pdf' | 'undo' | 'redo',
): void {
  activeSheetsWebContents?.send(IPC_CHANNELS.menuAction, action)
}

// ---- AI settings persistence (main process avoids renderer CORS for the chat/stream proxy) ----

function userDataPath(...parts: string[]): string {
  return join(app.getPath('userData'), ...parts)
}

// ── Crash recovery ──────────────────────────────────────────
// A dirty renderer asks for a recovery copy every 30s; it is written through the
// normal save pipeline (writeWorkbookTo) to a userData path, so it is a real .xlsx.
// A successful save removes it; opening a file whose copy is newer offers Restore.
const recoveryDir = () => userDataPath('sheets-autosave')
const recoveryPathFor = (filePath: string) =>
  join(recoveryDir(), `${createHash('sha1').update(filePath).digest('hex').slice(0, 16)}.xlsx`)

function clearWorkbookRecovery(filePath: string): void {
  try {
    unlinkSync(recoveryPathFor(filePath))
  } catch {
    /* nothing to clean */
  }
}

/** Recovery copy newer than the file itself, i.e. unsaved work from a lost session. */
function pendingRecoveryFor(filePath: string): string | null {
  const copy = recoveryPathFor(filePath)
  try {
    if (!existsSync(copy)) return null
    if (statSync(copy).mtimeMs <= statSync(filePath).mtimeMs) {
      unlinkSync(copy)
      return null
    }
    return copy
  } catch {
    return null
  }
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

const SETTINGS_PATH = () => userDataPath('ai-settings.json')

// Dev-only automation hooks: a fixed CDP port for driving the app from test
// scripts, and a workbook path that bypasses the native file dialog.
const debugPort = app.isPackaged ? undefined : process.env.XLSX_DEBUG_PORT
if (debugPort) app.commandLine.appendSwitch('remote-debugging-port', debugPort)
let forcedWorkbookPath = app.isPackaged ? undefined : process.env.XLSX_OPEN_PATH
/** true while a shell-queued path is waiting to be consumed (dev env/capture-server
 * paths stay sticky; shell-queued ones are one-shot so a later Open shows the dialog) */
let shellQueuedWorkbook = false

/** queue a workbook the next selectWorkbook call opens without a dialog (shell routing) */
export function setForcedWorkbookPath(path: string | undefined): void {
  forcedWorkbookPath = path
  shellQueuedWorkbook = path !== undefined
}

/** still waiting for the renderer to consume a shell-queued workbook? */
export function hasQueuedWorkbook(): boolean {
  return shellQueuedWorkbook
}

/** set by shell for home:new-sheet: renderer opens blank workbook instead of demo */
let pendingNewBlank = false

/** signal the next sheets renderer to open a new blank workbook (shell mode only) */
export function setSheetsNewBlank(): void {
  pendingNewBlank = true
}

// capturePage forces a renderer frame even when the window is occluded or on
// another Space, unlike CDP Page.captureScreenshot / macOS screencapture.
function startCaptureServer(): void {
  if (!debugPort) return
  const server = createServer((request, response) => {
    const url = new URL(request.url ?? '/', 'http://127.0.0.1')
    if (url.pathname === '/open') {
      forcedWorkbookPath = url.searchParams.get('path') ?? undefined
      response.writeHead(200)
      response.end('ok')
      return
    }
    // Drives the File menu from test scripts: CDP input can't reach native
    // menu accelerators, and osascript focus-stealing is flaky.
    if (url.pathname === '/menu') {
      const action = url.searchParams.get('action')
      if (
        action === 'open' ||
        action === 'save' ||
        action === 'save-as' ||
        action === 'export-pdf' ||
        action === 'undo' ||
        action === 'redo'
      ) {
        sendSheetsMenuAction(action)
        response.writeHead(200)
        response.end('ok')
      } else {
        response.writeHead(400)
        response.end('unknown action')
      }
      return
    }
    const webContents = getActiveSheetsWebContents()
    if (url.pathname !== '/capture' || !webContents) {
      response.writeHead(404)
      response.end()
      return
    }
    webContents
      .capturePage()
      .then((image) => {
        response.writeHead(200, { 'Content-Type': 'image/png' })
        response.end(image.toPNG())
      })
      .catch((error: unknown) => {
        response.writeHead(500)
        response.end(String(error))
      })
  })
  server.listen(Number(debugPort) + 1, '127.0.0.1')
}

const sidecarOpenResultSchema = workbookFileSchema.omit({
  sha256: true,
  readOnly: true,
})

export async function createSheetsWindow(
  options: { includeAiHandlers?: boolean } = {},
): Promise<BrowserWindow> {
  const client = sidecar ?? new XlsxSidecarClient(resolveSidecarPath())
  sidecar = client
  client.start()
  const window = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 680,
    show: false,
    title: 'GenOffice Sheets',
    // Traffic lights sit inside the toolbar row.
    ...(process.platform === 'darwin' ? { titleBarStyle: 'hiddenInset' as const } : {}),
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })
  mainWindow = window
  registerSheetsIpc()
  if (options.includeAiHandlers ?? true) registerSheetsAiIpc()
  if (options.includeAiHandlers ?? true) registerProjectIpc()
  registerSheetsSession(window.webContents, client)
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  window.webContents.on('will-navigate', (event) => event.preventDefault())
  if (!app.isPackaged) {
    window.webContents.on('console-message', (details) => {
      process.stderr.write(`[renderer:${details.level}] ${details.message}\n`)
    })
  }
  window.once('ready-to-show', () => window.show())
  window.on('close', (event) => {
    if (sheetsPendingEditCount(window.webContents.id) === 0) return
    event.preventDefault()
    void requestSheetsClose(window.webContents, window).then((proceed) => {
      // destroy() skips this handler on the way out (close() would re-enter
      // with the count possibly still non-zero after a discard).
      if (proceed && !window.isDestroyed()) window.destroy()
    })
  })
  window.on('closed', () => {
    mainWindow = null
  })

  if (runtime.rendererUrl) {
    await window.loadURL(runtime.rendererUrl)
  } else {
    await window.loadFile(runtime.rendererFile)
  }
  return window
}

/** tab-mode equivalent of createSheetsWindow: same runtime/IPC wiring, no BrowserWindow of its own. */
export function createSheetsView(options: { includeAiHandlers?: boolean } = {}): WebContentsView {
  const client = sidecar ?? new XlsxSidecarClient(resolveSidecarPath())
  sidecar = client
  client.start()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })
  registerSheetsIpc()
  if (options.includeAiHandlers ?? true) registerSheetsAiIpc()
  registerSheetsSession(view.webContents, client)
  view.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  view.webContents.on('will-navigate', (event) => event.preventDefault())
  if (!app.isPackaged) {
    view.webContents.on('console-message', (details) => {
      process.stderr.write(`[renderer:${details.level}] ${details.message}\n`)
    })
  }
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
  return view
}

// ---- Chat attachments: local files parsed and fed to the agent (copied from
// the apps/docs docs-main attachment pipeline) ----

const ATTACHMENT_MAX_BYTES = 50 * 1024 * 1024
/** Plain-text extensions, read as UTF-8 */
const ATTACHMENT_TEXT_EXTS = new Set([
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
/** office/pdf formats extract text via @genoffice/file-parse; images skip text
 * extraction and go multimodal (sheets:files-read-image) */
const ATTACHMENT_EXTS = new Set([
  ...ATTACHMENT_TEXT_EXTS,
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
/** Multimodal cap per image attachment (protects the context window) */
const ATTACHMENT_IMAGE_MAX_BYTES = 5 * 1024 * 1024

/** Extracted-text cache keyed by path; invalidated when mtime+size change */
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

/** Persists clipboard-pasted image bytes to a temp file (screenshots/bitmaps
 * without a local path); returns null for non-images or empty data */
let pastedImageSeq = 0
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
  const stamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '').replace('T', '-')
  const filePath = join(dir, `pasted-${stamp}-${++pastedImageSeq}.${cleanExt}`)
  writeFileSync(filePath, bytes)
  return filePath
}

/** Attachment text extraction via @genoffice/file-parse (docx/pdf/pptx/xlsx/plain text) */
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
  // The cache is bounded (keeping a few recent files is enough)
  if (attachmentTextCache.size > 8) {
    const oldest = attachmentTextCache.keys().next().value
    if (oldest) attachmentTextCache.delete(oldest)
  }
  return parsed.text
}

// Close guard: the renderer mirrors its pending-save count here, used to show a
// save confirmation before closing the window/tab.
const pendingEditCounts = new Map<number, number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
const trackedEditSenders = new Set<number>()

export function sheetsPendingEditCount(webContentsId: number): number {
  return pendingEditCounts.get(webContentsId) ?? 0
}

/**
 * Close guard for a sheets renderer: true means proceed with the close.
 * Clean → true; dirty → Save/Don't Save/Cancel dialog. Save asks the renderer to run
 * its journal save and waits for the outcome — a failed or canceled save
 * keeps the window open (the renderer already surfaced the error).
 */
/**
 * The app is shutting down (quit menu, SIGTERM from a restart/installer/killall,
 * SIGINT from a terminal). The close guard must not save then: nobody answered the
 * prompt, and a dialog raised during shutdown resolves to its default button, which
 * silently overwrote the user's original file. Unsaved work is covered
 * by the 30s recovery copy instead — the next launch offers to restore it.
 */
let appShuttingDown = false

export function markSheetsShuttingDown(): void {
  appShuttingDown = true
}

app.on('before-quit', markSheetsShuttingDown)
for (const signal of ['SIGTERM', 'SIGINT'] as const) {
  process.on(signal, () => {
    appShuttingDown = true
    app.quit()
  })
}

export async function requestSheetsClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  const count = pendingEditCounts.get(contents.id) ?? 0
  const decision = closeGuardDecision({
    pendingEdits: count,
    destroyed: contents.isDestroyed(),
    shuttingDown: appShuttingDown,
  })
  if (decision === 'proceed') return true
  const options = {
    // On macOS 'warning' shows the system warning triangle + app-icon badge
    type: 'warning' as const,
    message: tm('closeUnsavedMsg', { count }),
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
  if (response === 1) return true
  // The window went away (or a quit started) while the prompt was up: don't save
  if (appShuttingDown || contents.isDestroyed()) return true
  return await new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(IPC_CHANNELS.closeSaveRequest)
  })
}

let coreIpcRegistered = false

export function registerSheetsIpc(): void {
  if (coreIpcRegistered) return
  coreIpcRegistered = true

  ipcMain.on(IPC_CHANNELS.pendingEditsChanged, (event, count: unknown) => {
    if (typeof count !== 'number' || !Number.isFinite(count) || count < 0) return
    const senderId = event.sender.id
    pendingEditCounts.set(senderId, Math.floor(count))
    if (!trackedEditSenders.has(senderId)) {
      trackedEditSenders.add(senderId)
      event.sender.once('destroyed', () => {
        trackedEditSenders.delete(senderId)
        pendingEditCounts.delete(senderId)
        closeSaveWaiters.get(senderId)?.(false)
        closeSaveWaiters.delete(senderId)
      })
    }
  })

  ipcMain.on(IPC_CHANNELS.closeSaveResult, (event, ok: unknown) => {
    const waiter = closeSaveWaiters.get(event.sender.id)
    if (!waiter) return
    closeSaveWaiters.delete(event.sender.id)
    waiter(ok === true)
  })

  // shared with the other editor modules — last (identical) registration wins
  ipcMain.removeHandler('app:get-language')
  ipcMain.handle('app:get-language', () => getUiLang())

  /** returns true once when shell opened this tab for a new blank workbook */
  ipcMain.handle('sheets:consume-new-blank', () => {
    if (pendingNewBlank) {
      pendingNewBlank = false
      return true
    }
    return false
  })

  /**
   * Is a shell-queued workbook still waiting to be opened? The shell's 'open'
   * nudge loop gives up after 30s; on slow dev cold starts (vite compiles the
   * renderer on demand) Univer mounts later than that and the queued path
   * would strand the tab as a blank in-memory workbook. The renderer polls
   * this once it is ready and triggers the open itself.
   */
  ipcMain.handle('sheets:has-queued-workbook', () => hasQueuedWorkbook())

  ipcMain.handle(IPC_CHANNELS.selectWorkbook, async (event) => {
    const entry = sessionFor(event)
    let path = forcedWorkbookPath
    if (shellQueuedWorkbook) {
      // consume immediately (before the slow session open) so the shell's
      // retry loop stops re-sending 'open' for the same file
      forcedWorkbookPath = undefined
      shellQueuedWorkbook = false
    }
    if (!path) {
      const selection = await openFileDialog(event, {
        properties: ['openFile'],
        filters: [{ name: tm('filterSpreadsheets'), extensions: ['xlsx', 'xls', 'csv'] }],
      })
      if (selection.canceled || !selection.filePaths[0]) return null
      path = selection.filePaths[0]
    }
    const prepared = await prepareWorkbookForOpen(entry.client, path, dialogParent(event))
    const result = await openWorkbookSession(
      entry.client,
      prepared.openPath,
      entry.sessions,
      prepared.suggestSaveAs,
      prepared.csvImport,
    )
    if (result) workbookOpenedHook?.(event.sender, path)
    return result
  })

  ipcMain.handle(IPC_CHANNELS.readWorkbookRange, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookRangeRequestSchema.parse(input)
    if (!entry.sessions.has(request.sessionId)) throw new Error('Unknown workbook session.')
    const result = await entry.client.readRange(request)
    return workbookRangeResultSchema.parse(result)
  })

  ipcMain.handle(IPC_CHANNELS.readWorkbookFormulas, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookFormulaCellsRequestSchema.parse(input)
    if (!entry.sessions.has(request.sessionId)) throw new Error('Unknown workbook session.')
    const result = await entry.client.readFormulaCells(request)
    return workbookFormulaCellsResultSchema.parse(result)
  })

  // IronCalc recalculation: sheet ids resolve through the session's file
  // sheet names, so the renderer never sees paths and sheets added this
  // session (no file part) fail closed before reaching the engine.
  const sidecarRecalcResultSchema = z
    .object({
      cells: z.array(
        z
          .object({
            sheet: z.string(),
            row: z.number().int().nonnegative(),
            column: z.number().int().nonnegative(),
            formatted: z.string(),
            number: z.number().optional(),
            isFormula: z.boolean(),
          })
          .strict(),
      ),
      cached: z.boolean().optional(),
    })
    .strict()
  ipcMain.handle(IPC_CHANNELS.recalcWorkbook, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookRecalcRequestSchema.parse(input)
    const session = entry.sessions.get(request.sessionId)
    if (!session) throw new Error('Unknown workbook session.')
    const fileSheetName = (sheetId: string): string => {
      const name = session.sheetNames.get(sheetId)
      if (name === undefined) throw new Error(`Unknown sheet for recalculation: ${sheetId}`)
      return name
    }
    const result = sidecarRecalcResultSchema.parse(
      await entry.client.recalcCells({
        path: session.path,
        edits: request.edits.map((edit) => ({
          sheet: fileSheetName(edit.sheetId),
          row: edit.row,
          column: edit.column,
          input: edit.input,
        })),
        reads: request.reads.map((read) => ({
          sheet: fileSheetName(read.sheetId),
          range: read.range,
        })),
      }),
    )
    const idsByName = new Map([...session.sheetNames].map(([id, name]) => [name, id]))
    return workbookRecalcResultSchema.parse({
      cells: result.cells.flatMap((cell) => {
        const sheetId = idsByName.get(cell.sheet)
        if (sheetId === undefined) return []
        return [
          {
            sheetId,
            row: cell.row,
            column: cell.column,
            formatted: cell.formatted,
            ...(cell.number === undefined ? {} : { number: cell.number }),
            isFormula: cell.isFormula,
          },
        ]
      }),
    })
  })

  ipcMain.handle(IPC_CHANNELS.readWorkbookMedia, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookMediaRequestSchema.parse(input)
    if (!entry.sessions.has(request.sessionId)) throw new Error('Unknown workbook session.')
    const result = await entry.client.readMedia(request)
    return workbookMediaResultSchema.parse(result)
  })

  function sniffImageType(bytes: Buffer): 'image/png' | 'image/jpeg' | 'image/gif' | null {
    if (
      bytes.length >= 8 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47
    )
      return 'image/png'
    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
      return 'image/jpeg'
    }
    if (bytes.length >= 6 && bytes.subarray(0, 4).toString('latin1') === 'GIF8') return 'image/gif'
    return null
  }

  ipcMain.handle(IPC_CHANNELS.readLocalImage, async (event, input: unknown) => {
    sessionFor(event)
    const request = localImageRequestSchema.parse(input)
    const resolved = request.path.startsWith('~/')
      ? join(app.getPath('home'), request.path.slice(2))
      : request.path
    if (!isAbsolute(resolved)) throw new Error(tm('errImgAbsPath'))
    const info = await stat(resolved).catch(() => null)
    if (!info?.isFile()) throw new Error(tm('errImgNotFound', { path: request.path }))
    if (info.size > 20 * 1024 * 1024) throw new Error(tm('errImgTooLarge20'))
    const bytes = await readFile(resolved)
    const mediaType = sniffImageType(bytes)
    if (mediaType === null) {
      throw new Error(tm('errImgBadType'))
    }
    return localImageResultSchema.parse({ mediaType, base64: bytes.toString('base64') })
  })

  ipcMain.handle(IPC_CHANNELS.captureScreenSources, async (event) => {
    sessionFor(event)
    // macOS gates desktopCapturer behind the Screen Recording permission and
    // returns black frames instead of failing; surface a real denied state.
    if (process.platform === 'darwin') {
      const status = systemPreferences.getMediaAccessStatus('screen')
      if (status !== 'granted' && status !== 'not-determined') {
        return screenSourcesResultSchema.parse({ status: 'denied', sources: [] })
      }
    }
    const sources = await desktopCapturer.getSources({
      types: ['screen', 'window'],
      thumbnailSize: { width: 320, height: 200 },
      fetchWindowIcons: false,
    })
    if (
      process.platform === 'darwin' &&
      systemPreferences.getMediaAccessStatus('screen') !== 'granted'
    ) {
      return screenSourcesResultSchema.parse({ status: 'denied', sources: [] })
    }
    // In tab mode the sheets renderer is a WebContentsView, so fromWebContents
    // on the sender is null; the shell window is the one to exclude.
    const selfWindow = sheetsShellWindow ?? BrowserWindow.fromWebContents(event.sender)
    const selfId = selfWindow?.getMediaSourceId()
    return screenSourcesResultSchema.parse({
      status: 'ok',
      sources: sources
        .filter((source) => source.id !== selfId)
        .map((source) => ({
          id: source.id,
          name: source.name,
          kind: source.id.startsWith('screen') ? 'screen' : 'window',
          thumbnail: source.thumbnail.isEmpty() ? '' : source.thumbnail.toDataURL(),
        })),
    })
  })

  ipcMain.handle(IPC_CHANNELS.captureScreenSource, async (event, input: unknown) => {
    sessionFor(event)
    const request = screenCaptureRequestSchema.parse(input)
    // desktopCapturer only ever returns thumbnails, so a full-res capture is
    // a re-listing with the thumbnail sized to the largest physical display.
    const displays = screen.getAllDisplays()
    const captureSize = {
      width: Math.min(
        4096,
        Math.max(1920, ...displays.map((d) => Math.ceil(d.size.width * d.scaleFactor))),
      ),
      height: Math.min(
        4096,
        Math.max(1080, ...displays.map((d) => Math.ceil(d.size.height * d.scaleFactor))),
      ),
    }
    const sources = await desktopCapturer.getSources({
      types: ['screen', 'window'],
      thumbnailSize: captureSize,
      fetchWindowIcons: false,
    })
    const source = sources.find((candidate) => candidate.id === request.id)
    if (!source || source.thumbnail.isEmpty()) return null
    let image = source.thumbnail
    let png = image.toPNG()
    if (png.length > 20 * 1024 * 1024) {
      image = image.resize({ width: Math.round(image.getSize().width / 2) })
      png = image.toPNG()
    }
    const { width, height } = image.getSize()
    return screenCaptureResultSchema.parse({
      mediaType: 'image/png',
      base64: png.toString('base64'),
      width,
      height,
    })
  })

  ipcMain.handle(IPC_CHANNELS.readPivotDefinition, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookPivotRequestSchema.parse(input)
    const session = entry.sessions.get(request.sessionId)
    if (!session) throw new Error('Unknown workbook session.')
    const [pivotXml, cacheXml] = await Promise.all([
      readArchiveEntryText(entry.client, session.path, request.path),
      readArchiveEntryText(entry.client, session.path, request.cachePath),
    ])
    return workbookPivotDefinitionSchema.parse(parsePivotDefinition(pivotXml, cacheXml))
  })

  ipcMain.handle(IPC_CHANNELS.exportPdf, async (event, input: unknown) => {
    sessionFor(event)
    const request = workbookExportPdfRequestSchema.parse(input)
    return exportPdf(event, request)
  })

  ipcMain.handle(IPC_CHANNELS.saveWorkbook, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const client = entry.client
    const request = workbookSaveRequestSchema.parse(input)
    const session = entry.sessions.get(request.sessionId)
    if (!session) throw new Error('Unknown workbook session.')

    let targetPath = session.path
    // Converted imports (.xls/.csv) never save silently over the temp copy —
    // the first save always asks where the .xlsx should live.
    if (request.mode === 'save-as' || session.suggestSaveAs !== undefined) {
      const selection = await saveFileDialog(event, {
        defaultPath: session.suggestSaveAs ?? session.path,
        filters: [{ name: tm('filterXlsx'), extensions: ['xlsx'] }],
        // CSV import: explain why the save goes through .xlsx (CSV keeps values only)
        ...(session.csvImport
          ? { title: tm('csvSaveAsNotice'), message: tm('csvSaveAsNotice') }
          : {}),
      })
      if (selection.canceled || !selection.filePath) return { canceled: true }
      targetPath = selection.filePath.endsWith('.xlsx')
        ? selection.filePath
        : `${selection.filePath}.xlsx`
    }

    if ((await sha256File(session.path)) !== session.sha256) {
      throw new Error(tm('errDiskChanged'))
    }

    const mutation = await writeWorkbookTo(client, session, request, targetPath)

    // The sidecar session still streams the pre-save bytes; swap it for a
    // fresh session over the saved file so future reads match the disk state.
    entry.sessions.delete(request.sessionId)
    await client.close(request.sessionId).catch(() => undefined)
    const file = await openWorkbookSession(client, targetPath, entry.sessions)
    // Notify shell (if running) so it can update the tab title and record the
    // saved path in recent files (mirrors the open hook; covers Save As + first
    // save after converting an .xls/.csv import).
    workbookOpenedHook?.(event.sender, targetPath)
    // The file on disk now carries these edits
    clearWorkbookRecovery(targetPath)
    if (session.suggestSaveAs !== undefined) clearWorkbookRecovery(session.suggestSaveAs)
    return { canceled: false, file, touchedEntries: mutation.touchedEntries }
  })

  // Crash-recovery copy of a dirty workbook: the same save pipeline with a
  // userData target, no session swap and no dialogs — best-effort, silent on failure.
  ipcMain.handle(IPC_CHANNELS.writeWorkbookRecovery, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookSaveRequestSchema.parse(input)
    const session = entry.sessions.get(request.sessionId)
    // A converted import has no original file to recover into; its temp copy is enough
    if (!session || session.suggestSaveAs !== undefined) return { ok: false }
    try {
      await mkdir(recoveryDir(), { recursive: true })
      await writeWorkbookTo(entry.client, session, request, recoveryPathFor(session.path))
      return { ok: true }
    } catch (error) {
      console.warn('[sheets] recovery copy failed:', error)
      return { ok: false }
    }
  })

  ipcMain.handle(IPC_CHANNELS.closeWorkbook, async (event, sessionId: unknown) => {
    const entry = sessionFor(event)
    const validatedSessionId = z.string().uuid().parse(sessionId)
    if (!entry.sessions.delete(validatedSessionId)) return
    await entry.client.close(validatedSessionId)
  })

  // Content-derived naming for AI-generated workbooks (sheets' analog of slides'
  // deckName): the renderer proposes a base name after an AI run lands; the file
  // is renamed only while it still carries the shell's auto-created untitled name.
  ipcMain.handle(
    IPC_CHANNELS.autoRenameWorkbook,
    (event, sessionId: unknown, baseName: unknown) => {
      const entry = sessionFor(event)
      const validatedSessionId = z.string().uuid().parse(sessionId)
      const session = entry.sessions.get(validatedSessionId)
      if (!session || !untitledWorkbookPaths.has(session.path)) return { renamed: false }
      const base = sanitizeAutoRenameBase(z.string().min(1).max(100).parse(baseName))
      if (!base) return { renamed: false }
      const dir = dirname(session.path)
      let target = join(dir, `${base}.xlsx`)
      for (let i = 2; existsSync(target) && i < 100; i++) target = join(dir, `${base}-${i}.xlsx`)
      if (existsSync(target) || target === session.path) return { renamed: false }
      try {
        renameSync(session.path, target)
      } catch (err) {
        console.warn('[sheets] auto-rename failed:', err)
        return { renamed: false }
      }
      untitledWorkbookPaths.delete(session.path)
      entry.sessions.set(validatedSessionId, { ...session, path: target })
      event.sender.send(IPC_CHANNELS.workbookRenamed, basename(target))
      // Same contract as open/save: shell updates the tab title and recents
      workbookOpenedHook?.(event.sender, target)
      return { renamed: true, name: basename(target) }
    },
  )

  ipcMain.handle(IPC_CHANNELS.openExternal, async (event, url: unknown) => {
    sessionFor(event)
    const validatedUrl = safeExternalUrl(url)
    if (!validatedUrl) {
      throw new Error('Only http(s) links can be opened.')
    }
    await shell.openExternal(validatedUrl)
  })

  // ── Chat attachments (same structure as the docs/slides files:* pipeline) ──

  ipcMain.handle(IPC_CHANNELS.filesPick, async (event): Promise<AttachmentAddResult | null> => {
    sessionFor(event)
    const selection = await openFileDialog(event, {
      title: tm('dlgAddAttachment'),
      filters: [
        { name: tm('filterSupported'), extensions: [...ATTACHMENT_EXTS] },
        { name: tm('filterAll'), extensions: ['*'] },
      ],
      properties: ['openFile', 'multiSelections'],
    })
    if (selection.canceled || selection.filePaths.length === 0) return null
    return collectAttachments(selection.filePaths)
  })

  ipcMain.handle(IPC_CHANNELS.filesAdd, (event, paths: unknown): AttachmentAddResult => {
    sessionFor(event)
    return collectAttachments(z.array(z.string().min(1).max(1024)).max(50).parse(paths))
  })

  ipcMain.handle(
    IPC_CHANNELS.filesRead,
    async (
      event,
      filePath: unknown,
      offset: unknown,
      maxChars: unknown,
    ): Promise<AttachmentReadResult> => {
      sessionFor(event)
      const validatedPath = z.string().min(1).max(1024).parse(filePath)
      const name = basename(validatedPath)
      const ext = name.split('.').pop()?.toLowerCase() ?? ''
      if (!ATTACHMENT_EXTS.has(ext)) return { ok: false, error: tm('errUnsupportedExt', { ext }) }
      if (ATTACHMENT_IMAGE_EXTS.has(ext)) {
        return { ok: false, error: tm('errImageNoText') }
      }
      try {
        const text = await extractAttachmentText(validatedPath)
        const start = Math.max(0, Math.floor(Number(offset)) || 0)
        const size = Math.min(Math.max(1, Math.floor(Number(maxChars)) || 1), 48_000)
        return {
          ok: true,
          name,
          totalChars: text.length,
          offset: start,
          text: text.slice(start, start + size),
        }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  // Image attachments read raw bytes → base64; the renderer puts them into the
  // user message's images for multimodal input
  ipcMain.handle(IPC_CHANNELS.filesReadImage, (event, filePath: unknown): AttachmentImageResult => {
    sessionFor(event)
    const validatedPath = z.string().min(1).max(1024).parse(filePath)
    const name = basename(validatedPath)
    const ext = name.split('.').pop()?.toLowerCase() ?? ''
    const mime = ATTACHMENT_IMAGE_MIME[ext]
    if (!mime) return { ok: false, error: `${name}: ${tm('errNotImage')}` }
    try {
      const stat = statSync(validatedPath)
      if (stat.size > ATTACHMENT_IMAGE_MAX_BYTES) {
        return { ok: false, error: `${name}: ${tm('errImageTooLarge')}` }
      }
      return { ok: true, base64: readFileSync(validatedPath).toString('base64'), mime }
    } catch {
      return { ok: false, error: `${name}: ${tm('errUnreadable')}` }
    }
  })

  // Clipboard-pasted images (screenshots and other bitmaps without a local
  // path): persisted to a temp file, then go through the regular attachment flow
  ipcMain.handle(
    IPC_CHANNELS.filesAddPastedImage,
    (event, data: unknown, ext: unknown): AttachmentAddResult => {
      sessionFor(event)
      const filePath = savePastedImage(data, ext)
      return filePath
        ? collectAttachments([filePath])
        : { accepted: [], rejected: [tm('errNotImage')] }
    },
  )
}

let aiIpcRegistered = false

export function registerSheetsAiIpc(): void {
  if (aiIpcRegistered) return
  aiIpcRegistered = true

  ipcMain.handle(IPC_CHANNELS.aiGetSettings, (event): AiSettings => {
    sessionFor(event)
    const stored = readJson<Partial<AiSettings> & LegacyAiSettings>(SETTINGS_PATH(), {})
    const settings = resolveAiSettings(stored, defaultAiSettings())
    // AI features all go through Genspark (gsk login); legacy settings that chose
    // another provider are reset
    settings.provider = 'genspark'
    return settings
  })

  // Genspark account (gsk login state): the auth source for AI features; the
  // frontend uses it to guide sign-in when logged out
  ipcMain.handle(
    IPC_CHANNELS.aiGskStatus,
    async (_event, withEmail?: unknown): Promise<GenSparkAccountStatus> => {
      if (!hasGskAuth()) return { loggedIn: false }
      if (!withEmail) return { loggedIn: true }
      const info = await gskLoginInfo()
      return info?.email ? { loggedIn: true, email: info.email } : { loggedIn: true }
    },
  )

  ipcMain.handle(IPC_CHANNELS.aiGskLogin, () => {
    ensureGenofficeLogin((url) => void shell.openExternal(url))
  })

  ipcMain.handle(IPC_CHANNELS.aiSetSettings, (event, input: unknown) => {
    sessionFor(event)
    const settings = aiSettingsInputSchema.parse(input)
    writeJson(SETTINGS_PATH(), settings)
  })

  ipcMain.handle(IPC_CHANNELS.aiChat, async (event, input: unknown) => {
    sessionFor(event)
    const request = aiChatRequestSchema.parse(input)
    const provider = request.settings.provider as AiProviderId
    let config = request.settings.providers[provider]
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
      return await chatForProvider(provider, config, request.system, request.user)
    } catch (err) {
      return { ok: false, error: String(err) }
    }
  })

  ipcMain.handle(IPC_CHANNELS.aiStream, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = aiStreamRequestSchema.parse(input)
    const { requestId, system, messages } = request
    const tools = request.tools ?? []
    const maxTokens = request.maxTokens ?? 8192
    const provider = request.settings.provider as AiProviderId
    let config = request.settings.providers[provider]
    // Genspark's key never enters the settings file; it is read from the gsk
    // login state per request
    if (provider === 'genspark' && config && !config.apiKey) {
      config = { ...config, apiKey: gskApiKey() }
    }
    const send = (chunk: AiStreamChunk) => {
      if (!event.sender.isDestroyed()) event.sender.send(IPC_CHANNELS.aiStreamChunk, chunk)
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
    entry.aiStreams.set(requestId, controller)
    // wire-activity keepalive: lets the renderer's silence watchdog tell a slow turn from a dead one
    let lastPing = 0
    const ping = () => {
      const now = Date.now()
      if (now - lastPing < 5_000) return
      lastPing = now
      send({ requestId, type: 'ping' })
    }
    try {
      await streamForProvider(provider, config, system, messages, tools, maxTokens, {
        signal: controller.signal,
        onDelta: (text) => send({ requestId, type: 'delta', text }),
        onToolCall: (toolCall) => send({ requestId, type: 'tool-call', toolCall }),
        onActivity: ping,
      })
      send({ requestId, type: 'done' })
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
      entry.aiStreams.delete(requestId)
    }
  })

  ipcMain.handle(IPC_CHANNELS.aiStreamCancel, (event, requestId: unknown) => {
    const entry = sessionFor(event)
    entry.aiStreams.get(z.string().min(1).parse(requestId))?.abort()
  })

  // Shared search tools (content + images): Serper with DuckDuckGo fallback
  // (same source as slides/docs)
  ipcMain.handle('ai:web-search', async (_event, query: unknown, maxResults?: unknown) => {
    try {
      return await webSearch(
        z.string().parse(query),
        typeof maxResults === 'number' ? maxResults : 6,
      )
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })
  ipcMain.handle('ai:image-search', async (_event, query: unknown, maxResults?: unknown) => {
    try {
      return await imageSearch(
        z.string().parse(query),
        typeof maxResults === 'number' ? maxResults : 8,
      )
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })
}

// ── project-store IPC (standalone mode) ────────────────────────────────────
// In shell mode docs-main.registerProjectIpc has already registered it
// (idempotency guard).

let sheetsProjectStore: ProjectStore | null = null
let sheetsProjectIpcRegistered = false

function getSheetsProjectStore(): ProjectStore {
  if (!sheetsProjectStore) sheetsProjectStore = new ProjectStore(app.getPath('userData'))
  return sheetsProjectStore
}

export function registerProjectIpc(): void {
  if (sheetsProjectIpcRegistered) return
  sheetsProjectIpcRegistered = true

  ipcMain.handle(
    'project:resolveChat',
    (event, args: { filePath: string | null; tempChatId?: string; sessionId?: string }) => {
      const store = getSheetsProjectStore()
      store.ensureDefaultProject()

      // sheets mode: reverse-look up the file path via sessionId
      let resolvedPath = args.filePath
      if (!resolvedPath && args.sessionId) {
        const tabEntry = sheetsTabs.get(event.sender.id)
        if (tabEntry) {
          resolvedPath = tabEntry.sessions.get(args.sessionId)?.path ?? null
        }
      }

      if (!resolvedPath) {
        return { projectId: 'default', chatId: args.tempChatId ?? `unsaved-${Date.now()}` }
      }
      return store.resolveChatForFile(resolvedPath)
    },
  )

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
      const msg: Parameters<ProjectStore['appendChatMessage']>[2] = {
        role: args.role,
        text: args.text,
      }
      if (args.tools) msg.tools = args.tools
      if (args.attachments) msg.attachments = args.attachments
      getSheetsProjectStore().appendChatMessage(args.projectId, args.chatId, msg)
    },
  )

  ipcMain.handle(
    'project:loadChat',
    (_event, args: { projectId: string; chatId: string; limit?: number }) => {
      return getSheetsProjectStore().loadChat(args.projectId, args.chatId, args.limit ?? 200)
    },
  )

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
      const store = getSheetsProjectStore()
      let path = args.newFilePath ?? null
      if (!path && args.sessionId) {
        path = resolveSheetsSessionPath(event.sender.id, args.sessionId)
      }
      if (path) {
        return store.rebindChatToFile(args.projectId, args.tempChatId, path)
      }
      if (args.newChatId) store.rebindChat(args.projectId, args.tempChatId, args.newChatId)
      return { projectId: args.projectId, chatId: args.newChatId ?? args.tempChatId }
    },
  )
}

/**
 * sessionId → workbook file path reverse lookup (injected into docs-main's
 * project:resolveChat in shell mode). In standalone mode the handler registered
 * above queries sheetsTabs directly.
 */
export function resolveSheetsSessionPath(senderId: number, sessionId: string): string | null {
  return sheetsTabs.get(senderId)?.sessions.get(sessionId)?.path ?? null
}

/**
 * Resolve a save request's sheet ops / name mappings and write the workbook through
 * the sidecar. Split out of the save handler so a crash-recovery copy can reuse the
 * exact same pipeline with a different targetPath.
 */
async function writeWorkbookTo(
  client: XlsxSidecarClient,
  session: SessionInfo,
  request: WorkbookSaveRequest,
  targetPath: string,
): Promise<Awaited<ReturnType<typeof saveWorkbookViaSidecar>>> {
  // Sheet ops resolve first: added sheets have Univer ids the session map
  // doesn't know, so cell edits into them resolve through the op's name.
  const addedSheetNames = new Map<string, string>()
  // Added sheet id → file name of the sheet whose part seeds the new part.
  const duplicateSources = new Map<string, string>()
  const renames: { sheetName: string; newName: string }[] = []
  const removals: string[] = []
  const hiddenChanges: { sheetName: string; hidden: boolean }[] = []
  let orderChanged = false
  for (const op of request.sheetOps) {
    if (op.kind === 'add-sheet') {
      addedSheetNames.set(op.sheetId, op.name)
      continue
    }
    if (op.kind === 'duplicate-sheet') {
      // The renderer resolves duplicate chains to a sheet the file knows,
      // so the source must be in the session map.
      const sourceName = session.sheetNames.get(op.sourceSheetId)
      if (!sourceName) throw new Error(`Unknown duplicate source ${op.sourceSheetId}.`)
      addedSheetNames.set(op.sheetId, op.name)
      duplicateSources.set(op.sheetId, sourceName)
      continue
    }
    if (op.kind === 'reorder-sheets') {
      orderChanged = true
      continue
    }
    const sheetName = addedSheetNames.get(op.sheetId) ?? session.sheetNames.get(op.sheetId)
    if (!sheetName) throw new Error(`Unknown worksheet ${op.sheetId}.`)
    if (op.kind === 'rename-sheet') renames.push({ sheetName, newName: op.newName })
    else if (op.kind === 'set-sheet-hidden') {
      hiddenChanges.push({ sheetName, hidden: op.hidden })
    } else removals.push(sheetName)
  }
  const renameByOriginal = new Map(renames.map((rename) => [rename.sheetName, rename.newName]))
  const resolveSheetName = (sheetId: string): string => {
    const sheetName = addedSheetNames.get(sheetId) ?? session.sheetNames.get(sheetId)
    if (!sheetName) throw new Error(`Unknown worksheet ${sheetId}.`)
    return sheetName
  }
  let sheetPlan: SheetEditPlan | undefined
  if (request.sheetOps.length > 0) {
    sheetPlan = {
      renames,
      additions: [...addedSheetNames].map(([sheetId, name]) => ({
        name,
        sourceSheetName: duplicateSources.get(sheetId),
      })),
      removals,
      hiddenChanges,
      orderChanged,
      order: request.sheetOrder.map((sheetId) => {
        const original = resolveSheetName(sheetId)
        return addedSheetNames.has(sheetId)
          ? original
          : (renameByOriginal.get(original) ?? original)
      }),
    }
  }

  const edits: CellEdit[] = request.edits.map((edit) => ({
    sheetName: resolveSheetName(edit.sheetId),
    row: edit.row,
    column: edit.column,
    writeValue: edit.writeValue,
    cell: { value: edit.value, formula: edit.formula },
    style: edit.style,
    rich: edit.rich,
    styleReset: edit.styleReset,
  }))
  const opsBySheet = new Map<string, SheetStructuralOps['ops'][number][]>()
  for (const op of request.structuralOps) {
    const sheetName = resolveSheetName(op.sheetId)
    const sheetOps = opsBySheet.get(sheetName) ?? []
    if ('range' in op) {
      sheetOps.push({ kind: op.kind, range: op.range })
    } else if ('size' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, size: op.size })
    } else if ('level' in op) {
      sheetOps.push({
        kind: op.kind,
        start: op.start,
        end: op.end,
        level: op.level,
        ...(op.collapsed === undefined ? {} : { collapsed: op.collapsed }),
      })
    } else if ('hidden' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, hidden: op.hidden })
    } else if ('before' in op) {
      sheetOps.push({ kind: op.kind, index: op.index, count: op.count, before: op.before })
    } else {
      sheetOps.push({ kind: op.kind, index: op.index, count: op.count })
    }
    opsBySheet.set(sheetName, sheetOps)
  }
  const structuralOps: SheetStructuralOps[] = [...opsBySheet].map(([sheetName, ops]) => ({
    sheetName,
    ops,
  }))
  const filterStates = request.filterStates.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    filter: state.filter,
    hiddenRows: state.hiddenRows,
    visibilityRange: state.visibilityRange,
  }))
  const linksBySheet = new Map<string, { row: number; column: number; target: string | null }[]>()
  for (const link of request.hyperlinkEdits) {
    const sheetName = resolveSheetName(link.sheetId)
    const sheetLinks = linksBySheet.get(sheetName) ?? []
    sheetLinks.push({ row: link.row, column: link.column, target: link.target })
    linksBySheet.set(sheetName, sheetLinks)
  }
  const hyperlinkEdits = [...linksBySheet].map(([sheetName, links]) => ({
    sheetName,
    edits: links,
  }))
  const cfStates = request.cfStates.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    rules: state.rules,
  }))
  const dvStates = request.dvStates.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    rules: state.rules,
  }))
  const sheetProtections = request.sheetProtections.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    protected: state.protected,
  }))
  const pageSetupStates = request.pageSetupStates.map(({ sheetId, ...state }) => ({
    sheetName: resolveSheetName(sheetId),
    ...state,
  }))
  const noteStates = request.noteStates.map(({ sheetId, notes }) => ({
    sheetName: resolveSheetName(sheetId),
    notes,
  }))
  const visualAdditions = request.visualAdditions.map((addition) => ({
    sheetName: resolveSheetName(addition.sheetId),
    anchor: addition.anchor,
    chart: addition.chart,
    shape: addition.shape,
    image: addition.image,
  }))
  const tableAdditions = request.tableAdditions.map((table) => ({
    sheetName: resolveSheetName(table.sheetId),
    area: table.area,
    name: table.name,
    columnNames: table.columnNames,
    style: table.style,
    bandedRows: table.bandedRows,
  }))
  const pivotAdditions = request.pivotAdditions.map((pivot) => ({
    sheetName: resolveSheetName(pivot.sheetId),
    sourceSheetName: resolveSheetName(pivot.sourceSheetId),
    sourceArea: pivot.sourceArea,
    location: pivot.location,
    name: pivot.name,
    fieldNames: pivot.fieldNames,
    rowFieldIndices: pivot.rowFieldIndices,
    columnFieldIndex: pivot.columnFieldIndex,
    pageFieldIndices: pivot.pageFieldIndices,
    rowItems: pivot.rowItems,
    rowLevelItems: pivot.rowLevelItems,
    rowLines: pivot.rowLines,
    columnItems: pivot.columnItems,
    columnFieldIndices: pivot.columnFieldIndices,
    colLevelItems: pivot.colLevelItems,
    colLines: pivot.colLines,
    groupings: pivot.groupings,
    filters: pivot.filters,
    rowHiddenItems: pivot.rowHiddenItems,
    colHiddenItems: pivot.colHiddenItems,
    values: pivot.values,
  }))
  const sparklineAdditions = request.sparklineAdditions.map(({ sheetId, ...group }) => ({
    sheetName: resolveSheetName(sheetId),
    ...group,
  }))
  // Recalculated formula values: sheetId → file sheet name, the same
  // resolution the cell edits use.
  const formulaValuesBySheet = new Map<
    string,
    { row: number; column: number; value: string | number | boolean | null }[]
  >()
  for (const cell of request.formulaValues) {
    const sheetName = resolveSheetName(cell.sheetId)
    const list = formulaValuesBySheet.get(sheetName) ?? []
    list.push({ row: cell.row, column: cell.column, value: cell.value })
    formulaValuesBySheet.set(sheetName, list)
  }
  const formulaValues = [...formulaValuesBySheet].map(([sheetName, cells]) => ({
    sheetName,
    cells,
  }))
  const mutation = await saveWorkbookViaSidecar({
    client,
    sourcePath: session.path,
    targetPath,
    edits,
    structuralOps,
    chartEdits: request.chartEdits,
    // Located by package-absolute drawingPath, so no sheet-name mapping.
    visualEdits: request.visualEdits,
    sheetPlan,
    filterStates,
    hyperlinkEdits,
    cfStates,
    dvStates,
    sheetProtections,
    definedNamesState: request.definedNamesState,
    visualAdditions,
    pageSetupStates,
    noteStates,
    tableAdditions,
    pivotAdditions,
    sparklineAdditions,
    formulaValues,
    pivotCacheRefreshPaths: request.pivotCacheRefreshPaths,
    // Output-area expansion from layout growth: sheetId → sheet name; the part
    // path is resolved by the gateway.
    pivotRefreshUpdates: request.pivotRefreshUpdates.map((update) => ({
      cachePath: update.cachePath,
      sheetName: resolveSheetName(update.sheetId),
      newOutputRef: update.newOutputRef,
      ...(update.relayout === undefined
        ? {}
        : {
            relayout: (({ sheetId: _sheetId, sourceSheetId, ...rest }) => ({
              ...rest,
              sourceSheetName: resolveSheetName(sourceSheetId),
            }))(update.relayout),
          }),
    })),
  })
  return mutation
}

async function openWorkbookSession(
  client: XlsxSidecarClient,
  path: string,
  sessions: Map<string, SessionInfo>,
  suggestSaveAs?: string,
  csvImport?: boolean,
): Promise<WorkbookFile> {
  const [opened, digest] = await Promise.all([
    client.open(path, getUiLang()).then((result) => sidecarOpenResultSchema.parse(result)),
    sha256File(path),
  ])
  sessions.set(opened.sessionId, {
    path,
    sha256: digest,
    sheetNames: new Map(opened.sheets.map((sheet) => [sheet.id, sheet.name])),
    ...(suggestSaveAs === undefined ? {} : { suggestSaveAs }),
    ...(csvImport ? { csvImport } : {}),
  })
  return workbookFileSchema.parse({
    ...opened,
    path,
    sha256: digest,
    readOnly: false,
    needsSaveAs: suggestSaveAs !== undefined,
  })
}

/** which legacy charset an Excel CSV most likely uses, judged by the UI language */
function legacyCsvCharset(): string | undefined {
  const byLang: Partial<Record<Lang, string>> = {
    zh: 'gb18030',
    'zh-TW': 'big5',
    ja: 'shift_jis',
    ko: 'euc-kr',
  }
  return byLang[getUiLang()]
}

/// .xls and .csv open as a converted copy in the temp dir; the session
/// remembers the original's .xlsx sibling as the Save As default.
async function prepareWorkbookForOpen(
  client: XlsxSidecarClient,
  path: string,
  parent?: BrowserWindow | undefined,
): Promise<{ openPath: string; suggestSaveAs?: string; csvImport?: boolean }> {
  const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
  if (extension !== 'csv' && extension !== 'xls') {
    // Unsaved work from a lost session: offer the recovery copy. Restoring
    // opens it with suggestSaveAs pointing back at the original, so the first save asks
    // for confirmation instead of silently overwriting the file the user opened.
    const recovery = pendingRecoveryFor(path)
    if (recovery) {
      const options = {
        type: 'question' as const,
        buttons: [tm('autosaveRestore'), tm('autosaveDiscard')],
        defaultId: 0,
        cancelId: 1,
        message: tm('autosaveFoundTitle'),
        detail: tm('autosaveFoundBody'),
      }
      const answer = parent
        ? await dialog.showMessageBox(parent, options)
        : await dialog.showMessageBox(options)
      if (answer.response === 0) return { openPath: recovery, suggestSaveAs: path }
      clearWorkbookRecovery(path)
    }
    return { openPath: path }
  }
  const stem = basename(path).replace(/\.[^.]+$/, '')
  const directory = join(app.getPath('temp'), 'genoffice-imports', randomUUID())
  await mkdir(directory, { recursive: true })
  const openPath = join(directory, `${stem}.xlsx`)
  if (extension === 'csv') {
    await writeFile(
      openPath,
      await csvToXlsxBuffer(decodeCsvBuffer(await readFile(path), legacyCsvCharset())),
    )
  } else {
    await client.convertWorkbook({ path, targetPath: openPath })
  }
  return {
    openPath,
    suggestSaveAs: path.replace(/\.[^.]+$/, '.xlsx'),
    ...(extension === 'csv' ? { csvImport: true } : {}),
  }
}

/** shell-injected items appended to the File menu (e.g. Back to Home) */
let extraFileMenuItems: MenuItemConstructorOptions[] = []

export function setSheetsExtraFileMenuItems(items: MenuItemConstructorOptions[]): void {
  extraFileMenuItems = items
}

/** tab mode: closes the sheets tab instead of the whole shell window (Cmd+W / role:'close') */
let closeActiveTabHook: (() => void) | null = null
export function setSheetsCloseTabHook(fn: (() => void) | null): void {
  closeActiveTabHook = fn
}

/// The ribbon has no File tab; file commands live in
/// the application menu and are forwarded to the renderer.
function installApplicationMenu(): void {
  const sendMenuAction = sendSheetsMenuAction
  const labels = appMenuLabels(getUiLang())
  Menu.setApplicationMenu(
    Menu.buildFromTemplate([
      ...(process.platform === 'darwin' ? [{ role: 'appMenu' as const }] : []),
      {
        label: tm('menuFile'),
        submenu: [
          {
            label: tm('menuOpenWorkbook'),
            accelerator: 'CmdOrCtrl+O',
            click: () => sendMenuAction('open'),
          },
          ...(extraFileMenuItems.length > 0
            ? [{ type: 'separator' as const }, ...extraFileMenuItems]
            : []),
          { type: 'separator' },
          {
            label: tm('menuSave'),
            accelerator: 'CmdOrCtrl+S',
            click: () => sendMenuAction('save'),
          },
          {
            label: tm('menuSaveAs'),
            accelerator: 'Shift+CmdOrCtrl+S',
            click: () => sendMenuAction('save-as'),
          },
          {
            label: tm('menuExportPdf'),
            click: () => sendMenuAction('export-pdf'),
          },
          { type: 'separator' },
          closeActiveTabHook
            ? {
                label: process.platform === 'darwin' ? tm('menuClose') : tm('menuQuit'),
                accelerator: process.platform === 'darwin' ? 'CmdOrCtrl+W' : 'CmdOrCtrl+Q',
                click: () => closeActiveTabHook?.(),
              }
            : process.platform === 'darwin'
              ? { role: 'close' as const, label: tm('menuClose') }
              : { role: 'quit' as const, label: tm('menuQuit') },
        ],
      },
      {
        label: tm('menuEdit'),
        submenu: [
          // role: 'editMenu' would bind ⌘Z to webContents.undo(), a text-editing
          // no-op that starves Univer of the shortcut — forward it instead.
          {
            label: tm('menuUndo'),
            accelerator: 'CmdOrCtrl+Z',
            click: () => sendMenuAction('undo'),
          },
          {
            label: tm('menuRedo'),
            accelerator: 'Shift+CmdOrCtrl+Z',
            click: () => sendMenuAction('redo'),
          },
          { type: 'separator' },
          { role: 'cut', label: labels.cut },
          { role: 'copy', label: labels.copy },
          { role: 'paste', label: labels.paste },
          { type: 'separator' },
          { role: 'selectAll', label: labels.selectAll },
        ],
      },
      viewMenuTemplate(labels),
      windowMenuTemplate(process.platform, labels),
    ]),
  )
}

/** stop the Rust sidecar (shell calls this from its own before-quit hook) */
export function stopSheetsSidecar(): void {
  sidecar?.stop()
  sidecar = null
}

export {
  installApplicationMenu as installSheetsMenu,
  startCaptureServer as startSheetsCaptureServer,
}

/**
 * Attaches a proxy to the main process's global fetch (same source as
 * slides-main.applyMainProcessProxy): main-process Node fetch (undici) ignores
 * the system proxy by default, so direct connections from mainland networks to
 * overseas LLM endpoints like api.anthropic.com time out or get rejected by
 * egress region (403 Request not allowed). Environment variables take priority;
 * otherwise the system proxy is read via session.resolveProxy() after app ready.
 */
async function applyMainProcessProxy(): Promise<void> {
  const setDispatcher = async (proxyUrl: string) => {
    // spawned gsk CLI children do their own fetch and never see the
    // dispatcher below — forward the proxy to them via env
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
  const envProxy =
    process.env.HTTPS_PROXY ||
    process.env.https_proxy ||
    process.env.HTTP_PROXY ||
    process.env.http_proxy ||
    process.env.ALL_PROXY ||
    process.env.all_proxy
  if (envProxy) {
    await setDispatcher(envProxy)
    return
  }
  try {
    await app.whenReady()
    // PAC/rule proxies answer per-host: probe the host the login flow, the
    // Genspark LLM proxy and the gsk CLI actually target
    const resolved = await electronSession.defaultSession.resolveProxy('https://www.genspark.ai/')
    const m = /PROXY\s+([^;]+)/i.exec(resolved || '')
    if (m?.[1]) {
      await setDispatcher(`http://${m[1].trim()}`)
    } else {
      console.log('[proxy] system proxy = DIRECT, no dispatcher set')
    }
  } catch (e) {
    console.warn('[proxy] resolveProxy failed:', e)
  }
}

export function startSheetsStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  // GENOFFICE_USER_DATA: test drivers point this at a scratch dir so automated
  // instances get their own userData AND single-instance lock (the lock is scoped
  // to userData), allowing parallel instances alongside a normal dev run.
  // Same dev-only hook as apps/slides/src/main/slides-main.ts.
  if (!app.isPackaged && process.env.GENOFFICE_USER_DATA) {
    app.setPath('userData', process.env.GENOFFICE_USER_DATA)
  }
  void applyMainProcessProxy()
  app.whenReady().then(() => {
    setUiLang(normalizeLang(process.env.GENOFFICE_LANG ?? app.getLocale()))
    app.setAccessibilitySupportEnabled(true)
    installApplicationMenu()
    startCaptureServer()
    return createSheetsWindow()
  })
  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
  app.on('before-quit', () => {
    stopSheetsSidecar()
  })
  app.on('activate', () => {
    if (!mainWindow) void createSheetsWindow()
  })
}

function resolveSidecarPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  if (runtime.sidecarPath) return runtime.sidecarPath
  if (process.env.XLSX_SIDECAR_PATH) return process.env.XLSX_SIDECAR_PATH
  if (app.isPackaged) return join(process.resourcesPath, 'native', executable)
  return join(app.getAppPath(), 'native', 'xlsx-engine', 'target', 'release', executable)
}

async function sha256File(path: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const hash = createHash('sha256')
    const stream = createReadStream(path)
    stream.on('data', (chunk) => hash.update(chunk))
    stream.once('error', reject)
    stream.once('end', () => resolve(hash.digest('hex')))
  })
}

async function closeAllSessions(entry: {
  client: XlsxSidecarClient
  sessions: Map<string, SessionInfo>
}): Promise<void> {
  const sessionIds = [...entry.sessions.keys()]
  entry.sessions.clear()
  await Promise.allSettled(sessionIds.map((sessionId) => entry.client.close(sessionId)))
}
