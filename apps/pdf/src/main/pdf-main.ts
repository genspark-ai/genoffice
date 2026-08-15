import { existsSync } from 'node:fs'
import { readFile, rename, writeFile } from 'node:fs/promises'
import { userInfo } from 'node:os'
import { join } from 'node:path'
import { BrowserWindow, WebContentsView, app, dialog, ipcMain, shell } from 'electron'
import type { WebContents } from 'electron'
import {
  configuredDefaultSaveDir,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showOpenDialogWithMemory,
} from '@genoffice/electron-utils'
import { createI18n, getUiLang } from '@genoffice/i18n'
import { gskGenerateImage, hasGskAuth } from '@genoffice/ai-search'
import { PDF_CHANNELS } from '../shared/ipc'
import type {
  ExportImagesRequest,
  ExportImagesResult,
  ExtractPagesRequest,
  ExtractPagesResult,
  InsertBlankPageRequest,
  InsertBlankPageResult,
  InsertPdfRequest,
  InsertPdfResult,
  MergePagesRequest,
  MergePagesResult,
  MergePdfRequest,
  MergePdfResult,
  PagePreviewRequest,
  ReplacePagesRequest,
  ReplacePagesResult,
  SetPageSizeRequest,
  SetPageSizeResult,
  SplitPagesRequest,
  SplitPagesResult,
  SplitPdfRequest,
  SplitPdfResult,
  SavePdfRequest,
  SavePdfResult,
  CropPagesRequest,
  CropPagesResult,
  TextEditValidation,
  ValidateTextEditsRequest,
} from '../shared/ipc'
import type { SavedSignature } from '../shared/ipc'
import {
  cropPagesBytes,
  extractPagesBytes,
  insertBlankPageBytes,
  insertPdfBytes,
  mergePagesBytes,
  mergePdfBytes,
  readStaticFormFills,
  replacePagesBytes,
  savePdfToPath,
  setPageSizeBytes,
  splitPagesBytes,
  splitPdfBytes,
} from './save-pdf'
import {
  addSignature,
  isSignatureData,
  loadSignatures,
  removeSignature,
  saveSignatures,
} from './signature-store'
import { uniqueGeneratedPdfPath } from './generated-output'

const tDlg = createI18n({
  zh: {
    dlgExportImages: '导出图片到文件夹',
    dlgExtract: '抽取页面为 PDF',
    dlgInsert: '选择要导入的 PDF',
    dlgSplit: '拆分 PDF 到文件夹',
    dlgMerge: '选择要合并的 PDF',
    dlgMergeSave: '合并 PDF 保存为',
    dlgMergePages: '合并页面保存为',
    dlgReplace: '选择用于替换的 PDF',
    dlgSplitPages: '拆分页面保存为',
    filterPdf: 'PDF 文档',
    closeUnsavedMsg: '此 PDF 有未保存的更改。',
    closeUnsavedDetail: '关闭前是否保存？',
    btnSave: '保存',
    btnDontSave: '不保存',
    btnCancel: '取消',
  },
  en: {
    dlgExportImages: 'Export Images to Folder',
    dlgExtract: 'Extract Pages as PDF',
    dlgInsert: 'Choose a PDF to Import',
    dlgSplit: 'Split PDF into Folder',
    dlgMerge: 'Choose PDFs to Merge',
    dlgMergeSave: 'Save Merged PDF As',
    dlgMergePages: 'Save Merged Pages As',
    dlgReplace: 'Choose a Replacement PDF',
    dlgSplitPages: 'Save Split Pages As',
    filterPdf: 'PDF Documents',
    closeUnsavedMsg: 'This PDF has unsaved changes.',
    closeUnsavedDetail: 'Do you want to save them before closing?',
    btnSave: 'Save',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
  },
  ja: {
    dlgExportImages: '画像をフォルダに書き出す',
    dlgExtract: 'ページを PDF として抽出',
    dlgInsert: 'インポートする PDF を選択',
    dlgSplit: 'PDF をフォルダに分割',
    dlgMerge: '結合する PDF を選択',
    dlgMergeSave: '結合した PDF の保存先',
    dlgMergePages: '結合したページの保存先',
    dlgReplace: '差し替え用の PDF を選択',
    dlgSplitPages: '分割したページの保存先',
    filterPdf: 'PDF ドキュメント',
    closeUnsavedMsg: 'この PDF に未保存の変更があります。',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    btnSave: '保存',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
  },
  ko: {
    dlgExportImages: '이미지를 폴더로 내보내기',
    dlgExtract: '페이지를 PDF로 추출',
    dlgInsert: '가져올 PDF 선택',
    dlgSplit: 'PDF를 폴더로 분할',
    dlgMerge: '병합할 PDF 선택',
    dlgMergeSave: '병합된 PDF 저장',
    dlgMergePages: '합쳐진 페이지 저장',
    dlgReplace: '교체할 PDF 선택',
    dlgSplitPages: '분할된 페이지 저장',
    filterPdf: 'PDF 문서',
    closeUnsavedMsg: '이 PDF에 저장하지 않은 변경 사항이 있습니다.',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    btnSave: '저장',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
  },
  fr: {
    dlgExportImages: 'Exporter les images vers un dossier',
    dlgExtract: 'Extraire les pages en PDF',
    dlgInsert: 'Choisir un PDF à importer',
    dlgSplit: 'Diviser le PDF dans un dossier',
    dlgMerge: 'Choisir les PDF à fusionner',
    dlgMergeSave: 'Enregistrer le PDF fusionné sous',
    dlgMergePages: 'Enregistrer les pages fusionnées sous',
    dlgReplace: 'Choisir un PDF de remplacement',
    dlgSplitPages: 'Enregistrer les pages divisées sous',
    filterPdf: 'Documents PDF',
    closeUnsavedMsg: 'Ce PDF contient des modifications non enregistrées.',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    btnSave: 'Enregistrer',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
  },
  de: {
    dlgExportImages: 'Bilder in Ordner exportieren',
    dlgExtract: 'Seiten als PDF extrahieren',
    dlgInsert: 'Zu importierendes PDF wählen',
    dlgSplit: 'PDF in Ordner aufteilen',
    dlgMerge: 'Zu vereinende PDFs wählen',
    dlgMergeSave: 'Zusammengeführtes PDF speichern unter',
    dlgMergePages: 'Zusammengefasste Seiten speichern unter',
    dlgReplace: 'Ersatz-PDF wählen',
    dlgSplitPages: 'Geteilte Seiten speichern unter',
    filterPdf: 'PDF-Dokumente',
    closeUnsavedMsg: 'Dieses PDF enthält ungespeicherte Änderungen.',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    btnSave: 'Speichern',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
  },
  es: {
    dlgExportImages: 'Exportar imágenes a una carpeta',
    dlgExtract: 'Extraer páginas como PDF',
    dlgInsert: 'Elegir un PDF para importar',
    dlgSplit: 'Dividir PDF en una carpeta',
    dlgMerge: 'Elegir PDF para combinar',
    dlgMergeSave: 'Guardar PDF combinado como',
    dlgMergePages: 'Guardar páginas combinadas como',
    dlgReplace: 'Elegir un PDF de reemplazo',
    dlgSplitPages: 'Guardar páginas divididas como',
    filterPdf: 'Documentos PDF',
    closeUnsavedMsg: 'Este PDF tiene cambios sin guardar.',
    closeUnsavedDetail: '¿Quieres guardarlos antes de cerrar?',
    btnSave: 'Guardar',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
  },
  th: {
    dlgExportImages: 'ส่งออกรูปภาพไปยังโฟลเดอร์',
    dlgExtract: 'แยกหน้าเป็น PDF',
    dlgInsert: 'เลือก PDF ที่จะนำเข้า',
    dlgSplit: 'แยก PDF ไปยังโฟลเดอร์',
    dlgMerge: 'เลือก PDF ที่จะรวม',
    dlgMergeSave: 'บันทึก PDF ที่รวมแล้วเป็น',
    dlgMergePages: 'บันทึกหน้าที่รวมแล้วเป็น',
    dlgReplace: 'เลือก PDF สำหรับแทนที่',
    dlgSplitPages: 'บันทึกหน้าที่แยกแล้วเป็น',
    filterPdf: 'เอกสาร PDF',
    closeUnsavedMsg: 'PDF นี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    btnSave: 'บันทึก',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
  },
  id: {
    dlgExportImages: 'Ekspor gambar ke folder',
    dlgExtract: 'Ekstrak halaman sebagai PDF',
    dlgInsert: 'Pilih PDF untuk diimpor',
    dlgSplit: 'Pisahkan PDF ke folder',
    dlgMerge: 'Pilih PDF untuk digabung',
    dlgMergeSave: 'Simpan PDF gabungan sebagai',
    dlgMergePages: 'Simpan halaman gabungan sebagai',
    dlgReplace: 'Pilih PDF pengganti',
    dlgSplitPages: 'Simpan halaman terpisah sebagai',
    filterPdf: 'Dokumen PDF',
    closeUnsavedMsg: 'PDF ini memiliki perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  ru: {
    dlgExportImages: 'Экспорт изображений в папку',
    dlgExtract: 'Извлечь страницы в PDF',
    dlgInsert: 'Выберите PDF для импорта',
    dlgSplit: 'Разделить PDF в папку',
    dlgMerge: 'Выберите PDF для объединения',
    dlgMergeSave: 'Сохранить объединённый PDF как',
    dlgMergePages: 'Сохранить объединённые страницы как',
    dlgReplace: 'Выберите PDF для замены',
    dlgSplitPages: 'Сохранить разделённые страницы как',
    filterPdf: 'Документы PDF',
    closeUnsavedMsg: 'В этом PDF есть несохранённые изменения.',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    btnSave: 'Сохранить',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
  },
  ar: {
    dlgExportImages: 'تصدير الصور إلى مجلد',
    dlgExtract: 'استخراج الصفحات كملف PDF',
    dlgInsert: 'اختر PDF للاستيراد',
    dlgSplit: 'تقسيم PDF إلى مجلد',
    dlgMerge: 'اختر ملفات PDF للدمج',
    dlgMergeSave: 'حفظ PDF المدمج باسم',
    dlgMergePages: 'حفظ الصفحات المدمجة باسم',
    dlgReplace: 'اختر PDF بديلاً',
    dlgSplitPages: 'حفظ الصفحات المقسّمة باسم',
    filterPdf: 'مستندات PDF',
    closeUnsavedMsg: 'يحتوي هذا الـ PDF على تغييرات غير محفوظة.',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    btnSave: 'حفظ',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
  },
  pt: {
    dlgExportImages: 'Exportar imagens para pasta',
    dlgExtract: 'Extrair páginas como PDF',
    dlgInsert: 'Escolher um PDF para importar',
    dlgSplit: 'Dividir PDF em uma pasta',
    dlgMerge: 'Escolher PDFs para mesclar',
    dlgMergeSave: 'Salvar PDF mesclado como',
    dlgMergePages: 'Salvar páginas combinadas como',
    dlgReplace: 'Escolher um PDF de substituição',
    dlgSplitPages: 'Salvar páginas divididas como',
    filterPdf: 'Documentos PDF',
    closeUnsavedMsg: 'Este PDF tem alterações não salvas.',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    btnSave: 'Salvar',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
  },
  it: {
    dlgExportImages: 'Esporta immagini in una cartella',
    dlgExtract: 'Estrai pagine come PDF',
    dlgInsert: 'Scegli un PDF da importare',
    dlgSplit: 'Dividi il PDF in una cartella',
    dlgMerge: 'Scegli i PDF da unire',
    dlgMergeSave: 'Salva il PDF unito come',
    dlgMergePages: 'Salva le pagine combinate come',
    dlgReplace: 'Scegli un PDF sostitutivo',
    dlgSplitPages: 'Salva le pagine divise come',
    filterPdf: 'Documenti PDF',
    closeUnsavedMsg: 'Questo PDF contiene modifiche non salvate.',
    closeUnsavedDetail: 'Vuoi salvarle prima di chiudere?',
    btnSave: 'Salva',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
  },
  pl: {
    dlgExportImages: 'Eksportuj obrazy do folderu',
    dlgExtract: 'Wyodrębnij strony jako PDF',
    dlgInsert: 'Wybierz PDF do zaimportowania',
    dlgSplit: 'Podziel PDF do folderu',
    dlgMerge: 'Wybierz pliki PDF do scalenia',
    dlgMergeSave: 'Zapisz scalony PDF jako',
    dlgMergePages: 'Zapisz scalone strony jako',
    dlgReplace: 'Wybierz PDF zastępczy',
    dlgSplitPages: 'Zapisz podzielone strony jako',
    filterPdf: 'Dokumenty PDF',
    closeUnsavedMsg: 'Ten PDF ma niezapisane zmiany.',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    btnSave: 'Zapisz',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
  },
  nl: {
    dlgExportImages: 'Afbeeldingen naar map exporteren',
    dlgExtract: "Pagina's extraheren als PDF",
    dlgInsert: 'Kies een PDF om te importeren',
    dlgSplit: 'PDF splitsen naar map',
    dlgMerge: "Kies PDF's om samen te voegen",
    dlgMergeSave: 'Samengevoegde PDF opslaan als',
    dlgMergePages: "Gecombineerde pagina's opslaan als",
    dlgReplace: 'Kies een vervangende PDF',
    dlgSplitPages: "Gesplitste pagina's opslaan als",
    filterPdf: 'PDF-documenten',
    closeUnsavedMsg: 'Deze PDF bevat niet-opgeslagen wijzigingen.',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    btnSave: 'Opslaan',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
  },
  ms: {
    dlgExportImages: 'Eksport imej ke folder',
    dlgExtract: 'Ekstrak halaman sebagai PDF',
    dlgInsert: 'Pilih PDF untuk diimport',
    dlgSplit: 'Pisahkan PDF ke folder',
    dlgMerge: 'Pilih PDF untuk digabungkan',
    dlgMergeSave: 'Simpan PDF gabungan sebagai',
    dlgMergePages: 'Simpan halaman gabungan sebagai',
    dlgReplace: 'Pilih PDF pengganti',
    dlgSplitPages: 'Simpan halaman dipisah sebagai',
    filterPdf: 'Dokumen PDF',
    closeUnsavedMsg: 'PDF ini mempunyai perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  he: {
    dlgExportImages: 'ייצוא תמונות לתיקייה',
    dlgExtract: 'חילוץ עמודים כ-PDF',
    dlgInsert: 'בחרו PDF לייבוא',
    dlgSplit: 'פיצול PDF לתיקייה',
    dlgMerge: 'בחרו קובצי PDF למיזוג',
    dlgMergeSave: 'שמירת ה-PDF הממוזג בשם',
    dlgMergePages: 'שמירת העמודים המאוחדים בשם',
    dlgReplace: 'בחרו PDF חלופי',
    dlgSplitPages: 'שמירת העמודים המפוצלים בשם',
    filterPdf: 'מסמכי PDF',
    closeUnsavedMsg: 'ב-PDF הזה יש שינויים שלא נשמרו.',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    btnSave: 'שמירה',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
  },
  hi: {
    dlgExportImages: 'चित्र फ़ोल्डर में निर्यात करें',
    dlgExtract: 'पृष्ठों को PDF के रूप में निकालें',
    dlgInsert: 'आयात करने के लिए PDF चुनें',
    dlgSplit: 'PDF को फ़ोल्डर में विभाजित करें',
    dlgMerge: 'मर्ज करने के लिए PDF चुनें',
    dlgMergeSave: 'मर्ज किया गया PDF इस रूप में सहेजें',
    dlgMergePages: 'संयोजित पृष्ठ इस रूप में सहेजें',
    dlgReplace: 'प्रतिस्थापन के लिए PDF चुनें',
    dlgSplitPages: 'विभाजित पृष्ठ इस रूप में सहेजें',
    filterPdf: 'PDF दस्तावेज़',
    closeUnsavedMsg: 'इस PDF में सहेजे नहीं गए परिवर्तन हैं।',
    closeUnsavedDetail: 'क्या बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    btnSave: 'सहेजें',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
  },
  'zh-TW': {
    dlgExportImages: '匯出圖片到資料夾',
    dlgExtract: '擷取頁面為 PDF',
    dlgInsert: '選擇要匯入的 PDF',
    dlgSplit: '拆分 PDF 到資料夾',
    dlgMerge: '選擇要合併的 PDF',
    dlgMergeSave: '合併 PDF 儲存為',
    dlgMergePages: '合併頁面儲存為',
    dlgReplace: '選擇用於取代的 PDF',
    dlgSplitPages: '拆分頁面儲存為',
    filterPdf: 'PDF 文件',
    closeUnsavedMsg: '此 PDF 有未儲存的變更。',
    closeUnsavedDetail: '關閉前是否儲存？',
    btnSave: '儲存',
    btnDontSave: '不儲存',
    btnCancel: '取消',
  },
})
type DlgKey =
  | 'dlgExportImages'
  | 'dlgExtract'
  | 'dlgInsert'
  | 'dlgSplit'
  | 'dlgMerge'
  | 'dlgMergeSave'
  | 'dlgMergePages'
  | 'dlgReplace'
  | 'dlgSplitPages'
  | 'filterPdf'
  | 'closeUnsavedMsg'
  | 'closeUnsavedDetail'
  | 'btnSave'
  | 'btnDontSave'
  | 'btnCancel'
const tm = (key: DlgKey) => tDlg(getUiLang(), key)

interface RuntimePaths {
  preloadPath: string
  rendererUrl?: string
  rendererFile?: string
  /** Shell router used to open generated PDFs in a new GenOffice tab. */
  openGeneratedPath?: (path: string) => boolean
}

let runtime: RuntimePaths = { preloadPath: '' }

export function configurePdfRuntime(paths: RuntimePaths): void {
  runtime = paths
}

function openGeneratedPdf(path: string): void {
  try {
    if (runtime.openGeneratedPath?.(path)) return
  } catch (err) {
    // The file is already safely persisted; a tab-opening failure must not
    // report the merge itself as failed.
    console.warn('[pdf] Failed to open generated PDF:', err)
  }
  // Standalone PDF mode has no shell tab router. If routing is unavailable or
  // rejects the path, reveal the persisted output so success is never silent.
  shell.showItemInFolder(path)
}

/** Open path per view, queued at tab creation; the renderer consumes it after mount
 * (avoids did-finish-load races). Kept until the view is destroyed: a reload
 * (View > Reload) remounts the renderer and consumes again — a one-shot entry
 * would strand the tab on "No file to open". */
const openPathByWc = new Map<number, string>()
/** File paths granted to each view — readFile only allows these */
const allowedByWc = new Map<number, Set<string>>()
/** Unsaved-changes flags mirrored from the renderer; drives the save prompt before closing a tab/window */
const dirtyByWc = new Set<number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
const saveAsWaiters = new Map<number, (ok: boolean) => void>()
/** Save As destination granted per view (main-process dialog pick); the save handler refuses any other non-source target */
const saveAsTargetByWc = new Map<number, string>()

export function pdfIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

/**
 * Close guard for the pdf renderer: true means proceed with closing.
 * Clean → true; dirty → Save / Don't Save / Cancel. On Save, ask the renderer to
 * write to disk and await the result; on failure or timeout stay open (renderer
 * has already shown the error).
 */
export async function requestPdfClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  if (!dirtyByWc.has(contents.id) || contents.isDestroyed()) return true
  const options = {
    type: 'warning' as const,
    message: tm('closeUnsavedMsg'),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('btnSave'), tm('btnDontSave'), tm('btnCancel')],
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
  return await requestRendererSave(contents)
}

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
    contents.send(PDF_CHANNELS.closeSaveRequest)
  })
}

/** Menu Save: ask the renderer to write pending edits to disk; clean views resolve true immediately */
export function flushPdfSave(contents: WebContents): Promise<boolean> {
  if (contents.isDestroyed() || !dirtyByWc.has(contents.id)) return Promise.resolve(true)
  return requestRendererSave(contents)
}

/** Menu Print: ask the renderer to run its print flow (save, rasterize, system dialog) */
export function sendPdfPrintRequest(contents: WebContents): void {
  if (!contents.isDestroyed()) contents.send(PDF_CHANNELS.printRequest)
}

/**
 * Marks the whole Save As flow (dialog included) for the renderer, which pauses
 * autosave meanwhile: opening the save dialog blurs the window, and a
 * blur-triggered autosave would write the pending edits into the original file.
 */
export function setPdfSaveAsInFlight(contents: WebContents, inFlight: boolean): void {
  if (!contents.isDestroyed()) contents.send(PDF_CHANNELS.saveAsFlow, inFlight)
}

/**
 * Menu Save As: grant targetPath to the view, then ask the renderer to apply its
 * pending edits onto the source bytes and write the result to targetPath only.
 * The original file is never written (non-destructive Save As).
 */
export function requestPdfSaveAs(contents: WebContents, targetPath: string): Promise<boolean> {
  if (contents.isDestroyed()) return Promise.resolve(false)
  const wcId = contents.id
  saveAsTargetByWc.set(wcId, targetPath)
  return new Promise<boolean>((resolve) => {
    const done = (ok: boolean) => {
      saveAsTargetByWc.delete(wcId)
      resolve(ok)
    }
    const timer = setTimeout(() => {
      saveAsWaiters.delete(wcId)
      done(false)
    }, 120_000)
    saveAsWaiters.set(wcId, (ok) => {
      clearTimeout(timer)
      done(ok)
    })
    contents.send(PDF_CHANNELS.saveAsRequest, targetPath)
  })
}

/** Saved signatures live in userData, shared by all documents and windows */
const signaturesPath = () => join(app.getPath('userData'), 'pdf-signatures.json')

/** Serialize signature file read-modify-writes: several pdf views share one file */
let signatureQueue: Promise<unknown> = Promise.resolve()
function withSignatures(
  op: (list: SavedSignature[]) => Promise<SavedSignature[]>,
): Promise<SavedSignature[]> {
  const next = signatureQueue
    .catch(() => undefined)
    .then(async () => op(await loadSignatures(signaturesPath())))
  signatureQueue = next
  return next
}

let ipcRegistered = false

function registerPdfIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  ipcMain.handle(PDF_CHANNELS.consumePending, (e) => openPathByWc.get(e.sender.id) ?? null)

  ipcMain.handle(PDF_CHANNELS.getUsername, () => {
    try {
      return userInfo().username
    } catch {
      return ''
    }
  })

  ipcMain.handle(PDF_CHANNELS.readFile, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('pdf: path not granted to this view')
    }
    const buf = await readFile(path)
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
  })

  ipcMain.handle(PDF_CHANNELS.save, async (e, request: SavePdfRequest): Promise<SavePdfResult> => {
    const path = request?.path
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      return { ok: false, error: 'pdf: path not granted to this view' }
    }
    // Save As targets must have been granted by requestPdfSaveAs (main-process dialog pick)
    const target = typeof request.targetPath === 'string' ? request.targetPath : path
    if (target !== path && saveAsTargetByWc.get(e.sender.id) !== target) {
      return { ok: false, error: 'pdf: target path not granted to this view' }
    }
    try {
      const { skippedTextEdits, skippedTextInserts, skippedImageEdits } = await savePdfToPath(
        path,
        target,
        request,
      )
      return {
        ok: true,
        ...(skippedTextEdits.length > 0 ? { skippedTextEdits } : {}),
        ...(skippedTextInserts.length > 0 ? { skippedTextInserts } : {}),
        ...(skippedImageEdits.length > 0 ? { skippedImageEdits } : {}),
      }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  ipcMain.handle(PDF_CHANNELS.listPageImages, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('pdf: path not granted to this view')
    }
    // Lazy import like the text-edit paths: pdfium wasm only loads when the feature is used
    const { listPageImages } = await import('./image-edit')
    return listPageImages(new Uint8Array(await readFile(path)))
  })

  ipcMain.handle(PDF_CHANNELS.listStaticFormFills, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('pdf: path not granted to this view')
    }
    return readStaticFormFills(new Uint8Array(await readFile(path)))
  })

  ipcMain.handle(
    PDF_CHANNELS.pageImagePng,
    async (
      e,
      request: {
        path: string
        pageIndex: number
        rect: [number, number, number, number]
        scale?: number
      },
    ) => {
      const { path, pageIndex, rect, scale } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        typeof pageIndex !== 'number' ||
        !Array.isArray(rect)
      ) {
        throw new Error('pdf: path not granted to this view')
      }
      const { renderImagePng } = await import('./image-edit')
      return renderImagePng(
        new Uint8Array(await readFile(path)),
        pageIndex,
        rect,
        typeof scale === 'number' && Number.isFinite(scale) ? scale : 1,
      )
    },
  )

  ipcMain.handle(PDF_CHANNELS.pagePreviewPng, async (e, request: PagePreviewRequest) => {
    const { path, pageIndex, excludeRects, excludeAnnots, clip, pxWidth, rotate } = request ?? {}
    if (
      typeof path !== 'string' ||
      !allowedByWc.get(e.sender.id)?.has(path) ||
      typeof pageIndex !== 'number' ||
      !Array.isArray(excludeRects) ||
      (excludeAnnots !== undefined && !Array.isArray(excludeAnnots)) ||
      typeof clip !== 'object' ||
      typeof pxWidth !== 'number' ||
      typeof rotate !== 'number'
    ) {
      throw new Error('pdf: path not granted to this view')
    }
    const { renderPagePreviewPng } = await import('./image-edit')
    return renderPagePreviewPng(new Uint8Array(await readFile(path)), {
      pageIndex,
      excludeRects,
      excludeAnnots,
      clip,
      pxWidth,
      rotate,
    })
  })

  ipcMain.handle(
    PDF_CHANNELS.validateTextEdits,
    async (e, request: ValidateTextEditsRequest): Promise<TextEditValidation[]> => {
      const { path, edits } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(edits)
      ) {
        throw new Error('pdf: path not granted to this view')
      }
      // Same lazy import as the save path: the pdfium wasm only loads when text editing is used
      const { validateTextEdits } = await import('./text-edit')
      return validateTextEdits(new Uint8Array(await readFile(path)), edits)
    },
  )

  ipcMain.handle(PDF_CHANNELS.listEditFonts, async (): Promise<string[]> => {
    const { listEditFonts } = await import('./text-edit')
    return listEditFonts()
  })

  ipcMain.handle(
    PDF_CHANNELS.extractPages,
    async (e, request: ExtractPagesRequest): Promise<ExtractPagesResult> => {
      const { path, pages, suggestedName } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(pages)
      ) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      try {
        const bytes = await extractPagesBytes(new Uint8Array(await readFile(path)), pages)
        const targetPath = uniqueGeneratedPdfPath(
          configuredDefaultSaveDir(app),
          String(suggestedName || 'pages.pdf'),
        )
        await writeFile(targetPath, bytes)
        openGeneratedPdf(targetPath)
        return { ok: true, savedPath: targetPath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.insertPdf,
    async (e, request: InsertPdfRequest): Promise<InsertPdfResult> => {
      const { path, afterPageIndex } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgInsert'),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
        properties: ['openFile'],
      })
      const other = picked.filePaths[0]
      if (picked.canceled || !other) return { ok: true, canceled: true }
      try {
        const { merged, count } = await insertPdfBytes(
          new Uint8Array(await readFile(path)),
          new Uint8Array(await readFile(other)),
          typeof afterPageIndex === 'number' ? afterPageIndex : -1,
        )
        const tmp = `${path}.gensave-${process.pid}.tmp`
        await writeFile(tmp, merged)
        await rename(tmp, path)
        return { ok: true, insertedCount: count }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.insertBlankPage,
    async (e, request: InsertBlankPageRequest): Promise<InsertBlankPageResult> => {
      const { path, afterPageIndex } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      try {
        const bytes = await insertBlankPageBytes(
          new Uint8Array(await readFile(path)),
          typeof afterPageIndex === 'number' ? afterPageIndex : -1,
        )
        const tmp = `${path}.gensave-${process.pid}.tmp`
        await writeFile(tmp, bytes)
        await rename(tmp, path)
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.splitPdf,
    async (e, request: SplitPdfRequest): Promise<SplitPdfResult> => {
      const { path, chunkSize, baseName } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgSplit'),
        properties: ['openDirectory', 'createDirectory'],
      })
      const dir = picked.filePaths[0]
      if (picked.canceled || !dir) return { ok: true, canceled: true }
      try {
        const parts = await splitPdfBytes(
          new Uint8Array(await readFile(path)),
          typeof chunkSize === 'number' ? chunkSize : 1,
        )
        const safeBase = String(baseName || 'split').replace(/[/\\:*?"<>|]/g, '_')
        for (const [i, bytes] of parts.entries()) {
          await writeFile(join(dir, `${safeBase}-${i + 1}.pdf`), bytes)
        }
        // Many output files — don't open tabs; reveal them so success is never silent
        shell.showItemInFolder(join(dir, `${safeBase}-1.pdf`))
        return { ok: true, savedDir: dir, count: parts.length }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.mergePdf,
    async (e, request: MergePdfRequest): Promise<MergePdfResult> => {
      const { path, suggestedName } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgMerge'),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
        properties: ['openFile', 'multiSelections'],
      })
      if (picked.canceled || picked.filePaths.length === 0) return { ok: true, canceled: true }
      try {
        const others = await Promise.all(
          picked.filePaths.map(async (p) => new Uint8Array(await readFile(p))),
        )
        const { merged, appended } = await mergePdfBytes(
          new Uint8Array(await readFile(path)),
          others,
        )
        const targetPath = uniqueGeneratedPdfPath(
          configuredDefaultSaveDir(app),
          String(suggestedName || 'merged.pdf'),
        )
        await writeFile(targetPath, merged)
        openGeneratedPdf(targetPath)
        return { ok: true, savedPath: targetPath, appendedCount: appended }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.mergePages,
    async (e, request: MergePagesRequest): Promise<MergePagesResult> => {
      const { path, perSheet, direction, separator, suggestedName } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      if (!Number.isInteger(perSheet) || perSheet < 2 || perSheet > 16) {
        return { ok: false, error: 'pdf: pages-per-sheet must be 2-16' }
      }
      try {
        const bytes = await mergePagesBytes(new Uint8Array(await readFile(path)), {
          perSheet,
          direction: direction === 'horizontal' ? 'horizontal' : 'vertical',
          separator: separator === true,
        })
        const targetPath = uniqueGeneratedPdfPath(
          configuredDefaultSaveDir(app),
          String(suggestedName || 'merged-pages.pdf'),
        )
        await writeFile(targetPath, bytes)
        openGeneratedPdf(targetPath)
        return { ok: true, savedPath: targetPath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.replacePages,
    async (e, request: ReplacePagesRequest): Promise<ReplacePagesResult> => {
      const { path, pages } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(pages) ||
        pages.length === 0
      ) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgReplace'),
        filters: [{ name: tm('filterPdf'), extensions: ['pdf'] }],
        properties: ['openFile'],
      })
      const other = picked.filePaths[0]
      if (picked.canceled || !other) return { ok: true, canceled: true }
      try {
        const { merged, removed, inserted } = await replacePagesBytes(
          new Uint8Array(await readFile(path)),
          new Uint8Array(await readFile(other)),
          pages,
        )
        const tmp = `${path}.gensave-${process.pid}.tmp`
        await writeFile(tmp, merged)
        await rename(tmp, path)
        return { ok: true, removed, inserted }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.setPageSize,
    async (e, request: SetPageSizeRequest): Promise<SetPageSizeResult> => {
      const { path, width, height } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      if (!(Number.isFinite(width) && width > 0 && Number.isFinite(height) && height > 0)) {
        return { ok: false, error: 'pdf: invalid page size' }
      }
      try {
        const bytes = await setPageSizeBytes(new Uint8Array(await readFile(path)), width, height)
        const tmp = `${path}.gensave-${process.pid}.tmp`
        await writeFile(tmp, bytes)
        await rename(tmp, path)
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.splitPages,
    async (e, request: SplitPagesRequest): Promise<SplitPagesResult> => {
      const { path, perPage, suggestedName } = request ?? {}
      if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      if (perPage !== 2 && perPage !== 4 && perPage !== 9) {
        return { ok: false, error: 'pdf: unsupported split grid' }
      }
      try {
        const bytes = await splitPagesBytes(new Uint8Array(await readFile(path)), perPage)
        const targetPath = uniqueGeneratedPdfPath(
          configuredDefaultSaveDir(app),
          String(suggestedName || 'split-pages.pdf'),
        )
        await writeFile(targetPath, bytes)
        openGeneratedPdf(targetPath)
        return { ok: true, savedPath: targetPath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.cropPages,
    async (e, request: CropPagesRequest): Promise<CropPagesResult> => {
      const { path, pages, rect } = request ?? {}
      if (
        typeof path !== 'string' ||
        !allowedByWc.get(e.sender.id)?.has(path) ||
        !Array.isArray(pages) ||
        pages.length === 0 ||
        !rect
      ) {
        return { ok: false, error: 'pdf: path not granted to this view' }
      }
      try {
        const bytes = await cropPagesBytes(new Uint8Array(await readFile(path)), pages, rect)
        const tmp = `${path}.gensave-${process.pid}.tmp`
        await writeFile(tmp, bytes)
        await rename(tmp, path)
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    PDF_CHANNELS.exportImages,
    async (e, request: ExportImagesRequest): Promise<ExportImagesResult> => {
      const { images, pageNumbers, baseName } = request ?? {}
      if (!Array.isArray(images) || images.length === 0)
        return { ok: false, error: 'pdf: no images' }
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showOpenDialogWithMemory(dialog, win, {
        title: tm('dlgExportImages'),
        properties: ['openDirectory', 'createDirectory'],
      })
      const dir = picked.filePaths[0]
      if (picked.canceled || !dir) return { ok: true, canceled: true }
      try {
        const safeBase = String(baseName || 'page').replace(/[/\\:*?"<>|]/g, '_')
        for (const [i, b64] of images.entries()) {
          const no = pageNumbers?.[i] ?? i + 1
          await writeFile(join(dir, `${safeBase}-p${no}.png`), Buffer.from(b64, 'base64'))
        }
        // Reveal the exported images so success is never silent
        shell.showItemInFolder(join(dir, `${safeBase}-p${pageNumbers?.[0] ?? 1}.png`))
        return { ok: true, savedDir: dir, count: images.length }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  // pdf-owned (unlike ai:image-search / ai:fetch-image, which the shell registers app-wide):
  // slides' ai:generate-image is only registered once a slides view exists, so pdf needs its own
  ipcMain.handle(
    PDF_CHANNELS.generateImage,
    async (_e, op: { prompt?: unknown; aspectRatio?: unknown }) => {
      if (!hasGskAuth())
        return {
          error: 'Genspark account is not logged in on this machine; ask the user to log in first',
        }
      const prompt = String(op?.prompt ?? '').trim()
      if (!prompt) return { error: 'prompt must not be empty' }
      try {
        const r = await gskGenerateImage({
          prompt,
          aspectRatio: op?.aspectRatio ? String(op.aspectRatio) : undefined,
        })
        return { url: r.url }
      } catch (err) {
        return { error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(PDF_CHANNELS.listSignatures, () => withSignatures(async (list) => list))

  ipcMain.handle(PDF_CHANNELS.addSignature, (_e, data: unknown) =>
    withSignatures(async (list) => {
      if (!isSignatureData(data)) return list
      const next = addSignature(list, data)
      await saveSignatures(signaturesPath(), next)
      return next
    }),
  )

  ipcMain.handle(PDF_CHANNELS.removeSignature, (_e, id: unknown) =>
    withSignatures(async (list) => {
      if (typeof id !== 'string') return list
      const next = removeSignature(list, id)
      if (next.length !== list.length) await saveSignatures(signaturesPath(), next)
      return next
    }),
  )

  ipcMain.on(PDF_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(PDF_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  ipcMain.on(PDF_CHANNELS.saveAsResult, (e, ok: unknown) => {
    const waiter = saveAsWaiters.get(e.sender.id)
    saveAsWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // Language channel shared with other modules; removeHandler tolerates duplicate registration
  ipcMain.removeHandler(PDF_CHANNELS.getLanguage)
  ipcMain.handle(PDF_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, openPath?: string | null): void {
  const wcId = wc.id
  if (openPath && existsSync(openPath)) {
    openPathByWc.set(wcId, openPath)
    allowedByWc.set(wcId, new Set([openPath]))
  }
  // External links inside the PDF (Link annots with target=_blank) go to the system browser
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: ['http:', 'https:', 'mailto:'] })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    openPathByWc.delete(wcId)
    allowedByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    saveAsTargetByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
    saveAsWaiters.get(wcId)?.(false)
    saveAsWaiters.delete(wcId)
  })
}

export function createPdfView(openPath?: string | null): WebContentsView {
  registerPdfIpc()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  grantAndTrack(view.webContents, openPath)
  if (runtime.rendererUrl) void view.webContents.loadURL(runtime.rendererUrl)
  else if (runtime.rendererFile) void view.webContents.loadFile(runtime.rendererFile)
  return view
}

/** Standalone window mode: `npm run dev -w @genoffice/pdf`, pdf path passed via argv */
export function startPdfStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configurePdfRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    registerPdfIpc()
    const win = new BrowserWindow({
      width: 1200,
      height: 850,
      webPreferences: {
        preload: runtime.preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
      },
    })
    const argPath = process.argv.slice(1).find((a) => /\.pdf$/i.test(a) && existsSync(a))
    grantAndTrack(win.webContents, argPath)
    if (runtime.rendererUrl) void win.loadURL(runtime.rendererUrl)
    else if (runtime.rendererFile) void win.loadFile(runtime.rendererFile)
  })
  app.on('window-all-closed', () => app.quit())
}
