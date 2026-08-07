import { existsSync } from 'node:fs'
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, extname, join, resolve, sep } from 'node:path'
import { pathToFileURL } from 'node:url'
import {
  BrowserWindow,
  WebContentsView,
  app,
  dialog,
  ipcMain,
  net,
  protocol,
  shell,
} from 'electron'
import type { WebContents } from 'electron'
import {
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  safeExternalUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
} from '@genoffice/electron-utils'
import { createI18n, getUiLang } from '@genoffice/i18n'
import { atomicWriteFile } from './atomic-write'
import { MARKDOWN_CHANNELS } from '../shared/ipc'
import type {
  ExportDocxRequest,
  ExportFormat,
  ExportPdfRequest,
  ExportResult,
  ImageData,
  SaveMarkdownRequest,
  SaveMarkdownResult,
  SaveMode,
} from '../shared/ipc'

const tDlg = createI18n({
  zh: {
    dlgSaveTitle: '保存 Markdown 文档',
    filterMarkdown: 'Markdown 文档',
    dlgPickImage: '选择图片',
    filterImages: '图片',
    untitledFile: '未命名文档',
    closeUnsavedMsg: '此文档有未保存的更改。',
    closeUnsavedDetail: '关闭前是否保存？',
    btnSave: '保存',
    btnDontSave: '不保存',
    btnCancel: '取消',
  },
  en: {
    dlgSaveTitle: 'Save Markdown Document',
    filterMarkdown: 'Markdown Documents',
    dlgPickImage: 'Choose an Image',
    filterImages: 'Images',
    untitledFile: 'Untitled',
    closeUnsavedMsg: 'This document has unsaved changes.',
    closeUnsavedDetail: 'Do you want to save them before closing?',
    btnSave: 'Save',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
  },
  ja: {
    dlgSaveTitle: 'Markdown ドキュメントを保存',
    filterMarkdown: 'Markdown ドキュメント',
    dlgPickImage: '画像を選択',
    filterImages: '画像',
    untitledFile: '無題',
    closeUnsavedMsg: 'このドキュメントに未保存の変更があります。',
    closeUnsavedDetail: '閉じる前に保存しますか？',
    btnSave: '保存',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
  },
  ko: {
    dlgSaveTitle: 'Markdown 문서 저장',
    filterMarkdown: 'Markdown 문서',
    dlgPickImage: '이미지 선택',
    filterImages: '이미지',
    untitledFile: '제목 없음',
    closeUnsavedMsg: '이 문서에 저장하지 않은 변경 사항이 있습니다.',
    closeUnsavedDetail: '닫기 전에 저장하시겠습니까?',
    btnSave: '저장',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
  },
  fr: {
    dlgSaveTitle: 'Enregistrer le document Markdown',
    filterMarkdown: 'Documents Markdown',
    dlgPickImage: 'Choisir une image',
    filterImages: 'Images',
    untitledFile: 'Sans titre',
    closeUnsavedMsg: 'Ce document contient des modifications non enregistrées.',
    closeUnsavedDetail: 'Voulez-vous les enregistrer avant de fermer ?',
    btnSave: 'Enregistrer',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
  },
  de: {
    dlgSaveTitle: 'Markdown-Dokument speichern',
    filterMarkdown: 'Markdown-Dokumente',
    dlgPickImage: 'Bild auswählen',
    filterImages: 'Bilder',
    untitledFile: 'Unbenannt',
    closeUnsavedMsg: 'Dieses Dokument enthält ungespeicherte Änderungen.',
    closeUnsavedDetail: 'Vor dem Schließen speichern?',
    btnSave: 'Speichern',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
  },
  es: {
    dlgSaveTitle: 'Guardar documento Markdown',
    filterMarkdown: 'Documentos Markdown',
    dlgPickImage: 'Elegir imagen',
    filterImages: 'Imágenes',
    untitledFile: 'Sin título',
    closeUnsavedMsg: 'Este documento tiene cambios sin guardar.',
    closeUnsavedDetail: '¿Quieres guardarlos antes de cerrar?',
    btnSave: 'Guardar',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
  },
  th: {
    dlgSaveTitle: 'บันทึกเอกสาร Markdown',
    filterMarkdown: 'เอกสาร Markdown',
    dlgPickImage: 'เลือกรูปภาพ',
    filterImages: 'รูปภาพ',
    untitledFile: 'ไม่มีชื่อ',
    closeUnsavedMsg: 'เอกสารนี้มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก',
    closeUnsavedDetail: 'ต้องการบันทึกก่อนปิดหรือไม่?',
    btnSave: 'บันทึก',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
  },
  id: {
    dlgSaveTitle: 'Simpan dokumen Markdown',
    filterMarkdown: 'Dokumen Markdown',
    dlgPickImage: 'Pilih gambar',
    filterImages: 'Gambar',
    untitledFile: 'Tanpa judul',
    closeUnsavedMsg: 'Dokumen ini memiliki perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  ru: {
    dlgSaveTitle: 'Сохранить документ Markdown',
    filterMarkdown: 'Документы Markdown',
    dlgPickImage: 'Выберите изображение',
    filterImages: 'Изображения',
    untitledFile: 'Без названия',
    closeUnsavedMsg: 'В этом документе есть несохранённые изменения.',
    closeUnsavedDetail: 'Сохранить их перед закрытием?',
    btnSave: 'Сохранить',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
  },
  ar: {
    dlgSaveTitle: 'حفظ مستند Markdown',
    filterMarkdown: 'مستندات Markdown',
    dlgPickImage: 'اختر صورة',
    filterImages: 'صور',
    untitledFile: 'بدون عنوان',
    closeUnsavedMsg: 'يحتوي هذا المستند على تغييرات غير محفوظة.',
    closeUnsavedDetail: 'هل تريد حفظها قبل الإغلاق؟',
    btnSave: 'حفظ',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
  },
  pt: {
    dlgSaveTitle: 'Salvar documento Markdown',
    filterMarkdown: 'Documentos Markdown',
    dlgPickImage: 'Escolher imagem',
    filterImages: 'Imagens',
    untitledFile: 'Sem título',
    closeUnsavedMsg: 'Este documento tem alterações não salvas.',
    closeUnsavedDetail: 'Deseja salvá-las antes de fechar?',
    btnSave: 'Salvar',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
  },
  it: {
    dlgSaveTitle: 'Salva documento Markdown',
    filterMarkdown: 'Documenti Markdown',
    dlgPickImage: 'Scegli immagine',
    filterImages: 'Immagini',
    untitledFile: 'Senza titolo',
    closeUnsavedMsg: 'Questo documento contiene modifiche non salvate.',
    closeUnsavedDetail: 'Vuoi salvarle prima di chiudere?',
    btnSave: 'Salva',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
  },
  pl: {
    dlgSaveTitle: 'Zapisz dokument Markdown',
    filterMarkdown: 'Dokumenty Markdown',
    dlgPickImage: 'Wybierz obraz',
    filterImages: 'Obrazy',
    untitledFile: 'Bez tytułu',
    closeUnsavedMsg: 'Ten dokument ma niezapisane zmiany.',
    closeUnsavedDetail: 'Czy zapisać je przed zamknięciem?',
    btnSave: 'Zapisz',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
  },
  nl: {
    dlgSaveTitle: 'Markdown-document opslaan',
    filterMarkdown: 'Markdown-documenten',
    dlgPickImage: 'Kies een afbeelding',
    filterImages: 'Afbeeldingen',
    untitledFile: 'Naamloos',
    closeUnsavedMsg: 'Dit document bevat niet-opgeslagen wijzigingen.',
    closeUnsavedDetail: 'Wilt u ze opslaan voordat u sluit?',
    btnSave: 'Opslaan',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
  },
  ms: {
    dlgSaveTitle: 'Simpan dokumen Markdown',
    filterMarkdown: 'Dokumen Markdown',
    dlgPickImage: 'Pilih imej',
    filterImages: 'Imej',
    untitledFile: 'Tanpa tajuk',
    closeUnsavedMsg: 'Dokumen ini mempunyai perubahan yang belum disimpan.',
    closeUnsavedDetail: 'Simpan sebelum menutup?',
    btnSave: 'Simpan',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
  },
  he: {
    dlgSaveTitle: 'שמירת מסמך Markdown',
    filterMarkdown: 'מסמכי Markdown',
    dlgPickImage: 'בחרו תמונה',
    filterImages: 'תמונות',
    untitledFile: 'ללא שם',
    closeUnsavedMsg: 'במסמך הזה יש שינויים שלא נשמרו.',
    closeUnsavedDetail: 'האם לשמור אותם לפני הסגירה?',
    btnSave: 'שמירה',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
  },
  hi: {
    dlgSaveTitle: 'Markdown दस्तावेज़ सहेजें',
    filterMarkdown: 'Markdown दस्तावेज़',
    dlgPickImage: 'छवि चुनें',
    filterImages: 'छवियाँ',
    untitledFile: 'शीर्षकहीन',
    closeUnsavedMsg: 'इस दस्तावेज़ में सहेजे नहीं गए परिवर्तन हैं।',
    closeUnsavedDetail: 'क्या बंद करने से पहले उन्हें सहेजना चाहते हैं?',
    btnSave: 'सहेजें',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
  },
  'zh-TW': {
    dlgSaveTitle: '儲存 Markdown 文件',
    filterMarkdown: 'Markdown 文件',
    dlgPickImage: '選擇圖片',
    filterImages: '圖片',
    untitledFile: '未命名文件',
    closeUnsavedMsg: '此文件有未儲存的變更。',
    closeUnsavedDetail: '關閉前是否儲存？',
    btnSave: '儲存',
    btnDontSave: '不儲存',
    btnCancel: '取消',
  },
})
type DlgKey =
  | 'dlgSaveTitle'
  | 'filterMarkdown'
  | 'dlgPickImage'
  | 'filterImages'
  | 'untitledFile'
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
}

let runtime: RuntimePaths = { preloadPath: '' }

export function configureMarkdownRuntime(paths: RuntimePaths): void {
  runtime = paths
}

/** Open path per view, queued at tab creation; the renderer consumes it after mount.
 * Kept until the view is destroyed so a reload (View > Reload) consumes it again. */
const openPathByWc = new Map<number, string>()
/** File paths granted to each view — readFile/save only allow these */
const allowedByWc = new Map<number, Set<string>>()
/** Current save target per view; absent = untitled document */
const savePathByWc = new Map<number, string>()
/** Unsaved-changes flags mirrored from the renderer; drives the save prompt before closing a tab/window */
const dirtyByWc = new Set<number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
/** Resolvers for menu-triggered saves, resolved when the renderer's save invoke completes */
const saveWaiters = new Map<number, (ok: boolean) => void>()

/** Fired after a save lands on a NEW path (untitled first save / Save As) — the shell syncs tab title, recents, projects */
let fileSavedHook: ((wc: WebContents, path: string) => void) | null = null

export function setMarkdownFileSavedHook(hook: (wc: WebContents, path: string) => void): void {
  fileSavedHook = hook
}

/** Fired after a "convert & open in Docs" export — the shell routes the new .docx to a docs tab */
let docxExportedHook: ((path: string) => void) | null = null

export function setMarkdownDocxExportedHook(hook: (path: string) => void): void {
  docxExportedHook = hook
}

/** Shell menu export entry: ask the renderer to serialize and run the export flow */
export function sendMarkdownExportRequest(contents: WebContents, format: ExportFormat): void {
  if (!contents.isDestroyed()) contents.send(MARKDOWN_CHANNELS.exportRequest, format)
}

export function markdownIsDirty(webContentsId: number): boolean {
  return dirtyByWc.has(webContentsId)
}

export function markdownFilePath(webContentsId: number): string | undefined {
  return savePathByWc.get(webContentsId)
}

/** The file was renamed on disk — re-grant the new path and tell the renderer */
export function markdownFileRenamed(contents: WebContents, oldPath: string, newPath: string): void {
  const wcId = contents.id
  if (savePathByWc.get(wcId) === oldPath) savePathByWc.set(wcId, newPath)
  if (openPathByWc.get(wcId) === oldPath) openPathByWc.set(wcId, newPath)
  const allowed = allowedByWc.get(wcId)
  if (allowed?.has(oldPath)) allowed.add(newPath)
  if (!contents.isDestroyed()) contents.send(MARKDOWN_CHANNELS.fileRenamed, newPath)
}

/**
 * Close guard: true means proceed with closing. Clean → true; dirty →
 * Save / Don't Save / Cancel. On Save, ask the renderer to serialize + write
 * and await the result; a canceled untitled-save dialog keeps the tab open.
 */
export async function requestMarkdownClose(
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
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(MARKDOWN_CHANNELS.closeSaveRequest)
  })
}

/** Menu Save / Save As: ask the renderer to serialize and save; clean views resolve true immediately on plain save */
export function requestMarkdownSave(contents: WebContents, mode: SaveMode): Promise<boolean> {
  if (contents.isDestroyed()) return Promise.resolve(false)
  if (mode === 'save' && !dirtyByWc.has(contents.id) && savePathByWc.has(contents.id)) {
    return Promise.resolve(true)
  }
  return new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      saveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    saveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(MARKDOWN_CHANNELS.saveRequest, mode)
  })
}

async function writeTextAtomic(path: string, text: string): Promise<void> {
  await atomicWriteFile(path, Buffer.from(text, 'utf8'))
}

async function resolveSaveTarget(
  e: Electron.IpcMainInvokeEvent,
  mode: SaveMode,
  suggestedName?: string,
): Promise<string | null | 'canceled'> {
  const current = savePathByWc.get(e.sender.id)
  if (mode === 'save' && current) return current
  // AI auto-naming: silent first save of an untitled document
  if (mode === 'save' && !current && suggestedName) {
    const base = suggestedName
      .replace(/[/\\:*?"<>|]/g, '_')
      .slice(0, 80)
      .trim()
    if (base) {
      const dir = app.getPath('documents')
      let target = join(dir, `${base}.md`)
      for (let n = 1; existsSync(target); n++) target = join(dir, `${base}-${n}.md`)
      return target
    }
  }
  const win =
    BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
  const defaultPath = current
    ? join(dirname(current), basename(current))
    : join(app.getPath('documents'), `${tm('untitledFile')}.md`)
  const picked = await showSaveDialogWithMemory(dialog, win, {
    title: tm('dlgSaveTitle'),
    defaultPath,
    filters: [{ name: tm('filterMarkdown'), extensions: ['md', 'markdown'] }],
  })
  if (picked.canceled || !picked.filePath) return 'canceled'
  return picked.filePath
}

const DISPLAY_IMAGE_EXTS = new Set([
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
  '.webp',
  '.svg',
  '.bmp',
  '.avif',
])

/**
 * Serves authored image paths to the editor DOM. A plain file:// <img> URL is
 * blocked whenever the renderer page is served over http (dev server), so the
 * renderer resolves images to md-asset:// instead. Only image files inside an
 * open document's directory are served.
 */
function registerImageProtocol(): void {
  protocol.handle('md-asset', (request) => {
    let target: string
    try {
      target = decodeURIComponent(new URL(request.url).pathname)
    } catch {
      return new Response(null, { status: 400 })
    }
    if (/^\/[a-zA-Z]:\//.test(target)) target = target.slice(1)
    target = resolve(target)
    if (!DISPLAY_IMAGE_EXTS.has(extname(target).toLowerCase()) || !existsSync(target)) {
      return new Response(null, { status: 404 })
    }
    const inDocDir = [...new Set([...openPathByWc.values(), ...savePathByWc.values()])].some(
      (doc) => {
        const dir = resolve(dirname(doc))
        return target === dir || target.startsWith(dir + sep)
      },
    )
    if (!inDocDir) return new Response(null, { status: 403 })
    return net.fetch(pathToFileURL(target).toString())
  })
}

let ipcRegistered = false

function registerMarkdownIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  registerImageProtocol()

  ipcMain.handle(MARKDOWN_CHANNELS.consumePending, (e) => openPathByWc.get(e.sender.id) ?? null)

  ipcMain.handle(MARKDOWN_CHANNELS.readFile, async (e, path: unknown) => {
    if (typeof path !== 'string' || !allowedByWc.get(e.sender.id)?.has(path)) {
      throw new Error('markdown: path not granted to this view')
    }
    return await readFile(path, 'utf8')
  })

  ipcMain.handle(
    MARKDOWN_CHANNELS.save,
    async (e, request: SaveMarkdownRequest): Promise<SaveMarkdownResult> => {
      const waiter = saveWaiters.get(e.sender.id)
      saveWaiters.delete(e.sender.id)
      const done = (result: SaveMarkdownResult): SaveMarkdownResult => {
        waiter?.(result.ok && !('canceled' in result))
        return result
      }
      if (typeof request?.text !== 'string') {
        return done({ ok: false, error: 'markdown: bad save request' })
      }
      const mode: SaveMode = request.mode === 'saveAs' ? 'saveAs' : 'save'
      try {
        const suggestedName =
          typeof request.suggestedName === 'string' ? request.suggestedName : undefined
        const target = await resolveSaveTarget(e, mode, suggestedName)
        if (target === 'canceled') return done({ ok: true, canceled: true })
        if (!target) return done({ ok: false, error: 'markdown: no save target' })
        const isNewPath = savePathByWc.get(e.sender.id) !== target
        await writeTextAtomic(target, request.text)
        savePathByWc.set(e.sender.id, target)
        // keep the reload path in sync — a stale openPathByWc would make a
        // reloaded renderer load the OLD file and then save it over the new one
        openPathByWc.set(e.sender.id, target)
        const allowed = allowedByWc.get(e.sender.id) ?? new Set<string>()
        allowed.add(target)
        allowedByWc.set(e.sender.id, allowed)
        dirtyByWc.delete(e.sender.id)
        if (isNewPath) fileSavedHook?.(e.sender, target)
        return done({ ok: true, path: target })
      } catch (err) {
        return done({ ok: false, error: err instanceof Error ? err.message : String(err) })
      }
    },
  )

  ipcMain.handle(MARKDOWN_CHANNELS.pickImage, async (e): Promise<string | null> => {
    const docPath = savePathByWc.get(e.sender.id)
    if (!docPath) return null
    const win =
      BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
    const picked = await showOpenDialogWithMemory(dialog, win, {
      title: tm('dlgPickImage'),
      // only formats readImage/DOCX export can round-trip (docx-engine NewImage mimes)
      filters: [{ name: tm('filterImages'), extensions: ['png', 'jpg', 'jpeg', 'gif'] }],
      properties: ['openFile'],
    })
    const source = picked.filePaths[0]
    if (picked.canceled || !source) return null
    const assetsDir = join(dirname(docPath), 'assets')
    await mkdir(assetsDir, { recursive: true })
    const ext = extname(source)
    const base = basename(source, ext).replace(/[/\\:*?"<>|]/g, '_')
    let name = `${base}${ext}`
    for (let n = 1; existsSync(join(assetsDir, name)); n++) name = `${base}-${n}${ext}`
    await copyFile(source, join(assetsDir, name))
    return `assets/${name}`
  })

  ipcMain.handle(
    MARKDOWN_CHANNELS.saveImage,
    async (e, data: { base64?: unknown; ext?: unknown }): Promise<string | null> => {
      const docPath = savePathByWc.get(e.sender.id)
      const ext = String(data?.ext ?? '').toLowerCase()
      if (!docPath || typeof data?.base64 !== 'string' || !data.base64) return null
      // keep in sync with readImage's MIME map — every authored asset must stay DOCX-exportable
      if (!['png', 'jpg', 'jpeg', 'gif'].includes(ext)) return null
      const assetsDir = join(dirname(docPath), 'assets')
      await mkdir(assetsDir, { recursive: true })
      let name = `image.${ext}`
      for (let n = 1; existsSync(join(assetsDir, name)); n++) name = `image-${n}.${ext}`
      await writeFile(join(assetsDir, name), Buffer.from(data.base64, 'base64'))
      return `assets/${name}`
    },
  )

  const MIME_BY_EXT: Record<string, ImageData['mime']> = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
  }

  ipcMain.handle(
    MARKDOWN_CHANNELS.readImage,
    async (e, src: unknown): Promise<ImageData | null> => {
      const docPath = savePathByWc.get(e.sender.id)
      if (!docPath || typeof src !== 'string' || /^[a-z][a-z0-9+.-]*:/i.test(src)) return null
      const docDir = resolve(dirname(docPath))
      const target = resolve(docDir, src)
      // images must live inside the document's directory (assets/ convention)
      if (target !== docDir && !target.startsWith(docDir + sep)) return null
      const mime = MIME_BY_EXT[extname(target).toLowerCase()]
      if (!mime || !existsSync(target)) return null
      try {
        return { base64: (await readFile(target)).toString('base64'), mime }
      } catch {
        return null
      }
    },
  )

  ipcMain.handle(
    MARKDOWN_CHANNELS.exportDocx,
    async (e, request: ExportDocxRequest): Promise<ExportResult> => {
      if (typeof request?.base64 !== 'string' || !request.base64) {
        return { ok: false, error: 'markdown: bad export request' }
      }
      const safeName =
        String(request.suggestedName || tm('untitledFile'))
          .replace(/[/\\:*?"<>|]/g, '_')
          .slice(0, 80)
          .trim() || tm('untitledFile')
      try {
        const bytes = Buffer.from(request.base64, 'base64')
        if (request.mode === 'openInDocs') {
          // silent convert next to the .md (untitled documents go to Documents)
          const docPath = savePathByWc.get(e.sender.id)
          const dir = docPath ? dirname(docPath) : app.getPath('documents')
          let target = join(dir, `${safeName}.docx`)
          for (let n = 1; existsSync(target); n++) target = join(dir, `${safeName}-${n}.docx`)
          await writeFile(target, bytes)
          docxExportedHook?.(target)
          return { ok: true, path: target }
        }
        const win =
          BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
        const picked = await showSaveDialogWithMemory(dialog, win, {
          defaultPath: `${safeName}.docx`,
          filters: [{ name: 'Word', extensions: ['docx'] }],
        })
        if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
        await writeFile(picked.filePath, bytes)
        return { ok: true, path: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  ipcMain.handle(
    MARKDOWN_CHANNELS.exportPdf,
    async (e, request: ExportPdfRequest): Promise<ExportResult> => {
      if (typeof request?.html !== 'string' || !request.html) {
        return { ok: false, error: 'markdown: bad export request' }
      }
      const safeName =
        String(request.suggestedName || tm('untitledFile'))
          .replace(/[/\\:*?"<>|]/g, '_')
          .slice(0, 80)
          .trim() || tm('untitledFile')
      const win =
        BrowserWindow.fromWebContents(e.sender) ?? BrowserWindow.getFocusedWindow() ?? undefined
      const picked = await showSaveDialogWithMemory(dialog, win, {
        defaultPath: `${safeName}.pdf`,
        filters: [{ name: 'PDF', extensions: ['pdf'] }],
      })
      if (picked.canceled || !picked.filePath) return { ok: true, canceled: true }
      // sheets-style: render the print HTML in a hidden scripting-disabled window
      const workDir = await mkdtemp(join(tmpdir(), 'genoffice-md-pdf-'))
      const printWin = new BrowserWindow({
        show: false,
        webPreferences: { sandbox: true, javascript: false },
      })
      try {
        const htmlPath = join(workDir, 'print.html')
        await writeFile(htmlPath, request.html, 'utf8')
        await printWin.loadFile(htmlPath)
        const pdf = await printWin.webContents.printToPDF({
          pageSize: 'A4',
          printBackground: true,
          margins: { top: 0.6, bottom: 0.6, left: 0.6, right: 0.6 },
        })
        await writeFile(picked.filePath, pdf)
        return { ok: true, path: picked.filePath }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      } finally {
        printWin.destroy()
        await rm(workDir, { recursive: true, force: true })
      }
    },
  )

  ipcMain.on(MARKDOWN_CHANNELS.dirtyChanged, (e, dirty: unknown) => {
    if (dirty === true) dirtyByWc.add(e.sender.id)
    else dirtyByWc.delete(e.sender.id)
  })

  ipcMain.on(MARKDOWN_CHANNELS.closeSaveResult, (e, ok: unknown) => {
    const waiter = closeSaveWaiters.get(e.sender.id)
    closeSaveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // safety net for menu saves the renderer declined without invoking save()
  // (busy / still loading) — the save handler itself resolves the normal path
  ipcMain.on(MARKDOWN_CHANNELS.saveRequestAck, (e, ok: unknown) => {
    const waiter = saveWaiters.get(e.sender.id)
    saveWaiters.delete(e.sender.id)
    waiter?.(ok === true)
  })

  // Language channel shared with other modules; removeHandler tolerates duplicate registration
  ipcMain.removeHandler(MARKDOWN_CHANNELS.getLanguage)
  ipcMain.handle(MARKDOWN_CHANNELS.getLanguage, () => getUiLang())
}

function grantAndTrack(wc: WebContents, openPath?: string | null): void {
  const wcId = wc.id
  if (openPath && existsSync(openPath)) {
    openPathByWc.set(wcId, openPath)
    savePathByWc.set(wcId, openPath)
    allowedByWc.set(wcId, new Set([openPath]))
  }
  wc.setWindowOpenHandler(({ url }) => {
    const target = safeExternalUrl(url, { allowedProtocols: ['http:', 'https:', 'mailto:'] })
    if (target) void shell.openExternal(target)
    return { action: 'deny' }
  })
  wc.once('destroyed', () => {
    openPathByWc.delete(wcId)
    allowedByWc.delete(wcId)
    savePathByWc.delete(wcId)
    dirtyByWc.delete(wcId)
    closeSaveWaiters.get(wcId)?.(false)
    closeSaveWaiters.delete(wcId)
    saveWaiters.get(wcId)?.(false)
    saveWaiters.delete(wcId)
  })
}

export function createMarkdownView(openPath?: string | null): WebContentsView {
  registerMarkdownIpc()
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

/** Standalone window mode: `npm run dev -w @genoffice/markdown`, md path passed via argv */
export function startMarkdownStandalone(): void {
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  configureMarkdownRuntime({
    preloadPath: join(__dirname, '../preload/index.js'),
    rendererUrl: process.env.ELECTRON_RENDERER_URL,
    rendererFile: join(__dirname, '../renderer/index.html'),
  })
  void app.whenReady().then(() => {
    registerMarkdownIpc()
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
    const argPath = process.argv.slice(1).find((a) => /\.(md|markdown)$/i.test(a) && existsSync(a))
    grantAndTrack(win.webContents, argPath)
    if (runtime.rendererUrl) void win.loadURL(runtime.rendererUrl)
    else if (runtime.rendererFile) void win.loadFile(runtime.rendererFile)
  })
  app.on('window-all-closed', () => app.quit())
}
