/// Native right-click menu for surfaces without a self-drawn one (AI panel,
/// inputs, chrome). Renderer-drawn menus call preventDefault() on the DOM
/// contextmenu event, which suppresses this webContents event entirely
/// (verified on Electron 41), so the two never stack.
import type { App, ContextMenuParams, MenuItemConstructorOptions, WebContents } from 'electron'
import { saveImageFromUrl } from './save-image'

/** Native "View Image" hands the src to the renderer, which owns the viewer overlay */
export const VIEW_IMAGE_CHANNEL = 'genoffice:view-image'

export interface ContextMenuLabels {
  cut: string
  copy: string
  paste: string
  selectAll: string
  viewImage: string
  copyImage: string
  saveImageAs: string
}

const EN: ContextMenuLabels = {
  cut: 'Cut',
  copy: 'Copy',
  paste: 'Paste',
  selectAll: 'Select All',
  viewImage: 'View Image',
  copyImage: 'Copy Image',
  saveImageAs: 'Save Image As…',
}

// One shared table instead of 7 keys × 20 languages duplicated into every
// app's main-process dictionary; strings match the docs Edit menu.
const LABELS: Record<string, ContextMenuLabels> = {
  zh: {
    cut: '剪切',
    copy: '复制',
    paste: '粘贴',
    selectAll: '全选',
    viewImage: '查看图片',
    copyImage: '复制图片',
    saveImageAs: '图片另存为…',
  },
  en: EN,
  ja: {
    cut: '切り取り',
    copy: 'コピー',
    paste: '貼り付け',
    selectAll: 'すべて選択',
    viewImage: '画像を表示',
    copyImage: '画像をコピー',
    saveImageAs: '名前を付けて画像を保存…',
  },
  ko: {
    cut: '잘라내기',
    copy: '복사',
    paste: '붙여넣기',
    selectAll: '모두 선택',
    viewImage: '이미지 보기',
    copyImage: '이미지 복사',
    saveImageAs: '이미지를 다른 이름으로 저장…',
  },
  fr: {
    cut: 'Couper',
    copy: 'Copier',
    paste: 'Coller',
    selectAll: 'Tout sélectionner',
    viewImage: "Afficher l'image",
    copyImage: "Copier l'image",
    saveImageAs: "Enregistrer l'image sous…",
  },
  de: {
    cut: 'Ausschneiden',
    copy: 'Kopieren',
    paste: 'Einfügen',
    selectAll: 'Alles auswählen',
    viewImage: 'Bild anzeigen',
    copyImage: 'Bild kopieren',
    saveImageAs: 'Bild speichern unter…',
  },
  es: {
    cut: 'Cortar',
    copy: 'Copiar',
    paste: 'Pegar',
    selectAll: 'Seleccionar todo',
    viewImage: 'Ver imagen',
    copyImage: 'Copiar imagen',
    saveImageAs: 'Guardar imagen como…',
  },
  th: {
    cut: 'ตัด',
    copy: 'คัดลอก',
    paste: 'วาง',
    selectAll: 'เลือกทั้งหมด',
    viewImage: 'ดูรูปภาพ',
    copyImage: 'คัดลอกรูปภาพ',
    saveImageAs: 'บันทึกรูปภาพเป็น…',
  },
  id: {
    cut: 'Potong',
    copy: 'Salin',
    paste: 'Tempel',
    selectAll: 'Pilih Semua',
    viewImage: 'Lihat Gambar',
    copyImage: 'Salin Gambar',
    saveImageAs: 'Simpan Gambar Sebagai…',
  },
  ru: {
    cut: 'Вырезать',
    copy: 'Копировать',
    paste: 'Вставить',
    selectAll: 'Выделить все',
    viewImage: 'Открыть изображение',
    copyImage: 'Копировать изображение',
    saveImageAs: 'Сохранить изображение как…',
  },
  ar: {
    cut: 'قص',
    copy: 'نسخ',
    paste: 'لصق',
    selectAll: 'تحديد الكل',
    viewImage: 'عرض الصورة',
    copyImage: 'نسخ الصورة',
    saveImageAs: 'حفظ الصورة باسم…',
  },
  pt: {
    cut: 'Recortar',
    copy: 'Copiar',
    paste: 'Colar',
    selectAll: 'Selecionar Tudo',
    viewImage: 'Ver imagem',
    copyImage: 'Copiar imagem',
    saveImageAs: 'Salvar imagem como…',
  },
  it: {
    cut: 'Taglia',
    copy: 'Copia',
    paste: 'Incolla',
    selectAll: 'Seleziona tutto',
    viewImage: 'Visualizza immagine',
    copyImage: 'Copia immagine',
    saveImageAs: 'Salva immagine con nome…',
  },
  pl: {
    cut: 'Wytnij',
    copy: 'Kopiuj',
    paste: 'Wklej',
    selectAll: 'Zaznacz wszystko',
    viewImage: 'Wyświetl obraz',
    copyImage: 'Kopiuj obraz',
    saveImageAs: 'Zapisz obraz jako…',
  },
  cs: {
    cut: 'Vyjmout',
    copy: 'Kopírovat',
    paste: 'Vložit',
    selectAll: 'Vybrat vše',
    viewImage: 'Zobrazit obrázek',
    copyImage: 'Kopírovat obrázek',
    saveImageAs: 'Uložit obrázek jako…',
  },
  nl: {
    cut: 'Knippen',
    copy: 'Kopiëren',
    paste: 'Plakken',
    selectAll: 'Alles selecteren',
    viewImage: 'Afbeelding bekijken',
    copyImage: 'Afbeelding kopiëren',
    saveImageAs: 'Afbeelding opslaan als…',
  },
  ms: {
    cut: 'Potong',
    copy: 'Salin',
    paste: 'Tampal',
    selectAll: 'Pilih Semua',
    viewImage: 'Lihat Imej',
    copyImage: 'Salin Imej',
    saveImageAs: 'Simpan Imej Sebagai…',
  },
  he: {
    cut: 'גזור',
    copy: 'העתק',
    paste: 'הדבק',
    selectAll: 'בחר הכול',
    viewImage: 'הצג תמונה',
    copyImage: 'העתק תמונה',
    saveImageAs: 'שמור תמונה בשם…',
  },
  hi: {
    cut: 'काटें',
    copy: 'कॉपी करें',
    paste: 'चिपकाएँ',
    selectAll: 'सभी चुनें',
    viewImage: 'छवि देखें',
    copyImage: 'छवि कॉपी करें',
    saveImageAs: 'छवि इस रूप में सहेजें…',
  },
  'zh-TW': {
    cut: '剪下',
    copy: '複製',
    paste: '貼上',
    selectAll: '全選',
    viewImage: '檢視圖片',
    copyImage: '複製圖片',
    saveImageAs: '另存圖片為…',
  },
}

export function contextMenuLabels(lang: string): ContextMenuLabels {
  return LABELS[lang] ?? EN
}

export type ContextMenuItem =
  | { type: 'separator' }
  | { action: 'replaceMisspelling'; label: string }
  | { action: 'cut' | 'copy' | 'paste' | 'selectAll'; label: string; enabled: boolean }
  | { action: 'viewImage' | 'copyImage' | 'saveImageAs'; label: string }

type BuildParams = Pick<
  ContextMenuParams,
  'isEditable' | 'selectionText' | 'misspelledWord' | 'dictionarySuggestions' | 'editFlags'
> &
  Partial<Pick<ContextMenuParams, 'mediaType' | 'srcURL'>>

/** Empty result means: don't show a menu. */
export function buildContextMenuItems(
  params: BuildParams,
  labels: ContextMenuLabels,
): ContextMenuItem[] {
  const rest = buildEditItems(params, labels)
  if (params.mediaType !== 'image' || !params.srcURL) return rest
  const image: ContextMenuItem[] = [
    { action: 'viewImage', label: labels.viewImage },
    { action: 'copyImage', label: labels.copyImage },
    { action: 'saveImageAs', label: labels.saveImageAs },
  ]
  return rest.length ? [...image, { type: 'separator' }, ...rest] : image
}

function buildEditItems(params: BuildParams, labels: ContextMenuLabels): ContextMenuItem[] {
  const flags = params.editFlags
  if (params.isEditable) {
    const items: ContextMenuItem[] = []
    if (params.misspelledWord && params.dictionarySuggestions.length) {
      for (const word of params.dictionarySuggestions)
        items.push({ action: 'replaceMisspelling', label: word })
      items.push({ type: 'separator' })
    }
    items.push(
      { action: 'cut', label: labels.cut, enabled: flags.canCut },
      { action: 'copy', label: labels.copy, enabled: flags.canCopy },
      { action: 'paste', label: labels.paste, enabled: flags.canPaste },
      { type: 'separator' },
      { action: 'selectAll', label: labels.selectAll, enabled: flags.canSelectAll },
    )
    return items
  }
  if (params.selectionText.trim()) {
    return [
      { action: 'copy', label: labels.copy, enabled: flags.canCopy },
      { action: 'selectAll', label: labels.selectAll, enabled: flags.canSelectAll },
    ]
  }
  return []
}

// Symbol.for: survives multiple bundled copies (see navigation-guard.ts).
const INSTALLED = Symbol.for('genoffice.context-menu-installed')

export function installContextMenu(app: App, getLabels: () => ContextMenuLabels): void {
  const holder = app as unknown as Record<symbol, boolean | undefined>
  if (holder[INSTALLED]) return
  holder[INSTALLED] = true
  app.on('web-contents-created', (_event, contents) => {
    contents.on('context-menu', (_e, params) => {
      void popupMenu(contents, params, getLabels())
    })
  })
}

async function popupMenu(
  contents: WebContents,
  params: ContextMenuParams,
  labels: ContextMenuLabels,
): Promise<void> {
  const items = buildContextMenuItems(params, labels)
  if (!items.length) return
  // Dynamic import keeps this module loadable outside Electron (unit tests).
  const { Menu } = await import('electron')
  const template = items.map((item): MenuItemConstructorOptions => {
    if ('type' in item) return { type: 'separator' }
    if (item.action === 'replaceMisspelling')
      return { label: item.label, click: () => contents.replaceMisspelling(item.label) }
    if (!('enabled' in item)) {
      if (item.action === 'viewImage')
        return { label: item.label, click: () => contents.send(VIEW_IMAGE_CHANNEL, params.srcURL) }
      if (item.action === 'copyImage')
        return { label: item.label, click: () => contents.copyImageAt(params.x, params.y) }
      return {
        label: item.label,
        click: async () => {
          const { BrowserWindow } = await import('electron')
          await saveImageFromUrl(BrowserWindow.fromWebContents(contents), params.srcURL, {
            title: item.label.replace(/…$/, ''),
          })
        },
      }
    }
    const action = item.action
    return { label: item.label, enabled: item.enabled, click: () => contents[action]() }
  })
  Menu.buildFromTemplate(template).popup()
}
