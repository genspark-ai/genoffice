/// Native right-click menu for surfaces without a self-drawn one (AI panel,
/// inputs, chrome). Renderer-drawn menus call preventDefault() on the DOM
/// contextmenu event, which suppresses this webContents event entirely
/// (verified on Electron 41), so the two never stack.
import type { App, ContextMenuParams, MenuItemConstructorOptions, WebContents } from 'electron'

export interface ContextMenuLabels {
  cut: string
  copy: string
  paste: string
  selectAll: string
}

const EN: ContextMenuLabels = { cut: 'Cut', copy: 'Copy', paste: 'Paste', selectAll: 'Select All' }

// One shared table instead of 4 keys × 19 languages duplicated into every
// app's main-process dictionary; strings match the docs Edit menu.
const LABELS: Record<string, ContextMenuLabels> = {
  zh: { cut: '剪切', copy: '复制', paste: '粘贴', selectAll: '全选' },
  en: EN,
  ja: { cut: '切り取り', copy: 'コピー', paste: '貼り付け', selectAll: 'すべて選択' },
  ko: { cut: '잘라내기', copy: '복사', paste: '붙여넣기', selectAll: '모두 선택' },
  fr: { cut: 'Couper', copy: 'Copier', paste: 'Coller', selectAll: 'Tout sélectionner' },
  de: { cut: 'Ausschneiden', copy: 'Kopieren', paste: 'Einfügen', selectAll: 'Alles auswählen' },
  es: { cut: 'Cortar', copy: 'Copiar', paste: 'Pegar', selectAll: 'Seleccionar todo' },
  th: { cut: 'ตัด', copy: 'คัดลอก', paste: 'วาง', selectAll: 'เลือกทั้งหมด' },
  id: { cut: 'Potong', copy: 'Salin', paste: 'Tempel', selectAll: 'Pilih Semua' },
  ru: { cut: 'Вырезать', copy: 'Копировать', paste: 'Вставить', selectAll: 'Выделить все' },
  ar: { cut: 'قص', copy: 'نسخ', paste: 'لصق', selectAll: 'تحديد الكل' },
  pt: { cut: 'Recortar', copy: 'Copiar', paste: 'Colar', selectAll: 'Selecionar Tudo' },
  it: { cut: 'Taglia', copy: 'Copia', paste: 'Incolla', selectAll: 'Seleziona tutto' },
  pl: { cut: 'Wytnij', copy: 'Kopiuj', paste: 'Wklej', selectAll: 'Zaznacz wszystko' },
  nl: { cut: 'Knippen', copy: 'Kopiëren', paste: 'Plakken', selectAll: 'Alles selecteren' },
  ms: { cut: 'Potong', copy: 'Salin', paste: 'Tampal', selectAll: 'Pilih Semua' },
  he: { cut: 'גזור', copy: 'העתק', paste: 'הדבק', selectAll: 'בחר הכול' },
  hi: { cut: 'काटें', copy: 'कॉपी करें', paste: 'चिपकाएँ', selectAll: 'सभी चुनें' },
  'zh-TW': { cut: '剪下', copy: '複製', paste: '貼上', selectAll: '全選' },
}

export function contextMenuLabels(lang: string): ContextMenuLabels {
  return LABELS[lang] ?? EN
}

export type ContextMenuItem =
  | { type: 'separator' }
  | { action: 'replaceMisspelling'; label: string }
  | { action: 'cut' | 'copy' | 'paste' | 'selectAll'; label: string; enabled: boolean }

type BuildParams = Pick<
  ContextMenuParams,
  'isEditable' | 'selectionText' | 'misspelledWord' | 'dictionarySuggestions' | 'editFlags'
>

/** Empty result means: don't show a menu. */
export function buildContextMenuItems(
  params: BuildParams,
  labels: ContextMenuLabels,
): ContextMenuItem[] {
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
    const action = item.action
    return { label: item.label, enabled: item.enabled, click: () => contents[action]() }
  })
  Menu.buildFromTemplate(template).popup()
}
