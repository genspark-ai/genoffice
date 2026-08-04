/// Localized replacements for Electron role menus, whose built-in labels are
/// English-only and (for role:'windowMenu' on Windows/Linux) follow macOS
/// conventions (Zoom, Ctrl+M minimize, Bring All to Front).
import type { MenuItemConstructorOptions } from 'electron'
import { contextMenuLabels, type ContextMenuLabels } from './context-menu'

export interface AppMenuLabels extends ContextMenuLabels {
  window: string
  minimize: string
  closeWindow: string
  edit: string
  undo: string
  redo: string
  delete: string
  view: string
  reload: string
  forceReload: string
  toggleDevTools: string
  actualSize: string
  zoomIn: string
  zoomOut: string
  fullscreen: string
}

type Labels = Omit<AppMenuLabels, keyof ContextMenuLabels>

const EN: Labels = {
  window: 'Window',
  minimize: 'Minimize',
  closeWindow: 'Close Window',
  edit: 'Edit',
  undo: 'Undo',
  redo: 'Redo',
  delete: 'Delete',
  view: 'View',
  reload: 'Reload',
  forceReload: 'Force Reload',
  toggleDevTools: 'Developer Tools',
  actualSize: 'Actual Size',
  zoomIn: 'Zoom In',
  zoomOut: 'Zoom Out',
  fullscreen: 'Full Screen',
}

// Shared table, same rationale as context-menu.ts: one copy instead of
// 15 keys × 19 languages per app dictionary.
const LABELS: Record<string, Labels> = {
  zh: {
    window: '窗口',
    minimize: '最小化',
    closeWindow: '关闭窗口',
    edit: '编辑',
    undo: '撤销',
    redo: '重做',
    delete: '删除',
    view: '视图',
    reload: '重新加载',
    forceReload: '强制重新加载',
    toggleDevTools: '开发者工具',
    actualSize: '实际大小',
    zoomIn: '放大',
    zoomOut: '缩小',
    fullscreen: '全屏',
  },
  en: EN,
  ja: {
    window: 'ウィンドウ',
    minimize: '最小化',
    closeWindow: 'ウィンドウを閉じる',
    edit: '編集',
    undo: '元に戻す',
    redo: 'やり直す',
    delete: '削除',
    view: '表示',
    reload: '再読み込み',
    forceReload: '強制的に再読み込み',
    toggleDevTools: '開発者ツール',
    actualSize: '実際のサイズ',
    zoomIn: '拡大',
    zoomOut: '縮小',
    fullscreen: 'フルスクリーン',
  },
  ko: {
    window: '창',
    minimize: '최소화',
    closeWindow: '창 닫기',
    edit: '편집',
    undo: '실행 취소',
    redo: '다시 실행',
    delete: '삭제',
    view: '보기',
    reload: '새로고침',
    forceReload: '강제 새로고침',
    toggleDevTools: '개발자 도구',
    actualSize: '실제 크기',
    zoomIn: '확대',
    zoomOut: '축소',
    fullscreen: '전체 화면',
  },
  fr: {
    window: 'Fenêtre',
    minimize: 'Réduire',
    closeWindow: 'Fermer la fenêtre',
    edit: 'Édition',
    undo: 'Annuler',
    redo: 'Rétablir',
    delete: 'Supprimer',
    view: 'Affichage',
    reload: 'Recharger',
    forceReload: 'Forcer le rechargement',
    toggleDevTools: 'Outils de développement',
    actualSize: 'Taille réelle',
    zoomIn: 'Zoom avant',
    zoomOut: 'Zoom arrière',
    fullscreen: 'Plein écran',
  },
  de: {
    window: 'Fenster',
    minimize: 'Minimieren',
    closeWindow: 'Fenster schließen',
    edit: 'Bearbeiten',
    undo: 'Rückgängig',
    redo: 'Wiederholen',
    delete: 'Löschen',
    view: 'Ansicht',
    reload: 'Neu laden',
    forceReload: 'Erzwungenes Neuladen',
    toggleDevTools: 'Entwicklertools',
    actualSize: 'Originalgröße',
    zoomIn: 'Vergrößern',
    zoomOut: 'Verkleinern',
    fullscreen: 'Vollbild',
  },
  es: {
    window: 'Ventana',
    minimize: 'Minimizar',
    closeWindow: 'Cerrar ventana',
    edit: 'Edición',
    undo: 'Deshacer',
    redo: 'Rehacer',
    delete: 'Eliminar',
    view: 'Ver',
    reload: 'Recargar',
    forceReload: 'Forzar recarga',
    toggleDevTools: 'Herramientas de desarrollo',
    actualSize: 'Tamaño real',
    zoomIn: 'Acercar',
    zoomOut: 'Alejar',
    fullscreen: 'Pantalla completa',
  },
  th: {
    window: 'หน้าต่าง',
    minimize: 'ย่อเล็กสุด',
    closeWindow: 'ปิดหน้าต่าง',
    edit: 'แก้ไข',
    undo: 'เลิกทำ',
    redo: 'ทำซ้ำ',
    delete: 'ลบ',
    view: 'มุมมอง',
    reload: 'โหลดใหม่',
    forceReload: 'บังคับโหลดใหม่',
    toggleDevTools: 'เครื่องมือนักพัฒนา',
    actualSize: 'ขนาดจริง',
    zoomIn: 'ขยาย',
    zoomOut: 'ย่อ',
    fullscreen: 'เต็มหน้าจอ',
  },
  id: {
    window: 'Jendela',
    minimize: 'Minimalkan',
    closeWindow: 'Tutup Jendela',
    edit: 'Edit',
    undo: 'Urungkan',
    redo: 'Ulangi',
    delete: 'Hapus',
    view: 'Tampilan',
    reload: 'Muat Ulang',
    forceReload: 'Paksa Muat Ulang',
    toggleDevTools: 'Alat Pengembang',
    actualSize: 'Ukuran Sebenarnya',
    zoomIn: 'Perbesar',
    zoomOut: 'Perkecil',
    fullscreen: 'Layar Penuh',
  },
  ru: {
    window: 'Окно',
    minimize: 'Свернуть',
    closeWindow: 'Закрыть окно',
    edit: 'Правка',
    undo: 'Отменить',
    redo: 'Повторить',
    delete: 'Удалить',
    view: 'Вид',
    reload: 'Перезагрузить',
    forceReload: 'Принудительно перезагрузить',
    toggleDevTools: 'Инструменты разработчика',
    actualSize: 'Реальный размер',
    zoomIn: 'Увеличить',
    zoomOut: 'Уменьшить',
    fullscreen: 'Полноэкранный режим',
  },
  ar: {
    window: 'نافذة',
    minimize: 'تصغير',
    closeWindow: 'إغلاق النافذة',
    edit: 'تحرير',
    undo: 'تراجع',
    redo: 'إعادة',
    delete: 'حذف',
    view: 'عرض',
    reload: 'إعادة التحميل',
    forceReload: 'فرض إعادة التحميل',
    toggleDevTools: 'أدوات المطور',
    actualSize: 'الحجم الفعلي',
    zoomIn: 'تكبير',
    zoomOut: 'تصغير العرض',
    fullscreen: 'ملء الشاشة',
  },
  pt: {
    window: 'Janela',
    minimize: 'Minimizar',
    closeWindow: 'Fechar Janela',
    edit: 'Editar',
    undo: 'Desfazer',
    redo: 'Refazer',
    delete: 'Excluir',
    view: 'Exibir',
    reload: 'Recarregar',
    forceReload: 'Forçar Recarregamento',
    toggleDevTools: 'Ferramentas do Desenvolvedor',
    actualSize: 'Tamanho Real',
    zoomIn: 'Ampliar',
    zoomOut: 'Reduzir',
    fullscreen: 'Tela Cheia',
  },
  it: {
    window: 'Finestra',
    minimize: 'Riduci a icona',
    closeWindow: 'Chiudi finestra',
    edit: 'Modifica',
    undo: 'Annulla',
    redo: 'Ripeti',
    delete: 'Elimina',
    view: 'Visualizza',
    reload: 'Ricarica',
    forceReload: 'Forza ricarica',
    toggleDevTools: 'Strumenti di sviluppo',
    actualSize: 'Dimensioni effettive',
    zoomIn: 'Ingrandisci',
    zoomOut: 'Riduci',
    fullscreen: 'Schermo intero',
  },
  pl: {
    window: 'Okno',
    minimize: 'Minimalizuj',
    closeWindow: 'Zamknij okno',
    edit: 'Edycja',
    undo: 'Cofnij',
    redo: 'Ponów',
    delete: 'Usuń',
    view: 'Widok',
    reload: 'Załaduj ponownie',
    forceReload: 'Wymuś ponowne załadowanie',
    toggleDevTools: 'Narzędzia deweloperskie',
    actualSize: 'Rzeczywisty rozmiar',
    zoomIn: 'Powiększ',
    zoomOut: 'Pomniejsz',
    fullscreen: 'Pełny ekran',
  },
  nl: {
    window: 'Venster',
    minimize: 'Minimaliseren',
    closeWindow: 'Venster sluiten',
    edit: 'Bewerken',
    undo: 'Ongedaan maken',
    redo: 'Opnieuw',
    delete: 'Verwijderen',
    view: 'Beeld',
    reload: 'Opnieuw laden',
    forceReload: 'Geforceerd opnieuw laden',
    toggleDevTools: 'Ontwikkelaarstools',
    actualSize: 'Ware grootte',
    zoomIn: 'Inzoomen',
    zoomOut: 'Uitzoomen',
    fullscreen: 'Volledig scherm',
  },
  ms: {
    window: 'Tetingkap',
    minimize: 'Minimumkan',
    closeWindow: 'Tutup Tetingkap',
    edit: 'Edit',
    undo: 'Buat Asal',
    redo: 'Buat Semula',
    delete: 'Padam',
    view: 'Paparan',
    reload: 'Muat Semula',
    forceReload: 'Paksa Muat Semula',
    toggleDevTools: 'Alat Pembangun',
    actualSize: 'Saiz Sebenar',
    zoomIn: 'Zum Masuk',
    zoomOut: 'Zum Keluar',
    fullscreen: 'Skrin Penuh',
  },
  he: {
    window: 'חלון',
    minimize: 'מזער',
    closeWindow: 'סגור חלון',
    edit: 'עריכה',
    undo: 'בטל',
    redo: 'בצע שוב',
    delete: 'מחק',
    view: 'תצוגה',
    reload: 'טען מחדש',
    forceReload: 'טען מחדש בכפייה',
    toggleDevTools: 'כלי מפתחים',
    actualSize: 'גודל בפועל',
    zoomIn: 'הגדל',
    zoomOut: 'הקטן',
    fullscreen: 'מסך מלא',
  },
  hi: {
    window: 'विंडो',
    minimize: 'छोटा करें',
    closeWindow: 'विंडो बंद करें',
    edit: 'संपादन',
    undo: 'पूर्ववत करें',
    redo: 'फिर से करें',
    delete: 'हटाएँ',
    view: 'दृश्य',
    reload: 'पुनः लोड करें',
    forceReload: 'बलपूर्वक पुनः लोड करें',
    toggleDevTools: 'डेवलपर टूल',
    actualSize: 'वास्तविक आकार',
    zoomIn: 'ज़ूम इन',
    zoomOut: 'ज़ूम आउट',
    fullscreen: 'पूर्ण स्क्रीन',
  },
  'zh-TW': {
    window: '視窗',
    minimize: '最小化',
    closeWindow: '關閉視窗',
    edit: '編輯',
    undo: '復原',
    redo: '重做',
    delete: '刪除',
    view: '檢視',
    reload: '重新載入',
    forceReload: '強制重新載入',
    toggleDevTools: '開發人員工具',
    actualSize: '實際大小',
    zoomIn: '放大',
    zoomOut: '縮小',
    fullscreen: '全螢幕',
  },
}

export function appMenuLabels(lang: string): AppMenuLabels {
  return { ...contextMenuLabels(lang), ...(LABELS[lang] ?? EN) }
}

/** macOS keeps the native role (Minimize/Zoom/Front, window list); Windows/Linux
 * gets only conventional items — no Zoom/Front, and no Ctrl+M accelerator since
 * Windows has no menu shortcut for minimize. */
export function windowMenuTemplate(
  platform: NodeJS.Platform,
  labels: AppMenuLabels,
): MenuItemConstructorOptions {
  if (platform === 'darwin') return { role: 'windowMenu', label: labels.window }
  return {
    label: labels.window,
    submenu: [
      { label: labels.minimize, click: (_item, win) => win?.minimize() },
      { type: 'separator' },
      { label: labels.closeWindow, click: (_item, win) => win?.close() },
    ],
  }
}

/** macOS keeps role:'editMenu' (Speech/Substitutions submenus etc.); elsewhere
 * the same items Electron would generate, with localized labels. */
export function editMenuTemplate(
  platform: NodeJS.Platform,
  labels: AppMenuLabels,
): MenuItemConstructorOptions {
  if (platform === 'darwin') return { role: 'editMenu', label: labels.edit }
  return {
    label: labels.edit,
    submenu: [
      { role: 'undo', label: labels.undo },
      { role: 'redo', label: labels.redo },
      { type: 'separator' },
      { role: 'cut', label: labels.cut },
      { role: 'copy', label: labels.copy },
      { role: 'paste', label: labels.paste },
      { role: 'delete', label: labels.delete },
      { type: 'separator' },
      { role: 'selectAll', label: labels.selectAll },
    ],
  }
}

/** role:'viewMenu' expands identically on every platform, so no branch. */
export function viewMenuTemplate(labels: AppMenuLabels): MenuItemConstructorOptions {
  return {
    label: labels.view,
    submenu: [
      { role: 'reload', label: labels.reload },
      { role: 'forceReload', label: labels.forceReload },
      { role: 'toggleDevTools', label: labels.toggleDevTools },
      { type: 'separator' },
      { role: 'resetZoom', label: labels.actualSize },
      { role: 'zoomIn', label: labels.zoomIn },
      { role: 'zoomOut', label: labels.zoomOut },
      { type: 'separator' },
      { role: 'togglefullscreen', label: labels.fullscreen },
    ],
  }
}
