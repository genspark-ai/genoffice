/// Localized replacements for Electron role menus, whose built-in labels are
/// English-only and (for role:'windowMenu' on Windows/Linux) follow macOS
/// conventions (Zoom, Ctrl+M minimize, Bring All to Front).
import type { MenuItemConstructorOptions, WebContents } from 'electron'
import { baseLang, contextMenuLabels, type ContextMenuLabels } from './context-menu'

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
  help: string
  about: string
  checkUpdates: string
  version: string
  settings: string
  nameWithAi: string
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
  help: 'Help',
  about: 'About GenOffice',
  checkUpdates: 'Check for Updates…',
  version: 'Version',
  settings: 'Settings…',
  nameWithAi: 'Name with AI…',
}

// Shared table, same rationale as context-menu.ts: one copy instead of
// 15 keys × 20 languages per app dictionary.
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
    help: '帮助',
    about: '关于 GenOffice',
    checkUpdates: '检查更新…',
    version: '版本',
    nameWithAi: '用 AI 命名…',
    settings: '设置…',
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
    help: 'ヘルプ',
    about: 'GenOffice について',
    checkUpdates: '更新を確認…',
    version: 'バージョン',
    nameWithAi: 'AI で名前を変更…',
    settings: '設定…',
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
    help: '도움말',
    about: 'GenOffice 정보',
    checkUpdates: '업데이트 확인…',
    version: '버전',
    nameWithAi: 'AI로 이름 지정…',
    settings: '설정…',
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
    help: 'Aide',
    about: 'À propos de GenOffice',
    checkUpdates: 'Rechercher les mises à jour…',
    version: 'Version',
    nameWithAi: "Nommer avec l'IA…",
    settings: 'Réglages…',
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
    help: 'Hilfe',
    about: 'Über GenOffice',
    checkUpdates: 'Nach Updates suchen…',
    version: 'Version',
    nameWithAi: 'Mit KI benennen…',
    settings: 'Einstellungen…',
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
    help: 'Ayuda',
    about: 'Acerca de GenOffice',
    checkUpdates: 'Buscar actualizaciones…',
    version: 'Versión',
    nameWithAi: 'Nombrar con IA…',
    settings: 'Ajustes…',
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
    help: 'วิธีใช้',
    about: 'เกี่ยวกับ GenOffice',
    checkUpdates: 'ตรวจหาการอัปเดต…',
    version: 'เวอร์ชัน',
    nameWithAi: 'ตั้งชื่อด้วย AI…',
    settings: 'การตั้งค่า…',
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
    help: 'Bantuan',
    about: 'Tentang GenOffice',
    checkUpdates: 'Periksa Pembaruan…',
    version: 'Versi',
    nameWithAi: 'Beri nama dengan AI…',
    settings: 'Pengaturan…',
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
    help: 'Справка',
    about: 'О GenOffice',
    checkUpdates: 'Проверить обновления…',
    version: 'Версия',
    nameWithAi: 'Назвать с помощью ИИ…',
    settings: 'Настройки…',
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
    help: 'تعليمات',
    about: 'حول GenOffice',
    checkUpdates: 'التحقق من التحديثات…',
    version: 'الإصدار',
    nameWithAi: 'التسمية بالذكاء الاصطناعي…',
    settings: 'الإعدادات…',
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
    help: 'Ajuda',
    about: 'Sobre o GenOffice',
    checkUpdates: 'Procurar atualizações…',
    version: 'Versão',
    nameWithAi: 'Nomear com IA…',
    settings: 'Configurações…',
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
    help: 'Aiuto',
    about: 'Informazioni su GenOffice',
    checkUpdates: 'Controlla aggiornamenti…',
    version: 'Versione',
    nameWithAi: 'Assegna un nome con IA…',
    settings: 'Impostazioni…',
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
    help: 'Pomoc',
    about: 'O programie GenOffice',
    checkUpdates: 'Sprawdź aktualizacje…',
    version: 'Wersja',
    nameWithAi: 'Nazwij za pomocą AI…',
    settings: 'Ustawienia…',
  },
  cs: {
    window: 'Okno',
    minimize: 'Minimalizovat',
    closeWindow: 'Zavřít okno',
    edit: 'Úpravy',
    undo: 'Zpět',
    redo: 'Znovu',
    delete: 'Odstranit',
    view: 'Zobrazení',
    reload: 'Znovu načíst',
    forceReload: 'Vynutit znovunačtení',
    toggleDevTools: 'Nástroje pro vývojáře',
    actualSize: 'Skutečná velikost',
    zoomIn: 'Přiblížit',
    zoomOut: 'Oddálit',
    fullscreen: 'Celá obrazovka',
    help: 'Nápověda',
    about: 'O aplikaci GenOffice',
    checkUpdates: 'Zkontrolovat aktualizace…',
    version: 'Verze',
    nameWithAi: 'Pojmenovat pomocí AI…',
    settings: 'Nastavení…',
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
    help: 'Help',
    about: 'Over GenOffice',
    checkUpdates: 'Controleren op updates…',
    version: 'Versie',
    nameWithAi: 'Naam geven met AI…',
    settings: 'Instellingen…',
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
    help: 'Bantuan',
    about: 'Perihal GenOffice',
    checkUpdates: 'Semak Kemas Kini…',
    version: 'Versi',
    nameWithAi: 'Namakan dengan AI…',
    settings: 'Tetapan…',
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
    help: 'עזרה',
    about: 'אודות GenOffice',
    checkUpdates: 'בדוק עדכונים…',
    version: 'גרסה',
    nameWithAi: 'מתן שם באמצעות AI…',
    settings: 'הגדרות…',
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
    help: 'सहायता',
    about: 'GenOffice के बारे में',
    checkUpdates: 'अपडेट जांचें…',
    version: 'संस्करण',
    nameWithAi: 'AI से नाम दें…',
    settings: 'सेटिंग्स…',
  },
  vi: {
    window: 'Cửa sổ',
    minimize: 'Thu nhỏ',
    closeWindow: 'Đóng cửa sổ',
    edit: 'Chỉnh sửa',
    undo: 'Hoàn tác',
    redo: 'Làm lại',
    delete: 'Xóa',
    view: 'Xem',
    reload: 'Tải lại',
    forceReload: 'Buộc tải lại',
    toggleDevTools: 'Công cụ nhà phát triển',
    actualSize: 'Kích cỡ thực',
    zoomIn: 'Phóng to',
    zoomOut: 'Thu nhỏ',
    fullscreen: 'Toàn màn hình',
    help: 'Trợ giúp',
    about: 'Giới thiệu GenOffice',
    checkUpdates: 'Kiểm tra bản cập nhật…',
    version: 'Phiên bản',
    nameWithAi: 'Đặt tên bằng AI…',
    settings: 'Cài đặt…',
  },
  tr: {
    window: 'Pencere',
    minimize: 'Simge Durumuna Küçült',
    closeWindow: 'Pencereyi Kapat',
    edit: 'Düzenle',
    undo: 'Geri Al',
    redo: 'Yinele',
    delete: 'Sil',
    view: 'Görünüm',
    reload: 'Yeniden Yükle',
    forceReload: 'Zorla Yeniden Yükle',
    toggleDevTools: 'Geliştirici Araçları',
    actualSize: 'Gerçek Boyut',
    zoomIn: 'Yakınlaştır',
    zoomOut: 'Uzaklaştır',
    fullscreen: 'Tam Ekran',
    help: 'Yardım',
    about: 'GenOffice Hakkında',
    checkUpdates: 'Güncellemeleri Denetle…',
    version: 'Sürüm',
    nameWithAi: 'AI ile adlandır',
    settings: 'Ayarlar…',
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
    help: '說明',
    about: '關於 GenOffice',
    checkUpdates: '檢查更新…',
    version: '版本',
    nameWithAi: '用 AI 命名…',
    settings: '設定…',
  },
}

export function appMenuLabels(lang: string): AppMenuLabels {
  return { ...contextMenuLabels(lang), ...(LABELS[lang] ?? LABELS[baseLang(lang)] ?? EN) }
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

let lastDetachedDevToolsTarget: WebContents | undefined

/** role:'toggleDevTools' docks DevTools into the window, where the shell's
 * WebContentsView tabs are stacked above it and occlude it — open detached
 * instead, keeping the role's accelerator and toggle semantics. */
export function toggleDevToolsItem(labels: AppMenuLabels): MenuItemConstructorOptions {
  return {
    label: labels.toggleDevTools,
    accelerator: process.platform === 'darwin' ? 'Alt+Command+I' : 'Ctrl+Shift+I',
    click: async () => {
      const { webContents } = await import('electron')
      const focused = webContents.getFocusedWebContents()
      const previous =
        lastDetachedDevToolsTarget && !lastDetachedDevToolsTarget.isDestroyed()
          ? lastDetachedDevToolsTarget
          : undefined
      const wc = !focused || focused === previous?.devToolsWebContents ? previous : focused
      if (!wc) return
      if (wc.isDevToolsOpened()) {
        wc.closeDevTools()
        if (wc === lastDetachedDevToolsTarget) lastDetachedDevToolsTarget = undefined
      } else {
        wc.openDevTools({ mode: 'detach' })
        lastDetachedDevToolsTarget = wc
      }
    },
  }
}

export interface ViewMenuOptions {
  /** Reload / Force Reload / DevTools (default on) */
  readonly devItems?: boolean
  /** Actual Size / Zoom In / Zoom Out page-zoom roles (default on) */
  readonly pageZoom?: boolean
}

/**
 * role:'viewMenu' expands identically on every platform, so no branch. Menu
 * accelerators beat renderer key handlers on macOS, so apps whose document
 * shortcuts collide with ⌘R / ⌘0 / ⌘+ / ⌘- opt those groups out.
 */
export function viewMenuTemplate(
  labels: AppMenuLabels,
  { devItems = true, pageZoom = true }: ViewMenuOptions = {},
): MenuItemConstructorOptions {
  const groups: MenuItemConstructorOptions[][] = []
  if (devItems) {
    groups.push([
      { role: 'reload', label: labels.reload },
      { role: 'forceReload', label: labels.forceReload },
      toggleDevToolsItem(labels),
    ])
  }
  if (pageZoom) {
    groups.push([
      { role: 'resetZoom', label: labels.actualSize },
      { role: 'zoomIn', label: labels.zoomIn },
      { role: 'zoomOut', label: labels.zoomOut },
    ])
  }
  groups.push([{ role: 'togglefullscreen', label: labels.fullscreen }])
  return {
    label: labels.view,
    submenu: groups.flatMap((group, index) =>
      index === 0 ? group : [{ type: 'separator' }, ...group],
    ),
  }
}

/** The manual update check lives in the shell (electron-updater and its
 * result dialogs), while the menus that expose it are built here — the shell
 * injects the check at startup. Read at click time, so registration order
 * relative to menu construction doesn't matter; until registered the menu
 * entry no-ops and the About dialog doesn't offer the button. */
let updateCheckInvoker: (() => void) | null = null

export function setUpdateCheckInvoker(invoke: (() => void) | null): void {
  updateCheckInvoker = invoke
}

/** Help > Check for Updates…: user-triggered update check (sits right above
 * About, like Word). The shell-injected check owns all feedback: the update
 * window when newer exists, "you're up to date (version x)" otherwise. */
export function checkUpdatesMenuItem(labels: AppMenuLabels): MenuItemConstructorOptions {
  return {
    label: labels.checkUpdates,
    click: () => updateCheckInvoker?.(),
  }
}

/** Help > About: a native dialog with the app version — every window's menu
 * gets one, so users can report the exact build they run. */
export function aboutMenuItem(labels: AppMenuLabels): MenuItemConstructorOptions {
  return {
    label: labels.about,
    click: async () => {
      const { app, dialog, clipboard } = await import('electron')
      const version = app.getVersion()
      const canCheck = updateCheckInvoker !== null
      const { response } = await dialog.showMessageBox({
        type: 'info',
        title: 'GenOffice',
        message: 'GenOffice',
        detail: `${labels.version} ${version}`,
        buttons: ['OK', labels.copy, ...(canCheck ? [labels.checkUpdates] : [])],
        defaultId: 0,
        cancelId: 0,
      })
      if (response === 1) clipboard.writeText(`GenOffice ${version}`)
      if (response === 2) updateCheckInvoker?.()
    },
  }
}

/** The Settings entry, shared by every app: one label, one accelerator
 * (⌘, on macOS and Ctrl+, elsewhere), one action. The shell owns the
 * Settings view, so it registers the action; an app running on its own
 * falls back to its own preferences entry. */
let settingsInvoker: (() => void) | null = null

export function setSettingsInvoker(invoke: (() => void) | null): void {
  settingsInvoker = invoke
}

export function settingsMenuItem(
  labels: AppMenuLabels,
  fallback?: () => void,
): MenuItemConstructorOptions {
  return {
    label: labels.settings,
    accelerator: 'CmdOrCtrl+,',
    click: () => (settingsInvoker ?? fallback)?.(),
  }
}

/** macOS application menu — the one under the app name. Electron's
 * role:'appMenu' is a fixed template, so an entry of our own means spelling
 * the menu out; the native roles keep the standard order and labels. Only
 * macOS has this menu, so every other platform gets undefined. */
export function appMenuTemplate(
  platform: NodeJS.Platform,
  labels: AppMenuLabels,
  {
    extra = [],
    appLabel = 'GenOffice',
  }: { extra?: MenuItemConstructorOptions[]; appLabel?: string } = {},
): MenuItemConstructorOptions | undefined {
  if (platform !== 'darwin') return undefined
  return {
    // Electron rejects a menu item that has neither label, role nor type
    label: appLabel,
    submenu: [
      { role: 'about', label: labels.about },
      ...extra,
      { type: 'separator' },
      { role: 'services' },
      { type: 'separator' },
      { role: 'hide' },
      { role: 'hideOthers' },
      { role: 'unhide' },
      { type: 'separator' },
      { role: 'quit' },
    ],
  }
}

/** Help menu with Check for Updates… + About; extra app-specific items go
 * before the separator. */
export function helpMenuTemplate(
  labels: AppMenuLabels,
  extraItems: MenuItemConstructorOptions[] = [],
): MenuItemConstructorOptions {
  return {
    role: 'help',
    label: labels.help,
    submenu: [
      ...extraItems,
      ...(extraItems.length > 0 ? [{ type: 'separator' } as const] : []),
      checkUpdatesMenuItem(labels),
      aboutMenuItem(labels),
    ],
  }
}
