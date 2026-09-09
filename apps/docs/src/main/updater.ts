import { app, dialog } from 'electron'
import type { BrowserWindow } from 'electron'
import { autoUpdater } from 'electron-updater'
import type { UpdateInfo } from 'electron-updater'
import { createI18n, getUiLang } from '@genoffice/i18n'

/**
 * Full-package auto-update for the standalone GenOffice Docs app over the generic
 * provider (Azure CDN).
 *
 * electron-builder bakes the publish URL into resources/app-update.yml at
 * package time and emits the update feed (latest.yml on Windows,
 * latest-mac.yml on macOS) next to the installers; the release pipeline
 * uploads both to the CDN. Standalone docs builds currently ship without a
 * publish config, so packaged runs simply log the missing-feed error.
 *
 * Unlike the shell's modal update window, docs keeps the flow minimal:
 * updates download silently in the background and a single native dialog
 * offers "Restart & Install" once the payload is ready. Declining defers the
 * install to normal app quit (autoInstallOnAppQuit).
 */

const tUpd = createI18n({
  zh: {
    updTitle: '软件更新',
    updHeadline: '发现新版本',
    updDesc: '新版本包含性能改进与问题修复，建议立即更新。',
    updInstall: '立即重启安装',
    updLater: '稍后再说',
  },
  en: {
    updTitle: 'Software Update',
    updHeadline: 'A new version is available',
    updDesc:
      'This update includes performance improvements and bug fixes. We recommend updating now.',
    updInstall: 'Restart & Install',
    updLater: 'Remind me later',
  },
  ja: {
    updTitle: 'ソフトウェアアップデート',
    updHeadline: '新しいバージョンがあります',
    updDesc:
      'このアップデートにはパフォーマンス改善とバグ修正が含まれます。今すぐの更新をおすすめします。',
    updInstall: '再起動してインストール',
    updLater: '後で通知',
  },
  ko: {
    updTitle: '소프트웨어 업데이트',
    updHeadline: '새 버전이 있습니다',
    updDesc:
      '이 업데이트에는 성능 개선과 버그 수정이 포함되어 있습니다. 지금 업데이트하는 것을 권장합니다.',
    updInstall: '다시 시작 및 설치',
    updLater: '나중에 알림',
  },
  fr: {
    updTitle: 'Mise à jour logicielle',
    updHeadline: 'Une nouvelle version est disponible',
    updDesc:
      'Cette mise à jour apporte des améliorations de performances et des corrections de bogues. Nous vous recommandons de mettre à jour maintenant.',
    updInstall: 'Redémarrer et installer',
    updLater: 'Plus tard',
  },
  de: {
    updTitle: 'Softwareaktualisierung',
    updHeadline: 'Eine neue Version ist verfügbar',
    updDesc:
      'Dieses Update enthält Leistungsverbesserungen und Fehlerbehebungen. Wir empfehlen, jetzt zu aktualisieren.',
    updInstall: 'Neu starten und installieren',
    updLater: 'Später erinnern',
  },
  es: {
    updTitle: 'Actualización de software',
    updHeadline: 'Hay una nueva versión disponible',
    updDesc:
      'Esta actualización incluye mejoras de rendimiento y correcciones de errores. Recomendamos actualizar ahora.',
    updInstall: 'Reiniciar e instalar',
    updLater: 'Recordar más tarde',
  },
  th: {
    updTitle: 'อัปเดตซอฟต์แวร์',
    updHeadline: 'มีเวอร์ชันใหม่พร้อมใช้งาน',
    updDesc: 'การอัปเดตนี้มีการปรับปรุงประสิทธิภาพและแก้ไขข้อบกพร่อง แนะนำให้อัปเดตทันที',
    updInstall: 'รีสตาร์ทและติดตั้ง',
    updLater: 'เตือนภายหลัง',
  },
  id: {
    updTitle: 'Pembaruan Perangkat Lunak',
    updHeadline: 'Versi baru tersedia',
    updDesc:
      'Pembaruan ini mencakup peningkatan kinerja dan perbaikan bug. Kami menyarankan untuk memperbarui sekarang.',
    updInstall: 'Mulai Ulang & Pasang',
    updLater: 'Ingatkan nanti',
  },
  ru: {
    updTitle: 'Обновление программы',
    updHeadline: 'Доступна новая версия',
    updDesc:
      'Это обновление содержит улучшения производительности и исправления ошибок. Рекомендуем обновиться сейчас.',
    updInstall: 'Перезапустить и установить',
    updLater: 'Напомнить позже',
  },
  ar: {
    updTitle: 'تحديث البرنامج',
    updHeadline: 'يتوفر إصدار جديد',
    updDesc: 'يتضمن هذا التحديث تحسينات في الأداء وإصلاحات للأخطاء. نوصي بالتحديث الآن.',
    updInstall: 'إعادة التشغيل والتثبيت',
    updLater: 'ذكّرني لاحقًا',
  },
  pt: {
    updTitle: 'Atualização de Software',
    updHeadline: 'Uma nova versão está disponível',
    updDesc:
      'Esta atualização inclui melhorias de desempenho e correções de erros. Recomendamos atualizar agora.',
    updInstall: 'Reiniciar e instalar',
    updLater: 'Lembrar mais tarde',
  },
  it: {
    updTitle: 'Aggiornamento software',
    updHeadline: 'È disponibile una nuova versione',
    updDesc:
      'Questo aggiornamento include miglioramenti delle prestazioni e correzioni di bug. Consigliamo di aggiornare subito.',
    updInstall: 'Riavvia e installa',
    updLater: 'Ricordamelo più tardi',
  },
  pl: {
    updTitle: 'Aktualizacja oprogramowania',
    updHeadline: 'Dostępna jest nowa wersja',
    updDesc:
      'Ta aktualizacja zawiera ulepszenia wydajności i poprawki błędów. Zalecamy aktualizację teraz.',
    updInstall: 'Uruchom ponownie i zainstaluj',
    updLater: 'Przypomnij później',
  },
  cs: {
    updTitle: 'Aktualizace softwaru',
    updHeadline: 'Je k dispozici nová verze',
    updDesc:
      'Tato aktualizace obsahuje vylepšení výkonu a opravy chyb. Doporučujeme aktualizovat nyní.',
    updInstall: 'Restartovat a nainstalovat',
    updLater: 'Připomenout později',
  },
  nl: {
    updTitle: 'Software-update',
    updHeadline: 'Er is een nieuwe versie beschikbaar',
    updDesc:
      'Deze update bevat prestatieverbeteringen en foutoplossingen. We raden aan nu bij te werken.',
    updInstall: 'Opnieuw starten en installeren',
    updLater: 'Later herinneren',
  },
  ms: {
    updTitle: 'Kemas Kini Perisian',
    updHeadline: 'Versi baharu tersedia',
    updDesc:
      'Kemas kini ini merangkumi penambahbaikan prestasi dan pembetulan pepijat. Kami syorkan kemas kini sekarang.',
    updInstall: 'Mula Semula & Pasang',
    updLater: 'Ingatkan kemudian',
  },
  he: {
    updTitle: 'עדכון תוכנה',
    updHeadline: 'גרסה חדשה זמינה',
    updDesc: 'עדכון זה כולל שיפורי ביצועים ותיקוני באגים. מומלץ לעדכן עכשיו.',
    updInstall: 'הפעל מחדש והתקן',
    updLater: 'הזכר לי מאוחר יותר',
  },
  hi: {
    updTitle: 'सॉफ़्टवेयर अपडेट',
    updHeadline: 'नया संस्करण उपलब्ध है',
    updDesc:
      'इस अपडेट में प्रदर्शन सुधार और बग फ़िक्स शामिल हैं। हम अभी अपडेट करने की सलाह देते हैं।',
    updInstall: 'पुनरारंभ करें और इंस्टॉल करें',
    updLater: 'बाद में याद दिलाएँ',
  },
  'zh-TW': {
    updTitle: '軟體更新',
    updHeadline: '發現新版本',
    updDesc: '新版本包含效能改進與問題修復，建議立即更新。',
    updInstall: '立即重新啟動安裝',
    updLater: '稍後再說',
  },
})

const FIRST_CHECK_DELAY_MS = 15_000
const RECHECK_INTERVAL_MS = 4 * 60 * 60 * 1000

let started = false
// version the user declined this session — don't nag again until next launch
let dismissedVersion: string | null = null

function log(...args: unknown[]): void {
  console.log('[docs-updater]', ...args)
}

async function promptInstall(win: BrowserWindow | null, info: UpdateInfo): Promise<void> {
  const lang = getUiLang()
  const options = {
    type: 'info' as const,
    title: tUpd(lang, 'updTitle'),
    message: `${tUpd(lang, 'updHeadline')} (${info.version})`,
    detail: tUpd(lang, 'updDesc'),
    buttons: [tUpd(lang, 'updInstall'), tUpd(lang, 'updLater')],
    defaultId: 0,
    cancelId: 1,
  }
  const { response } =
    win && !win.isDestroyed()
      ? await dialog.showMessageBox(win, options)
      : await dialog.showMessageBox(options)
  if (response === 0) {
    autoUpdater.quitAndInstall(true, true)
  } else {
    // autoInstallOnAppQuit still applies the update on normal quit
    dismissedVersion = info.version
  }
}

export function initDocsAutoUpdater(getWindow: () => BrowserWindow | null): void {
  if (started) return
  started = true

  // Unpacked runs have no app-update.yml and must not hit the CDN with a dev
  // version. Windows updates via NSIS (latest.yml); macOS via the zip target
  // (latest-mac.yml) — Squirrel.Mac additionally requires a signed app, so
  // unsigned local builds check but never install.
  if (!app.isPackaged) return
  if (process.platform !== 'win32' && process.platform !== 'darwin') return

  autoUpdater.autoDownload = true
  autoUpdater.autoInstallOnAppQuit = true
  // full-package policy: never attempt blockmap differential downloads
  // (the publish pipeline does not upload .blockmap files)
  autoUpdater.disableDifferentialDownload = true

  autoUpdater.on('error', (err) => {
    // background check/download failures are expected offline; log only
    log('error:', err instanceof Error ? err.message : err)
  })

  autoUpdater.on('update-downloaded', (info: UpdateInfo) => {
    if (info.version === dismissedVersion) return
    log('downloaded:', info.version)
    promptInstall(getWindow(), info).catch((err) => {
      log('install prompt failed:', err instanceof Error ? err.message : err)
    })
  })

  const check = (): void => {
    autoUpdater
      .checkForUpdates()
      .catch((err) => log('check failed:', err instanceof Error ? err.message : err))
  }
  setTimeout(check, FIRST_CHECK_DELAY_MS)
  setInterval(check, RECHECK_INTERVAL_MS)
}
