import { app, shell } from 'electron'
import type { BrowserWindow } from 'electron'
import { autoUpdater } from 'electron-updater'
import type { UpdateInfo } from 'electron-updater'
import { createI18n, getUiLang, htmlLang } from '@genoffice/i18n'
import type { UpdateChannel, UpdateUiState, UpdateUiStrings } from '../shared/update-api'
import {
  closeUpdateWindow,
  isUpdateWindowOpen,
  pushUpdateState,
  showUpdateWindow,
} from './update-window'

/**
 * Full-package auto-update over the generic provider (Azure CDN).
 *
 * The release pipeline publishes `latest.yml` + the versioned installer to
 * the update channel prefix (production builds only). The packaged app reads
 * that URL from resources/app-update.yml, which electron-builder bakes in
 * from the `publish` config in apps/shell/electron-builder.cjs — the URL
 * itself is injected at build time via the GENOFFICE_UPDATE_URL env var and
 * is intentionally not committed to the repo.
 *
 * UX is the strong-guidance modal card (update-window.ts), not a native
 * dialog. Windows updates through the NSIS installer (latest.yml); macOS
 * through the zip target (latest-mac.yml); Linux through the AppImage
 * target (latest-linux.yml) — all published by the internal release
 * pipeline. On Linux only AppImage runs self-update (electron-updater
 * replaces the .AppImage file in place, no root needed); deb installs have
 * no updater — users upgrade via `apt install ./<new>.deb`.
 *
 * Dev preview: GENOFFICE_FAKE_UPDATE=<version> in an unpacked run opens the
 * window with a simulated download so the UI can be exercised end to end.
 */

const tUpd = createI18n({
  zh: {
    updTitle: '软件更新',
    updHeadline: '发现新版本',
    updDesc: '新版本包含性能改进与问题修复，建议立即更新。',
    updDownload: '立即更新',
    updLater: '稍后再说',
    updInstall: '立即重启安装',
    updDownloading: '正在下载更新…',
    updFailed: '更新下载失败，请检查网络后重试。',
    updRetry: '重试',
    updManual: '自动更新失败，请从下载页面获取最新版本并手动安装。',
    updOpenDownload: '前往下载页面',
  },
  en: {
    updTitle: 'Software Update',
    updHeadline: 'A new version is available',
    updDesc:
      'This update includes performance improvements and bug fixes. We recommend updating now.',
    updDownload: 'Update Now',
    updLater: 'Remind me later',
    updInstall: 'Restart & Install',
    updDownloading: 'Downloading update…',
    updFailed: 'Update download failed. Check your network and try again.',
    updRetry: 'Retry',
    updManual:
      'Automatic update failed. Please get the latest version from the download page and install it manually.',
    updOpenDownload: 'Open Download Page',
  },
  ja: {
    updTitle: 'ソフトウェアアップデート',
    updHeadline: '新しいバージョンがあります',
    updDesc:
      'このアップデートにはパフォーマンス改善とバグ修正が含まれます。今すぐの更新をおすすめします。',
    updDownload: '今すぐ更新',
    updLater: '後で通知',
    updInstall: '再起動してインストール',
    updDownloading: 'アップデートをダウンロード中…',
    updFailed: 'ダウンロードに失敗しました。ネットワークを確認して再試行してください。',
    updRetry: '再試行',
    updManual:
      '自動更新に失敗しました。ダウンロードページから最新バージョンを取得して手動でインストールしてください。',
    updOpenDownload: 'ダウンロードページを開く',
  },
  ko: {
    updTitle: '소프트웨어 업데이트',
    updHeadline: '새 버전이 있습니다',
    updDesc:
      '이 업데이트에는 성능 개선과 버그 수정이 포함되어 있습니다. 지금 업데이트하는 것을 권장합니다.',
    updDownload: '지금 업데이트',
    updLater: '나중에 알림',
    updInstall: '다시 시작 및 설치',
    updDownloading: '업데이트 다운로드 중…',
    updFailed: '업데이트 다운로드에 실패했습니다. 네트워크를 확인한 후 다시 시도하세요.',
    updRetry: '다시 시도',
    updManual:
      '자동 업데이트에 실패했습니다. 다운로드 페이지에서 최신 버전을 받아 직접 설치해 주세요.',
    updOpenDownload: '다운로드 페이지 열기',
  },
  fr: {
    updTitle: 'Mise à jour logicielle',
    updHeadline: 'Une nouvelle version est disponible',
    updDesc:
      'Cette mise à jour apporte des améliorations de performances et des corrections de bogues. Nous vous recommandons de mettre à jour maintenant.',
    updDownload: 'Mettre à jour',
    updLater: 'Plus tard',
    updInstall: 'Redémarrer et installer',
    updDownloading: 'Téléchargement de la mise à jour…',
    updFailed: 'Échec du téléchargement. Vérifiez votre réseau et réessayez.',
    updRetry: 'Réessayer',
    updManual:
      'La mise à jour automatique a échoué. Téléchargez la dernière version depuis la page de téléchargement et installez-la manuellement.',
    updOpenDownload: 'Ouvrir la page de téléchargement',
  },
  de: {
    updTitle: 'Softwareaktualisierung',
    updHeadline: 'Eine neue Version ist verfügbar',
    updDesc:
      'Dieses Update enthält Leistungsverbesserungen und Fehlerbehebungen. Wir empfehlen, jetzt zu aktualisieren.',
    updDownload: 'Jetzt aktualisieren',
    updLater: 'Später erinnern',
    updInstall: 'Neu starten und installieren',
    updDownloading: 'Update wird heruntergeladen…',
    updFailed:
      'Download fehlgeschlagen. Prüfen Sie Ihre Netzwerkverbindung und versuchen Sie es erneut.',
    updRetry: 'Erneut versuchen',
    updManual:
      'Automatisches Update fehlgeschlagen. Laden Sie die neueste Version von der Download-Seite herunter und installieren Sie sie manuell.',
    updOpenDownload: 'Download-Seite öffnen',
  },
  es: {
    updTitle: 'Actualización de software',
    updHeadline: 'Hay una nueva versión disponible',
    updDesc:
      'Esta actualización incluye mejoras de rendimiento y correcciones de errores. Recomendamos actualizar ahora.',
    updDownload: 'Actualizar ahora',
    updLater: 'Recordar más tarde',
    updInstall: 'Reiniciar e instalar',
    updDownloading: 'Descargando la actualización…',
    updFailed: 'Error al descargar. Compruebe su red e inténtelo de nuevo.',
    updRetry: 'Reintentar',
    updManual:
      'La actualización automática falló. Descargue la última versión desde la página de descargas e instálela manualmente.',
    updOpenDownload: 'Abrir página de descargas',
  },
  th: {
    updTitle: 'อัปเดตซอฟต์แวร์',
    updHeadline: 'มีเวอร์ชันใหม่พร้อมใช้งาน',
    updDesc: 'การอัปเดตนี้มีการปรับปรุงประสิทธิภาพและแก้ไขข้อบกพร่อง แนะนำให้อัปเดตทันที',
    updDownload: 'อัปเดตเลย',
    updLater: 'เตือนภายหลัง',
    updInstall: 'รีสตาร์ทและติดตั้ง',
    updDownloading: 'กำลังดาวน์โหลดอัปเดต…',
    updFailed: 'ดาวน์โหลดไม่สำเร็จ โปรดตรวจสอบเครือข่ายแล้วลองอีกครั้ง',
    updRetry: 'ลองอีกครั้ง',
    updManual:
      'การอัปเดตอัตโนมัติล้มเหลว โปรดดาวน์โหลดเวอร์ชันล่าสุดจากหน้าดาวน์โหลดแล้วติดตั้งด้วยตนเอง',
    updOpenDownload: 'เปิดหน้าดาวน์โหลด',
  },
  id: {
    updTitle: 'Pembaruan Perangkat Lunak',
    updHeadline: 'Versi baru tersedia',
    updDesc:
      'Pembaruan ini mencakup peningkatan kinerja dan perbaikan bug. Kami menyarankan untuk memperbarui sekarang.',
    updDownload: 'Perbarui Sekarang',
    updLater: 'Ingatkan nanti',
    updInstall: 'Mulai Ulang & Pasang',
    updDownloading: 'Mengunduh pembaruan…',
    updFailed: 'Unduhan gagal. Periksa jaringan Anda dan coba lagi.',
    updRetry: 'Coba Lagi',
    updManual:
      'Pembaruan otomatis gagal. Silakan unduh versi terbaru dari halaman unduhan dan pasang secara manual.',
    updOpenDownload: 'Buka Halaman Unduhan',
  },
  ru: {
    updTitle: 'Обновление программы',
    updHeadline: 'Доступна новая версия',
    updDesc:
      'Это обновление содержит улучшения производительности и исправления ошибок. Рекомендуем обновиться сейчас.',
    updDownload: 'Обновить сейчас',
    updLater: 'Напомнить позже',
    updInstall: 'Перезапустить и установить',
    updDownloading: 'Загрузка обновления…',
    updFailed: 'Не удалось загрузить обновление. Проверьте сеть и повторите попытку.',
    updRetry: 'Повторить',
    updManual:
      'Автоматическое обновление не удалось. Скачайте последнюю версию со страницы загрузки и установите её вручную.',
    updOpenDownload: 'Открыть страницу загрузки',
  },
  ar: {
    updTitle: 'تحديث البرنامج',
    updHeadline: 'يتوفر إصدار جديد',
    updDesc: 'يتضمن هذا التحديث تحسينات في الأداء وإصلاحات للأخطاء. نوصي بالتحديث الآن.',
    updDownload: 'التحديث الآن',
    updLater: 'ذكّرني لاحقًا',
    updInstall: 'إعادة التشغيل والتثبيت',
    updDownloading: 'جارٍ تنزيل التحديث…',
    updFailed: 'فشل تنزيل التحديث. تحقق من الشبكة وحاول مرة أخرى.',
    updRetry: 'إعادة المحاولة',
    updManual: 'فشل التحديث التلقائي. يرجى تنزيل أحدث إصدار من صفحة التنزيل وتثبيته يدويًا.',
    updOpenDownload: 'فتح صفحة التنزيل',
  },
  pt: {
    updTitle: 'Atualização de Software',
    updHeadline: 'Uma nova versão está disponível',
    updDesc:
      'Esta atualização inclui melhorias de desempenho e correções de erros. Recomendamos atualizar agora.',
    updDownload: 'Atualizar agora',
    updLater: 'Lembrar mais tarde',
    updInstall: 'Reiniciar e instalar',
    updDownloading: 'Baixando a atualização…',
    updFailed: 'Falha no download. Verifique sua rede e tente novamente.',
    updRetry: 'Tentar novamente',
    updManual:
      'A atualização automática falhou. Baixe a versão mais recente na página de download e instale manualmente.',
    updOpenDownload: 'Abrir página de download',
  },
  it: {
    updTitle: 'Aggiornamento software',
    updHeadline: 'È disponibile una nuova versione',
    updDesc:
      'Questo aggiornamento include miglioramenti delle prestazioni e correzioni di bug. Consigliamo di aggiornare subito.',
    updDownload: 'Aggiorna ora',
    updLater: 'Ricordamelo più tardi',
    updInstall: 'Riavvia e installa',
    updDownloading: "Download dell'aggiornamento…",
    updFailed: 'Download non riuscito. Controlla la rete e riprova.',
    updRetry: 'Riprova',
    updManual:
      "Aggiornamento automatico non riuscito. Scarica l'ultima versione dalla pagina di download e installala manualmente.",
    updOpenDownload: 'Apri pagina di download',
  },
  pl: {
    updTitle: 'Aktualizacja oprogramowania',
    updHeadline: 'Dostępna jest nowa wersja',
    updDesc:
      'Ta aktualizacja zawiera ulepszenia wydajności i poprawki błędów. Zalecamy aktualizację teraz.',
    updDownload: 'Aktualizuj teraz',
    updLater: 'Przypomnij później',
    updInstall: 'Uruchom ponownie i zainstaluj',
    updDownloading: 'Pobieranie aktualizacji…',
    updFailed: 'Pobieranie nie powiodło się. Sprawdź sieć i spróbuj ponownie.',
    updRetry: 'Spróbuj ponownie',
    updManual:
      'Automatyczna aktualizacja nie powiodła się. Pobierz najnowszą wersję ze strony pobierania i zainstaluj ją ręcznie.',
    updOpenDownload: 'Otwórz stronę pobierania',
  },
  nl: {
    updTitle: 'Software-update',
    updHeadline: 'Er is een nieuwe versie beschikbaar',
    updDesc:
      'Deze update bevat prestatieverbeteringen en foutoplossingen. We raden aan nu bij te werken.',
    updDownload: 'Nu bijwerken',
    updLater: 'Later herinneren',
    updInstall: 'Opnieuw starten en installeren',
    updDownloading: 'Update wordt gedownload…',
    updFailed: 'Download mislukt. Controleer uw netwerk en probeer het opnieuw.',
    updRetry: 'Opnieuw proberen',
    updManual:
      'Automatische update mislukt. Download de nieuwste versie via de downloadpagina en installeer deze handmatig.',
    updOpenDownload: 'Downloadpagina openen',
  },
  ms: {
    updTitle: 'Kemas Kini Perisian',
    updHeadline: 'Versi baharu tersedia',
    updDesc:
      'Kemas kini ini merangkumi penambahbaikan prestasi dan pembetulan pepijat. Kami syorkan kemas kini sekarang.',
    updDownload: 'Kemas Kini Sekarang',
    updLater: 'Ingatkan kemudian',
    updInstall: 'Mula Semula & Pasang',
    updDownloading: 'Memuat turun kemas kini…',
    updFailed: 'Muat turun gagal. Semak rangkaian anda dan cuba lagi.',
    updRetry: 'Cuba Lagi',
    updManual:
      'Kemas kini automatik gagal. Sila muat turun versi terkini dari halaman muat turun dan pasang secara manual.',
    updOpenDownload: 'Buka Halaman Muat Turun',
  },
  he: {
    updTitle: 'עדכון תוכנה',
    updHeadline: 'גרסה חדשה זמינה',
    updDesc: 'עדכון זה כולל שיפורי ביצועים ותיקוני באגים. מומלץ לעדכן עכשיו.',
    updDownload: 'עדכן עכשיו',
    updLater: 'הזכר לי מאוחר יותר',
    updInstall: 'הפעל מחדש והתקן',
    updDownloading: 'מוריד את העדכון…',
    updFailed: 'ההורדה נכשלה. בדוק את הרשת ונסה שוב.',
    updRetry: 'נסה שוב',
    updManual: 'העדכון האוטומטי נכשל. הורד את הגרסה העדכנית מדף ההורדות והתקן אותה ידנית.',
    updOpenDownload: 'פתח את דף ההורדות',
  },
  hi: {
    updTitle: 'सॉफ़्टवेयर अपडेट',
    updHeadline: 'नया संस्करण उपलब्ध है',
    updDesc:
      'इस अपडेट में प्रदर्शन सुधार और बग फ़िक्स शामिल हैं। हम अभी अपडेट करने की सलाह देते हैं।',
    updDownload: 'अभी अपडेट करें',
    updLater: 'बाद में याद दिलाएँ',
    updInstall: 'पुनरारंभ करें और इंस्टॉल करें',
    updDownloading: 'अपडेट डाउनलोड हो रहा है…',
    updFailed: 'डाउनलोड विफल रहा। अपना नेटवर्क जाँचें और पुनः प्रयास करें।',
    updRetry: 'पुनः प्रयास करें',
    updManual:
      'स्वचालित अपडेट विफल रहा। कृपया डाउनलोड पृष्ठ से नवीनतम संस्करण प्राप्त करें और मैन्युअल रूप से इंस्टॉल करें।',
    updOpenDownload: 'डाउनलोड पृष्ठ खोलें',
  },
  'zh-TW': {
    updTitle: '軟體更新',
    updHeadline: '發現新版本',
    updDesc: '新版本包含效能改進與問題修復，建議立即更新。',
    updDownload: '立即更新',
    updLater: '稍後再說',
    updInstall: '立即重新啟動安裝',
    updDownloading: '正在下載更新…',
    updFailed: '更新下載失敗，請檢查網路後重試。',
    updRetry: '重試',
    updManual: '自動更新失敗，請從下載頁面取得最新版本並手動安裝。',
    updOpenDownload: '前往下載頁面',
  },
})

const FIRST_CHECK_DELAY_MS = 15_000
const RECHECK_INTERVAL_MS = 4 * 60 * 60 * 1000

// After this many failed download/apply attempts for the same version the
// dialog stops offering "retry" and guides the user to a manual download
// instead. Covers permanently broken update paths — most importantly a
// code-signing identity (Apple Team ID) change, which Squirrel.Mac rejects
// on every retry while the error looks like a download failure to the user.
const MANUAL_FALLBACK_AFTER = 2
const DOWNLOAD_PAGE_URL = 'https://github.com/genspark-ai/genoffice/releases/latest'

let started = false
// version the user declined this session — don't nag again until next launch
let dismissedVersion: string | null = null

// electron-updater feed name per user-facing channel. The platform suffix is
// appended by electron-updater itself: 'beta' resolves to beta.yml on
// Windows, beta-mac.yml on macOS, beta-linux.yml on Linux x64.
const CHANNEL_FEED: Record<UpdateChannel, string> = { stable: 'latest', beta: 'beta' }

// true once the packaged-run updater is configured; channel switches before
// that (or in dev runs) must not touch electron-updater
let updaterActive = false

function log(...args: unknown[]): void {
  console.log('[updater]', ...args)
}

function uiStrings(): UpdateUiStrings {
  const lang = getUiLang()
  return {
    title: tUpd(lang, 'updTitle'),
    headline: tUpd(lang, 'updHeadline'),
    desc: tUpd(lang, 'updDesc'),
    download: tUpd(lang, 'updDownload'),
    later: tUpd(lang, 'updLater'),
    install: tUpd(lang, 'updInstall'),
    downloading: tUpd(lang, 'updDownloading'),
    failed: tUpd(lang, 'updFailed'),
    retry: tUpd(lang, 'updRetry'),
    manualDesc: tUpd(lang, 'updManual'),
    openDownload: tUpd(lang, 'updOpenDownload'),
  }
}

function initialState(version: string): UpdateUiState {
  return {
    phase: 'available',
    version,
    currentVersion: app.getVersion(),
    percent: 0,
    lang: htmlLang(getUiLang()),
    strings: uiStrings(),
  }
}

export function applyUpdateChannel(channel: UpdateChannel): void {
  if (!updaterActive) return
  autoUpdater.channel = CHANNEL_FEED[channel]
  // the channel setter unconditionally flips allowDowngrade to true; force it
  // back off since a beta user switching to stable must not downgrade
  autoUpdater.allowDowngrade = false
  log('channel switched:', channel)
  autoUpdater.checkForUpdates().catch((err) => log('check failed:', err?.message ?? err))
}

export function initAutoUpdater(
  getWindow: () => BrowserWindow | null,
  initialChannel: UpdateChannel = 'stable',
): void {
  if (started) return
  started = true

  // dev preview of the update window with a simulated download
  if (!app.isPackaged && process.env.GENOFFICE_FAKE_UPDATE) {
    initFakeUpdate(getWindow, process.env.GENOFFICE_FAKE_UPDATE)
    return
  }
  // Unpacked runs have no app-update.yml and must not hit the CDN with a
  // dev version. Windows updates via NSIS (latest.yml), macOS via the zip
  // target + latest-mac.yml (Squirrel.Mac requires a signed, notarized app
  // — dmg is first-install only), Linux via the AppImage target +
  // latest-linux.yml. On Linux the updater only works for AppImage runs
  // (electron-updater's AppImageUpdater needs the APPIMAGE env var the
  // AppImage runtime sets); deb installs update manually via apt.
  if (!app.isPackaged) return
  const isLinuxAppImage = process.platform === 'linux' && Boolean(process.env.APPIMAGE)
  if (process.platform !== 'win32' && process.platform !== 'darwin' && !isLinuxAppImage) return

  updaterActive = true
  autoUpdater.channel = CHANNEL_FEED[initialChannel]
  // the channel setter unconditionally flips allowDowngrade to true; force it
  // back off since a beta user switching to stable must not downgrade
  autoUpdater.allowDowngrade = false
  autoUpdater.autoDownload = false
  // if the user picked "later" after download, install on normal quit
  autoUpdater.autoInstallOnAppQuit = true
  // full-package policy: never attempt blockmap differential downloads
  // (CI does not publish .blockmap files)
  autoUpdater.disableDifferentialDownload = true

  let latestSeenVersion: string | null = null
  // consecutive failed attempts for latestSeenVersion; a download can fail
  // through the downloadUpdate() rejection OR only through the 'error' event
  // (macOS: Squirrel.Mac reports signature/apply failures natively), so both
  // paths funnel into failDownload() and the in-flight flag dedupes them
  let failedAttempts = 0
  let downloadInFlight = false

  const failDownload = (): void => {
    if (!downloadInFlight) return
    downloadInFlight = false
    failedAttempts += 1
    pushUpdateState({ phase: failedAttempts >= MANUAL_FALLBACK_AFTER ? 'manual' : 'error' })
  }

  const actions = {
    onDownload: () => {
      downloadInFlight = true
      pushUpdateState({ phase: 'downloading', percent: 0 })
      autoUpdater.downloadUpdate().catch((err) => {
        log('download failed:', err?.message ?? err)
        failDownload()
      })
    },
    onInstall: () => {
      closeUpdateWindow()
      // let the window fully close before tearing the app down
      setImmediate(() => autoUpdater.quitAndInstall(true, true))
    },
    onLater: () => {
      dismissedVersion = latestSeenVersion
      closeUpdateWindow()
    },
    onOpenDownload: () => {
      void shell.openExternal(DOWNLOAD_PAGE_URL)
    },
  }

  autoUpdater.on('error', (err) => {
    // network failures during background checks are expected; only surface
    // when the user is watching a download
    log('error:', err?.message ?? err)
    failDownload()
  })

  autoUpdater.on('update-available', (info: UpdateInfo) => {
    if (info.version === dismissedVersion) return
    const sameVersionRecheck = info.version === latestSeenVersion
    if (!sameVersionRecheck) failedAttempts = 0
    latestSeenVersion = info.version
    log('update available:', info.version)
    // a periodic recheck resolving to the version the open dialog already
    // shows must not reset its phase to 'available' — that would wipe an
    // in-progress download or a terminal 'manual' fallback back to the
    // "Update Now" offer
    if (sameVersionRecheck && isUpdateWindowOpen()) return
    showUpdateWindow(getWindow(), initialState(info.version), actions)
  })

  autoUpdater.on('download-progress', (progress) => {
    pushUpdateState({ phase: 'downloading', percent: progress.percent })
  })

  autoUpdater.on('update-downloaded', (info: UpdateInfo) => {
    log('downloaded:', info.version)
    downloadInFlight = false
    failedAttempts = 0
    pushUpdateState({ phase: 'downloaded', percent: 100 })
  })

  const check = (): void => {
    autoUpdater.checkForUpdates().catch((err) => log('check failed:', err?.message ?? err))
  }
  setTimeout(check, FIRST_CHECK_DELAY_MS)
  setInterval(check, RECHECK_INTERVAL_MS)
}

/** unpacked-run simulation: real window + IPC, fake download that completes */
function initFakeUpdate(getWindow: () => BrowserWindow | null, version: string): void {
  let timer: NodeJS.Timeout | null = null
  const actions = {
    onDownload: () => {
      let pct = 0
      pushUpdateState({ phase: 'downloading', percent: 0 })
      timer = setInterval(() => {
        pct += 4
        if (pct >= 100) {
          if (timer) clearInterval(timer)
          pushUpdateState({ phase: 'downloaded', percent: 100 })
        } else {
          pushUpdateState({ phase: 'downloading', percent: pct })
        }
      }, 100)
    },
    onInstall: () => {
      log('[fake] install requested')
      closeUpdateWindow()
    },
    onLater: () => {
      if (timer) clearInterval(timer)
      closeUpdateWindow()
    },
    onOpenDownload: () => {
      log('[fake] open download page requested')
    },
  }
  setTimeout(() => showUpdateWindow(getWindow(), initialState(version), actions), 1500)
}
