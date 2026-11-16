import type { Lang } from '@genoffice/i18n'

/** Status-bar web-share chip (all editors): three strings — the off-state
 *  label, its tooltip, and the stop button's tooltip. The running chip shows
 *  the port itself (locale-free) and copies the share URL on click. */
export interface WebShareLabels {
  /** chip label while the service is off */
  share: string
  /** tooltip while off: what pressing will do */
  shareTip: string
  /** tooltip on the ✕: stop serving */
  stopTip: string
}

export const WEB_SHARE_LABELS: Record<Lang, WebShareLabels> = {
  zh: {
    share: '共享',
    shareTip: '在本地网络开启浏览器访问,并复制分享链接',
    stopTip: '停止共享(浏览器会话将断开)',
  },
  en: {
    share: 'Share',
    shareTip: 'Start browser access on the local network and copy the share link',
    stopTip: 'Stop sharing (browser sessions disconnect)',
  },
  ja: {
    share: '共有',
    shareTip: 'ローカルネットワークでブラウザアクセスを開始し、リンクをコピーします',
    stopTip: '共有を停止(ブラウザセッションが切断されます)',
  },
  ko: {
    share: '공유',
    shareTip: '로컬 네트워크에서 브라우저 접근을 시작하고 공유 링크를 복사합니다',
    stopTip: '공유 중지(브라우저 세션 연결 해제)',
  },
  fr: {
    share: 'Partager',
    shareTip: "Activer l'accès navigateur sur le réseau local et copier le lien",
    stopTip: 'Arrêter le partage (les sessions navigateur sont déconnectées)',
  },
  de: {
    share: 'Freigeben',
    shareTip: 'Browser-Zugriff im lokalen Netzwerk starten und Link kopieren',
    stopTip: 'Freigabe beenden (Browsersitzungen werden getrennt)',
  },
  es: {
    share: 'Compartir',
    shareTip: 'Iniciar acceso por navegador en la red local y copiar el enlace',
    stopTip: 'Dejar de compartir (las sesiones del navegador se desconectan)',
  },
  th: {
    share: 'แชร์',
    shareTip: 'เปิดให้เข้าถึงผ่านเบราว์เซอร์ในเครือข่ายภายในและคัดลอกลิงก์',
    stopTip: 'หยุดแชร์ (เซสชันเบราว์เซอร์จะถูกตัดการเชื่อมต่อ)',
  },
  id: {
    share: 'Bagikan',
    shareTip: 'Mulai akses peramban di jaringan lokal dan salin tautan',
    stopTip: 'Berhenti membagikan (sesi peramban terputus)',
  },
  ru: {
    share: 'Поделиться',
    shareTip: 'Включить доступ из браузера в локальной сети и скопировать ссылку',
    stopTip: 'Остановить общий доступ (сеансы браузера будут отключены)',
  },
  ar: {
    share: 'مشاركة',
    shareTip: 'بدء الوصول عبر المتصفح على الشبكة المحلية ونسخ الرابط',
    stopTip: 'إيقاف المشاركة (ستُغلق جلسات المتصفح)',
  },
  pt: {
    share: 'Compartilhar',
    shareTip: 'Iniciar acesso pelo navegador na rede local e copiar o link',
    stopTip: 'Parar de compartilhar (as sessões do navegador serão desconectadas)',
  },
  it: {
    share: 'Condividi',
    shareTip: 'Attiva l accesso dal browser sulla rete locale e copia il collegamento',
    stopTip: 'Interrompi la condivisione (le sessioni del browser si disconnettono)',
  },
  pl: {
    share: 'Udostępnij',
    shareTip: 'Włącz dostęp z przeglądarki w sieci lokalnej i skopiuj link',
    stopTip: 'Zatrzymaj udostępnianie (sesje przeglądarki zostaną rozłączone)',
  },
  cs: {
    share: 'Sdílet',
    shareTip: 'Zapnout přístup z prohlížeče v místní síti a zkopírovat odkaz',
    stopTip: 'Ukončit sdílení (relace prohlížeče se odpojí)',
  },
  nl: {
    share: 'Delen',
    shareTip: 'Browser toegang op het lokale netwerk starten en de link kopiëren',
    stopTip: 'Delen stoppen (browsersessies worden verbroken)',
  },
  ms: {
    share: 'Kongsi',
    shareTip: 'Mula akses pelayar pada rangkaian setempat dan salin pautan',
    stopTip: 'Berhenti kongsi (sesi pelayar akan diputuskan)',
  },
  he: {
    share: 'שיתוף',
    shareTip: 'הפעלת גישת דפדפן ברשת המקומית והעתקת הקישור',
    stopTip: 'הפסקת שיתוף (פעילויות הדפדפן ינותקו)',
  },
  hi: {
    share: 'साझा करें',
    shareTip: 'स्थानीय नेटवर्क पर ब्राउज़र पहुँच शुरू करें और लिंक कॉपी करें',
    stopTip: 'साझा करना बंद करें (ब्राउज़र सत्र डिस्कनेक्ट हो जाएँगे)',
  },
  'zh-TW': {
    share: '共享',
    shareTip: '在區域網路開啟瀏覽器存取,並複製分享連結',
    stopTip: '停止共享(瀏覽器工作階段將中斷)',
  },
  vi: {
    share: 'Chia sẻ',
    shareTip: 'Bắt đầu truy cập trình duyệt trên mạng cục bộ và sao chép liên kết',
    stopTip: 'Dừng chia sẻ (các phiên trình duyệt sẽ ngắt kết nối)',
  },
  tr: {
    share: 'Paylaş',
    shareTip: 'Yerel ağda tarayıcı erişimini başlat ve bağlantıyı kopyala',
    stopTip: 'Paylaşımı durdur (tarayıcı oturumlarının bağlantısı kesilir)',
  },
  hu: {
    share: 'Megosztás',
    shareTip: 'Indíts böngészős elérést a helyi hálózaton, és másolja a hivatkozást',
    stopTip: 'Megosztás leállítása (a böngészős munkamenetek lekapcsolnak)',
  },
}
