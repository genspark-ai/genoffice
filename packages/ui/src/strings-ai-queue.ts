import type { Lang } from '@genoffice/i18n'

/** Labels for the queued-messages strip above the AI composer (all editors share them). */
export interface AiQueueStripLabels {
  /** collapsed strip title, e.g. "3 queued" */
  queuedCount: (n: number) => string
  clearTitle: string
  editTitle: string
  removeTitle: string
  /** composer placeholder while a reply runs and Enter queues instead of sending */
  queuePlaceholder: string
}

export const AI_QUEUE_LABELS: Record<Lang, AiQueueStripLabels> = {
  zh: {
    queuedCount: (n) => `${n} 条排队中`,
    clearTitle: '清空排队消息',
    editTitle: '编辑排队消息',
    removeTitle: '移除排队消息',
    queuePlaceholder: 'Enter 排队 · Esc 停止',
  },
  en: {
    queuedCount: (n) => `${n} queued`,
    clearTitle: 'Clear queued messages',
    editTitle: 'Edit queued message',
    removeTitle: 'Remove queued message',
    queuePlaceholder: 'Enter to queue · Esc to stop',
  },
  ja: {
    queuedCount: (n) => `${n} 件が待機中`,
    clearTitle: '待機中のメッセージを削除',
    editTitle: '待機中のメッセージを編集',
    removeTitle: '待機中のメッセージを取り除く',
    queuePlaceholder: 'Enter で待機 · Esc で停止',
  },
  ko: {
    queuedCount: (n) => `${n}개 대기 중`,
    clearTitle: '대기 중인 메시지 비우기',
    editTitle: '대기 중인 메시지 편집',
    removeTitle: '대기 중인 메시지 제거',
    queuePlaceholder: 'Enter 대기 · Esc 중지',
  },
  fr: {
    queuedCount: (n) => `${n} en attente`,
    clearTitle: 'Vider les messages en attente',
    editTitle: 'Modifier le message en attente',
    removeTitle: 'Retirer le message en attente',
    queuePlaceholder: 'Entrée pour mettre en attente · Échap pour arrêter',
  },
  de: {
    queuedCount: (n) => `${n} in Warteschlange`,
    clearTitle: 'Warteschlange leeren',
    editTitle: 'Nachricht in Warteschlange bearbeiten',
    removeTitle: 'Nachricht aus Warteschlange entfernen',
    queuePlaceholder: 'Enter zum Einreihen · Esc zum Stoppen',
  },
  es: {
    queuedCount: (n) => `${n} en cola`,
    clearTitle: 'Vaciar mensajes en cola',
    editTitle: 'Editar mensaje en cola',
    removeTitle: 'Quitar mensaje en cola',
    queuePlaceholder: 'Entrada para encolar · Esc para detener',
  },
  th: {
    queuedCount: (n) => `กำลังรออยู่ ${n} รายการ`,
    clearTitle: 'ล้างข้อความที่รออยู่',
    editTitle: 'แก้ไขข้อความที่รออยู่',
    removeTitle: 'นำข้อความที่รออยู่ออก',
    queuePlaceholder: 'Enter เพื่อจัดคิว · Esc เพื่อหยุด',
  },
  id: {
    queuedCount: (n) => `${n} dalam antrean`,
    clearTitle: 'Kosongkan pesan yang mengantre',
    editTitle: 'Edit pesan yang mengantre',
    removeTitle: 'Hapus pesan yang mengantre',
    queuePlaceholder: 'Enter untuk mengantre · Esc untuk berhenti',
  },
  ru: {
    queuedCount: (n) => `${n} в очереди`,
    clearTitle: 'Очистить очередь сообщений',
    editTitle: 'Изменить сообщение в очереди',
    removeTitle: 'Убрать сообщение из очереди',
    queuePlaceholder: 'Enter — в очередь · Esc — остановить',
  },
  ar: {
    queuedCount: (n) => `${n} في قائمة الانتظار`,
    clearTitle: 'إفراغ الرسائل المنتظرة',
    editTitle: 'تعديل الرسالة المنتظرة',
    removeTitle: 'إزالة الرسالة المنتظرة',
    queuePlaceholder: 'Enter للانتظار · Esc للإيقاف',
  },
  pt: {
    queuedCount: (n) => `${n} na fila`,
    clearTitle: 'Limpar mensagens na fila',
    editTitle: 'Editar mensagem na fila',
    removeTitle: 'Remover mensagem da fila',
    queuePlaceholder: 'Enter para enfileirar · Esc para parar',
  },
  it: {
    queuedCount: (n) => `${n} in coda`,
    clearTitle: 'Svuota i messaggi in coda',
    editTitle: 'Modifica il messaggio in coda',
    removeTitle: 'Rimuovi il messaggio in coda',
    queuePlaceholder: 'Invio per accodare · Esc per fermare',
  },
  pl: {
    queuedCount: (n) => `${n} w kolejce`,
    clearTitle: 'Wyczyść wiadomości w kolejce',
    editTitle: 'Edytuj wiadomość w kolejce',
    removeTitle: 'Usuń wiadomość z kolejki',
    queuePlaceholder: 'Enter, aby zakolejkować · Esc, aby zatrzymać',
  },
  cs: {
    queuedCount: (n) => `${n} ve frontě`,
    clearTitle: 'Vymazat zprávy ve frontě',
    editTitle: 'Upravit zprávu ve frontě',
    removeTitle: 'Odebrat zprávu z fronty',
    queuePlaceholder: 'Enter zařadí · Esc zastaví',
  },
  nl: {
    queuedCount: (n) => `${n} in wachtrij`,
    clearTitle: 'Wachtrij leegmaken',
    editTitle: 'Bericht in wachtrij bewerken',
    removeTitle: 'Bericht uit wachtrij verwijderen',
    queuePlaceholder: 'Enter om te wachtrijken · Esc om te stoppen',
  },
  ms: {
    queuedCount: (n) => `${n} dalam baris gilir`,
    clearTitle: 'Kosongkan mesej dalam baris gilir',
    editTitle: 'Edit mesej dalam baris gilir',
    removeTitle: 'Buang mesej dalam baris gilir',
    queuePlaceholder: 'Enter untuk baris gilir · Esc untuk berhenti',
  },
  he: {
    queuedCount: (n) => `${n} בתור`,
    clearTitle: 'נקה הודעות בתור',
    editTitle: 'ערוך הודעה בתור',
    removeTitle: 'הסר הודעה מהתור',
    queuePlaceholder: 'Enter לתור · Esc לעצירה',
  },
  hi: {
    queuedCount: (n) => `${n} कतार में`,
    clearTitle: 'कतार में संदेश साफ़ करें',
    editTitle: 'कतार में संदेश संपादित करें',
    removeTitle: 'कतार में संदेश हटाएँ',
    queuePlaceholder: 'कतार के लिए Enter · रोकने के लिए Esc',
  },
  'zh-TW': {
    queuedCount: (n) => `${n} 則排隊中`,
    clearTitle: '清空排隊訊息',
    editTitle: '編輯排隊訊息',
    removeTitle: '移除排隊訊息',
    queuePlaceholder: 'Enter 排隊 · Esc 停止',
  },
  vi: {
    queuedCount: (n) => `${n} đang chờ`,
    clearTitle: 'Xoá các tin nhắn đang chờ',
    editTitle: 'Sửa tin nhắn đang chờ',
    removeTitle: 'Bỏ tin nhắn đang chờ',
    queuePlaceholder: 'Enter để xếp hàng · Esc để dừng',
  },
}
