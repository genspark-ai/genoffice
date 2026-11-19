import type { Lang } from '@genoffice/i18n'

/** Tooltips for the token-usage chips under assistant replies and on the composer
 * (all editors and the kb-ask panel share them; the chips themselves are numeric). */
export interface AiUsageLabels {
  /** per-reply chip tooltip */
  runTitle: string
  /** session-total chip tooltip */
  sessionTitle: string
  /** tooltip suffix: share of the input served from the prompt cache ({pct}) */
  cacheHit: string
}

export const AI_USAGE_LABELS: Record<Lang, AiUsageLabels> = {
  zh: {
    runTitle: '本回复 token 用量',
    sessionTitle: '本次会话累计用量',
    cacheHit: '缓存命中 {pct}%',
  },
  en: {
    runTitle: 'Token usage for this reply',
    sessionTitle: 'Session token usage',
    cacheHit: '{pct}% from cache',
  },
  ja: {
    runTitle: 'この回答のトークン使用量',
    sessionTitle: 'このセッションの累計',
    cacheHit: 'キャッシュ利用 {pct}%',
  },
  ko: {
    runTitle: '이 답변의 토큰 사용량',
    sessionTitle: '세션 누적 사용량',
    cacheHit: '캐시 적중 {pct}%',
  },
  fr: {
    runTitle: 'Tokens utilisés pour cette réponse',
    sessionTitle: 'Total de tokens de la session',
    cacheHit: '{pct} % depuis le cache',
  },
  de: {
    runTitle: 'Token-Verbrauch dieser Antwort',
    sessionTitle: 'Token-Verbrauch der Sitzung',
    cacheHit: '{pct} % aus dem Cache',
  },
  es: {
    runTitle: 'Uso de tokens de esta respuesta',
    sessionTitle: 'Uso de tokens de la sesión',
    cacheHit: '{pct} % desde caché',
  },
  th: {
    runTitle: 'การใช้โทเค็นของคำตอบนี้',
    sessionTitle: 'โทเค็นรวมของเซสชันนี้',
    cacheHit: 'จากแคช {pct}%',
  },
  id: {
    runTitle: 'Penggunaan token balasan ini',
    sessionTitle: 'Penggunaan token sesi ini',
    cacheHit: '{pct}% dari cache',
  },
  ru: {
    runTitle: 'Токены, использованные в этом ответе',
    sessionTitle: 'Токены за всю сессию',
    cacheHit: '{pct}% из кэша',
  },
  ar: {
    runTitle: 'استخدام الرموز لهذه الإجابة',
    sessionTitle: 'إجمالي الرموز لهذه الجلسة',
    cacheHit: '{pct}% من ذاكرة التخزين المؤقت',
  },
  pt: {
    runTitle: 'Uso de tokens desta resposta',
    sessionTitle: 'Uso de tokens da sessão',
    cacheHit: '{pct}% do cache',
  },
  it: {
    runTitle: 'Token usati per questa risposta',
    sessionTitle: 'Token usati nella sessione',
    cacheHit: '{pct}% dalla cache',
  },
  pl: {
    runTitle: 'Tokeny użyte w tej odpowiedzi',
    sessionTitle: 'Tokeny użyte w tej sesji',
    cacheHit: '{pct}% z pamięci podręcznej',
  },
  cs: {
    runTitle: 'Tokeny použité v této odpovědi',
    sessionTitle: 'Tokeny použité v této relaci',
    cacheHit: '{pct} % z mezipaměti',
  },
  nl: {
    runTitle: 'Tokengebruik van dit antwoord',
    sessionTitle: 'Tokengebruik van deze sessie',
    cacheHit: '{pct}% uit cache',
  },
  ms: {
    runTitle: 'Penggunaan token balasan ini',
    sessionTitle: 'Penggunaan token sesi ini',
    cacheHit: '{pct}% daripada cache',
  },
  he: {
    runTitle: 'צריכת אסימונים לתשובה זו',
    sessionTitle: 'צריכת אסימונים להפעלה זו',
    cacheHit: '{pct}% מהמטמון',
  },
  hi: {
    runTitle: 'इस उत्तर का टोकन उपयोग',
    sessionTitle: 'इस सत्र का कुल टोकन उपयोग',
    cacheHit: 'कैश से {pct}%',
  },
  vi: {
    runTitle: 'Token dùng cho câu trả lời này',
    sessionTitle: 'Token dùng trong phiên này',
    cacheHit: '{pct}% từ bộ nhớ đệm',
  },
  tr: {
    runTitle: 'Bu yanıtın token kullanımı',
    sessionTitle: 'Bu oturumun token kullanımı',
    cacheHit: '%{pct} önbellekten',
  },
  hu: {
    runTitle: 'A válasz tokenfelhasználása',
    sessionTitle: 'A munkamenet tokenfelhasználása',
    cacheHit: '{pct}% a gyorsítótárból',
  },
  'zh-TW': {
    runTitle: '本回覆 token 用量',
    sessionTitle: '本次工作階段累計用量',
    cacheHit: '快取命中 {pct}%',
  },
}
