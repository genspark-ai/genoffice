import type { Lang } from '@genoffice/i18n'

/** Tooltips for the token-usage chips under assistant replies and on the composer
 * (all editors and the kb-ask panel share them; the chips themselves are numeric). */
export interface AiUsageLabels {
  /** per-reply chip tooltip */
  runTitle: string
  /** session-total chip tooltip */
  sessionTitle: string
}

export const AI_USAGE_LABELS: Record<Lang, AiUsageLabels> = {
  zh: {
    runTitle: '本回复 token 用量',
    sessionTitle: '本次会话累计用量',
  },
  en: {
    runTitle: 'Token usage for this reply',
    sessionTitle: 'Session token usage',
  },
  ja: {
    runTitle: 'この回答のトークン使用量',
    sessionTitle: 'このセッションの累計',
  },
  ko: {
    runTitle: '이 답변의 토큰 사용량',
    sessionTitle: '세션 누적 사용량',
  },
  fr: {
    runTitle: 'Tokens utilisés pour cette réponse',
    sessionTitle: 'Total de tokens de la session',
  },
  de: {
    runTitle: 'Token-Verbrauch dieser Antwort',
    sessionTitle: 'Token-Verbrauch der Sitzung',
  },
  es: {
    runTitle: 'Uso de tokens de esta respuesta',
    sessionTitle: 'Uso de tokens de la sesión',
  },
  th: {
    runTitle: 'การใช้โทเค็นของคำตอบนี้',
    sessionTitle: 'โทเค็นรวมของเซสชันนี้',
  },
  id: {
    runTitle: 'Penggunaan token balasan ini',
    sessionTitle: 'Penggunaan token sesi ini',
  },
  ru: {
    runTitle: 'Токены, использованные в этом ответе',
    sessionTitle: 'Токены за всю сессию',
  },
  ar: {
    runTitle: 'استخدام الرموز لهذه الإجابة',
    sessionTitle: 'إجمالي الرموز لهذه الجلسة',
  },
  pt: {
    runTitle: 'Uso de tokens desta resposta',
    sessionTitle: 'Uso de tokens da sessão',
  },
  it: {
    runTitle: 'Token usati per questa risposta',
    sessionTitle: 'Token usati nella sessione',
  },
  pl: {
    runTitle: 'Tokeny użyte w tej odpowiedzi',
    sessionTitle: 'Tokeny użyte w tej sesji',
  },
  cs: {
    runTitle: 'Tokeny použité v této odpovědi',
    sessionTitle: 'Tokeny použité v této relaci',
  },
  nl: {
    runTitle: 'Tokengebruik van dit antwoord',
    sessionTitle: 'Tokengebruik van deze sessie',
  },
  ms: {
    runTitle: 'Penggunaan token balasan ini',
    sessionTitle: 'Penggunaan token sesi ini',
  },
  he: {
    runTitle: 'צריכת אסימונים לתשובה זו',
    sessionTitle: 'צריכת אסימונים להפעלה זו',
  },
  hi: {
    runTitle: 'इस उत्तर का टोकन उपयोग',
    sessionTitle: 'इस सत्र का कुल टोकन उपयोग',
  },
  vi: {
    runTitle: 'Token dùng cho câu trả lời này',
    sessionTitle: 'Token dùng trong phiên này',
  },
  'zh-TW': {
    runTitle: '本回覆 token 用量',
    sessionTitle: '本次工作階段累計用量',
  },
}
