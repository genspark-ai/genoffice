import { createContext, useCallback, useContext, useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import {
  createI18n,
  htmlDir,
  htmlLang,
  type Lang,
  type LangDicts,
  type Params,
} from '@genoffice/i18n'
import { zh } from './lang/zh'
import { en } from './lang/en'

type Dict = typeof zh

export type StringKey = keyof Dict
export type TFunc = (key: StringKey, params?: Params) => string

// zh and en stay in the entry bundle: createI18n's fallback chain and the
// module-level t() below must answer before any lazy shard resolves. Every
// other locale registers into this registry once its chunk has loaded.
const dicts: Partial<Record<Lang, Record<StringKey, string>>> = { zh, en }
const translate = createI18n(dicts as LangDicts<Dict>)

const LOADERS: { [L in Exclude<Lang, 'zh' | 'en'>]: () => Promise<Record<StringKey, string>> } = {
  ja: () => import('./lang/ja').then((m) => m.ja),
  ko: () => import('./lang/ko').then((m) => m.ko),
  fr: () => import('./lang/fr').then((m) => m.fr),
  de: () => import('./lang/de').then((m) => m.de),
  es: () => import('./lang/es').then((m) => m.es),
  th: () => import('./lang/th').then((m) => m.th),
  id: () => import('./lang/id').then((m) => m.id),
  ru: () => import('./lang/ru').then((m) => m.ru),
  ar: () => import('./lang/ar').then((m) => m.ar),
  pt: () => import('./lang/pt').then((m) => m.pt),
  it: () => import('./lang/it').then((m) => m.it),
  pl: () => import('./lang/pl').then((m) => m.pl),
  cs: () => import('./lang/cs').then((m) => m.cs),
  nl: () => import('./lang/nl').then((m) => m.nl),
  ms: () => import('./lang/ms').then((m) => m.ms),
  he: () => import('./lang/he').then((m) => m.he),
  hi: () => import('./lang/hi').then((m) => m.hi),
  vi: () => import('./lang/vi').then((m) => m.vi),
  'zh-TW': () => import('./lang/zh-TW').then((m) => m.zhTW),
}

const pending = new Map<Lang, Promise<void>>()

/** Register a locale's dictionary before it becomes active; never rejects. */
export function ensureLocale(lang: Lang): Promise<void> {
  if (dicts[lang]) return Promise.resolve()
  const inFlight = pending.get(lang)
  if (inFlight) return inFlight
  const load = LOADERS[lang as Exclude<Lang, 'zh' | 'en'>]()
    .then((dict) => {
      dicts[lang] = dict
    })
    .catch(() => {
      pending.delete(lang)
    })
  pending.set(lang, load)
  return load
}

/** Appended to the AI system prompt: reply in the user's message language, falling back to the UI language (same wording as docs/slides/pdf) */
const AI_LANG_DIRECTIVES: Record<Lang, string> = {
  zh: '\n\n用与用户消息相同的语言回复；无法判断用户消息的语言时，用简体中文回复。',
  en: "\n\nReply in the same language as the user's message; if it cannot be determined, reply in English.",
  ja: '\n\nユーザーのメッセージと同じ言語で返信してください。言語を判別できない場合は日本語で返信してください。',
  ko: '\n\n사용자 메시지와 같은 언어로 답변하세요. 언어를 판단할 수 없으면 한국어로 답변하세요.',
  fr: "\n\nRéponds dans la même langue que le message de l'utilisateur ; si elle ne peut pas être déterminée, réponds en français.",
  de: '\n\nAntworte in derselben Sprache wie die Nachricht des Benutzers; lässt sie sich nicht bestimmen, antworte auf Deutsch.',
  es: '\n\nResponde en el mismo idioma que el mensaje del usuario; si no se puede determinar, responde en español.',
  th: '\n\nตอบเป็นภาษาเดียวกับข้อความของผู้ใช้ หากไม่สามารถระบุได้ ให้ตอบเป็นภาษาไทย',
  id: '\n\nBalas dalam bahasa yang sama dengan pesan pengguna; jika tidak dapat ditentukan, balas dalam bahasa Indonesia.',
  ru: '\n\nОтвечай на том же языке, что и сообщение пользователя; если его невозможно определить, отвечай на русском.',
  ar: '\n\nأجب بنفس لغة رسالة المستخدم؛ وإذا تعذر تحديدها، فأجب باللغة العربية.',
  pt: '\n\nResponda no mesmo idioma da mensagem do usuário; se não for possível determiná-lo, responda em português.',
  it: "\n\nRispondi nella stessa lingua del messaggio dell'utente; se non può essere determinata, rispondi in italiano.",
  pl: '\n\nOdpowiadaj w tym samym języku, co wiadomość użytkownika; jeśli nie da się go ustalić, odpowiadaj po polsku.',
  cs: '\n\nOdpovídej ve stejném jazyce jako zpráva uživatele; pokud ho nelze určit, odpovídej česky.',
  nl: '\n\nAntwoord in dezelfde taal als het bericht van de gebruiker; als die niet te bepalen is, antwoord dan in het Nederlands.',
  ms: '\n\nBalas dalam bahasa yang sama dengan mesej pengguna; jika tidak dapat ditentukan, balas dalam bahasa Melayu.',
  he: '\n\nהשב באותה שפה של הודעת המשתמש; אם לא ניתן לקבוע אותה, השב בעברית.',
  hi: '\n\nउपयोगकर्ता के संदेश की भाषा में ही उत्तर दें; यदि भाषा निर्धारित न हो सके, तो हिंदी में उत्तर दें।',
  'zh-TW': '\n\n用與使用者訊息相同的語言回覆；無法判斷使用者訊息的語言時，用繁體中文回覆。',
  vi: '\n\nTrả lời bằng ngôn ngữ của tin nhắn người dùng; nếu không xác định được, hãy trả lời bằng tiếng Việt.',
}

export function aiLangDirective(lang: Lang): string {
  return AI_LANG_DIRECTIVES[lang]
}

/** The product this panel lives in, in the user's language (the APP skill variable). Same wording as the shell's app titles. */
const AI_APP_NAMES: Record<Lang, string> = {
  zh: 'HTML 编辑器',
  'zh-TW': 'HTML 編輯器',
  en: 'The HTML editor',
  it: "L'editor HTML",
  pl: 'Edytor HTML',
  cs: 'Editor HTML',
  nl: 'De HTML-editor',
  ms: 'Editor HTML',
  ja: 'HTML エディター',
  ko: 'HTML 편집기',
  fr: 'L’éditeur HTML',
  de: 'Der HTML-Editor',
  es: 'El editor HTML',
  th: 'ตัวแก้ไข HTML',
  id: 'Editor HTML',
  ru: 'Редактор HTML',
  ar: 'محرّر HTML',
  pt: 'O editor HTML',
  he: 'עורך ה-HTML',
  hi: 'HTML एडिटर',
  vi: 'Trình soạn thảo HTML',
}

export function aiAppName(lang: Lang): string {
  return AI_APP_NAMES[lang]
}

const LocaleContext = createContext<{ lang: Lang; dictsRev: number }>({
  lang: 'zh',
  dictsRev: 0,
})

/** Module-level current language: for code outliving render closures, kept in sync with the Provider */
let moduleLang: Lang = 'zh'

export function t(key: StringKey, params?: Params): string {
  return translate(moduleLang, key, params)
}

export function LocaleProvider({ initial, children }: { initial: Lang; children: ReactNode }) {
  const [lang, setLang] = useState<Lang>(initial)
  const [dictsRev, setDictsRev] = useState(0)
  const latest = useRef<Lang>(initial)
  moduleLang = lang

  // a locale only becomes active once its dictionary is registered, so no
  // consumer ever renders fallback strings for a locale that has a shard
  const activate = useCallback((next: Lang) => {
    latest.current = next
    if (dicts[next]) {
      setLang(next)
      return
    }
    void ensureLocale(next).then(() => {
      if (latest.current !== next) return
      setLang(next)
      setDictsRev((rev) => rev + 1)
    })
  }, [])

  useEffect(() => {
    activate(initial)
  }, [activate, initial])
  useEffect(
    () =>
      window.htmlApi.onLanguageChanged((next) => {
        document.documentElement.lang = htmlLang(next)
        document.documentElement.dir = htmlDir(next)
        activate(next)
      }),
    [activate],
  )
  return <LocaleContext.Provider value={{ lang, dictsRev }}>{children}</LocaleContext.Provider>
}

export function useI18n(): { lang: Lang; t: TFunc } {
  const { lang } = useContext(LocaleContext)
  return { lang, t: (key, params) => translate(lang, key, params) }
}
