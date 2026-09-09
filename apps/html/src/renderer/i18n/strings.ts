import { aiStrings } from './strings-ai'
import { appStrings } from './strings-app'

export const strings = {
  zh: { ...appStrings.zh, ...aiStrings.zh },
  en: { ...appStrings.en, ...aiStrings.en },
  ja: { ...appStrings.ja, ...aiStrings.ja },
  ko: { ...appStrings.ko, ...aiStrings.ko },
  fr: { ...appStrings.fr, ...aiStrings.fr },
  de: { ...appStrings.de, ...aiStrings.de },
  es: { ...appStrings.es, ...aiStrings.es },
  th: { ...appStrings.th, ...aiStrings.th },
  id: { ...appStrings.id, ...aiStrings.id },
  ru: { ...appStrings.ru, ...aiStrings.ru },
  ar: { ...appStrings.ar, ...aiStrings.ar },
  pt: { ...appStrings.pt, ...aiStrings.pt },
  it: { ...appStrings.it, ...aiStrings.it },
  pl: { ...appStrings.pl, ...aiStrings.pl },
  cs: { ...appStrings.cs, ...aiStrings.cs },
  nl: { ...appStrings.nl, ...aiStrings.nl },
  ms: { ...appStrings.ms, ...aiStrings.ms },
  he: { ...appStrings.he, ...aiStrings.he },
  hi: { ...appStrings.hi, ...aiStrings.hi },
  'zh-TW': { ...appStrings['zh-TW'], ...aiStrings['zh-TW'] },
}
