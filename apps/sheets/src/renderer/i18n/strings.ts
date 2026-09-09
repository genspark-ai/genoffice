import { aiStrings } from './strings-ai'
import { appStrings } from './strings-app'
import { dialogStrings } from './strings-dialogs'

export const strings = {
  zh: { ...appStrings.zh, ...dialogStrings.zh, ...aiStrings.zh },
  en: { ...appStrings.en, ...dialogStrings.en, ...aiStrings.en },
  ja: { ...appStrings.ja, ...dialogStrings.ja, ...aiStrings.ja },
  ko: { ...appStrings.ko, ...dialogStrings.ko, ...aiStrings.ko },
  fr: { ...appStrings.fr, ...dialogStrings.fr, ...aiStrings.fr },
  de: { ...appStrings.de, ...dialogStrings.de, ...aiStrings.de },
  es: { ...appStrings.es, ...dialogStrings.es, ...aiStrings.es },
  th: { ...appStrings.th, ...dialogStrings.th, ...aiStrings.th },
  id: { ...appStrings.id, ...dialogStrings.id, ...aiStrings.id },
  ru: { ...appStrings.ru, ...dialogStrings.ru, ...aiStrings.ru },
  ar: { ...appStrings.ar, ...dialogStrings.ar, ...aiStrings.ar },
  pt: { ...appStrings.pt, ...dialogStrings.pt, ...aiStrings.pt },
  it: { ...appStrings.it, ...dialogStrings.it, ...aiStrings.it },
  pl: { ...appStrings.pl, ...dialogStrings.pl, ...aiStrings.pl },
  cs: { ...appStrings.cs, ...dialogStrings.cs, ...aiStrings.cs },
  nl: { ...appStrings.nl, ...dialogStrings.nl, ...aiStrings.nl },
  ms: { ...appStrings.ms, ...dialogStrings.ms, ...aiStrings.ms },
  he: { ...appStrings.he, ...dialogStrings.he, ...aiStrings.he },
  hi: { ...appStrings.hi, ...dialogStrings.hi, ...aiStrings.hi },
  'zh-TW': { ...appStrings['zh-TW'], ...dialogStrings['zh-TW'], ...aiStrings['zh-TW'] },
}
