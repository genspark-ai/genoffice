import { aiStrings } from './strings-ai'
import { appStrings } from './strings-app'
import { paneStrings } from './strings-panes'
import { ribbonStrings } from './strings-ribbon'

export const strings = {
  zh: { ...appStrings.zh, ...ribbonStrings.zh, ...paneStrings.zh, ...aiStrings.zh },
  en: { ...appStrings.en, ...ribbonStrings.en, ...paneStrings.en, ...aiStrings.en },
  ja: { ...appStrings.ja, ...ribbonStrings.ja, ...paneStrings.ja, ...aiStrings.ja },
  ko: { ...appStrings.ko, ...ribbonStrings.ko, ...paneStrings.ko, ...aiStrings.ko },
  fr: { ...appStrings.fr, ...ribbonStrings.fr, ...paneStrings.fr, ...aiStrings.fr },
  de: { ...appStrings.de, ...ribbonStrings.de, ...paneStrings.de, ...aiStrings.de },
  es: { ...appStrings.es, ...ribbonStrings.es, ...paneStrings.es, ...aiStrings.es },
  th: { ...appStrings.th, ...ribbonStrings.th, ...paneStrings.th, ...aiStrings.th },
  id: { ...appStrings.id, ...ribbonStrings.id, ...paneStrings.id, ...aiStrings.id },
  ru: { ...appStrings.ru, ...ribbonStrings.ru, ...paneStrings.ru, ...aiStrings.ru },
  ar: { ...appStrings.ar, ...ribbonStrings.ar, ...paneStrings.ar, ...aiStrings.ar },
  pt: { ...appStrings.pt, ...ribbonStrings.pt, ...paneStrings.pt, ...aiStrings.pt },
  it: { ...appStrings.it, ...ribbonStrings.it, ...paneStrings.it, ...aiStrings.it },
  pl: { ...appStrings.pl, ...ribbonStrings.pl, ...paneStrings.pl, ...aiStrings.pl },
  cs: { ...appStrings.cs, ...ribbonStrings.cs, ...paneStrings.cs, ...aiStrings.cs },
  nl: { ...appStrings.nl, ...ribbonStrings.nl, ...paneStrings.nl, ...aiStrings.nl },
  ms: { ...appStrings.ms, ...ribbonStrings.ms, ...paneStrings.ms, ...aiStrings.ms },
  he: { ...appStrings.he, ...ribbonStrings.he, ...paneStrings.he, ...aiStrings.he },
  hi: { ...appStrings.hi, ...ribbonStrings.hi, ...paneStrings.hi, ...aiStrings.hi },
  'zh-TW': {
    ...appStrings['zh-TW'],
    ...ribbonStrings['zh-TW'],
    ...paneStrings['zh-TW'],
    ...aiStrings['zh-TW'],
  },
}
