import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { parseLangAttr, type LangSlots } from './proofing-language'

const PREF_KEY = 'genoffice.docs.detectLanguage'

/** off unless turned on in Tools > Language: marking text changes what a save writes */
export function detectLanguageEnabled(): boolean {
  return typeof localStorage !== 'undefined' && localStorage.getItem(PREF_KEY) === '1'
}

export function setDetectLanguageEnabled(on: boolean): void {
  if (typeof localStorage === 'undefined') return
  if (on) localStorage.setItem(PREF_KEY, '1')
  else localStorage.removeItem(PREF_KEY)
}

/** the words that give a Latin-script language away; each list is that language's commonest */
const STOPWORDS: Readonly<Record<string, string>> = {
  'en-US':
    'the and of to in is that it was for with as his on be at by this had not are but from or have an they which you were her all she there would their we him been has when who will more no if out so said what',
  'fr-FR':
    'le la les de des et est un une du en que qui dans pour pas au sur ne se ce il elle sont avec plus par mais nous vous ont cette comme été leur tout aux être fait',
  'de-DE':
    'der die das und ist nicht ein eine zu den mit von auf für sich dem des im auch als es an werden aus er hat dass sie nach wird bei einer um am sind noch wie einem über einen so zum war',
  'es-ES':
    'el la los las de que y en un una es por con no para su al lo como más pero sus le ya fue este ha sí porque esta son entre cuando muy sin sobre también hasta hay donde quien desde todo nos',
  'pt-BR':
    'o a os as de que e em um uma é não para com por se na no da do mais como mas foi ao ele das tem à seu sua ou ser quando muito há nos já está eu também só pelo pela até isso ela entre',
  'it-IT':
    'il lo la i gli le di che e è un una in per non con si da sono come ma anche più del della dei nel alla ha questo se ci era al delle loro tutto ho cui essere',
  'nl-NL':
    'de het een en van ik te dat die in is niet zijn op aan met als voor had er maar om hem dan zou of wat mijn men dit zo door over ze zich bij ook tot je naar heb heeft hebben deze',
  'pl-PL':
    'i w nie na z że się do to jest o jak a po co tak ale za od przez go czy tylko jego ich był ma już być są może dla tym które który przy',
  'cs-CZ':
    'a se na je v že to s z do o i ale jako by si po tak jsem jsou pro když už byl jen která který být ve ze nebo jeho jak co má',
  'hu-HU':
    'a az és hogy nem is egy meg de van volt már csak el még mint ki ez azt aki ha vagy amely fel pedig úgy volna lesz minden után ezt',
  'tr-TR':
    've bir bu da de için ile ne çok daha gibi en var ama o ben sen biz olarak kadar sonra ya değil mi her şey olan ise veya',
  'sv-SE':
    'och i att det som en på är av för med till den har de inte om ett han men var jag sig från vi så kan man när år hon under också efter eller nu sin där vid mot',
  'da-DK':
    'og i at det en den til er som på de med han af for ikke der var mig sig men et har om vi min havde ham hun nu over da fra du ud sin dem os op man hans hvor eller hvad skal',
  'nb-NO':
    'og i det er på en at til som av for med har ikke den de et om var seg han fra men jeg vi så kan hun nå etter eller også ved skal bli ble være ut opp',
  'fi-FI':
    'ja on ei että oli se hän mutta niin kun olen sen myös ovat tai kuin joka mitä vain jos ole nyt sitä tämä siitä hänen kanssa olla minä sinä',
  'vi-VN':
    'và của là có không được trong cho một người này đã với các những để khi đến như cũng về từ ra thì nhiều sẽ nó tôi',
  'id-ID':
    'yang dan di dengan untuk ini itu dari tidak dalam akan pada juga saya ke karena ada mereka bisa oleh adalah kami kita sudah atau telah lebih',
}

const WORDS = new Map<string, Set<string>>(
  Object.entries(STOPWORDS).map(([tag, list]) => [tag, new Set(list.split(' '))]),
)

/** written the same way: the tag a text already carries wins over the list's own */
const SAME_BASE = (a: string | undefined, b: string): boolean =>
  !!a && a.split('-')[0].toLowerCase() === b.split('-')[0].toLowerCase()

function latinLanguage(text: string): string | null {
  const words = text.toLowerCase().match(/\p{Script=Latin}+/gu) ?? []
  if (words.length < 4) return null
  let best: string | null = null
  let bestHits = 0
  let secondHits = 0
  for (const [tag, list] of WORDS) {
    let hits = 0
    for (const w of words) if (list.has(w)) hits++
    if (hits > bestHits) {
      secondHits = bestHits
      bestHits = hits
      best = tag
    } else if (hits > secondHits) secondHits = hits
  }
  // a clear lead: a few shared words (de, en, in) must not decide it
  return bestHits >= 3 && bestHits >= secondHits + 2 ? best : null
}

const count = (text: string, re: RegExp): number => text.match(re)?.length ?? 0

/** the languages a text is written in, one per w:lang slot it has enough of to tell */
export function detectLanguages(text: string): LangSlots {
  const out: LangSlots = {}
  if (count(text, /[\p{Script=Hiragana}\p{Script=Katakana}]/gu) > 0) out.eastAsia = 'ja-JP'
  else if (count(text, /\p{Script=Hangul}/gu) > 0) out.eastAsia = 'ko-KR'
  else if (count(text, /\p{Script=Han}/gu) >= 2) out.eastAsia = 'zh-CN'
  if (count(text, /\p{Script=Arabic}/gu) >= 3) out.bidi = 'ar-SA'
  else if (count(text, /\p{Script=Hebrew}/gu) >= 3) out.bidi = 'he-IL'
  if (count(text, /\p{Script=Cyrillic}/gu) >= 3)
    out.val = /[\u0456\u0457\u0454\u0491]/i.test(text) ? 'uk-UA' : 'ru-RU'
  else if (count(text, /\p{Script=Greek}/gu) >= 3) out.val = 'el-GR'
  else if (count(text, /\p{Script=Devanagari}/gu) >= 3) out.val = 'hi-IN'
  else if (count(text, /\p{Script=Thai}/gu) >= 3) out.val = 'th-TH'
  else {
    const latin = latinLanguage(text)
    if (latin) out.val = latin
  }
  return out
}

/** `own` with the detected slots laid over it; null when that changes no language */
export function mergedLangSlots(own: LangSlots | null, found: LangSlots): LangSlots | null {
  const next: LangSlots = { ...own }
  let changed = false
  for (const slot of ['val', 'eastAsia', 'bidi'] as const) {
    const tag = found[slot]
    if (!tag || SAME_BASE(own?.[slot], tag)) continue
    next[slot] = tag
    changed = true
  }
  return changed ? next : null
}

const detectKey = new PluginKey('languageDetect')

/**
 * Word's "Detect language automatically": as text is typed, the paragraph
 * at the caret is marked with the language it reads as.
 */
export const LanguageDetectExtension = Extension.create({
  name: 'languageDetect',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: detectKey,
        appendTransaction(transactions, _old, state) {
          if (!detectLanguageEnabled()) return null
          const typed = transactions.some(
            (tr) =>
              tr.docChanged &&
              !tr.getMeta(detectKey) &&
              !tr.getMeta('history$') &&
              !tr.getMeta('paste'),
          )
          if (!typed || !state.selection.empty) return null
          const { $from } = state.selection
          const block = $from.parent
          const mark = state.schema.marks.docTextStyle
          if (!block.isTextblock || !mark || block.textContent.length > 4000) return null
          const found = detectLanguages(block.textContent)
          if (!found.val && !found.eastAsia && !found.bidi) return null
          const start = $from.start()
          const tr = state.tr
          block.forEach((child, offset) => {
            if (!child.isText) return
            const own = mark.isInSet(child.marks)
            const next = mergedLangSlots(parseLangAttr(own?.attrs.lang), found)
            if (!next) return
            tr.addMark(
              start + offset,
              start + offset + child.nodeSize,
              mark.create({ ...own?.attrs, lang: JSON.stringify(next) }),
            )
          })
          if (!tr.docChanged) return null
          // typing goes on in the language just found
          tr.setStoredMarks(state.storedMarks)
          return tr.setMeta(detectKey, true).setMeta('addToHistory', false)
        },
      }),
    ]
  },
})
