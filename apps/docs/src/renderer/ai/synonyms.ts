const MAX_SYNONYMS = 8
const MAX_LENGTH = 40

const SYSTEM =
  'You are a thesaurus. Reply with up to 8 synonyms or near-synonyms of the word or short phrase the user sends, ' +
  'in the same language as that word, most common first, one per line. ' +
  'No numbering, no notes, no quotation marks. Reply with nothing when there are none.'

/** One candidate per line or comma; list markers, quotes, repeats and the word itself are dropped. */
export function parseSynonyms(content: string, word: string): string[] {
  const seen = new Set([word.trim().toLowerCase()])
  const out: string[] = []
  for (const raw of content.split(/[\n,;\u3001\uff0c\uff1b]/)) {
    const entry = raw
      .replace(/^\s*(?:[-*\u2022\u00b7]|\d+[.)\u3001])\s*/, '')
      .replace(/^["'\u201c\u2018\u300c]+|["'\u201d\u2019\u300d.\u3002]+$/g, '')
      .trim()
    const key = entry.toLowerCase()
    if (!entry || entry.length > MAX_LENGTH || seen.has(key)) continue
    seen.add(key)
    out.push(entry)
    if (out.length === MAX_SYNONYMS) break
  }
  return out
}

/** Empty when no provider is set up or the request fails: the menu then reads "No Suggestions". */
export async function fetchSynonyms(word: string): Promise<string[]> {
  try {
    const settings = await window.desktop.getAiSettings()
    const reply = await window.desktop.aiChat({ settings, system: SYSTEM, user: word })
    return reply.ok && reply.content ? parseSynonyms(reply.content, word) : []
  } catch {
    return []
  }
}
