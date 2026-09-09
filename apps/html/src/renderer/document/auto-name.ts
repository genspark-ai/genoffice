const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
}

function decode(s: string): string {
  return s.replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === '#') {
      const code =
        e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)
      return Number.isFinite(code) && code <= 0x10ffff ? String.fromCodePoint(code) : m
    }
    return ENTITIES[e.toLowerCase()] ?? m
  })
}

function textOf(html: string, tag: string): string {
  const m = new RegExp(`<${tag}(?:\\s[^>]*)?>([\\s\\S]*?)</${tag}\\s*>`, 'i').exec(html)
  if (!m) return ''
  return decode(m[1]!.replace(/<[^>]*>/g, ''))
    .replace(/\s+/g, ' ')
    .trim()
}

const LEADING_FILLER =
  /^(?:\u8bf7|\u9ebb\u70e6|\u5e2e\u6211|\u7ed9\u6211|please\s+|help me\s+|can you\s+|could you\s+)+/i
const SENTENCE_END = /[\u3002\uff01\uff1f!?]|\.(?=\s|$)/
const CLAUSE_BREAK = /[\uff0c,\uff1b;\uff1a:\u3001]/
const MAX_PROMPT_NAME = 40

/** Provisional name for an untitled document, taken from the user's first request: its first clause, trimmed of politeness */
export function deriveNameFromPrompt(prompt: string): string {
  let s = prompt.split(/\r?\n/).find((line) => line.trim()) ?? ''
  s = s.replace(/\s+/g, ' ').trim()
  s = s.split(SENTENCE_END)[0]!
  s = s.replace(LEADING_FILLER, '')
  const clause = CLAUSE_BREAK.exec(s.slice(6))
  if (clause) s = s.slice(0, 6 + clause.index)
  if (s.length > MAX_PROMPT_NAME) {
    const cut = s.lastIndexOf(' ', MAX_PROMPT_NAME)
    s = s.slice(0, cut > 20 ? cut : MAX_PROMPT_NAME)
  }
  s = s.replace(/[\s\p{P}]+$/u, '').trim()
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : ''
}

/** File name for an AI-generated untitled document: <title>, else the first <h1>, else the first words of the body */
export function deriveAutoFileName(html: string): string {
  const title = textOf(html, 'title') || textOf(html, 'h1')
  if (title) return title.slice(0, 60)
  const body = textOf(html, 'body') || html.replace(/<[^>]*>/g, ' ')
  return decode(body).replace(/\s+/g, ' ').trim().split(' ').slice(0, 8).join(' ').slice(0, 60)
}
