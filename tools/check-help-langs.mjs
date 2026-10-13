#!/usr/bin/env node
/**
 * Check the manual's language coverage: `npm run check:help-langs`
 *
 * Plain node, not tsx, so CI can run it after `npm ci` without fetching a
 * runner — which is the same reason the other repository checks are `.mjs`.
 *
 * Several workers write these files in parallel, one per language, and the
 * failure this guards is the one nobody notices by eye: a language that is
 * quietly a copy of another, or an article that lost a section, or a `help://`
 * link whose target got translated and now resolves to nothing.
 *
 * The checks are all mechanical on purpose. Judging prose is a reviewer's job;
 * this only asserts that every file exists and that its *structure* is identical
 * to the English source, which is exactly the class of mistake that survives a
 * skim.
 */
import { readFileSync, readdirSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = resolve(fileURLToPath(new URL('.', import.meta.url)), '..')
const TOPICS = join(ROOT, 'apps/shell/src/renderer/src/i18n/help/topics')
const REGISTRY = join(TOPICS, '..', 'help-registry.ts')
const I18N = join(ROOT, 'packages/i18n/src/index.ts')

/**
 * The topic ids, read out of the registry source.
 *
 * The registry cannot be imported: it uses `import.meta.glob`, which only a
 * bundler provides. Reading the ids out of the text is the lesser evil to a
 * second hand-kept list, and a registry that stops looking like this fails
 * loudly rather than quietly checking fewer articles than the app has.
 */
function topicIdsFrom(source) {
  const array = source.slice(source.indexOf('export const HELP_TOPICS'))
  if (!array) throw new Error('HELP_TOPICS not found in the registry')
  const ids = [...array.matchAll(/^\s*id: '([a-z0-9-]+)',$/gm)].map((m) => m[1])
  if (ids.length === 0) throw new Error('no topic ids parsed out of HELP_TOPICS')
  return ids
}

/**
 * The languages the manual has to ship, read from the i18n package rather than
 * listed here: a hand-kept copy of the 21 codes is a 21-code list that goes
 * stale the day a language is added, and the failure is silent.
 *
 * `@genoffice/i18n` exports its TypeScript source, so node cannot import it
 * either — same reason as the ids above. Read as text, and fail loudly rather
 * than quietly checking fewer languages than the app ships.
 */
function langsFrom(source) {
  const decl = source.indexOf('export const LANGS')
  if (decl < 0) throw new Error('LANGS not found in the i18n package')
  // start after the `= [`, not after the declaration: the type annotation is
  // `readonly Lang[]`, whose own `]` would otherwise bound the slice to nothing
  const open = source.indexOf('[', source.indexOf('=', decl))
  const end = source.indexOf(']', open)
  if (open < 0 || end < 0) throw new Error('unterminated LANGS array in the i18n package')
  // everything after it is other tables (HTML_LANGS, MAC_KEY_NAMES, …) full of
  // quoted strings, and an unbounded match would check hundreds of "languages"
  // that do not exist
  const langs = [...source.slice(open, end).matchAll(/'([a-zA-Z-]+)'/g)].map((m) => m[1])
  if (langs.length === 0) throw new Error('no languages parsed out of LANGS')
  return langs
}

const SOURCE = 'en'
const IDS = topicIdsFrom(readFileSync(REGISTRY, 'utf-8'))
const LANGS = langsFrom(readFileSync(I18N, 'utf-8'))

/** Latin- or Cyrillic-script languages: any Han character in one of these is a
 *  worker that pasted the CJK source instead of translating it. */
const NO_HAN = new Set([
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'cs',
  'nl',
  'ms',
  'hi',
  'vi',
])

/**
 * The script a language actually writes in.
 *
 * A file in one of these that contains none of its own script is a copy of
 * something else, whatever its heading and bullet counts say. Structure is
 * cheap to match and proves nothing: an untranslated file matches English
 * perfectly, because it *is* English.
 */
const SCRIPTS = {
  zh: /[\u4E00-\u9FFF]/,
  'zh-TW': /[\u4E00-\u9FFF]/,
  ja: /[\u3040-\u30FF\u4E00-\u9FFF]/,
  ko: /[가-힯]/,
  ru: /[Ѐ-ӿ]/,
  ar: /[؀-ۿ]/,
  he: /[֐-׿]/,
  th: /[฀-๿]/,
  hi: /[ऀ-ॿ]/,
}

/**
 * The longest run of text the two files share, ignoring whitespace and code.
 *
 * Markdown tables are excluded along with fenced code, and not as a
 * convenience: a table's cells are proper nouns and version numbers — macOS,
 * `.dmg`, Windows 10+ — identical in every translation by construction, and its
 * `--- | ---` filler is punctuation that matches itself. Counting those made
 * this report "not translated" for a Polish article whose only shared run was a
 * separator row. What this is meant to catch is prose, so only prose is
 * compared.
 */
function longestSharedRun(a, b) {
  const norm = (t) =>
    t
      .replace(/```[\s\S]*?```/g, ' ')
      .replace(/^\s*\|.*$/gm, ' ')
      .replace(/\s+/g, ' ')
      .trim()
  const sa = norm(a)
  const sb = norm(b)
  const short = sa.length < sb.length ? sa : sb
  const long = sa.length < sb.length ? sb : sa
  let best = 0
  for (let i = 0; i < short.length && best < 120; i++) {
    if (short[i] === ' ') continue
    let j = 0
    while (i + j < short.length && j < 120 && short[i + j] === long[i + j]) j++
    if (j > best) best = j
  }
  return best
}

const read = (id, lang) => {
  try {
    return readFileSync(join(TOPICS, `${id}.${lang}.md`), 'utf8')
  } catch {
    return null
  }
}

const headings = (t) => (t.match(/^#{1,6} /gm) ?? []).length
const tableRows = (t) => (t.match(/^\|/gm) ?? []).length
const images = (t) => [...t.matchAll(/!\[[^\]]*\]\(([^)]+)\)/g)].map((m) => m[1]).sort()
const links = (t) =>
  [...new Set([...t.matchAll(/\]\((help:\/\/[^)]+)\)/g)].map((m) => m[1]))].sort()
const codeFences = (t) => (t.match(/^```/gm) ?? []).length
const bullets = (t) => (t.match(/^\s*[-*] /gm) ?? []).length
const numbered = (t) => (t.match(/^\s*\d+\. /gm) ?? []).length
const HAN = /[\u4E00-\u9FFF]/

let problems = 0
const fail = (msg) => {
  problems++
  console.log(`  FAIL  ${msg}`)
}

console.log(`manual: ${IDS.length} topics x ${LANGS.length} languages\n`)

for (const lang of LANGS) {
  const missing = []
  const structural = []
  const content = []
  for (const id of IDS) {
    const text = read(id, lang)
    if (text === null) {
      missing.push(id)
      continue
    }
    const src = read(id, SOURCE)
    if (headings(text) !== headings(src)) {
      structural.push(`${id}: ${headings(text)} headings vs ${headings(src)}`)
    }
    if (tableRows(text) !== tableRows(src)) {
      structural.push(`${id}: ${tableRows(text)} table rows vs ${tableRows(src)}`)
    }
    if (bullets(text) !== bullets(src)) structural.push(`${id}: bullet count`)
    if (numbered(text) !== numbered(src)) structural.push(`${id}: numbered-step count`)
    if (codeFences(text) !== codeFences(src)) structural.push(`${id}: code fence count`)
    if (images(text).join() !== images(src).join()) {
      structural.push(`${id}: images ${images(text).join()} vs ${images(src).join()}`)
    }
    if (links(text).join() !== links(src).join()) {
      structural.push(`${id}: help:// links ${links(text).join()} vs ${links(src).join()}`)
    }
    if (NO_HAN.has(lang) && HAN.test(text)) {
      content.push(`${id}: contains Han characters`)
    }
    // uppercase-only on purpose: "todo" is an ordinary word in Spanish and
    // Portuguese (ignorar todo), and a case-insensitive match flags half a
    // language edition for writing its own language correctly
    if (/\bTODO\b|\bLorem\b|\bTBD\b|\[translate/.test(text)) {
      content.push(`${id}: placeholder text`)
    }
    const script = SCRIPTS[lang]
    if (script && !script.test(text)) {
      content.push(`${id}: no ${lang} script anywhere — the file is not translated`)
    }
    if (lang !== SOURCE && longestSharedRun(text, src) >= 120) {
      content.push(`${id}: a long run is identical to English — not translated`)
    }
    if (text.trim().length < src.trim().length * 0.35) {
      content.push(`${id}: suspiciously short (${text.trim().length} vs ${src.trim().length})`)
    }
  }
  const bad = missing.length + structural.length + content.length
  if (bad === 0) {
    console.log(`  ok    ${lang.padEnd(6)} ${IDS.length} topics`)
  } else {
    console.log(`  BAD   ${lang.padEnd(6)} ${bad} problems`)
    for (const m of missing) fail(`${lang}: missing ${m}`)
    for (const m of structural) fail(`${lang}: ${m}`)
    for (const m of content) fail(`${lang}: ${m}`)
  }
}

// an orphan is a file no registry entry will ever load
const known = new Set(IDS.flatMap((id) => LANGS.map((l) => `${id}.${l}.md`)))
for (const f of readdirSync(TOPICS)) {
  if (f.endsWith('.md') && !known.has(f)) fail(`orphan: topics/${f} matches no topic+language`)
}

console.log(
  problems === 0
    ? `\nall ${LANGS.length} languages complete and structurally identical to English`
    : `\n${problems} problems`,
)
process.exit(problems === 0 ? 0 : 1)
