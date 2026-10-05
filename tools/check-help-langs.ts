#!/usr/bin/env node
/**
 * Check the manual's language coverage: `npx tsx tools/check-help-langs.ts`
 *
 * Four workers write these files in parallel, one per language, and the
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

const TOPICS = resolve(__dirname, '../apps/shell/src/renderer/src/i18n/help/topics')
const SOURCE = 'en'

/** the 21 UI languages the manual has to ship */
const LANGS = [
  'zh',
  'zh-TW',
  'en',
  'ja',
  'ko',
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
  'he',
  'hi',
  'vi',
]

/** Latin- or Cyrillic-script languages: any Han character in one of these is a
 *  worker that pasted the Chinese source instead of translating it. */
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

const IDS = [
  'getting-started',
  'home-screen',
  'tabs-and-windows',
  'docs',
  'sheets',
  'slides',
  'pdf',
  'markdown',
  'html',
  'ai-panel',
  'ai-models',
  'fonts',
  'file-ops',
  'settings-integrations',
]

const read = (id: string, lang: string): string | null => {
  try {
    return readFileSync(join(TOPICS, `${id}.${lang}.md`), 'utf8')
  } catch {
    return null
  }
}

const headings = (t: string): number => (t.match(/^#{1,6} /gm) ?? []).length
const tableRows = (t: string): number => (t.match(/^\|/gm) ?? []).length
const images = (t: string): string[] =>
  [...t.matchAll(/!\[[^\]]*\]\(([^)]+)\)/g)].map((m) => m[1]!).sort()
const links = (t: string): string[] =>
  [...new Set([...t.matchAll(/\]\((help:\/\/[^)]+)\)/g)].map((m) => m[1]!))].sort()
const codeFences = (t: string): number => (t.match(/^```/gm) ?? []).length
const bullets = (t: string): number => (t.match(/^\s*[-*] /gm) ?? []).length
const numbered = (t: string): number => (t.match(/^\s*\d+\. /gm) ?? []).length
const HAN = /[一-鿿]/

let problems = 0
const fail = (msg: string): void => {
  problems++
  console.log(`  FAIL  ${msg}`)
}

console.log(`manual: ${IDS.length} topics x ${LANGS.length} languages\n`)

for (const lang of LANGS) {
  const missing: string[] = []
  const structural: string[] = []
  const content: string[] = []
  for (const id of IDS) {
    const text = read(id, lang)
    if (text === null) {
      missing.push(id)
      continue
    }
    const src = read(id, SOURCE)!
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
    ? '\nall 21 languages complete and structurally identical to English'
    : `\n${problems} problems`,
)
process.exit(problems === 0 ? 0 : 1)
