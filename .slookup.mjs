#!/usr/bin/env node
// Usage: node .slookup.mjs <stringsFile> <key1> [key2...]
// Picks the top-level `  <lang>: {` block (depth 1) so files with helper
// objects earlier in the file don't confuse the match.
import { readFileSync } from 'node:fs'

const file = process.argv[2]
const keys = process.argv.slice(3)
const src = readFileSync(file, 'utf8')
const LANGS = ['ja', 'ko', 'fr', 'de', 'es']

function lineStart(i) {
  return i === 0 ? 0 : src.lastIndexOf('\n', i - 1) + 1
}
function lineEnd(i) {
  const n = src.indexOf('\n', i)
  return n === -1 ? src.length : n
}

function blockRange(lang) {
  const re = new RegExp(`^  '?${lang}'?: \\{$`, 'gm')
  let m
  let best = null
  while ((m = re.exec(src)) !== null) {
    const start = m.index + m[0].length
    let depth = 1
    let i = start
    while (i < src.length && depth > 0) {
      const c = src[i]
      if (c === '{') depth++
      else if (c === '}') depth--
      i++
    }
    const range = { start, end: i }
    if (!best || range.start > best.start) best = range // last wins
  }
  return best
}

const blocks = {}
for (const l of LANGS) {
  const r = blockRange(l)
  blocks[l] = r ? src.slice(r.start, r.end) : ''
}

function val(block, key) {
  const re = new RegExp(`^\\s*'?${key}'?:\\s*((?:'(?:[^'\\\\]|\\\\.)*'\\s*)+)`, 'm')
  const m = re.exec(block)
  if (!m) return ''
  return m[1]
    .replace(/'\s*'/g, '')
    .replace(/^'|'$/g, '')
    .replace(/\\'/g, "'")
    .replace(/\\n/g, ' ')
}

process.stdout.write('key\t' + LANGS.join('\t') + '\n')
for (const k of keys) {
  const row = LANGS.map((l) => val(blocks[l], k))
  process.stdout.write(k + '\t' + row.join('\t') + '\n')
}
