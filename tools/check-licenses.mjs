/**
 * Dependency license gate: every production npm dependency must
 * carry a license from the permissive allowlist, so a copyleft dependency
 * cannot slip into a release. Reads license fields from package-lock.json
 * (no install needed); the Rust sidecar equivalent is cargo-deny
 * (apps/sheets/native/xlsx-engine/deny.toml).
 */
import { readFileSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')

const ALLOWED = new Set([
  'MIT',
  'MIT-0',
  'Apache-2.0',
  'ISC',
  'BSD-2-Clause',
  'BSD-3-Clause',
  '0BSD',
  'BlueOak-1.0.0',
  'CC0-1.0',
  'CC-BY-4.0',
  'Zlib',
  'Unlicense',
  'Python-2.0',
  'Unicode-3.0',
  'OFL-1.1',
])

/** Packages whose published package.json lacks a license field; license
 * verified manually against the LICENSE file shipped in the package. */
const EXCEPTIONS = {
  '@univerjs/telemetry': 'Apache-2.0',
}

/** Minimal SPDX expression check: OR passes if any branch is allowed,
 * AND requires every branch, WITH falls back to the base license. */
function isAllowed(expr) {
  let s = expr.trim()
  while (s.startsWith('(') && s.endsWith(')')) {
    const inner = s.slice(1, -1)
    let depth = 0
    let balanced = true
    for (const ch of inner) {
      if (ch === '(') depth++
      else if (ch === ')') depth--
      if (depth < 0) balanced = false
    }
    if (!balanced || depth !== 0) break
    s = inner.trim()
  }
  const splitTop = (sep) => {
    const parts = []
    let depth = 0
    let start = 0
    for (let i = 0; i <= s.length - sep.length; i++) {
      if (s[i] === '(') depth++
      else if (s[i] === ')') depth--
      else if (depth === 0 && s.startsWith(sep, i)) {
        parts.push(s.slice(start, i))
        start = i + sep.length
        i += sep.length - 1
      }
    }
    parts.push(s.slice(start))
    return parts
  }
  const orParts = splitTop(' OR ')
  if (orParts.length > 1) return orParts.some(isAllowed)
  const andParts = splitTop(' AND ')
  if (andParts.length > 1) return andParts.every(isAllowed)
  // legacy "A/B" dual-license shorthand
  const slashParts = s.includes('/') ? s.split('/') : [s]
  if (slashParts.length > 1) return slashParts.some(isAllowed)
  const withParts = s.split(' WITH ')
  return ALLOWED.has(withParts[0].trim())
}

const lock = JSON.parse(readFileSync(join(ROOT, 'package-lock.json'), 'utf8'))
const violations = []

for (const [path, info] of Object.entries(lock.packages)) {
  const idx = path.lastIndexOf('node_modules/')
  if (idx === -1) continue // workspace roots
  if (info.dev || info.link) continue
  const name = path.slice(idx + 'node_modules/'.length)
  const license = info.license || EXCEPTIONS[name]
  if (!license) {
    violations.push(`${name}: no license field in lockfile (add to EXCEPTIONS after verifying)`)
  } else if (!isAllowed(license)) {
    violations.push(`${name}: ${license}`)
  }
}

if (violations.length > 0) {
  console.error('Disallowed or unknown licenses in production dependencies:\n')
  for (const v of violations) console.error(`  ${v}`)
  console.error('\nAllowlist lives in tools/check-licenses.mjs.')
  process.exit(1)
}

console.log('All production npm dependency licenses are within the allowlist.')
