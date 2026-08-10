/**
 * Aggregates licenses of third-party components shipped with the installer
 * into apps/shell/build/THIRD-PARTY-NOTICES.txt.
 *
 * The shipped set is derived from what the source actually imports, then closed
 * over the dependency graph — not from `dependencies` vs `devDependencies`,
 * which in this repo says nothing about what ships: electron-vite bundles the
 * renderer/main graph into out/ and no node_modules directory is packaged, so
 * several genuinely-bundled libraries are declared as devDependencies.
 *
 * The Chromium part is too large (18MB); electron-builder extraResources ships
 * LICENSES.chromium.html separately, referenced from the body text.
 */
import { readFileSync, writeFileSync, mkdirSync, existsSync, readdirSync } from 'node:fs'
import { join, dirname, relative } from 'node:path'
import { fileURLToPath } from 'node:url'
import { execFileSync } from 'node:child_process'
import { builtinModules, createRequire } from 'node:module'

const require = createRequire(import.meta.url)

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const BUILTIN = new Set(builtinModules)

/**
 * Source trees whose imports end up in a shipped bundle. Deliberately only
 * `src/` — vite configs and scripts pull in the whole build toolchain, which
 * is not distributed.
 */
const SRC_GLOBS = [
  'apps/docs/src',
  'apps/markdown/src',
  'apps/pdf/src',
  'apps/sheets/src',
  'apps/shell/src',
  'apps/slides/src',
  ...readdirSync(join(ROOT, 'packages'), { withFileTypes: true })
    .filter((e) => e.isDirectory())
    .map((e) => `packages/${e.name}/src`),
]
const CODE_EXT = /\.(ts|tsx|js|jsx|mjs|cjs)$/
const SKIP_DIR = /^(node_modules|out|dist|release|tests|__tests__|target)$/

/** Runtime that is always present but never imported by specifier */
const IMPLICIT = ['electron']

/**
 * Packages copied into the installer verbatim by electron-builder rather than
 * bundled — the gsk CLI and its runtime deps are spawned, never imported, so
 * they are invisible to the import scan.
 */
function extraResourceSeeds() {
  // the electron-builder config lives in its own cjs module (not package.json
  // "build") so the publish URL can be injected from the environment
  const build = require(join(ROOT, 'apps/shell/electron-builder.cjs'))
  const entries = [build.extraResources, build.mac?.extraResources, build.win?.extraResources]
  const names = []
  for (const e of entries.flat()) {
    const m = e?.from && /node_modules\/((?:@[^/]+\/)?[^/]+)$/.exec(e.from)
    if (m) names.push(m[1])
  }
  return names
}

/** license file lives somewhere non-obvious */
const LICENSE_PATH = { electron: 'dist/LICENSE' }

/** license text is not published to npm; supply the notice by hand */
const NOTE = {
  '@fluentui/react-icons':
    'Copyright (c) Microsoft Corporation. Licensed under the MIT License.\nhttps://github.com/microsoft/fluentui-system-icons',
  '@genspark/cli':
    'Copyright (c) Genspark. Licensed under the MIT License.\nhttps://www.npmjs.com/package/@genspark/cli',
}

/** SPDX strings that need a word on which side of a dual license we take */
const SPDX_NOTE = {
  jszip: 'MIT (dual MIT/GPL-3.0-or-later, used under MIT)',
  'r-efi': 'MIT (tri-licensed MIT/Apache-2.0/LGPL-2.1-or-later, used under MIT)',
}

const walk = (dir, out = []) => {
  if (!existsSync(dir)) return out
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    if (e.isDirectory()) {
      if (!SKIP_DIR.test(e.name)) walk(join(dir, e.name), out)
    } else if (CODE_EXT.test(e.name)) out.push(join(dir, e.name))
  }
  return out
}

// the lookbehind keeps Array.from('x') / obj.import out of the results
const SPECIFIER = /(?<![.\w])(?:from|import|require)\s*\(?\s*['"]([^'"\n]+)['"]/g

const IN_NODE_MODULES = /node_modules\/((?:@[^/]+\/)?[^/?]+)/

/** 'foo/bar' -> 'foo'; '@a/b/c' -> '@a/b'; relative/builtin/workspace -> null */
function bareName(spec) {
  if (!spec || spec.startsWith('node:')) return null
  // a few deep imports reach into node_modules by relative path (harfbuzzjs wasm)
  const viaPath = IN_NODE_MODULES.exec(spec)
  if (viaPath) return viaPath[1]
  if (spec.startsWith('.') || spec.startsWith('/')) return null
  const parts = spec.split('/')
  const name = spec.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0]
  if (BUILTIN.has(name) || name.startsWith('@genoffice/')) return null
  // the specifier regex also fires on prose inside string concatenations
  return NPM_NAME.test(name) ? name : null
}

const NPM_NAME = /^(@[a-z0-9-~][a-z0-9-._~]*\/)?[a-z0-9-~][a-z0-9-._~]*$/

function importedNames() {
  const names = new Set([...IMPLICIT, ...extraResourceSeeds()])
  for (const g of SRC_GLOBS) {
    for (const file of walk(join(ROOT, g))) {
      const text = readFileSync(file, 'utf8')
      for (const m of text.matchAll(SPECIFIER)) {
        const n = bareName(m[1])
        if (n) names.add(n)
      }
    }
  }
  return names
}

const WORKSPACE_MODULES = SRC_GLOBS.map((g) => join(ROOT, g.replace(/\/src$/, ''), 'node_modules'))

/** installed package names, unscoped first so existing resolutions keep winning */
function topLevelPackages() {
  const dirs = (at) => readdirSync(at, { withFileTypes: true }).filter((e) => e.isDirectory())
  const top = dirs(join(ROOT, 'node_modules')).map((e) => e.name)
  const scoped = top.filter((n) => n.startsWith('@'))
  return [
    ...top.filter((n) => !n.startsWith('@')),
    ...scoped.flatMap((s) => dirs(join(ROOT, 'node_modules', s)).map((e) => `${s}/${e.name}`)),
  ]
}

/** npm hoists, so the root copy is the usual answer; fall back to a nested one */
function pkgDir(name) {
  const root = join(ROOT, 'node_modules', name)
  if (existsSync(join(root, 'package.json'))) return root
  for (const owner of topLevelPackages()) {
    const nested = join(ROOT, 'node_modules', owner, 'node_modules', name)
    if (existsSync(join(nested, 'package.json'))) return nested
  }
  for (const modules of WORKSPACE_MODULES) {
    const own = join(modules, name)
    if (existsSync(join(own, 'package.json'))) return own
  }
  return null
}

const manifest = (dir) => JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'))

/** transitive closure over runtime deps (build-only devDependencies of a dep never ship) */
function closure(seed) {
  const seen = new Map()
  const queue = [...seed].map((name) => ({ name, optional: false }))
  const missing = new Set()
  while (queue.length > 0) {
    const { name, optional } = queue.shift()
    if (seen.has(name)) continue
    const dir = pkgDir(name)
    if (dir === null) {
      // an uninstalled optional dep is a platform variant for another host, not a gap
      if (!optional) missing.add(name)
      seen.set(name, null)
      continue
    }
    const pkg = manifest(dir)
    seen.set(name, { dir, pkg })
    for (const dep of Object.keys(pkg.dependencies ?? {})) {
      if (!seen.has(dep)) queue.push({ name: dep, optional: false })
    }
    for (const dep of Object.keys(pkg.optionalDependencies ?? {})) {
      if (!seen.has(dep)) queue.push({ name: dep, optional: true })
    }
  }
  return { resolved: [...seen].filter(([, v]) => v !== null), missing }
}

const LICENSE_FILE = /^(LICENSE|LICENCE|COPYING)(\.|$)/i

function licenseText(name, dir) {
  const override = LICENSE_PATH[name]
  if (override && existsSync(join(dir, override))) {
    return readFileSync(join(dir, override), 'utf8').trim()
  }
  const files = readdirSync(dir).filter((f) => LICENSE_FILE.test(f))
  files.sort((a, b) => a.length - b.length)
  if (files.length > 0) return readFileSync(join(dir, files[0]), 'utf8').trim()
  return null
}

function noticeText(dir) {
  for (const f of ['NOTICE', 'NOTICE.txt', 'NOTICE.md']) {
    if (existsSync(join(dir, f))) return readFileSync(join(dir, f), 'utf8').trim()
  }
  return null
}

function spdxOf(name, pkg) {
  if (SPDX_NOTE[name]) return SPDX_NOTE[name]
  const raw = typeof pkg.license === 'string' ? pkg.license : pkg.license?.type
  if (raw) return raw
  // some packages ship the file but omit the field (e.g. @univerjs/telemetry)
  return 'see license text below'
}

function repoOf(pkg) {
  const r = pkg.repository
  const url = typeof r === 'string' ? r : r?.url
  return url ? url.replace(/^git\+/, '').replace(/\.git$/, '') : null
}

/** Rust crates statically linked into xlsx-sidecar */
function rustCrates() {
  const target = join(ROOT, 'apps/sheets/native/xlsx-engine/Cargo.toml')
  try {
    const raw = execFileSync(
      'cargo',
      ['metadata', '--format-version', '1', '--manifest-path', target],
      { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] },
    )
    return JSON.parse(raw)
      .packages.filter((p) => p.name !== 'xlsx-sidecar')
      .map((p) => ({
        name: p.name,
        version: p.version,
        spdx: SPDX_NOTE[p.name] ?? p.license ?? 'see repository',
        url: p.repository ?? `https://crates.io/crates/${p.name}`,
        dir: p.manifest_path ? dirname(p.manifest_path) : null,
      }))
      .sort((a, b) => a.name.localeCompare(b.name))
  } catch {
    return null
  }
}

const hr = (title) => `\n${'='.repeat(72)}\n${title}\n${'='.repeat(72)}\n`
const sub = (title) => `\n${'-'.repeat(72)}\n${title}\n${'-'.repeat(72)}\n`

const seed = importedNames()
const { resolved, missing } = closure(seed)
resolved.sort(([a], [b]) => a.localeCompare(b))

let out = `GenOffice — Third-Party Software Notices

This application includes third-party software components under the licenses
reproduced below.

Chromium (bundled via Electron) is licensed under BSD-style and other
licenses — see LICENSES.chromium.html next to this file.
`

out += hr(`1. npm packages (${resolved.length})`)
const noText = []
for (const [name, { dir, pkg }] of resolved) {
  out += sub(`${name} v${pkg.version} — ${spdxOf(name, pkg)}`)
  const text = licenseText(name, dir) ?? NOTE[name]
  if (text) {
    out += text + '\n'
  } else {
    noText.push(name)
    const repo = repoOf(pkg)
    out += `Licensed under ${spdxOf(name, pkg)}.${repo ? `\n${repo}` : ''}\n`
  }
  const notice = noticeText(dir)
  if (notice) out += `\nNOTICE:\n${notice}\n`
}

const crates = rustCrates()
out += hr(`2. Rust crates (xlsx-sidecar native component, statically linked)`)
if (crates === null) {
  out += '\ncargo metadata unavailable at generation time; see'
  out += ' apps/sheets/native/xlsx-engine/Cargo.lock for the full crate list.\n'
} else {
  out += `\n${crates.length} crates, all under permissive terms:\n\n`
  for (const c of crates) out += `  ${c.name} ${c.version}  —  ${c.spdx}\n    ${c.url}\n`
  const texts = new Map()
  for (const c of crates) {
    if (!c.dir) continue
    const text = licenseText(c.name, c.dir)
    if (text && !texts.has(text)) texts.set(text, c.name)
  }
  out += sub('Crate license texts (deduplicated)')
  for (const [text, first] of texts) out += `\n[first seen in ${first}]\n${text}\n`
}

/** Bundled fonts (for docs rendering; all metric-compatible replacements for Microsoft fonts) */
const FONTS = [
  [
    'Liberation Sans / Serif / Mono 2.1.5',
    'SIL OFL 1.1',
    'Digitized data © 2010 Google Corporation with Reserved Font Names Arimo, Tinos and Cousine.\n© 2012 Red Hat, Inc. with Reserved Font Name Liberation.',
  ],
  [
    'Carlito',
    'SIL OFL 1.1',
    '© 2010-2013 tyPoland Lukasz Dziedzic with Reserved Font Name "Carlito".',
  ],
  [
    'Caladea',
    'Apache-2.0',
    '© Huerta Tipográfica. Licensed under the Apache License, Version 2.0.',
  ],
  [
    'Noto Sans CJK SC / Noto Serif CJK SC (subset)',
    'SIL OFL 1.1',
    '© Adobe / Google. This bundle ships a subset of the original fonts (reduced glyph coverage for size);\nno other modifications were made.',
  ],
]

out += hr('3. Bundled fonts')
for (const [name, spdx, copyright] of FONTS) out += sub(`${name} — ${spdx}`) + copyright + '\n'
out += sub('SIL Open Font License 1.1 — full text')
out +=
  readFileSync(join(ROOT, 'apps/docs/src/renderer/fonts/LICENSE-OFL.txt'), 'utf8').trim() + '\n'

out += hr('4. Unicode Character Database data')
out += `
apps/pdf/src/shared/radicals.ts contains a generated mapping derived from
Unicode Character Database 17.0.0, EquivalentUnifiedIdeograph.txt
(2025-08-01):
https://www.unicode.org/Public/17.0.0/ucd/EquivalentUnifiedIdeograph.txt

`
out += readFileSync(join(ROOT, 'LICENSE-UNICODE.txt'), 'utf8').trim() + '\n'

const dest = join(ROOT, 'apps/shell/build/THIRD-PARTY-NOTICES.txt')
mkdirSync(dirname(dest), { recursive: true })
writeFileSync(dest, out)
console.log(
  `written: ${relative(ROOT, dest)} (${(out.length / 1024).toFixed(0)} KB) — ` +
    `${resolved.length} npm packages, ${crates?.length ?? 0} crates`,
)
if (noText.length > 0) console.warn(`no license file published: ${noText.join(', ')}`)
if (missing.size > 0) console.warn(`not installed, skipped: ${[...missing].join(', ')}`)
