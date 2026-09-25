#!/usr/bin/env node
/**
 * CI path filter: map changed files onto the vitest workspaces whose tests can be
 * affected. Emits GitHub Actions outputs:
 *
 *   mode=all      full `npm test` (base unknown, shared infra touched, or push to main)
 *   mode=some     `workspaces` lists the `-w` targets, in root `npm test` order
 *   mode=none     nothing to run (no unit-testable file changed, e.g. docs/e2e only)
 *
 * The mapping over-approximates on purpose: when a path's blast radius is unclear
 * it falls back to ALL, so the filter can only over-run, never under-run. Pull
 * requests get the filtered set; main and release branches always run the full
 * suite, which is the backstop for anything the mapping misses.
 *
 * Usage: node tools/ci-affected-tests.mjs [--base <ref>] [--files <path>...]
 * `--files` bypasses git (used by the self-test); with no base and no files the
 * result is ALL.
 */
import { execFileSync } from 'node:child_process'

// Workspaces in root `npm test` order — filtered runs preserve this order so
// fixture generation and sidecar builds still happen before their dependents.
const ORDER = [
  '@genoffice/i18n',
  '@genoffice/electron-utils',
  '@genoffice/font-metrics',
  '@genoffice/docx-engine',
  '@genoffice/pdf2docx',
  '@genoffice/html2docx',
  '@genoffice/file-parse',
  '@genoffice/pptx-engine',
  '@genoffice/xlsx-gateway',
  '@genoffice/pptx-ops',
  '@genoffice/pptx-render',
  '@genoffice/pipelines',
  '@genoffice/ai-search',
  '@genoffice/agent-core',
  '@genoffice/ai-provider',
  '@genoffice/project-store',
  '@genoffice/cli',
  '@genoffice/ui',
  '@genoffice/docs',
  '@genoffice/sheets',
  '@genoffice/shell',
  '@genoffice/slides',
  '@genoffice/pdf',
  '@genoffice/markdown',
  '@genoffice/html',
]

const ALL = Symbol('all')

/**
 * Path prefix → workspaces to run. Longer/more specific prefixes first is not
 * required (all matching rules union), but app rules come before package rules
 * that could also prefix-match. `ALL` marks shared infrastructure: anything the
 * whole suite touches.
 */
const RULES = [
  // Apps: app code is a leaf, only its own tests can be affected.
  ['apps/docs/', ['@genoffice/docs']],
  ['apps/sheets/', ['@genoffice/sheets']],
  ['apps/slides/', ['@genoffice/slides']],
  ['apps/pdf/', ['@genoffice/pdf']],
  ['apps/markdown/', ['@genoffice/markdown']],
  ['apps/html/', ['@genoffice/html']],
  ['apps/shell/', ['@genoffice/shell']],

  // Engine packages: apps and sibling packages consume them, so their blast
  // radius is the engine itself plus every consumer with tests over it.
  [
    'packages/docx-engine/',
    [
      '@genoffice/docx-engine',
      '@genoffice/pdf2docx',
      '@genoffice/file-parse',
      '@genoffice/cli',
      '@genoffice/docs',
      '@genoffice/pdf',
      '@genoffice/markdown',
      '@genoffice/html',
    ],
  ],
  [
    'packages/pptx-engine/',
    [
      '@genoffice/pptx-engine',
      '@genoffice/pptx-render',
      '@genoffice/pptx-ops',
      '@genoffice/cli',
      '@genoffice/slides',
    ],
  ],
  ['packages/pptx-render/', ['@genoffice/pptx-render', '@genoffice/pptx-ops', '@genoffice/slides']],
  ['packages/pptx-ops/', ['@genoffice/pptx-ops', '@genoffice/slides']],
  ['packages/pdf2docx/', ['@genoffice/pdf2docx', '@genoffice/pdf']],
  ['packages/html2docx/', ['@genoffice/html2docx', '@genoffice/html']],
  ['packages/xlsx-gateway/', ['@genoffice/xlsx-gateway', '@genoffice/sheets']],
  [
    'packages/font-metrics/',
    ['@genoffice/font-metrics', '@genoffice/docs', '@genoffice/sheets', '@genoffice/slides'],
  ],
  ['packages/cli/', ['@genoffice/cli']],

  // Shared infrastructure: every app consumes these, so run everything.
  ['packages/i18n/', ALL],
  ['packages/ui/', ALL],
  ['packages/electron-utils/', ALL],
  ['packages/agent-core/', ALL],
  ['packages/ai-provider/', ALL],
  ['packages/ai-search/', ALL],
  ['packages/project-store/', ALL],
  ['packages/pipelines/', ALL],
  ['packages/file-parse/', ALL],

  // Repo-level infrastructure: could affect anything.
  ['.github/', ALL],
  ['tools/', ALL],
  ['scripts/', ALL],
  ['fixtures/', ALL],
  ['package.json', ALL],
  ['package-lock.json', ALL],
  ['tsconfig', ALL],
  ['vitest.config.ts', ALL],
  ['eslint.config.mjs', ALL],

  // e2e-only changes: the unit suite has nothing to say about them.
  ['e2e/', []],
]

function parseArgs(argv) {
  const opts = { base: process.env.FORMAT_BASE_REF || '', files: [] }
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--base') opts.base = argv[++i] ?? ''
    else if (argv[i] === '--files') opts.files.push(...argv.slice(i + 1))
    else if (argv[i] === '--self-test') opts.selfTest = true
  }
  if (/^0+$/.test(opts.base)) opts.base = ''
  return opts
}

function changedFiles(base) {
  if (!base) return []
  const git = (args) => execFileSync('git', args, { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 })
  git(['rev-parse', '--verify', `${base}^{commit}`])
  return git(['diff', '--name-only', '--diff-filter=ACMRT', `${base}...HEAD`])
    .split('\n')
    .filter(Boolean)
}

function workspacesFor(changed) {
  const selected = new Set()
  let forceAll = false
  for (const path of changed) {
    const rule = RULES.find(([prefix]) => path === prefix || path.startsWith(prefix))
    if (!rule) {
      // Unknown path: over-approximate to the full suite.
      forceAll = true
      continue
    }
    if (rule[1] === ALL) forceAll = true
    else for (const ws of rule[1]) selected.add(ws)
  }
  if (forceAll) return ALL
  return ORDER.filter((ws) => selected.has(ws))
}

function selfTest() {
  const run = (files) => {
    const r = workspacesFor(files)
    return r === ALL ? 'ALL' : r.join(' ')
  }
  const cases = [
    ['apps/docs/src/renderer/main.tsx', '@genoffice/docs'],
    ['apps/sheets/tests/xlsx-structure.test.ts', '@genoffice/sheets'],
    [
      'packages/docx-engine/src/index.ts',
      '@genoffice/docx-engine @genoffice/pdf2docx @genoffice/file-parse @genoffice/cli @genoffice/docs @genoffice/pdf @genoffice/markdown @genoffice/html',
    ],
    ['packages/pptx-ops/src/font-size.ts', '@genoffice/pptx-ops @genoffice/slides'],
    ['packages/i18n/src/index.ts', 'ALL'],
    ['package-lock.json', 'ALL'],
    ['e2e/open-focus-typing.spec.ts', ''],
    ['README.md', 'ALL'], // unknown path: over-approximate
    [['apps/docs/src/a.ts', 'apps/x/src/b.ts'], 'ALL'],
  ]
  for (const [files, expected] of cases) {
    const got = run(Array.isArray(files) ? files : [files])
    if (got !== expected) {
      console.error(
        `self-test FAILED\n  files: ${JSON.stringify(files)}\n  expected: ${expected}\n  got:      ${got}`,
      )
      process.exit(1)
    }
  }
  console.log('self-test OK')
}

const opts = parseArgs(process.argv.slice(2))
if (opts.selfTest) {
  selfTest()
  process.exit(0)
}

const changed = opts.files.length ? opts.files : changedFiles(opts.base)
const result = workspacesFor(changed)
const mode = result === ALL ? 'all' : result.length === 0 ? 'none' : 'some'
console.log(`mode=${mode}`)
if (mode === 'some') console.log(`workspaces=${result.join(' ')}`)
