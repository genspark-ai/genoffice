#!/usr/bin/env node
// PR-time mirror of the English-only guard in tools/sync-public-repo.mjs:
// no Chinese may appear in code comments or documentation prose, because the
// public-repo sync aborts on it. Catching it here surfaces the violation on
// the PR that introduces it instead of on the next sync run.
//
// Functional CJK string literals are fine (i18n resources, test fixture
// text, zh-UI matchers), as are the AI prompt guides (runtime resources that
// legitimately show CJK examples).
//
// Keep the file-set and comment-extraction logic identical to the guard in
// tools/sync-public-repo.mjs; paths that never reach the public snapshot
// (its EXCLUDE list) are skipped here too.
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'

// Mirror of EXCLUDE in tools/sync-public-repo.mjs (update both together).
const SYNC_EXCLUDED = [
  '.cursor/',
  '.github/dependabot.yml',
  '.github/workflows/mac-build.yml',
  '.github/workflows/windows-build.yml',
  '.github/workflows/linux-build.yml',
  '.github/workflows/promote-stable.yml',
  'scripts/linux-release-upload.cjs',
  'scripts/mac-release-upload.cjs',
  'scripts/promote-stable.cjs',
  'scripts/win-sign.cjs',
  'scripts/win-update-meta.cjs',
  'scripts/mac-release-tracks.cjs',
  'apps/shell/tests/mac-release-tracks.test.ts',
  'tools/sync-public-repo.mjs',
  'apps/sheets/docs/roadmap-gates.json',
  'apps/sheets/scripts/gate-roadmap.ts',
  'docs/superpowers/',
]

const HAN = /[\u3400-\u9fff]/

const git = spawnSync('git', ['ls-files'], { encoding: 'utf8' })
if (git.status !== 0) {
  console.error(git.stderr)
  process.exit(git.status ?? 1)
}
const root = spawnSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).stdout.trim()

const violations = []
for (const file of git.stdout.trim().split('\n')) {
  if (SYNC_EXCLUDED.some((p) => (p.endsWith('/') ? file.startsWith(p) : file === p))) continue
  const isCode = /\.(ts|tsx|mjs|cjs|js)$/.test(file)
  const isDoc = /\.(md|html?)$/.test(file) && !file.includes('/ai/prompts/')
  if (!isCode && !isDoc) continue
  const lines = readFileSync(join(root, file), 'utf8').split('\n')
  lines.forEach((line, index) => {
    const text = isDoc
      ? line
      : (line.match(/(?:^|[^:'"])\/\/(.*)$/) ??
          line.match(/^\s*\*(.*)$/) ??
          line.match(/\/\*(.*)$/))?.[1]
    if (text !== undefined && HAN.test(text)) {
      violations.push(`  ${file}:${index + 1}: ${line.trim()}`)
    }
  })
}

if (violations.length > 0) {
  console.error(
    `Chinese text found in code comments or docs:\n${violations.join('\n')}\n` +
      'Comments and documentation must be English-only (the public-repo sync ' +
      'aborts on these lines — see tools/sync-public-repo.mjs). Move CJK text ' +
      'into string literals or rewrite the comment in English.',
  )
  process.exit(1)
}
console.log('check-english-comments: OK')
