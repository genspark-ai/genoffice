#!/usr/bin/env node

import { execFileSync } from 'node:child_process'
import {
  chmodSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path'
import { tmpdir } from 'node:os'

const DEFAULT_REMOTE = 'upstream'
const DEFAULT_BRANCH = 'main'

export function parseArgs(argv) {
  const options = {
    apply: false,
    remote: DEFAULT_REMOTE,
    branch: DEFAULT_BRANCH,
    settingsPath: process.env.GENOFFICE_SETTINGS_PATH ?? null,
    credentialsPath: process.env.GENOFFICE_CREDENTIALS_PATH ?? null,
    help: false,
  }
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]
    if (arg === '--apply') options.apply = true
    else if (arg === '--help' || arg === '-h') options.help = true
    else if (arg === '--remote') options.remote = requiredValue(argv, ++i, arg)
    else if (arg === '--branch') options.branch = requiredValue(argv, ++i, arg)
    else if (arg === '--settings-path') options.settingsPath = requiredValue(argv, ++i, arg)
    else if (arg === '--credentials-path') options.credentialsPath = requiredValue(argv, ++i, arg)
    else throw new Error(`Unknown option: ${arg}`)
  }
  if (
    options.settingsPath &&
    options.credentialsPath &&
    resolve(options.settingsPath) === resolve(options.credentialsPath)
  ) {
    throw new Error('Settings and credentials paths must be different files.')
  }
  return options
}

function requiredValue(argv, index, option) {
  const value = argv[index]
  if (!value || value.startsWith('-')) throw new Error(`${option} requires a value.`)
  return value
}

export function protectedPaths(options) {
  return [options.settingsPath, options.credentialsPath]
    .filter((value) => typeof value === 'string' && value.length > 0)
    .map((value) => resolve(value))
}

export function captureProtectedFiles(paths) {
  return paths.map((path) => {
    if (!existsSync(path)) return { path, exists: false, bytes: null, mode: null }
    const stat = lstatSync(path)
    if (!stat.isFile()) throw new Error(`Protected path is not a regular file: ${path}`)
    return {
      path,
      exists: true,
      bytes: readFileSync(path),
      mode: stat.mode & 0o777,
    }
  })
}

export function restoreProtectedFiles(snapshots) {
  for (const snapshot of snapshots) {
    if (!snapshot.exists || !snapshot.bytes) continue
    mkdirSync(dirname(snapshot.path), { recursive: true, mode: 0o700 })
    writeFileSync(snapshot.path, snapshot.bytes, { mode: snapshot.mode ?? 0o600 })
    chmodSync(snapshot.path, snapshot.mode ?? 0o600)
  }
}

function git(args, options = {}) {
  const output = execFileSync('git', args, {
    cwd: options.cwd ?? process.cwd(),
    encoding: 'utf8',
    stdio: options.inherit ? 'inherit' : ['ignore', 'pipe', 'pipe'],
  })
  return typeof output === 'string' ? output.trim() : ''
}

function gitTracked(path) {
  const repoRoot = resolve(git(['rev-parse', '--show-toplevel']))
  const rel = relative(repoRoot, resolve(path))
  if (!rel || rel.startsWith('..') || isAbsolute(rel)) return false
  try {
    git(['ls-files', '--error-unmatch', '--', rel])
    return true
  } catch {
    return false
  }
}

function assertSafeProtectedPaths(paths) {
  for (const path of paths) {
    if (existsSync(path) && lstatSync(path).isSymbolicLink()) {
      throw new Error(`Refusing a symbolic-link protected path: ${path}`)
    }
    if (gitTracked(path)) {
      throw new Error(
        `Refusing a tracked protected path; keep user settings outside the repository: ${path}`,
      )
    }
  }
}

function assertCleanWorktree() {
  const status = git(['status', '--porcelain'])
  if (status) {
    throw new Error('Working tree is not clean. Commit or stash changes before upstream sync.')
  }
}

function timestamp() {
  return new Date()
    .toISOString()
    .replace(/[-:]/g, '')
    .replace(/\.\d{3}Z$/, 'Z')
}

export function usage() {
  return `Usage: node tools/sync-upstream.mjs [options]

Fetch upstream and show the planned sync (default):
  node tools/sync-upstream.mjs --remote upstream --branch main

Apply a rebase after making a local backup ref:
  node tools/sync-upstream.mjs --apply --settings-path /path/to/ai-settings.json \\
    --credentials-path /path/to/ai-credentials.json

Options:
  --apply                    Rebase the current branch onto remote/branch.
  --remote <name>            Git remote (default: upstream).
  --branch <name>            Remote branch (default: main).
  --settings-path <path>     Preserve this user settings file byte-for-byte.
  --credentials-path <path>  Preserve this encrypted credential file byte-for-byte.
  --help                     Show this message.
`
}

export function syncUpstream(options) {
  const paths = protectedPaths(options)
  assertSafeProtectedPaths(paths)
  assertCleanWorktree()

  const currentBranch = git(['branch', '--show-current'])
  if (!currentBranch) throw new Error('Cannot sync from a detached HEAD.')
  const target = `${options.remote}/${options.branch}`

  console.log(`Fetching ${target}…`)
  git(['fetch', options.remote, options.branch, '--prune'], { inherit: true })

  if (!options.apply) {
    console.log(`Dry run complete. Re-run with --apply to rebase ${currentBranch} onto ${target}.`)
    return { applied: false, currentBranch, backupBranch: null, protectedCount: paths.length }
  }

  const backupBranch = `backup/upstream-sync-${timestamp()}`
  git(['branch', backupBranch, currentBranch])
  console.log(`Created local backup ref ${backupBranch}.`)

  const snapshots = captureProtectedFiles(paths)
  const tempDir = mkdtempSync(join(tmpdir(), 'genoffice-upstream-sync-'))
  chmodSync(tempDir, 0o700)
  for (const snapshot of snapshots) {
    if (snapshot.exists && snapshot.bytes) {
      const backupPath = join(tempDir, basename(snapshot.path))
      writeFileSync(backupPath, snapshot.bytes, { mode: 0o600 })
      chmodSync(backupPath, 0o600)
    }
  }

  let rebaseError = null
  try {
    console.log(`Rebasing ${currentBranch} onto ${target}…`)
    try {
      git(['rebase', target], { inherit: true })
    } catch (error) {
      rebaseError = error
    }
  } finally {
    // Restore after both success and conflict. A conflict is intentionally left
    // active for the developer to resolve; user data never participates in it.
    restoreProtectedFiles(snapshots)
    rmSync(tempDir, { recursive: true, force: true })
  }
  if (rebaseError) {
    throw new Error(
      `Rebase stopped with conflicts. Resolve them, then run git rebase --continue. Backup ref: ${backupBranch}`,
    )
  }
  console.log(`Upstream sync complete. Protected ${paths.length} user file(s).`)
  return { applied: true, currentBranch, backupBranch, protectedCount: paths.length }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  try {
    const options = parseArgs(process.argv.slice(2))
    if (options.help) console.log(usage())
    else syncUpstream(options)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exitCode = 1
  }
}
