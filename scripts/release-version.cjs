#!/usr/bin/env node
/**
 * scripts/release-version.cjs — resolve the release version and enforce the
 * forward-only guard before any packaging starts.
 *
 * Version source: an explicit --version argument (workflow_dispatch input),
 * else the pushed tag in GITHUB_REF_NAME with its leading "v" stripped.
 *
 * The guard reuses the release-upload decision from update-feed-utils.cjs:
 * the candidate must be strictly newer than the published stable release
 * (same version is allowed — a re-run against an already-published release
 * republishes its artifacts; an older version is rejected like the update
 * feed would be). The "current" version is read from the repository's latest
 * GitHub Release via `gh api` (GH_TOKEN required), or from --current for
 * offline checks.
 *
 * On success the resolved bare version is printed to stdout (single line) and
 * appended to $GITHUB_OUTPUT as `version=...` when running in Actions.
 */

const { execFileSync } = require('node:child_process')
const { releaseUploadDecision } = require('./update-feed-utils.cjs')

const SEMVER = /^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/

function parseArgs(argv) {
  const out = { force: false, version: null, current: null }
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === '--force') out.force = true
    else if (arg === '--version' || arg === '--current') {
      const value = argv[++i]
      if (!value) {
        console.error(`missing value after ${arg}`)
        process.exit(2)
      }
      out[arg === '--version' ? 'version' : 'current'] = value
    } else {
      console.error(`unknown argument: ${arg}`)
      process.exit(2)
    }
  }
  return out
}

function stripV(value) {
  return value.replace(/^v/, '')
}

function resolveCandidate(explicit) {
  if (explicit) return { value: stripV(explicit), source: '--version' }
  const ref = process.env.GITHUB_REF_NAME
  if (ref && /^v\d/.test(ref)) return { value: stripV(ref), source: `tag ${ref}` }
  return null
}

function currentPublishedVersion() {
  const repo = process.env.GITHUB_REPOSITORY
  if (!repo) {
    console.error('no GITHUB_REPOSITORY in the environment and no --current given')
    console.error('pass --current <tag> for an offline check')
    process.exit(2)
  }
  if (!process.env.GH_TOKEN && !process.env.GITHUB_TOKEN) {
    console.error('GH_TOKEN (or GITHUB_TOKEN) is required to query the latest release')
    process.exit(2)
  }
  try {
    const tag = execFileSync('gh', ['api', `repos/${repo}/releases/latest`, '--jq', '.tag_name'], {
      encoding: 'utf8',
      env: { ...process.env, GH_TOKEN: process.env.GH_TOKEN || process.env.GITHUB_TOKEN },
    }).trim()
    return tag || null
  } catch (err) {
    const message = `${err.stderr ?? ''}${err.message ?? ''}`
    if (/404|Not Found/i.test(message)) return null
    throw err
  }
}

const args = parseArgs(process.argv.slice(2))
const candidate = resolveCandidate(args.version)
if (!candidate) {
  console.error('no release version: run on a v* tag push or pass --version <x.y.z>')
  process.exit(2)
}
if (!SEMVER.test(candidate.value)) {
  console.error(`release version "${candidate.value}" is not a plain x.y.z semver`)
  process.exit(1)
}

const current = args.current !== null ? stripV(args.current) : currentPublishedVersion()
const decision = releaseUploadDecision(candidate.value, current, {
  force: args.force,
  allowExisting: true,
})
if (decision.action === 'reject') {
  console.error(`release guard: ${decision.reason}`)
  console.error('(re-publishing an old version: delete the newer release first, or --force)')
  process.exit(1)
}

const note = args.force
  ? `${candidate.value} (force: forward-only check bypassed)`
  : decision.action === 'skip'
    ? `${candidate.value} matches the published release — republish allowed`
    : `${candidate.value} is newer than the published ${current ?? '(first release)'}`
console.error(`[release-version] ${candidate.source} -> ${candidate.value}; ${note}`)

const output = process.env.GITHUB_OUTPUT
if (output) {
  const { appendFileSync } = require('node:fs')
  appendFileSync(output, `version=${candidate.value}\n`)
}
console.log(candidate.value)
