#!/usr/bin/env node
/**
 * scripts/update-feed-utils.cjs — shared helpers for the electron-updater
 * feed files (latest*.yml / beta*.yml): version parsing and the
 * forward-only promote/upload guard. Used by mac-release-upload.cjs and
 * promote-stable.cjs; kept dependency-free so vitest can require it directly.
 */

function ymlVersion(text) {
  const m = /^version:\s*(\S+)/m.exec(text)
  return m ? m[1] : null
}

const NUMERIC = /^\d+$/

function parseVersion(value) {
  const core = String(value).split('+')[0]
  const dash = core.indexOf('-')
  const release = dash === -1 ? core : core.slice(0, dash)
  const parts = release.split('.')
  if (parts.length === 0 || parts.some((p) => !NUMERIC.test(p))) return null
  return {
    parts: parts.map(Number),
    prerelease: dash === -1 ? null : core.slice(dash + 1),
  }
}

function comparePrerelease(a, b) {
  if (a === null && b === null) return 0
  if (a === null) return 1
  if (b === null) return -1
  const left = a.split('.')
  const right = b.split('.')
  for (let i = 0; i < Math.max(left.length, right.length); i++) {
    const x = left[i]
    const y = right[i]
    if (x === undefined) return -1
    if (y === undefined) return 1
    const numericX = NUMERIC.test(x)
    const numericY = NUMERIC.test(y)
    if (numericX && numericY) {
      const delta = Number(x) - Number(y)
      if (delta !== 0) return delta
      continue
    }
    if (numericX !== numericY) return numericX ? -1 : 1
    if (x !== y) return x < y ? -1 : 1
  }
  return 0
}

function semverNewer(a, b) {
  const pa = parseVersion(a)
  const pb = parseVersion(b)
  if (!pa || !pb) return false
  for (let i = 0; i < 3; i++) {
    const x = pa.parts[i] ?? 0
    const y = pb.parts[i] ?? 0
    if (x !== y) return x > y
  }
  return comparePrerelease(pa.prerelease, pb.prerelease) > 0
}

function assertPromotable(candidate, currentStable, force) {
  if (force || !currentStable || semverNewer(candidate, currentStable)) return { ok: true }
  return {
    ok: false,
    reason: `${candidate} is not newer than the published stable ${currentStable}; pass --force to roll back`,
  }
}

function releaseUploadDecision(candidate, current, { force = false, allowExisting = false } = {}) {
  if (candidate === current && allowExisting) return { action: 'skip' }
  if (force || !current || semverNewer(candidate, current)) return { action: 'upload' }
  return {
    action: 'reject',
    reason: `${candidate} is not newer than published ${current}; bump the release version or pass --force to roll back`,
  }
}

module.exports = { ymlVersion, semverNewer, assertPromotable, releaseUploadDecision }
