import { afterEach, describe, it } from 'node:test'
import assert from 'node:assert/strict'
import { chmodSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import {
  captureProtectedFiles,
  parseArgs,
  protectedPaths,
  restoreProtectedFiles,
} from './sync-upstream.mjs'

const tempDirs = []

afterEach(() => {
  for (const dir of tempDirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

function tempDir() {
  const dir = mkdtempSync(join(tmpdir(), 'genoffice-upstream-test-'))
  tempDirs.push(dir)
  return dir
}

describe('upstream sync safety helpers', () => {
  it('parses a dry-run with explicit protected paths', () => {
    const options = parseArgs([
      '--remote',
      'upstream',
      '--branch',
      'main',
      '--settings-path',
      '/private/settings/ai-settings.json',
      '--credentials-path',
      '/private/settings/ai-credentials.json',
    ])
    assert.equal(options.apply, false)
    assert.deepEqual(protectedPaths(options), [
      '/private/settings/ai-settings.json',
      '/private/settings/ai-credentials.json',
    ])
  })

  it('rejects using one file for both protected paths', () => {
    assert.throws(
      () =>
        parseArgs(['--settings-path', '/tmp/same.json', '--credentials-path', '/tmp/same.json']),
      /different files/,
    )
  })

  it('captures and restores encrypted settings bytes and file mode', () => {
    const dir = tempDir()
    const path = join(dir, 'ai-credentials.json')
    const original = Buffer.from('{"version":1,"values":{"openai":"ciphertext"}}\n')
    writeFileSync(path, original, { mode: 0o600 })
    const snapshot = captureProtectedFiles([path])
    writeFileSync(path, 'rebase changed this')
    chmodSync(path, 0o644)

    restoreProtectedFiles(snapshot)

    assert.deepEqual(readFileSync(path), original)
    assert.equal((captureProtectedFiles([path])[0].mode ?? 0) & 0o777, 0o600)
  })

  it('does not create a missing user file during restore', () => {
    const path = join(tempDir(), 'missing-ai-settings.json')
    const snapshot = captureProtectedFiles([path])
    restoreProtectedFiles(snapshot)
    assert.equal(existsSync(path), false)
  })
})
