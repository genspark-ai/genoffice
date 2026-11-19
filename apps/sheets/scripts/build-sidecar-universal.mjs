#!/usr/bin/env node
// Build the xlsx-sidecar as a macOS universal (fat) binary for release
// packaging: compile both Apple targets and lipo-merge them into
// native/xlsx-engine/target/release/xlsx-sidecar — the exact path the shell's
// electron-builder mac extraResources reads, so the host-arch dev build and
// this fat release build are interchangeable at packaging time. Both mac
// arch packages (arm64 + x64) ship the same fat sidecar.
//
// Dev builds stay on plain `native:build` (host arch only, no extra rustup
// targets required); this script is only called by the release packaging
// paths (root dist:mac and mac-build.yml). It needs both targets installed:
//   rustup target add x86_64-apple-darwin aarch64-apple-darwin

import { execFileSync } from 'node:child_process'
import { mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

function fatal(msg) {
  console.error(`[sidecar-universal] ERROR: ${msg}`)
  process.exit(1)
}

if (process.platform !== 'darwin') fatal('macOS only (needs lipo)')

// cargo resolves --config relative to the cwd, not --manifest-path (see the
// note in native/xlsx-engine/.cargo/config.toml), so run from apps/sheets
// exactly like the native:build script does.
const sheetsDir = join(dirname(fileURLToPath(import.meta.url)), '..')
const TARGETS = ['x86_64-apple-darwin', 'aarch64-apple-darwin']

for (const target of TARGETS) {
  console.log(`[sidecar-universal] cargo build --target ${target}`)
  execFileSync(
    'cargo',
    [
      'build',
      '--release',
      '--manifest-path',
      'native/xlsx-engine/Cargo.toml',
      '--config',
      'native/xlsx-engine/.cargo/config.toml',
      '--target',
      target,
    ],
    { cwd: sheetsDir, stdio: 'inherit' },
  )
}

const outDir = join(sheetsDir, 'native/xlsx-engine/target/release')
mkdirSync(outDir, { recursive: true })
const out = join(outDir, 'xlsx-sidecar')
execFileSync(
  'lipo',
  [
    '-create',
    ...TARGETS.map((t) => join(sheetsDir, 'native/xlsx-engine/target', t, 'release/xlsx-sidecar')),
    '-output',
    out,
  ],
  { stdio: 'inherit' },
)

const archs = execFileSync('lipo', ['-archs', out], { encoding: 'utf8' }).trim()
if (!archs.includes('x86_64') || !archs.includes('arm64')) {
  fatal(`merge produced "${archs}" — expected x86_64 + arm64`)
}
console.log(`[sidecar-universal] ${out} (${archs})`)
