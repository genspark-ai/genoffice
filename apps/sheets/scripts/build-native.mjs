#!/usr/bin/env node
import { spawnSync } from 'node:child_process'

const check = spawnSync(process.platform === 'win32' ? 'where.exe' : 'which', ['cargo'], {
  stdio: 'ignore',
})

if (check.status !== 0) {
  console.log('[sheets] Rust/cargo not found on PATH; skipping xlsx-engine native build.')
  process.exit(0)
}

const r = spawnSync(
  'cargo',
  [
    'build',
    '--release',
    '--manifest-path',
    'native/xlsx-engine/Cargo.toml',
    '--config',
    'native/xlsx-engine/.cargo/config.toml',
  ],
  { stdio: 'inherit', shell: process.platform === 'win32' },
)

process.exit(r.status ?? 0)
