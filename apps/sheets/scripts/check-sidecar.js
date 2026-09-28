const { existsSync } = require('fs')
const { execSync } = require('child_process')
const { join } = require('path')

const sidecar = join(__dirname, '../native/xlsx-engine/target/release/xlsx-sidecar.exe')
if (existsSync(sidecar)) {
  console.log('Prebuilt xlsx-sidecar.exe found. Skipping cargo build.')
  process.exit(0)
}

console.log('Building xlsx-sidecar via cargo...')
execSync(
  'cargo build --release --manifest-path native/xlsx-engine/Cargo.toml --config native/xlsx-engine/.cargo/config.toml',
  { stdio: 'inherit' },
)
