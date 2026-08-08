/**
 * Run cargo for the sheets Rust sidecar, finding it even when ~/.cargo/bin is
 * not on PATH.
 *
 * rustup installs cargo into ~/.cargo/bin and adds that to the *user* PATH,
 * which a shell started before the install — or one that inherited its
 * environment from a longer-lived parent — never picks up. `npm test` and
 * `npm run build:all` then die on `'cargo' is not recognized`, which says
 * nothing about the Rust toolchain being the missing piece.
 *
 * Resolution order is PATH first, so an explicit rustup override or a CI
 * toolchain still wins; only then the default install location.
 */
import { spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'

const EXE = process.platform === 'win32' ? 'cargo.exe' : 'cargo'

function resolveCargo() {
  // a real probe rather than a PATH scan: this is what the spawn below will do
  if (!spawnSync(EXE, ['--version'], { stdio: 'ignore' }).error) return EXE
  const candidates = [
    process.env.CARGO_HOME ? join(process.env.CARGO_HOME, 'bin', EXE) : null,
    join(homedir(), '.cargo', 'bin', EXE),
  ]
  return candidates.find((candidate) => candidate && existsSync(candidate)) ?? null
}

const cargo = resolveCargo()

if (!cargo) {
  const hint =
    process.platform === 'win32'
      ? '  $env:PATH = "$HOME\\.cargo\\bin;$env:PATH"   # PowerShell\n' +
        '  export PATH="$HOME/.cargo/bin:$PATH"        # Git Bash\n'
      : '  export PATH="$HOME/.cargo/bin:$PATH"\n'
  process.stderr.write(
    '\ncargo was not found.\n\n' +
      'apps/sheets builds a Rust sidecar (native/xlsx-engine), so `npm test` and\n' +
      '`npm run build:all` need the Rust toolchain: https://rustup.rs\n\n' +
      "If it is already installed, ~/.cargo/bin is just not on this shell's PATH.\n" +
      'Open a new terminal, or add it for this one:\n\n' +
      hint +
      '\n',
  )
  process.exit(1)
}

process.exit(spawnSync(cargo, process.argv.slice(2), { stdio: 'inherit' }).status ?? 1)
