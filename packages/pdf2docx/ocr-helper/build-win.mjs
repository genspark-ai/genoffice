/**
 * Compile the Windows system-OCR helper (win-ocr.cs -> win-ocr.exe) with the
 * in-box .NET Framework compiler — no SDK install needed beyond the Windows
 * Kits metadata (present on CI runners and any dev box with VS/SDK).
 *
 * Used by apps/shell/electron-builder.cjs (packaging preflight) and
 * smoke-win.mjs (CI). Idempotent: exits 0 immediately when the exe is newer
 * than the source.
 */
import { execFileSync } from 'node:child_process'
import { existsSync, readdirSync, statSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const src = join(here, 'win-ocr.cs')
const out = join(here, 'win-ocr.exe')

if (process.platform !== 'win32') {
  console.error('build-win.mjs only runs on Windows')
  process.exit(1)
}

if (existsSync(out) && statSync(out).mtimeMs >= statSync(src).mtimeMs) {
  console.log('win-ocr.exe is up to date')
  process.exit(0)
}

/** newest version-sorted subdirectory matching `dirPattern` that contains
 * `file`. The pattern keeps non-version siblings out of the ranking:
 * UnionMetadata holds a `Facade` dir whose Windows.winmd is type-forwarding
 * only — plain lexical sort ranks 'Facade' above every '10.0.x' and CI then
 * compiles against an empty API surface (CS0234 Windows.Media.Ocr missing).
 * GAC version dirs are 'v4.0_…'-shaped, hence per-call patterns. */
function newestWith(dir, file, dirPattern) {
  if (!existsSync(dir)) return null
  const hits = readdirSync(dir)
    .filter((name) => dirPattern.test(name))
    .filter((name) => existsSync(join(dir, name, file)))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
  return hits.length > 0 ? join(dir, hits[hits.length - 1], file) : null
}

const windir = process.env.WINDIR ?? 'C:\\Windows'
const csc = join(windir, 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe')
if (!existsSync(csc)) {
  console.error(`csc.exe not found at ${csc} (.NET Framework 4.x missing?)`)
  process.exit(1)
}

const kits = 'C:\\Program Files (x86)\\Windows Kits\\10\\UnionMetadata'
const winmd = newestWith(kits, 'Windows.winmd', /^\d/) ?? join(kits, 'Windows.winmd')
if (!existsSync(winmd)) {
  console.error(`Windows.winmd not found under ${kits} (install the Windows 10/11 SDK)`)
  process.exit(1)
}

const gac = join(windir, 'Microsoft.NET', 'assembly', 'GAC_MSIL', 'System.Runtime.WindowsRuntime')
const winrtDll =
  newestWith(gac, 'System.Runtime.WindowsRuntime.dll', /^v\d/) ??
  'C:\\Program Files (x86)\\Reference Assemblies\\Microsoft\\Framework\\.NETFramework\\v4.8\\Facades\\System.Runtime.WindowsRuntime.dll'
if (!existsSync(winrtDll)) {
  console.error('System.Runtime.WindowsRuntime.dll not found (GAC or v4.8 facades)')
  process.exit(1)
}

// The union winmd's type signatures lean on the System.Runtime facades
// (CS0012 'System.Attribute is defined in an assembly that is not
// referenced' without them). Prefer the newest .NET Framework reference
// assemblies; fall back to the GAC copies.
const netfx = 'C:\\Program Files (x86)\\Reference Assemblies\\Microsoft\\Framework\\.NETFramework'
const facades = [
  ['Facades\\System.Runtime.dll', 'System.Runtime'],
  [
    'Facades\\System.Runtime.InteropServices.WindowsRuntime.dll',
    'System.Runtime.InteropServices.WindowsRuntime',
  ],
]
  .map(([refPath, gacName]) => {
    const fromRefAsm = newestWith(netfx, refPath, /^v4/)
    if (fromRefAsm) return fromRefAsm
    return newestWith(
      join(windir, 'Microsoft.NET', 'assembly', 'GAC_MSIL', gacName),
      `${gacName}.dll`,
      /^v\d/,
    )
  })
  .filter(Boolean)

// NOTE: UnionMetadata\Facade\Windows.winmd must NOT be referenced together
// with the versioned union metadata — both carry the assembly identity
// 'Windows' and csc rejects the pair with CS1704. The versioned winmd alone
// is the complete API surface.
const refs = [winmd, winrtDll, ...facades]
console.log(
  `compiling win-ocr.exe\n  csc:   ${csc}\n  refs:\n${refs.map((r) => `    ${r}`).join('\n')}`,
)
execFileSync(csc, ['/nologo', '/optimize+', ...refs.map((r) => `/r:${r}`), `/out:${out}`, src], {
  stdio: 'inherit',
})
console.log('built', out)
