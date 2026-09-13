// Copies the runtime dependencies the bundle leaves external (jsdom and its
// tree) into dist/node_modules, mirroring their layout under the checkout's
// node_modules so nested versions keep resolving. The packaged app ships the
// result beside genoffice.cjs (Resources/cli/node_modules); src/dom.ts resolves
// jsdom from there.
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs'
import { dirname, join, relative, resolve, sep } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const rootModules = resolve(here, '../../node_modules')
const out = join(here, 'dist', 'node_modules')
const EXTERNALS = ['jsdom']

function resolvePkg(name, from) {
  let dir = from
  for (;;) {
    const candidate = join(dir, 'node_modules', name)
    if (existsSync(join(candidate, 'package.json'))) return candidate
    const parent = dirname(dir)
    if (parent === dir) return null
    dir = parent
  }
}

const seen = new Set()
function walk(name, from) {
  const dir = resolvePkg(name, from)
  if (!dir) throw new Error(`genoffice deps: cannot resolve ${name} from ${from}`)
  if (seen.has(dir)) return
  seen.add(dir)
  const pkg = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'))
  for (const dep of Object.keys(pkg.dependencies ?? {})) walk(dep, dir)
  for (const dep of Object.keys(pkg.optionalDependencies ?? {})) {
    if (resolvePkg(dep, dir)) walk(dep, dir)
  }
}
for (const name of EXTERNALS) walk(name, resolve(here, '..', '..'))

rmSync(out, { recursive: true, force: true })
mkdirSync(out, { recursive: true })
for (const dir of seen) {
  const rel = relative(rootModules, dir)
  if (rel.startsWith('..'))
    throw new Error(`genoffice deps: ${dir} is outside the root node_modules`)
  cpSync(dir, join(out, rel), {
    recursive: true,
    dereference: true,
    // nested node_modules are copied by their own closure entries
    filter: (src) => !relative(dir, src).split(sep).includes('node_modules'),
  })
}
console.log(`genoffice deps: ${seen.size} packages → ${relative(here, out)}`)
