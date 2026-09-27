// Bundles the CLI into one CommonJS file so the packaged app can run it with
// ELECTRON_RUN_AS_NODE (no node_modules ship with the app). Runtime assets
// (pdfium wasm, xlsx sidecar, OCR helper) are located by src/resources.ts.
import { build } from 'esbuild'
import { readFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const { version } = JSON.parse(await readFile(join(here, 'package.json'), 'utf8'))

/**
 * A packaged build overrides the app version through electron-builder
 * extraMetadata, which never reaches this file; GENOFFICE_APP_VERSION carries
 * it in, so the bundled `genoffice --version` cannot drift from the app that
 * shipped it. Unset (a plain workspace build) keeps the CLI package version.
 */
export const CLI_VERSION_ENV = 'GENOFFICE_APP_VERSION'
export const CLI_BUNDLE = join(here, 'dist/genoffice.cjs')
export const CLI_VERSION_BANNER_PREFIX = 'const __cliAppVersion = '

export function resolveCliVersion(env, packageVersion) {
  const injected = typeof env[CLI_VERSION_ENV] === 'string' ? env[CLI_VERSION_ENV].trim() : ''
  return injected || packageVersion
}

/** The banner esbuild copies verbatim, so the packaged bundle states its own version. */
export function cliVersionBanner(cliVersion) {
  return `${CLI_VERSION_BANNER_PREFIX}${JSON.stringify(cliVersion)};`
}

export function readBundledCliVersion(bundleText) {
  const line = bundleText.split('\n').find((row) => row.startsWith(CLI_VERSION_BANNER_PREFIX))
  if (!line) return null
  try {
    const value = JSON.parse(line.slice(CLI_VERSION_BANNER_PREFIX.length).replace(/;\s*$/, ''))
    return typeof value === 'string' ? value : null
  } catch {
    return null
  }
}

// Vite's `?raw` text imports (the pptx op guides are markdown files)
const rawText = {
  name: 'raw-text',
  setup(b) {
    b.onResolve({ filter: /\?raw$/ }, (args) => ({
      path: resolve(args.resolveDir, args.path.replace(/\?raw$/, '')),
      namespace: 'raw-text',
    }))
    b.onLoad({ filter: /.*/, namespace: 'raw-text' }, async (args) => ({
      contents: await readFile(args.path, 'utf8'),
      loader: 'text',
    }))
  },
}

/**
 * The bundle runs under ELECTRON_RUN_AS_NODE, where `require('electron')` has
 * no module to resolve to. `electron` is external below, so a top-level import
 * anywhere in the dependency graph (a package barrel that re-exports a preload
 * or main-process helper, for instance) survives bundling and crashes the CLI
 * on startup instead of failing the build. Returns the offending module paths
 * esbuild recorded above each `require("electron")`.
 */
export function findElectronRequires(bundleText) {
  const modules = []
  let current = '(unknown module)'
  for (const line of bundleText.split('\n')) {
    const header = /^\s*\/\/ ((?:\.\.\/|src\/|node_modules\/)\S+)$/.exec(line)
    if (header) current = header[1]
    else if (/\brequire\((["'])electron\1\)/.test(line) && !modules.includes(current)) {
      modules.push(current)
    }
  }
  return modules
}

async function bundle() {
  const cliVersion = resolveCliVersion(process.env, version)
  await build({
    entryPoints: [join(here, 'src/cli.ts')],
    outfile: CLI_BUNDLE,
    bundle: true,
    platform: 'node',
    target: 'node22',
    format: 'cjs',
    logLevel: 'info',
    // esbuild leaves import.meta empty in cjs output; deps use import.meta.url
    // for asset lookups, so give them the bundle's own location.
    define: {
      'import.meta.url': '__cliImportMetaUrl',
      __GENOFFICE_VERSION__: JSON.stringify(cliVersion),
    },
    banner: {
      js: [
        "const __cliImportMetaUrl = require('node:url').pathToFileURL(__filename).href;",
        cliVersionBanner(cliVersion),
      ].join('\n'),
    },
    jsx: 'automatic',
    // jsdom reads its own stylesheet from disk and cannot be inlined; it ships beside the bundle
    // mermaid is only reached by the markdown editor's lazy diagram renderer, which never runs headless
    external: ['electron', 'jsdom', 'mermaid'],
    loader: {
      '.css': 'empty',
      '.ttf': 'empty',
      '.woff': 'empty',
      '.woff2': 'empty',
      '.svg': 'empty',
      '.png': 'empty',
    },
    plugins: [rawText],
  })
  const offenders = findElectronRequires(await readFile(CLI_BUNDLE, 'utf8'))
  if (offenders.length > 0) {
    throw new Error(
      `CLI bundle requires "electron" (unavailable under ELECTRON_RUN_AS_NODE) via:\n` +
        offenders.map((module) => `  ${module}`).join('\n') +
        '\nImport from an electron-free subpath instead of the package barrel.',
    )
  }
}

// imported by tests and by packaging checks, which only want the helpers above
if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) await bundle()
