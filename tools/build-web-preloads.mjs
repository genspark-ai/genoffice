/**
 * Builds the web-mode browser bundles into apps/shell/web-preloads/ (gitignored):
 *
 *   client.js          — the WebSocket bridge client + dialog overlay
 *   preload-<app>.js   — each app's own preload source compiled for the
 *                        browser, with `electron` aliased to the web shim
 *
 * The renderer bundles themselves are untouched: the same electron-vite build
 * the packaged app ships is served statically by the web server. Run after
 * `npm run build:all` when preparing a web-mode session.
 */
import { createRequire } from 'node:module'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { mkdirSync } from 'node:fs'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const require = createRequire(join(root, 'packages', 'cli', 'package.json'))
const esbuild = require('esbuild')

const outDir = join(root, 'apps', 'shell', 'web-preloads')
mkdirSync(outDir, { recursive: true })

const common = {
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: 'es2022',
  minify: false,
  sourcemap: false,
  logLevel: 'info',
}

await esbuild.build({
  ...common,
  entryPoints: [join(root, 'apps', 'shell', 'src', 'web', 'shim', 'client.ts')],
  outfile: join(outDir, 'client.js'),
})

const apps = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html']
for (const app of apps) {
  await esbuild.build({
    ...common,
    entryPoints: [join(root, 'apps', app, 'src', 'preload', 'index.ts')],
    alias: { electron: join(root, 'apps', 'shell', 'src', 'web', 'shim', 'electron-stub.ts') },
    // the sheets preload's debug hook reads process.env; the browser has none
    define: { 'process.env.GENOFFICE_DEBUG_HOOKS': 'undefined' },
    outfile: join(outDir, `preload-${app}.js`),
  })
}

console.log('web preloads written to', outDir)
