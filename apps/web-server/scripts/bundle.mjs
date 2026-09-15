// Bundle the web-server with esbuild.
//
// This script exists because esbuild's CLI struggles with the `--banner:js=`
// quoting when invoked from a `package.json` script on macOS zsh, and
// because we need a small, repeatable process for the standalone web build.
//
// The banner re-creates `require` via `node:module`'s `createRequire` so
// CommonJS dependencies (most notably `word-extractor`, used by
// `@genoffice/file-parse/src/doc.ts` for legacy .doc files, which does
// `const { Buffer } = require('buffer')`) keep working inside an ESM bundle.

import { build } from 'esbuild'
import { mkdirSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const root = resolve(here, '..')
const outdir = resolve(root, 'dist/bundle')

mkdirSync(outdir, { recursive: true })

await build({
  entryPoints: [resolve(root, 'src/index.ts')],
  outfile: resolve(outdir, 'index.js'),
  bundle: true,
  platform: 'node',
  format: 'esm',
  target: 'node22',
  packages: 'bundle',
  external: ['node:*', 'ws'],
  banner: {
    js: [
      "import { createRequire as __genofficeCreateRequire } from 'node:module';",
      'const require = __genofficeCreateRequire(import.meta.url);',
    ].join('\n'),
  },
  logLevel: 'info',
})
