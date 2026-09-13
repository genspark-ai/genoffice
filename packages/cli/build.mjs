// Bundles the CLI into one CommonJS file so the packaged app can run it with
// ELECTRON_RUN_AS_NODE (no node_modules ship with the app). Runtime assets
// (pdfium wasm, xlsx sidecar, OCR helper) are located by src/resources.ts.
import { build } from 'esbuild'
import { readFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const { version } = JSON.parse(await readFile(join(here, 'package.json'), 'utf8'))

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

await build({
  entryPoints: [join(here, 'src/cli.ts')],
  outfile: join(here, 'dist/genoffice.cjs'),
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'cjs',
  logLevel: 'info',
  // esbuild leaves import.meta empty in cjs output; deps use import.meta.url
  // for asset lookups, so give them the bundle's own location.
  define: {
    'import.meta.url': '__cliImportMetaUrl',
    __GENOFFICE_VERSION__: JSON.stringify(version),
  },
  banner: { js: "const __cliImportMetaUrl = require('node:url').pathToFileURL(__filename).href;" },
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
