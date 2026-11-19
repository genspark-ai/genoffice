import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, normalizePath } from 'vite'
import { viteStaticCopy } from 'vite-plugin-static-copy'

// renderer-only dev server (embedded by shell via PDF_RENDERER_URL for HMR; no standalone Electron)
const require = createRequire(import.meta.url)
const pdfjsRoot = dirname(dirname(require.resolve('pdfjs-dist/package.json')))
// vite-plugin-static-copy globs require POSIX separators; join() breaks on Windows
const pdfjsDir = (sub: string) => normalizePath(join(pdfjsRoot, 'pdfjs-dist', sub))

export default defineConfig({
  root: 'src/renderer',
  plugins: [
    react(),
    viteStaticCopy({
      targets: [
        { src: pdfjsDir('cmaps'), dest: 'pdfjs' },
        { src: pdfjsDir('standard_fonts'), dest: 'pdfjs' },
        { src: pdfjsDir('wasm'), dest: 'pdfjs' },
      ],
    }),
  ],
  server: {
    port: Number(process.env.PDF_DEV_PORT) || 5176,
    strictPort: true,
  },
})
