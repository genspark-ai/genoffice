import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import { dirname, join } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import { normalizePath } from 'vite'
import { viteStaticCopy } from 'vite-plugin-static-copy'

// Non-embedded CMaps/standard fonts (e.g. CJK) need pdfjs data dirs, shipped with renderer output
const require = createRequire(import.meta.url)
const pdfjsRoot = dirname(dirname(require.resolve('pdfjs-dist/package.json')))
// vite-plugin-static-copy globs require POSIX separators; join() breaks on Windows
const pdfjsDir = (sub: string) => normalizePath(join(pdfjsRoot, 'pdfjs-dist', sub))

export default defineConfig({
  // @genoffice/i18n ships as TS source; pdf-lib's package only includes out/** — both must be bundled
  main: {
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@genoffice/i18n',
          'pdf-lib',
          '@genoffice/electron-utils',
          '@genoffice/font-metrics',
          '@genoffice/ipc-bridge',
        ],
      }),
    ],
    resolve: {
      alias: {
        // subpath before the bare name: string aliases are prefix replacements
        '@genoffice/ipc-bridge/client': resolve(__dirname, '../../packages/ipc-bridge/src/client.ts'),
  '@genoffice/ipc-bridge/web-native': resolve(__dirname, '../../packages/ipc-bridge/src/web-native.ts'),
        '@genoffice/ipc-bridge': resolve(__dirname, '../../packages/ipc-bridge/src/index.ts'),
      },
    },
  },
  preload: {
    // i18n and electron-utils ship as TS source — must be bundled, not left external
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/i18n', '@genoffice/electron-utils', '@genoffice/ipc-bridge'] })],
  },
  renderer: {
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
      strictPort: Boolean(process.env.PDF_DEV_PORT),
      // web version: same-origin proxy to the HTTP IPC bridge inside the
      // running Electron main process (keeps the page CSP's connect-src 'self')
      proxy: {
        '/api': {
          target: `http://127.0.0.1:${Number(process.env.PDF_IPC_PORT) || 5276}`,
          changeOrigin: true,
        },
      },
    },
  },
})
