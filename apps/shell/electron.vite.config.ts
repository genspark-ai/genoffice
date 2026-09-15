import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import { fileURLToPath } from 'node:url'
import { dirname } from 'node:path'

const __dirname_local = dirname(fileURLToPath(import.meta.url))
const aiProviderBrowserStub = resolve(
  __dirname_local,
  '../../packages/ai-provider/src/browser/codex-app-server.ts',
)

export default defineConfig({
  // Bundle everything into the shell main (same policy as apps/docs): the
  // imported docs/sheets main modules are TS source with no build artifacts,
  // so externalizing them would break Node ESM resolution at runtime.
  // externalizeDepsPlugin keeps the previous auto-externalization of runtime
  // deps but forces @genoffice/ipc-bridge back into the bundle — it ships as
  // raw TS source that Node cannot load at runtime.
  main: {
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/ipc-bridge'] })],
  },
  preload: {
    // Sandboxed preload scripts cannot require arbitrary npm packages at
    // runtime, so the ipc-bridge/electron-utils TS sources must be bundled.
    plugins: [
      externalizeDepsPlugin({ exclude: ['@genoffice/electron-utils', '@genoffice/ipc-bridge'] }),
    ],
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/preload/index.ts'),
          // dedicated preload for the auto-update window
          update: resolve(__dirname, 'src/preload/update.ts'),
          // dedicated preload for the PDF password prompt window
          'pdf-password': resolve(__dirname, 'src/preload/pdf-password.ts'),
        },
      },
    },
  },
  renderer: {
    // Renderer is browser-only. externalizeDepsPlugin keeps Node-only
    // modules out of the browser bundle so vite's browser shim never
    // has to resolve `stat` / `readline` / `node:*` imports.
    resolve: {
      alias: [
        // codex-app-server.ts is Node-only (uses node:crypto/fs/readline).
        // Alias any variant of the path to a browser stub so renderer
        // builds don't try to resolve `stat` against __vite-browser-external.
        { find: /.*codex-app-server\.ts$/, replacement: aiProviderBrowserStub },
        { find: '@genoffice/ai-provider/codex-app-server', replacement: aiProviderBrowserStub },
      ],
    },
    plugins: [
      react(),
      externalizeDepsPlugin({
        exclude: ['@genoffice/ipc-bridge'],
        // Keep these Node-only modules out of the renderer bundle even
        // when the renderer entry chain still resolves them through type
        // re-exports. Without this, rollup fails to bundle
        // codex-app-server.ts because it imports node:fs/stat which is
        // not exported by __vite-browser-external.
      }),
    ],
    build: {
      rollupOptions: {
        external: (id) => id.includes('codex-app-server') || id.startsWith('@genoffice/ai-provider/codex-app-server'),
        input: {
          index: resolve(__dirname, 'src/renderer/index.html'),
          // strong-guidance update window (see src/main/update-window.ts)
          update: resolve(__dirname, 'src/renderer/update.html'),
          // PDF password prompt window (see src/main/pdf-password-dialog.ts)
          'pdf-password': resolve(__dirname, 'src/renderer/pdf-password.html'),
        },
      },
    },
    server: {
      port: Number(process.env.SHELL_DEV_PORT) || 5199,
      strictPort: Boolean(process.env.SHELL_DEV_PORT),
      // web version: same-origin proxy to the HTTP IPC bridge inside the
      // running Electron main process (keeps the page CSP's connect-src 'self')
      proxy: {
        '/api': {
          target: `http://127.0.0.1:${Number(process.env.SHELL_IPC_PORT) || 5299}`,
          changeOrigin: true,
        },
      },
    },
  },
})
