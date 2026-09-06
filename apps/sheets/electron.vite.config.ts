import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

// Resolve workspace packages from this checkout's sources (worktree-safe).
const ipcBridgeAlias = {
  // subpath before the bare name: string aliases are prefix replacements
  '@genoffice/ipc-bridge/client': resolve(__dirname, '../../packages/ipc-bridge/src/client.ts'),
  '@genoffice/ipc-bridge/web-native': resolve(__dirname, '../../packages/ipc-bridge/src/web-native.ts'),
  '@genoffice/ipc-bridge': resolve(__dirname, '../../packages/ipc-bridge/src/index.ts'),
}

export default defineConfig({
  main: {
    // @genoffice/* workspace packages ship TS source (no build step, no
    // compiled entry point) — externalizing them makes Node's ESM loader try
    // to resolve their relative imports at runtime and fail. Bundle those;
    // externalize everything else (Electron, zod, node builtins).
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@genoffice/ai-provider',
          '@genoffice/agent-core',
          '@genoffice/ai-search',
          '@genoffice/docx-engine',
          '@genoffice/file-parse',
          '@genoffice/electron-utils',
          '@genoffice/i18n',
          '@genoffice/pptx-render',
          '@genoffice/xlsx-gateway',
          '@genoffice/ipc-bridge',
        ],
      }),
    ],
    resolve: { alias: ipcBridgeAlias },
  },
  preload: {
    // Sandboxed preload scripts cannot require arbitrary npm packages at
    // runtime, so the drop-open bridge must be bundled, not externalized.
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/electron-utils', '@genoffice/ipc-bridge'] })],
  },
  renderer: {
    plugins: [react()],
    server: {
      port: Number(process.env.SHEETS_DEV_PORT) || 5174,
      strictPort: Boolean(process.env.SHEETS_DEV_PORT),
      // web version: same-origin proxy to the HTTP IPC bridge inside the
      // running Electron main process (keeps the page CSP's connect-src 'self')
      proxy: {
        '/api': {
          target: `http://127.0.0.1:${Number(process.env.SHEETS_IPC_PORT) || 5274}`,
          changeOrigin: true,
        },
      },
    },
  },
})
