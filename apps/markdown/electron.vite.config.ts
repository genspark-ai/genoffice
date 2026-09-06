import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

// npm hoists some @tiptap packages to the repo root (shared with docs at a
// different version) and nests others under this app — dedupe forces every
// import onto this app's single copy so the bundle never carries two cores.
const TIPTAP_DEDUPE = [
  '@tiptap/core',
  '@tiptap/pm',
  '@tiptap/react',
  '@tiptap/extensions',
  '@tiptap/extension-list',
  '@tiptap/extension-table',
  '@tiptap/extension-image',
  '@tiptap/suggestion',
  '@tiptap/markdown',
  '@tiptap/extension-highlight',
  '@tiptap/extension-code-block',
]

const localAlias = {
  // subpath before the bare name: string aliases are prefix replacements
  '@genoffice/ipc-bridge/client': resolve(__dirname, '../../packages/ipc-bridge/src/client.ts'),
  '@genoffice/ipc-bridge/web-native': resolve(__dirname, '../../packages/ipc-bridge/src/web-native.ts'),
  '@genoffice/ipc-bridge': resolve(__dirname, '../../packages/ipc-bridge/src/index.ts'),
}

// Web dual-protocol: the browser talks to the Electron main process over
// HTTP/SSE served by @genoffice/ipc-bridge on this loopback port (renderer dev
// port +100, env-overridable).
const ipcBridgeProxy = {
  '/api': {
    target: `http://127.0.0.1:${Number(process.env.MARKDOWN_IPC_PORT) || 5277}`,
    changeOrigin: true,
  },
}

export default defineConfig({
  // @genoffice/i18n and @genoffice/electron-utils ship as TS source — must be bundled
  main: {
    plugins: [
      externalizeDepsPlugin({
        exclude: ['@genoffice/i18n', '@genoffice/electron-utils', '@genoffice/ipc-bridge'],
      }),
    ],
    resolve: { alias: localAlias },
  },
  preload: {
    // same bundling requirement as main (see comment above)
    plugins: [
      externalizeDepsPlugin({
        exclude: ['@genoffice/i18n', '@genoffice/electron-utils', '@genoffice/ipc-bridge'],
      }),
    ],
    resolve: { alias: localAlias },
  },
  renderer: {
    plugins: [react()],
    resolve: { dedupe: TIPTAP_DEDUPE, alias: localAlias },
    server: {
      port: Number(process.env.MARKDOWN_DEV_PORT) || 5177,
      strictPort: Boolean(process.env.MARKDOWN_DEV_PORT),
      proxy: ipcBridgeProxy,
    },
  },
})
