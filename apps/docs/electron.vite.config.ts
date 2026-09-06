import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

// Resolve workspace packages from this checkout's sources: in a git worktree
// node_modules is a symlink into the main checkout, so bare specifiers would
// silently bundle the other checkout's (possibly stale) code.
const localAlias = {
  '@genoffice/docx-engine': resolve(__dirname, '../../packages/docx-engine/src/index.ts'),
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
    target: `http://127.0.0.1:${Number(process.env.DOCS_IPC_PORT) || 5273}`,
    changeOrigin: true,
  },
}

export default defineConfig({
  // Main and preload use only electron + node builtins; bundle everything so
  // the packaged app doesn't rely on node_modules at runtime.
  // @genoffice/* deps ship as raw TS source with extensionless imports, so they
  // must be bundled — externalizing them yields ERR_MODULE_NOT_FOUND under Node
  // (same setup as apps/slides).
  main: {
    plugins: [
      externalizeDepsPlugin({
        exclude: ['@genoffice/electron-utils', '@genoffice/font-metrics', '@genoffice/ipc-bridge'],
      }),
    ],
    resolve: { alias: localAlias },
  },
  preload: {
    // Sandboxed preload scripts cannot require arbitrary npm packages at
    // runtime, so the drop-open bridge must be bundled, not externalized.
    plugins: [
      externalizeDepsPlugin({ exclude: ['@genoffice/electron-utils', '@genoffice/ipc-bridge'] }),
    ],
    resolve: { alias: localAlias },
  },
  renderer: {
    plugins: [react()],
    resolve: { alias: localAlias },
    server: {
      // Overridable so multiple genoffice dev instances can coexist (default 5173).
      port: Number(process.env.DOCS_DEV_PORT) || 5173,
      strictPort: Boolean(process.env.DOCS_DEV_PORT),
      proxy: ipcBridgeProxy,
    },
  },
})
