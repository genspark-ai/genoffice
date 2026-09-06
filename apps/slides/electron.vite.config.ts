import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'

const here = dirname(fileURLToPath(import.meta.url))

// Pin resolution to this repo's workspace sources (matches tsconfig paths;
// avoids bundling stale implementations when node_modules links point elsewhere)
const workspaceAlias = {
  // Subpath before the bare name: string aliases are prefix replacements
  '@genoffice/pptx-engine/table-grid': resolve(
    here,
    '../../packages/pptx-engine/src/table-grid.ts',
  ),
  '@genoffice/pptx-engine/identity': resolve(here, '../../packages/pptx-engine/src/identity.ts'),
  '@genoffice/pptx-engine/custgeom': resolve(here, '../../packages/pptx-engine/src/custgeom.ts'),
  '@genoffice/pptx-engine/background-promote': resolve(
    here,
    '../../packages/pptx-engine/src/background-promote.ts',
  ),
  '@genoffice/pptx-engine': resolve(here, '../../packages/pptx-engine/src/index.ts'),
  '@genoffice/pptx-ops/op-docs': resolve(here, '../../packages/pptx-ops/src/op-docs.ts'),
  '@genoffice/pptx-ops/font-size': resolve(here, '../../packages/pptx-ops/src/font-size.ts'),
  '@genoffice/pptx-ops': resolve(here, '../../packages/pptx-ops/src/index.ts'),
  '@genoffice/pptx-render/preset-geometry': resolve(
    here,
    '../../packages/pptx-render/src/preset-geometry.ts',
  ),
  '@genoffice/pptx-render': resolve(here, '../../packages/pptx-render/src/index.ts'),
  '@genoffice/pipelines/slides/layout-audit': resolve(
    here,
    '../../packages/pipelines/src/slides/layout-audit.ts',
  ),
  '@genoffice/pipelines/slides': resolve(here, '../../packages/pipelines/src/slides/index.ts'),
  // Metafile (EMF/WMF) rasterizer shared with the docs engine (renderer-only: needs canvas)
  '@genoffice/docx-engine/metafile': resolve(here, '../../packages/docx-engine/src/metafile.ts'),
  // HTTP IPC bridge (web dual-protocol server)
  '@genoffice/ipc-bridge/client': resolve(here, '../../packages/ipc-bridge/src/client.ts'),
  '@genoffice/ipc-bridge/web-native': resolve(here, '../../packages/ipc-bridge/src/web-native.ts'),
  '@genoffice/ipc-bridge': resolve(here, '../../packages/ipc-bridge/src/index.ts'),
}

export default defineConfig({
  // Main process/preload must bundle @genoffice/* sources (they are pulled in as TS
  // source with extensionless relative imports; externalizing them under Node
  // yields ERR_MODULE_NOT_FOUND).
  main: {
    resolve: { alias: workspaceAlias },
    // Bundle opentype.js too (the packaged app ships only out/**, so external deps are unresolvable at runtime)
    plugins: [
      externalizeDepsPlugin({
        exclude: [
          '@genoffice/pptx-engine',
          '@genoffice/pptx-ops',
          '@genoffice/pptx-render',
          '@genoffice/pipelines',
          '@genoffice/ai-search',
          '@genoffice/file-parse',
          '@genoffice/electron-utils',
          '@genoffice/ipc-bridge',
          'opentype.js',
        ],
      }),
    ],
  },
  preload: {
    // electron-utils ships raw TS source — must be bundled, not left external
    plugins: [externalizeDepsPlugin({ exclude: ['@genoffice/electron-utils', '@genoffice/ipc-bridge'] })],
  },
  renderer: {
    resolve: { alias: workspaceAlias },
    plugins: [react()],
    server: {
      port: Number(process.env.SLIDES_DEV_PORT) || 5175,
      strictPort: Boolean(process.env.SLIDES_DEV_PORT),
      // web version: same-origin proxy to the HTTP IPC bridge inside the
      // running Electron main process (keeps the page CSP's connect-src 'self')
      proxy: {
        '/api': {
          target: `http://127.0.0.1:${Number(process.env.SLIDES_IPC_PORT) || 5275}`,
          changeOrigin: true,
        },
      },
    },
  },
})
