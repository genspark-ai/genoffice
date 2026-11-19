import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  base: './',
  plugins: [react()],
  // Linked worktree dependencies must not select another checkout's engine.
  resolve: {
    alias: {
      '@genoffice/docx-engine/lazy-media': resolve(
        __dirname,
        '../../packages/docx-engine/src/lazy-media.ts',
      ),
      '@genoffice/docx-engine/zip-splice': resolve(
        __dirname,
        '../../packages/docx-engine/src/zip-splice.ts',
      ),
      '@genoffice/docx-engine': resolve(__dirname, '../../packages/docx-engine/src/index.ts'),
    },
  },
  build: {
    outDir: 'out/embedded',
    target: 'es2022',
    assetsInlineLimit: 0,
    cssCodeSplit: false,
    modulePreload: { polyfill: false },
    rollupOptions: {
      preserveEntrySignatures: 'strict',
      input: resolve(__dirname, 'src/renderer/embedded.tsx'),
      output: {
        format: 'es',
        entryFileNames: 'editor.js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: (asset) =>
          asset.names.some((name) => name.endsWith('.css'))
            ? 'editor.css'
            : 'assets/[name]-[hash][extname]',
      },
    },
  },
})
