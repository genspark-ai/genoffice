import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: {
      // resolve the sibling source package by path (not via node_modules), so a
      // git worktree whose node_modules is linked to another checkout still
      // tests against the local docx-engine edits
      '@genoffice/docx-engine': fileURLToPath(
        new URL('../docx-engine/src/index.ts', import.meta.url),
      ),
    },
  },
  test: {
    include: ['tests/**/*.test.ts'],
    testTimeout: 20000,
  },
})
