import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: [
      // resolve the sibling packages by exact path (not via node_modules), so a git
      // worktree whose node_modules is linked to another checkout still tests
      // against the local sources. An exact regex keeps a string find
      // from swallowing subpath imports.
      {
        find: /^@genoffice\/compound-file$/,
        replacement: fileURLToPath(new URL('../compound-file/src/index.ts', import.meta.url)),
      },
      {
        find: /^@genoffice\/docx-engine$/,
        replacement: fileURLToPath(new URL('../docx-engine/src/index.ts', import.meta.url)),
      },
    ],
  },
  test: {
    include: ['tests/**/*.test.ts'],
    testTimeout: 20000,
  },
})
