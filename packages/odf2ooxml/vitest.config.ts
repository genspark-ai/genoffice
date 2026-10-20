import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: [
      // resolve the sibling source packages by exact path (not via node_modules), so
      // a git worktree whose node_modules is linked to another checkout still
      // tests against the local engine edits. Exact regexes: a string find would
      // also swallow the engine's own '@genoffice/pptx-engine/custgeom' subpath.
      {
        find: /^@genoffice\/docx-engine$/,
        replacement: fileURLToPath(new URL('../docx-engine/src/index.ts', import.meta.url)),
      },
      {
        find: /^@genoffice\/pptx-engine$/,
        replacement: fileURLToPath(new URL('../pptx-engine/src/index.ts', import.meta.url)),
      },
      {
        find: /^@genoffice\/pptx-engine\/custgeom$/,
        replacement: fileURLToPath(new URL('../pptx-engine/src/custgeom.ts', import.meta.url)),
      },
    ],
  },
  test: {
    include: ['tests/**/*.test.ts'],
    testTimeout: 20000,
  },
})
