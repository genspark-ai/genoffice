import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

const here = dirname(fileURLToPath(import.meta.url))

export default defineConfig({
  resolve: {
    alias: {
      // Resolve sibling sources by path (not via node_modules) so a worktree
      // whose node_modules is linked to another checkout still tests local edits.
      '@genoffice/pptx-engine/table-grid': resolve(here, '../pptx-engine/src/table-grid.ts'),
      '@genoffice/pptx-engine/identity': resolve(here, '../pptx-engine/src/identity.ts'),
      '@genoffice/pptx-engine/background-promote': resolve(
        here,
        '../pptx-engine/src/background-promote.ts',
      ),
      '@genoffice/pptx-engine': resolve(here, '../pptx-engine/src/index.ts'),
      '@genoffice/docx-engine/math': resolve(here, '../docx-engine/src/math.ts'),
    },
  },
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'node',
  },
})
