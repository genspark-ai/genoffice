import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    // `node:sqlite` is experimental on Node 22 and absent from
    // `builtinModules`, so Vite's client (jsdom) bundler tries to bundle
    // the static import in src/main/store/sqlite.ts and fails. Route it
    // through a runtime-require shim (tests/shims/node-sqlite.cjs) that
    // resolves in the Node worker on every supported engine.
    alias: [
      {
        find: /^node:sqlite$/,
        replacement: fileURLToPath(new URL('./tests/shims/node-sqlite.cjs', import.meta.url)),
      },
    ],
  },
  test: {
    include: ['tests/**/*.test.ts', 'tests/**/*.test.tsx'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
