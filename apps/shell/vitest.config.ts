import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    setupFiles: ['./tests/helpers/i18n-preload.ts'],
    environment: 'node',
    testTimeout: 20000,
  },
})
