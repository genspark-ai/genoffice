import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    setupFiles: ['tests/setup-equality.ts', 'tests/setup-strict-xml.ts'],
    testTimeout: 20000,
  },
})
