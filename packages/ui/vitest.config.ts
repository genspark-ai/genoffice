import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.tsx'],
    environment: 'jsdom',
    testTimeout: 20000,
  },
})
