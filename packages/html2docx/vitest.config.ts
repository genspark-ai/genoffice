import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    // every case renders in a real Chromium
    testTimeout: 90000,
    hookTimeout: 60000,
  },
})
