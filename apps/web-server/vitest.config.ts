import { defineConfig } from 'vitest/config'

/**
 * Vitest config for the web-server package.
 *
 * Tests live under tests/ and exercise the IPC handlers (skills,
 * marketplace, AI) directly. The marketplace-e2e suite boots the full
 * bundle on a random port and drives it end-to-end to prove that the
 * install → pi-loader → uninstall chain still works after every change.
 */
export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'node',
    testTimeout: 60_000,
    hookTimeout: 60_000,
    // Mirror the bundle's --external list so server modules resolve cleanly.
    server: {
      deps: {
        external: ['ws'],
      },
    },
  },
})
