import { defineConfig } from '@playwright/test'

const port = Number(process.env.MARKDOWN_DEV_PORT) || 5177

/** Browser-only renderer coverage: it must not boot the Electron shell. */
export default defineConfig({
  testDir: './tests/browser',
  testMatch: 'roundtrip-serializer.spec.ts',
  outputDir: '../../test-results/markdown-roundtrip',
  timeout: 30_000,
  workers: 1,
  webServer: {
    command: 'npm run dev:renderer -w @genoffice/markdown',
    url: `http://localhost:${port}`,
    reuseExistingServer: true,
    timeout: 60_000,
  },
})
