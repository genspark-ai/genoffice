import { defineConfig } from '@playwright/test'

/**
 * E2E config for the GenOffice Electron shell.
 *
 * Tests launch the real built app (electron.launch), so they run serially —
 * parallel Electron instances fight over the GPU cache and dock on macOS.
 * Run with: `npm run test:e2e` (after `npm run build:all`).
 */
export default defineConfig({
  testDir: '.',
  outputDir: './test-results',
  timeout: 90_000,
  fullyParallel: false,
  workers: 1,
  retries: 0,
  // flat, platform-suffix-free names: baselines are Linux-CI-only (the visual
  // suite skips elsewhere), and {arg} already carries the corpus doc name
  snapshotPathTemplate: '{testDir}/visual-baselines/{arg}{ext}',
  expect: { timeout: 15_000 },
  reporter: [['list'], ['html', { outputFolder: './playwright-report', open: 'never' }]],
})
