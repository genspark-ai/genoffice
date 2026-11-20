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
  // e2e/local holds research probes that must never burn CI budget — they
  // wedge the Electron CDP attach and run for many minutes per launch (see
  // e2e/local/large-doc-parking.spec.ts for the local run command)
  testIgnore: ['**/local/**'],
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
