import { defineConfig } from '@playwright/test'
import base from '../playwright.config'

/**
 * Local-only runner for the research probes in e2e/local. The main config
 * (e2e/playwright.config.ts) testIgnores this directory so the CI sweep
 * never collects it; this config exists to collect just it:
 *
 *   npx playwright test --config e2e/local/playwright.local.config.ts
 *
 * Everything else (workers, timeouts, reporter) inherits from the main
 * config; testDir is this directory.
 */
export default defineConfig({
  ...base,
  testDir: '.',
  testIgnore: [],
})
