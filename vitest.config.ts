import { defineConfig } from 'vitest/config'

/// Root aggregation only: each workspace keeps its own environment and
/// resolves fixtures relative to its own directory. `npm test` still runs
/// per-workspace scripts (including the Rust sidecar build).
export default defineConfig({
  test: {
    projects: ['apps/*/vitest.config.ts', 'packages/*/vitest.config.ts'],
  },
})
