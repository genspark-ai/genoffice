// Test-only shim for `node:sqlite` (referenced via `resolve.alias` in
// apps/mail/vitest.config.ts). It re-exports the real driver through a
// runtime require so the static `node:sqlite` specifier never reaches
// Vite's client (jsdom) bundler: on Node 22 `node:sqlite` is experimental
// and absent from `builtinModules`, so Vite tries to bundle the static
// import and the suite fails. A runtime require resolves in the Node
// worker on every supported engine (>=22.12).
module.exports = require('node:sqlite')
