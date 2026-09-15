/**
 * Browser stub for codex-app-server (Node-only).
 *
 * Replaces the real `codex-app-server.ts` in browser bundles. The renderer
 * imports `streamCodexAppServer` / `chatCodexAppServer` from the parent entry;
 * vite's renderer build externalizes the real module and routes the import
 * here so rollup never has to resolve `node:fs` / `node:crypto` /
 * `node:readline` against __vite-browser-external.
 *
 * Browser callers should never reach this code path: the registry only
 * routes 'codex-app-server' protocol in main process. If a renderer somehow
 * triggers it, throw a clear error rather than silently succeeding.
 */

export async function streamCodexAppServer(): Promise<void> {
  throw new Error(
    'codex-app-server streaming is not available in the renderer; ' +
      'route requests through the main process IPC bridge.',
  )
}

export async function chatCodexAppServer(): Promise<void> {
  throw new Error(
    'codex-app-server chat is not available in the renderer; ' +
      'route requests through the main process IPC bridge.',
  )
}
