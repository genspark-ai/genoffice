/**
 * Rebuild app preloads whose sources are newer than the built artifact.
 *
 * In dev mode the shell loads each app's preload from apps/<app>/out/preload/
 * straight off disk; `npm run dev` only starts renderer vite servers and never
 * rebuilds preloads, so after a pull the artifact can silently miss newly added
 * preload APIs (calls fail with "... is not a function" in the renderer console
 * only). Runs as the root `predev` hook: a fresh tree is a fast no-op, a stale
 * app gets a full `electron-vite build` (the only entry point that emits the
 * preload bundle).
 */
import { existsSync, readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { spawnSync } from 'node:child_process'

const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown']

/** Newest mtime (ms) under dir, 0 when missing. */
function newestMtime(dir) {
  if (!existsSync(dir)) return 0
  let newest = 0
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, entry.name)
    newest = Math.max(newest, entry.isDirectory() ? newestMtime(p) : statSync(p).mtimeMs)
  }
  return newest
}

const stale = APPS.filter((app) => {
  const artifact = join('apps', app, 'out', 'preload', 'index.js')
  const built = existsSync(artifact) ? statSync(artifact).mtimeMs : 0
  // shared/ is included: preloads import IPC channel/type modules from there
  const src = Math.max(
    newestMtime(join('apps', app, 'src', 'preload')),
    newestMtime(join('apps', app, 'src', 'shared')),
  )
  return src > built
})

if (stale.length) {
  console.log(`Rebuilding stale preloads: ${stale.join(', ')}`)
  for (const app of stale) {
    const r = spawnSync('npm', ['run', 'build', '-w', `@genoffice/${app}`], { stdio: 'inherit' })
    if (r.status !== 0) process.exit(r.status ?? 1)
  }
}
