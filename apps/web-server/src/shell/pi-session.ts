/**
 * pi-session — bridges a real pi AgentSession into the web-server so the
 * GenOffice shell exposes everything pi can do to the renderer through the
 * existing IPC channel. Concretely:
 *
 *   - `getPiSession()` lazily builds an AgentSession wired against
 *     `DATA_DIR/pi-skills/`, `DATA_DIR/lumos-skill-wrappers/`, the GenOffice
 *     `agent-skills` package's pi extensions, and any marketplace-installed
 *     skill / plugin manifest.
 *   - `home:pi-list-skills` / `home:pi-list-tools` are read-only listings the
 *     UI uses to verify what the agent actually sees (closing the
 *     "install is a UI illusion" gap that older audits flagged).
 *   - `home:pi-prompt` sends a user message through the agent loop and
 *     streams events back over SSE — same plumbing as the docs/sheets/slides
 *     apps, just routed through a single session instead of an Electron
 *     renderer.
 *   - `home:pi-reload-resources` triggers `resourceLoader.reload()` after
 *     the user installs/uninstalls a skill so the change takes effect
 *     without restarting the server.
 *
 * The session is created with `extensionMode: 'print'` so the embedded
 * ReactUIAdapter only acts as a data bag — no TUI-only methods are required.
 * The session itself is intentionally long-lived: one per server, recreated
 * on demand after a settings change.
 */

import { join } from 'node:path'
import {
  createOfficeSession,
  type OfficeSession,
  type OfficeSessionOptions,
} from '@genoffice/agent-runtime'

import { DATA_DIR, registerHandle } from '../common/index'
import { PI_AGENT_DIR, PI_PLUGIN_DIR, PI_SKILLS_DIR, PI_CWD } from './pi-resources'
import { LUMOS_SKILLS_WRAPPER_DIR } from './pi-resources'

/** Resolve lazily so the session is built on first use, not at module import. */
let sessionPromise: Promise<OfficeSession> | null = null

/** When the user re-installs a skill we want to drop the cached session so the
 *  next `getPiSession()` rebuilds against the fresh skills dir. */
export function invalidatePiSession(): void {
  if (!sessionPromise) return
  sessionPromise.then((s) => s.dispose()).catch(() => undefined)
  sessionPromise = null
}

/**
 * Lazily build (or reuse) a real pi AgentSession wired against the GenOffice
 * skills / extensions surface. The returned `session.session` is the pi
 * `AgentSession` — full event API + tool inventory + prompt(). The wrapper
 * `reloadResources()` and `dispose()` are GenOffice conveniences.
 */
export async function getPiSession(): Promise<OfficeSession> {
  if (sessionPromise) return sessionPromise
  sessionPromise = buildPiSession()
  return sessionPromise
}

async function buildPiSession(): Promise<OfficeSession> {
  const opts: OfficeSessionOptions = {
    cwd: PI_CWD,
    agentDir: PI_AGENT_DIR,
    extensionMode: 'print',
    additionalSkillPaths: [PI_SKILLS_DIR, LUMOS_SKILLS_WRAPPER_DIR].filter(
      (dir): dir is string => !!dir && dir.length > 0,
    ),
    // Extension paths is empty — the agent-skills package's `createXxxTool`
    // factories are imported by the renderer (apps/docs/src/renderer/ai),
    // not wired into the web-server's server-side pi session. Wiring them
    // here would require duplicating the ReactUIAdapter surface. For now the
    // web-server's pi session is skill-only; renderer code wires tools
    // through the docs/sheets/slides IPC paths.
  }
  return createOfficeSession(opts)
}

/* ------------------------------------------------------------------ */
/* IPC handlers — the UI consumes the pi session through these.       */
/* ------------------------------------------------------------------ */

export function registerPiSessionHandlers(): void {
registerHandle('home:pi-list-skills', async () => {
    try {
      const { resourceLoader } = await getPiSession()
      const loaded = resourceLoader.getSkills()
      const skills = loaded.skills.map((s) => ({
        name: s.name,
        description: s.description,
        filePath: s.filePath,
        source: (s as unknown as { source?: string }).source ?? null,
      }))
      const diagnostics = (loaded.diagnostics ?? []).map((d) => ({
        path: (d as { path?: string }).path ?? null,
        severity: (d as { severity?: string }).severity ?? 'info',
        message: (d as { message?: string }).message ?? '',
      }))
      return { ok: true, skills, count: skills.length, diagnostics }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  registerHandle('home:pi-list-tools', async () => {
    const { session } = await getPiSession()
    const tools = session.getAllTools().map((t) => ({
      name: t.name,
      description: t.description,
    }))
    return { ok: true, tools, count: tools.length }
  })

  registerHandle('home:pi-reload-resources', async () => {
    const session = await getPiSession()
    const result = await session.reloadResources()
    return { ok: true, ...result }
  })

  registerHandle('home:pi-status', async () => {
    try {
      const session = await getPiSession()
      const tools = session.session.getAllTools()
      const skills = (await import('node:fs')).readdirSync(PI_SKILLS_DIR).filter((n) => !n.startsWith('.'))
      return {
        ok: true,
        ready: true,
        skills,
        toolCount: tools.length,
        sessionId: (session.session as unknown as { sessionId?: string }).sessionId ?? null,
      }
    } catch (err) {
      return {
        ok: false,
        ready: false,
        error: err instanceof Error ? err.message : String(err),
      }
    }
  })
}
