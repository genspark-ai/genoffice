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
  ReactUIAdapter,
  createOfficeSession,
  type OfficeSession,
  type OfficeSessionOptions,
} from '@genoffice/agent-runtime'

import { DATA_DIR, registerHandle } from '../common/index'
import { createWebSearchExtension } from '@genoffice/agent-skills/extensions/web-search-skill'
import { createImageSearchExtension } from '@genoffice/agent-skills/extensions/image-search-skill'
import { createOcrExtension } from '@genoffice/agent-skills/extensions/ocr-skill'
import { installAgentTeam } from '@genoffice/agent-skills/extensions/agent-team'
import { JsonlAuditSink, installAuditLog } from '@genoffice/agent-skills/extensions/audit-log'
import { installLocalModels } from '@genoffice/agent-skills/extensions/local-models'
import { getEnabledBuiltinIds } from './skills'
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
  const enabled = getEnabledBuiltinIds()
  const enabledSet = new Set<string>([...enabled.skills, ...enabled.plugins])
  // One shared UI adapter so every built-in extension and the host session
  // see the same data bag (DialogRequest / NotificationItem / etc.).
  const uiAdapter = new ReactUIAdapter()
  // In-process extension factories — preferred over file paths because the
  // web-server bundle already has every agent-skills extension loaded.
  // The translate-skill is the unification point: when this factory is in
  // the list, the same 6 tools (translate_text / translate_file /
  // build_dictionary / kb_search / kb_upsert / kb_remove) become visible
  // to the embedded AgentSession AND to the UI through the home:translate-*
  // IPC handlers registered below. No more `translate-http.ts` bypass.
  const { createTranslateSkillExtension } = await import(
    '@genoffice/agent-skills/extensions/translate-skill'
  )
  const extensionFactories: NonNullable<OfficeSessionOptions['extensionFactories']> = [
    createTranslateSkillExtension(),
  ]
  if (enabledSet.has('web-search')) extensionFactories.push(createWebSearchExtension())
  if (enabledSet.has('image-search')) extensionFactories.push(createImageSearchExtension())
  if (enabledSet.has('ocr')) extensionFactories.push(createOcrExtension())
  if (enabledSet.has('agent-team')) extensionFactories.push((pi) => installAgentTeam(pi, {}))
  if (enabledSet.has('audit-log')) extensionFactories.push((pi) =>
    installAuditLog(pi, {
      sink: new JsonlAuditSink({
        filePath: process.env.GENOFFICE_AUDIT_LOG ?? `${process.env.DATA_DIR ?? '.genoffice'}/audit-log.jsonl`,
      }),
    }),
  )
  if (enabledSet.has('local-models')) extensionFactories.push((pi) => installLocalModels(pi, {}))

  const opts: OfficeSessionOptions = {
    cwd: PI_CWD,
    agentDir: PI_AGENT_DIR,
    extensionMode: 'print',
    uiAdapter,
    additionalSkillPaths: [PI_SKILLS_DIR, LUMOS_SKILLS_WRAPPER_DIR].filter(
      (dir): dir is string => !!dir && dir.length > 0,
    ),
    extensionFactories,
  }
  return createOfficeSession(opts)
}

/* ------------------------------------------------------------------ */
/* IPC handlers — the UI consumes the pi session through these.       */
/* ------------------------------------------------------------------ */

/**
 * Look up a tool by name in the live pi session and execute it. The
 * shared entry point for every UI / IPC handler that needs to talk to
 * a translate-skill tool (kb_list / kb_search / kb_upsert / kb_remove /
 * translate_text / translate_file / build_dictionary). The same function
 * the agent uses — one source of truth for the whole translation surface.
 */
export async function callTranslateTool(name: string, args: unknown): Promise<unknown> {
    const { session: agent } = await getPiSession()
    const tool = agent.getToolDefinition(name)
    if (!tool) {
      return { ok: false, error: `translate-skill tool "${name}" not registered in pi session` }
    }
    // The IPC transport delivers args as an array. If the caller passed a single
  // object, unwrap it. If they passed an array of positional args, take the
  // first one as the params object.
  const params = Array.isArray(args)
    ? (args.length > 0 ? args[0] : {})
    : (args ?? {})
  const result = await (tool as unknown as { execute: (id: string, params: Record<string, unknown>, signal: AbortSignal | undefined) => Promise<{ content: Array<{ type: string; text?: string }>; details: unknown }> }).execute(`ui-${Date.now()}`, params as Record<string, unknown>, undefined)
    const first = (result?.content ?? []).find((c: { type?: string }) => c.type === 'text') as { text?: string } | undefined
    return {
      ok: (result?.details as { ok?: boolean } | undefined)?.ok ?? true,
      details: result?.details ?? null,
      summary: first?.text ?? '',
    }
  }

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
    const office = await getPiSession()
    const result = await office.reloadResources()
    return { ok: true, ...result }
  })

  registerHandle('home:pi-status', async () => {
    try {
      const { session } = await getPiSession()
      const tools = session.getAllTools()
      const skills = (await import('node:fs')).readdirSync(PI_SKILLS_DIR).filter((n) => !n.startsWith('.'))
      return {
        ok: true,
        ready: true,
        skills,
        toolCount: tools.length,
        sessionId: (session as unknown as { sessionId?: string }).sessionId ?? null,
      }
    } catch (err) {
      return {
        ok: false,
        ready: false,
        error: err instanceof Error ? err.message : String(err),
      }
    }
  })

  // ------------------------------------------------------------------
  // UI-facing translation handlers — every call goes through the
  // translate-skill pi extension registered above. The legacy
  // chat.ts handlers (`ai:translate`, `ai:translate-file`, ...) remain
  // untouched so existing renderer code keeps working while we migrate.
  // ------------------------------------------------------------------



  registerHandle('home:translate-text', async (_event, args) => {
    return callTranslateTool('translate_text', args)
  })
  registerHandle('home:translate-build-dictionary', async (_event, args) => {
    return callTranslateTool('build_dictionary', args)
  })
  registerHandle('home:translate-kb-list', async (_event, args) => {
    return callTranslateTool('kb_list', args)
  })
  registerHandle('home:translate-kb-search', async (_event, args) => {
    return callTranslateTool('kb_search', args)
  })
  registerHandle('home:translate-kb-upsert', async (_event, args) => {
    return callTranslateTool('kb_upsert', args)
  })
  registerHandle('home:translate-kb-remove', async (_event, args) => {
    return callTranslateTool('kb_remove', args)
  })
  registerHandle('home:translate-file', async (_event, args) => {
    return callTranslateTool('translate_file', args)
  })
}
