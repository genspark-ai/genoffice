import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'
import type { AgentSkill } from './skill'

/**
 * The shell-level desktop loop (charter workstream D3): the skills of the
 * *other* open apps, proxied into this loop.
 *
 * An app's agent executors live in its own renderer; this skill never touches
 * them directly. The host injects a manifest — which app windows are open and
 * which read tools each app exposes — and an `execute` that forwards one
 * prefixed tool call (e.g. `docs:read_document`) into the shell's main
 * process, which routes it to the target window over the app's own bridge.
 *
 * READ-ONLY BY DESIGN. The proxied surface is reading open documents (live
 * content, unsaved edits included). Write tools are deliberately absent: a
 * cross-window write would have to carry the target app's confirmation loop
 * (risk:'high' + confirmTool + the reviewable-edit UI) across a window
 * boundary, and the honest v1 is "read here, write in the app's own panel".
 * The prompt states that boundary to the model, and the main-side proxy
 * (apps/shell/src/main/mcp/app-skill-proxy.ts) only registers read executors.
 *
 * The manifest is a synchronous snapshot the host refreshes (per turn and on
 * tab changes); tool names carry the app prefix so `composeSkills` never sees
 * a collision, and a tool exists only while at least one window of its app is
 * open — a closed app degrades to "tool absent", which the model reads as
 * "nothing open there" rather than an error to retry.
 */

/** one app window the proxy can reach (a document tab, shell or detached) */
export interface ProxiedAppWindow {
  /** the tab id: the `document` argument read tools accept */
  id: string
  /** the owning app ('docs' | 'sheets' | 'slides' | 'markdown' | 'html') */
  kind: string
  title: string
  /** absolute path; absent while the document is untitled */
  filePath?: string
  active: boolean
  /** the tab carries unsaved changes (the enumeration reports it when known) */
  dirty?: boolean
  /** lives in its own OS window rather than the shell window */
  detached?: boolean
}

/** one read tool an open app exposes through the proxy */
export interface ProxiedAppTool {
  /** prefixed name, e.g. 'docs:read_document' — unique across composed skills */
  name: string
  /** the app the tool belongs to (the prefix) */
  kind: string
  description: string
  /** JSON Schema (object) describing the tool input */
  inputSchema: Record<string, unknown>
}

/** what the main-side proxy reports about the open apps right now */
export interface ProxiedAppManifest {
  windows: ProxiedAppWindow[]
  tools: ProxiedAppTool[]
}

/** the proxy's answer to one read call */
export interface ProxiedAppExecuteResult {
  /** the window that was actually read (the model cited a target; echo it) */
  target: { id: string; kind: string; title: string; path?: string }
  /** the app's own read result; formatProxiedAppResult turns it into output text */
  result: unknown
}

export interface ProxiedAppsDeps {
  /** the last manifest snapshot; null before the first refresh or when the bridge is unavailable */
  manifest(): ProxiedAppManifest | null
  /** forward one prefixed tool call into the main-process proxy */
  execute(tool: string, input: Record<string, unknown>): Promise<ProxiedAppExecuteResult>
}

const PROXIED_APPS_PROMPT = `## Open documents in other GenOffice apps (read-only)
Documents the user has open in GenOffice's other apps are readable here. Each
turn's context lists them (tab id | app | title | path).
- Read tools are prefixed by app (docs:…, sheets:…, slides:…, markdown:…,
  html:…). Pass a tab id or path as \`document\`; omit it to read that app's
  active document. A tool absent from this turn was not offered because no
  document of that app is open — do not retry it.
- These tools are READ-ONLY: they never edit, save, insert or close anything.
  When the user asks for a change to an open document, give the answer and say
  to apply it in that app's own AI panel, whose edits run through its
  confirmation flow. Never claim a proxied tool changed a document.
- Reading is not a substitute for the knowledge base: the corpus tools search
  all indexed files; the proxied tools see only what is open right now, with
  unsaved edits included.`

/**
 * Turn a proxied read result into model-facing text: the families return
 * plain text (docs, markdown, html) or JSON-serializable structures (sheets
 * workbook overview, slides deck); a structure is pretty-printed rather than
 * stringified inline so block/element references stay greppable.
 */
export function formatProxiedAppResult(result: unknown): string {
  if (typeof result === 'string') return result
  if (
    result &&
    typeof result === 'object' &&
    typeof (result as { text?: unknown }).text === 'string'
  ) {
    return (result as { text: string }).text
  }
  try {
    return JSON.stringify(result, null, 2)
  } catch {
    return String(result)
  }
}

/** the `document` input field every proxied read tool accepts (optional) */
export const PROXIED_APP_DOCUMENT_FIELD: Record<string, unknown> = {
  document: {
    type: 'string',
    description:
      'tab id or absolute path of the open document to read (from the context list); ' +
      'omit to read the app’s active document',
  },
}

const DEGRADED = (tool: string): ToolExecution => ({
  output:
    `${tool} is not available: no document of that app is open right now. ` +
    'The tool list follows the open windows, so it may appear again once one is open.',
  isError: true,
  summary: tool,
})

export function createProxiedAppsSkill(deps: ProxiedAppsDeps): AgentSkill {
  const windowLine = (w: ProxiedAppWindow): string => {
    const flags = [
      w.active ? '(active)' : null,
      w.dirty ? 'unsaved' : null,
      w.detached ? 'own window' : null,
    ]
    return [w.id, w.kind, w.title, w.filePath ?? '(never saved)', ...flags.filter(Boolean)].join(
      ' | ',
    )
  }
  return {
    id: 'proxied-apps',
    // no window of any app open: no prompt section and no tools, so the model
    // is not offered a read it cannot aim (the same degradation the knowledge
    // base uses for an empty corpus)
    get systemPrompt() {
      return deps.manifest()?.tools.length ? PROXIED_APPS_PROMPT : ''
    },
    get tools(): AgentToolDef[] {
      const manifest = deps.manifest()
      if (!manifest) return []
      return manifest.tools.map((tool) => ({
        name: tool.name,
        description: tool.description,
        inputSchema: tool.inputSchema,
      }))
    },
    buildContext: () => {
      const manifest = deps.manifest()
      if (!manifest || manifest.windows.length === 0) return ''
      const lines = manifest.windows.map(windowLine).join('\n')
      return `Open documents in other GenOffice apps (tab id | app | title | path):\n${lines}`
    },
    executeTool: async (call: AgentToolCall): Promise<ToolExecution> => {
      const manifest = deps.manifest()
      const tool = manifest?.tools.find((t) => t.name === call.name)
      if (!manifest || !tool) return DEGRADED(call.name)
      const input = { ...call.input }
      if (input.document !== undefined && typeof input.document !== 'string') {
        return {
          output: 'the `document` argument must be a tab id or path string',
          isError: true,
          summary: call.name,
        }
      }
      if (input.document === '') delete input.document
      try {
        const { target, result } = await deps.execute(call.name, input)
        return {
          output: formatProxiedAppResult(result),
          summary: `read ${target.title}`,
        }
      } catch (error) {
        // a closed tab or a busy bridge: the reason is the tool result, so the
        // model can pick another target instead of retrying the same one
        const message = error instanceof Error ? error.message : String(error)
        return { output: message, isError: true, summary: call.name }
      }
    },
  }
}
