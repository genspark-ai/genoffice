import { isAbsolute } from 'node:path'
import { z } from 'zod'
import type { McpToolDefinition } from '../mcp-server'
import { familyLabel, withSaveExtension, type SessionFamily } from './formats'

/**
 * Shared session lifecycle for the visible-editing tools.
 *
 * The three families (docs / slides / sheets) each open a real tab and drive it
 * through their own bridge, but the *lifecycle* is identical: open a blank tab,
 * edit it, save it to a path, end the session. Registering one pair of
 * create_session / save_session tools instead of six near-identical ones keeps
 * the tool list the agent sees small, while each family module keeps owning its
 * own tab + editor protocol behind a FamilyDriver.
 *
 * Family ids, labels and the saveable extensions all come from the format
 * registry (./formats) so this surface tracks the editor instead of drifting
 * from it. One session is active at a time: creating a new session replaces the
 * previous one, so the content tools always address an unambiguous target.
 */

export type { SessionFamily }

/**
 * Resolve a document reference (tab id, or path) to the webContents id of that
 * open tab. The content tools call this when the caller passed `document`, which
 * redirects the edit from the session's own blank tab to one the user has open —
 * the bridges are addressed by webContents id, so no session is involved.
 *
 * With `focus`, the tab is also brought into view (activated, and the window
 * revealed) so the user watches the change land instead of it happening behind
 * a tab they cannot see. Mutating tools ask for it; reads do not, because an
 * agent sweeping several tabs for context should not make the UI jump along.
 */
export type TargetResolver = (
  target: string,
  family: SessionFamily,
  options?: { focus?: boolean },
) => Promise<number>

/**
 * One family's visible-session lifecycle, implemented by the family's tool
 * module. `openBlankTab` opens the tab; `save` writes that tab to disk through
 * the family's own bridge.
 */
export interface FamilyDriver {
  family: SessionFamily
  openBlankTab: () => Promise<number>
  save: (wcId: number, path: string, overwrite: boolean) => Promise<unknown>
}

interface ActiveSession {
  family: SessionFamily
  wcId: number
}

/**
 * The one active session, shared between the generic lifecycle tools and the
 * per-family content tools. Created per server so tests and restarts start
 * clean.
 */
export interface SessionHost {
  begin: (family: SessionFamily, wcId: number) => void
  /** the active session itself, or null */
  current: () => ActiveSession | null
  /** the active session's tab id; throws a guided error when there is none or it is another family */
  require: (family: SessionFamily) => number
  end: () => void
}

/**
 * A driver's save result that reports failure as data instead of throwing
 * (sheets' `SaveOutcome`). Returns the message to raise, or null on success.
 */
function saveFailure(result: unknown): string | null {
  if (!result || typeof result !== 'object') return null
  const { ok, reason, error } = result as { ok?: unknown; reason?: unknown; error?: unknown }
  if (ok !== false) return null
  const detail = typeof reason === 'string' ? reason : typeof error === 'string' ? error : undefined
  return detail ? `the file could not be saved: ${detail}` : 'the file could not be saved'
}

export function createSessionHost(): SessionHost {
  let active: ActiveSession | null = null
  return {
    begin(family, wcId) {
      active = { family, wcId }
    },
    current: () => active,
    require(family) {
      if (!active) throw new Error('no session is open — call create_session first')
      if (active.family !== family) {
        throw new Error(
          `the active session is a ${familyLabel(active.family)} — call create_session with family "${family}" first`,
        )
      }
      return active.wcId
    },
    end() {
      active = null
    },
  }
}

/**
 * Where a content tool should send its command: the open document the caller
 * named, or the session's own tab when they named none.
 *
 * Naming a document is the escape hatch from the single-session model — an agent
 * can edit the workbook the user is looking at without a `create_session`, and
 * without that session's blank tab being replaced. The tab keeps its own
 * dirty/save lifecycle: MCP edits land in the live document and the user decides
 * when to save (or the caller uses `open_documents {action:"save"}`).
 */
export function contentTarget(
  host: SessionHost,
  resolveTarget: TargetResolver | undefined,
  family: SessionFamily,
): (document?: unknown, options?: { focus?: boolean }) => Promise<number> {
  return async (document, options) => {
    if (typeof document !== 'string' || !document) return host.require(family)
    if (!resolveTarget) {
      throw new Error(
        'this build cannot target an open document; use create_session and its own tab instead',
      )
    }
    return resolveTarget(document, family, options)
  }
}

export function createSessionTools(
  drivers: FamilyDriver[],
  host: SessionHost,
): McpToolDefinition[] {
  if (drivers.length === 0) return []
  const byFamily = new Map(drivers.map((d) => [d.family, d]))
  const families = drivers.map((d) => d.family) as [SessionFamily, ...SessionFamily[]]
  const listing = drivers.map((d) => `${d.family} (${familyLabel(d.family)})`).join(', ')

  return [
    {
      name: 'create_session',
      description:
        `Open a new empty file in a visible GenOffice tab and start an editing session; family picks the editor: ${listing}. ` +
        "Then add content with that family's tools (docs: insert_content / replace_blocks / apply_ops; slides: apply_slide_ops; sheets: apply_sheet_ops) and finish with save_session. " +
        'One session is active at a time — a new create_session replaces it. The user watches each step in the app.',
      inputSchema: {
        family: z.enum(families).describe('which editor to open'),
      },
      handler: async (args) => {
        const family = args.family as SessionFamily
        const driver = byFamily.get(family)
        if (!driver) throw new Error(`unknown family "${family}"`)
        const wcId = await driver.openBlankTab()
        host.begin(family, wcId)
        return {
          ok: true,
          family,
          sessionId: wcId,
          message: `A new empty ${familyLabel(family)} is open in GenOffice. Add content, then call save_session.`,
        }
      },
    },
    {
      name: 'save_session',
      description:
        'Save the active session to an absolute path and end it. A path with no extension gets the ' +
        "active family's file format appended (docx: .docx; xlsx: .xlsx; pptx: .pptx); a different " +
        'extension is refused. Refuses to replace an existing file unless overwrite is true. This is ' +
        'the output step.',
      inputSchema: {
        path: z.string().describe('absolute output path for the file'),
        overwrite: z
          .boolean()
          .optional()
          .describe('allow replacing an existing file at `path`; default false'),
      },
      handler: async (args) => {
        const session = host.current()
        if (!session) throw new Error('no session is open — call create_session first')
        const driver = byFamily.get(session.family)
        if (!driver) throw new Error(`no driver for the active "${session.family}" session`)
        const requested = String(args.path ?? '')
        if (!isAbsolute(requested)) throw new Error('path must be absolute')
        const filePath = withSaveExtension(session.family, requested)
        const result = await driver.save(session.wcId, filePath, args.overwrite === true)
        // Some drivers report a failed write as data rather than by throwing
        // (sheets returns the renderer's SaveOutcome `{ok:false}`), so returning
        // it verbatim would tell the agent the file was written when it was not.
        // Raise a tool error and keep the session open: the edits are still in
        // the tab, so the caller can correct the path and retry instead of
        // rebuilding the document from scratch.
        const failure = saveFailure(result)
        if (failure) throw new Error(failure)
        host.end()
        // Normalize the success shape: drivers differ (docs returns `{ok,path}`,
        // slides just `{path}`), and an agent that checks `ok` on every family
        // would read the slides save as a failure. A driver that returned a
        // non-object is reported as-is under `result` rather than being spread
        // into the response as indexed characters.
        const detail =
          result && typeof result === 'object' ? (result as Record<string, unknown>) : { result }
        return { ok: true, ...detail }
      },
    },
  ]
}
