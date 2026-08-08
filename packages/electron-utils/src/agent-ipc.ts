/**
 * The ipcMain half of the agent's shared web skill: the channels
 * `@genoffice/agent-core`'s `createWebSkill` calls at run time.
 *
 * These lived only in docs-main, which the shell calls once for every app.
 * That is fine in the shell and wrong everywhere else: slides and sheets can
 * each run as their own process (`npm run dev -w @genoffice/slides`), where
 * docs-main is never loaded, and every one of these channels answered "No
 * handler registered" — so browsing, user skills, memory and page capture were
 * dead in exactly the mode used to develop them.
 *
 * Registering them from here rather than exporting docs-main's copy keeps the
 * dependency honest: slides does not depend on the docs app, it depends on the
 * agent having a backend.
 *
 * Call once per process, and only where the app owns the generic `ai:*`
 * channels — in the shell that is docs-main alone, so the standalone entry
 * points must stay behind whatever flag already gates their AI handlers
 * (`includeAiHandlers` in sheets). Registering twice throws.
 */
import { ipcMain } from 'electron'
import { buildInstructionsPrompt, skillBodyForTool, type AppSurface } from '@genoffice/agent-core'
import { tavilyExtract } from '@genoffice/ai-search'
import type { AgentInstructionsStore } from './agent-instructions-store'
import { browsePage } from './browser'
import { currentProxyUrl } from './network-settings'

/**
 * @param store lazily resolved, because the store is constructed from
 * `app.getPath('userData')` and registration can run before the app is ready.
 */
export function registerAgentToolIpc(store: () => AgentInstructionsStore): void {
  /**
   * Prompt section for one surface, assembled in main so every editor gets the
   * same scope filtering rather than each renderer reimplementing it. Read at
   * the start of a turn, so an edit in the settings window applies to the next
   * message without reopening the document.
   */
  ipcMain.handle('ai:instructions-prompt', (_event, surface: AppSurface): string =>
    buildInstructionsPrompt(
      store().readRules(),
      store().listSkills(),
      surface,
      store().readMemories(),
    ),
  )

  /** backs the load_skill tool: the body, only if that skill is in scope here */
  ipcMain.handle('ai:skill-body', (_event, surface: AppSurface, id: string): string =>
    skillBodyForTool(store().listSkills(), surface, String(id ?? '')),
  )

  /** backs the remember tool; false when the text was not worth storing */
  ipcMain.handle(
    'ai:remember',
    (_event, text: unknown): boolean => store().addMemory(text) !== null,
  )

  /**
   * Backs the forget tool. The model only knows a memory by the wording it can
   * see in the prompt, so match on that rather than exposing ids to it.
   */
  ipcMain.handle('ai:forget', (_event, text: unknown): boolean => {
    const wanted = String(text ?? '')
      .trim()
      .toLowerCase()
    const hit = store()
      .readMemories()
      .find((m) => m.text.toLowerCase() === wanted)
    return hit ? store().deleteMemory(hit.id) : false
  })

  /**
   * Backs view_page: render what the sender is showing and hand back a PNG.
   *
   * capturePage forces a frame even when the window is occluded or on another
   * desktop, which a CDP screenshot does not — an agent run must not depend on
   * the document being the frontmost window. The rect comes from the renderer
   * because only it knows where the page element sits; an empty one captures
   * the whole view.
   */
  ipcMain.handle(
    'ai:capture-page',
    async (
      event,
      rect?: { x: number; y: number; width: number; height: number },
    ): Promise<string> => {
      try {
        const usable =
          rect && rect.width >= 1 && rect.height >= 1
            ? {
                x: Math.max(0, Math.round(rect.x)),
                y: Math.max(0, Math.round(rect.y)),
                width: Math.round(rect.width),
                height: Math.round(rect.height),
              }
            : undefined
        const image = await (usable ? event.sender.capturePage(usable) : event.sender.capturePage())
        // an empty capture is a real outcome (occluded, zero-sized); '' lets
        // the tool report it instead of shipping a blank image to the model
        return image.isEmpty() ? '' : image.toPNG().toString('base64')
      } catch (err) {
        console.warn('[ai] capture-page failed:', err)
        return ''
      }
    },
  )

  // ── agent browsing + page extraction ─────────────────────────────

  ipcMain.handle(
    'ai:browse-page',
    async (_event, url: string, opts?: { maxChars?: number; includeLinks?: boolean }) => {
      try {
        const page = await browsePage(url, {
          ...(opts?.maxChars ? { maxChars: opts.maxChars } : {}),
          includeLinks: opts?.includeLinks === true,
          proxyUrl: currentProxyUrl(),
        })
        return { ok: true as const, page }
      } catch (err) {
        return { ok: false as const, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  /** Tavily's server-side extraction: cheaper than browsing and beats bot walls */
  ipcMain.handle('ai:extract-pages', async (_event, urls: string[], advanced?: boolean) => {
    try {
      const result = await tavilyExtract(Array.isArray(urls) ? urls.map(String) : [], {
        advanced: advanced === true,
      })
      return { ok: true as const, ...result }
    } catch (err) {
      return { ok: false as const, error: err instanceof Error ? err.message : String(err) }
    }
  })
}
