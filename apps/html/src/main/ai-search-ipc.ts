/**
 * Standalone search IPC for the html main process: the shared ai:* channels are
 * registered app-wide by docs-main registerAiIpc in shell mode, so tab mode
 * must not register them again (a second handler for the same channel throws).
 * Standalone html has no shell or docs process, so startHtmlStandalone wires
 * the search tools here — same source as sheets/docs/slides (Serper and the
 * other keyed providers with the DuckDuckGo fallback, configured through
 * ai-settings.json; image search additionally rides the Genspark chain when
 * the user is signed in with cloud tools on).
 */
import { app, ipcMain } from 'electron'
import { join } from 'node:path'
import { imageSearchTool, webSearchTool } from '@genoffice/ai-search'

const AI_SETTINGS_PATH = () => join(app.getPath('userData'), 'ai-settings.json')

let searchIpcRegistered = false

export function registerHtmlAiSearchIpc(): void {
  if (searchIpcRegistered) return
  searchIpcRegistered = true

  // Shared search tools (content + images): Serper with DuckDuckGo fallback
  // (same shape as sheets-main / docs-main / slides ai-ipc)
  ipcMain.handle('ai:web-search', async (_event, query: unknown, maxResults?: unknown) => {
    try {
      return await webSearchTool(
        AI_SETTINGS_PATH(),
        String(query),
        typeof maxResults === 'number' ? maxResults : 6,
      )
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })
  ipcMain.handle('ai:image-search', async (_event, query: unknown, maxResults?: unknown) => {
    try {
      return await imageSearchTool(
        AI_SETTINGS_PATH(),
        String(query),
        typeof maxResults === 'number' ? maxResults : 8,
      )
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })
}
