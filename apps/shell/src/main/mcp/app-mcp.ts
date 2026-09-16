import { McpServerService, DEFAULT_MCP_PORT, type McpToolDefinition } from './mcp-server'
import { McpLogger } from './mcp-logger'
import type { CliRunner } from './cli-runner'
import { createDocumentTools, documentDriver, type DocsControl } from './tools/document-tools'
import { createSlidesTools, slidesDriver, type SlidesControl } from './tools/slides-tools'
import { createSheetsTools, sheetsDriver, type SheetsControl } from './tools/sheets-tools'
import { createSessionHost, createSessionTools, type FamilyDriver } from './tools/session-tools'
import { createOpenDocumentTools, type OpenDocumentsControl } from './tools/open-documents-tools'

/**
 * Main-process wiring for the MCP server.
 *
 * Owns the single service instance, reads the enable/port settings, and exposes
 * the small surface the shell's IPC handlers and lifecycle hooks need. Deps are
 * injected (`configureMcpRuntime`) so this module never imports `index.ts`
 * back, and stays testable.
 */

export interface McpRuntimeDeps {
  /** app version reported by get_app_info */
  version: string
  /** default folder for generated files */
  defaultSaveDir: () => string
  /** open a file in the UI (routed to the matching tab) */
  openPath: (filePath: string) => boolean
  /** drive a visible docs editor (live document session); absent in headless runs */
  docsControl?: DocsControl
  /** drive a visible slides deck (main-process session); absent in headless runs */
  slidesControl?: SlidesControl
  /** drive a visible sheets grid (renderer workbook session); absent in headless runs */
  sheetsControl?: SheetsControl
  /** the bundled genoffice CLI, backing the headless create/read tools; absent when unavailable */
  cliRunner?: CliRunner
  /**
   * documents the user has open (list/read/close); absent in headless runs,
   * where there is no tab manager to ask
   */
  openDocumentsControl?: OpenDocumentsControl
  /** where the MCP log file lives (userData); logging is unavailable without it */
  logFilePath?: string
}

export interface McpSettings {
  enabled: boolean
  port: number
  /** headless create_docx (no UI) is exposed to clients; default off */
  background: boolean
  /** server/tool activity is written to the log file; default off */
  logging: boolean
}

export interface McpStatus {
  running: boolean
  enabled: boolean
  port: number
  background: boolean
  logging: boolean
  url: string | null
  /** capability families the current build exposes (settings pane rows) */
  capabilities: string[]
}

let deps: McpRuntimeDeps | null = null
let service: McpServerService | null = null
let logger: McpLogger | null = null
let currentSettings: McpSettings = {
  enabled: false,
  port: DEFAULT_MCP_PORT,
  background: false,
  logging: false,
}

export function configureMcpRuntime(runtimeDeps: McpRuntimeDeps): void {
  deps = runtimeDeps
  logger = runtimeDeps.logFilePath ? new McpLogger(runtimeDeps.logFilePath) : null
}

/** drops a line into the log ring + file when logging is on; no-op otherwise */
function mcpLogger(message: string): void {
  if (!currentSettings.logging) return
  logger?.append(message)
}

/** the configured log file path, when logging is available at all */
export function mcpLogFilePath(): string | null {
  return deps?.logFilePath ?? null
}

/** recent log lines for the settings pane */
export function getMcpRecentLogs(): string[] {
  return logger?.recent() ?? []
}

export function clearMcpLogs(): void {
  logger?.clear()
}

/** reveal the log file in the file manager; creates it when missing */
export function revealMcpLogFile(): void {
  logger?.ensureFile()
}

function buildTools(): McpToolDefinition[] {
  if (!deps) throw new Error('MCP runtime not configured')
  // get_app_info advertises what the registered tool families can generate
  const extraFormats = [
    ...(deps.slidesControl ? ['pptx'] : []),
    ...(deps.sheetsControl ? ['xlsx'] : []),
  ]
  // one session host per tool set: create_session / save_session drive whichever
  // family is active, and each family's content tools address that same tab.
  // buildTools runs once per client session (see the server's toolsFactory), so
  // each connected client gets its own active session rather than sharing one.
  const host = createSessionHost()
  const drivers: FamilyDriver[] = [
    ...(deps.docsControl ? [documentDriver(deps.docsControl)] : []),
    ...(deps.slidesControl ? [slidesDriver(deps.slidesControl)] : []),
    ...(deps.sheetsControl ? [sheetsDriver(deps.sheetsControl)] : []),
  ]
  const cli = deps.cliRunner
  return [
    // the session entry point first: an agent picking a tool sees create_session
    ...createSessionTools(drivers, host),
    ...createDocumentTools(
      {
        version: deps.version,
        defaultSaveDir: deps.defaultSaveDir,
        // headless create_docx is opt-in: off by default so the default surface is
        // the visible document session
        background: currentSettings.background,
        // open_in_genoffice reports ok only when the file routed to a tab
        openInTab: (filePath) => {
          if (!deps) return
          const opened = deps.openPath(filePath)
          if (!opened) throw new Error(`could not open ${filePath} in GenOffice`)
        },
        docs: deps.docsControl,
        extraFormats,
        ...(cli ? { cli } : {}),
      },
      host,
    ),
    ...createSlidesTools(
      {
        defaultSaveDir: deps.defaultSaveDir,
        background: currentSettings.background,
        slides: deps.slidesControl,
        ...(cli ? { cli } : {}),
      },
      host,
    ),
    ...createSheetsTools(
      {
        defaultSaveDir: deps.defaultSaveDir,
        background: currentSettings.background,
        sheets: deps.sheetsControl,
        ...(cli ? { cli } : {}),
      },
      host,
    ),
    // documents the user has open, independent of the session above
    ...createOpenDocumentTools({
      defaultSaveDir: deps.defaultSaveDir,
      ...(deps.openDocumentsControl ? { control: deps.openDocumentsControl } : {}),
    }),
  ]
}

function ensureService(): McpServerService {
  if (!deps) throw new Error('MCP runtime not configured')
  if (!service) {
    service = new McpServerService({
      port: currentSettings.port,
      // built per client session so concurrent clients do not share one
      // active editing session; the factory re-reads settings at connect time
      toolsFactory: buildTools,
      logger: mcpLogger,
      version: deps.version,
    })
  }
  return service
}

/** Start the server if settings say it should be on. Safe to call at boot. */
export async function startMcpFromSettings(settings: McpSettings): Promise<void> {
  currentSettings = normalize(settings)
  if (!currentSettings.enabled) return
  await ensureService().start(currentSettings.port)
}

/**
 * Apply a settings change: persist semantics are the caller's job (index.ts
 * writes app-settings.json); here we reconcile the running server with the new
 * values — start, stop, or restart when the port or the exposed tool set changes.
 */
export async function applyMcpSettings(settings: McpSettings): Promise<McpStatus> {
  const next = normalize(settings)
  const wasRunning = service?.isRunning() ?? false
  // both comparisons read the OLD settings: they must happen before currentSettings is overwritten
  const portChanged = next.port !== currentSettings.port
  const backgroundChanged = next.background !== currentSettings.background
  currentSettings = next

  if (!next.enabled) {
    await service?.stop()
    return mcpStatus()
  }
  // a background flip changes the exposed tool list, so live sessions must be
  // dropped and rebuilt — same treatment as a port change
  if (!wasRunning) {
    await ensureService().start(next.port)
  } else if (portChanged || backgroundChanged) {
    await service?.stop()
    service = null
    await ensureService().start(next.port)
  }
  return mcpStatus()
}

export async function stopMcp(): Promise<void> {
  await service?.stop()
}

export function stopMcpSync(): void {
  service?.stopSync()
}

export function mcpStatus(): McpStatus {
  const running = service?.isRunning() ?? false
  return {
    running,
    enabled: currentSettings.enabled,
    port: currentSettings.port,
    background: currentSettings.background,
    logging: currentSettings.logging,
    url: running ? service!.getUrl() : null,
    capabilities: [
      'docs',
      ...(deps?.slidesControl ? ['slides'] : []),
      ...(deps?.sheetsControl ? ['sheets'] : []),
    ],
  }
}

function normalize(settings: McpSettings): McpSettings {
  const port =
    Number.isInteger(settings.port) && settings.port > 0 && settings.port < 65536
      ? settings.port
      : DEFAULT_MCP_PORT
  return {
    enabled: settings.enabled === true,
    port,
    background: settings.background === true,
    logging: settings.logging === true,
  }
}
