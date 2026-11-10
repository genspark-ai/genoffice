/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi } from '../src/shared/home-api'
import { LocaleProvider, useI18n } from '../src/renderer/src/locale'
import { McpServerSection } from '../src/renderer/src/McpServerSection'

/** the section as Settings → Integrations → MCP → B mounts it */
function Section() {
  const { t } = useI18n()
  return createElement(McpServerSection, { t })
}

const actEnvironment = globalThis as typeof globalThis & {
  IS_REACT_ACT_ENVIRONMENT?: boolean
}
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root
let setMcpSettings: ReturnType<typeof vi.fn>
let getMcpLogs: ReturnType<typeof vi.fn>

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

async function click(button: HTMLButtonElement): Promise<void> {
  await act(async () => {
    button.click()
    await Promise.resolve()
  })
}

async function renderModal(
  mcpStatus: {
    running: boolean
    enabled: boolean
    port: number
    background?: boolean
    logging?: boolean
    url: string | null
  },
  logLines?: string[],
): Promise<void> {
  setMcpSettings = vi.fn().mockResolvedValue({
    ...mcpStatus,
    running: mcpStatus.running,
  })
  getMcpLogs = vi
    .fn()
    .mockResolvedValue(
      logLines ??
        (mcpStatus.logging ? ['[mcp] listening on', '[mcp] tool get_app_info ok (0ms)'] : []),
    )
  window.aiOffice = {
    getTheme: async () => 'system',
    getDefaultSaveDir: async () => '',
    getAnalyticsEnabled: async () => true,
    getAutoSaveDefault: async () => ({ on: false, updatedAt: 0 }),
    getAiPanelPrefs: async () => ({ fontSize: 'default', spellcheck: true }),
    getUpdateChannel: async () => 'stable',
    getAppVersion: async () => '1.0.0',
    githubStars: async () => null,
    getMcpStatus: async () => mcpStatus,
    setMcpSettings,
    getMcpLogs,
    clearMcpLogs: async () => undefined,
    openMcpLogFile: async () => undefined,
  } as unknown as HomeApi

  await act(async () => {
    root.render(createElement(LocaleProvider, { initial: 'en' }, createElement(Section)))
    await Promise.resolve()
  })
}

describe('MCP local server section', () => {
  it('shows the running status, connection info and a copyable config example', async () => {
    const writeText = vi.fn<(text: string) => Promise<void>>().mockResolvedValue(undefined)
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })

    await renderModal({
      running: true,
      enabled: true,
      port: 3093,
      url: 'http://127.0.0.1:3093',
    })

    // status line: green dot + running text
    expect(host.querySelector('.set-status-dot.running')).not.toBeNull()
    expect(host.textContent).toContain('Running')

    // connection rows carry both transport URLs
    expect(host.textContent).toContain('http://127.0.0.1:3093/mcp')
    expect(host.textContent).toContain('http://127.0.0.1:3093/sse')

    // config example names the server apart from the stdio entry, with the streamable URL
    const code = host.querySelector('.set-code')
    expect(code?.textContent).toContain('"genoffice-editor"')
    expect(code?.textContent).not.toContain('"genoffice"')
    expect(code?.textContent).toContain('http://127.0.0.1:3093/mcp')

    // the first copy button copies the config example and flashes "Copied"
    const copyButtons = Array.from(host.querySelectorAll<HTMLButtonElement>('.set-btn'))
    expect(copyButtons.length).toBeGreaterThanOrEqual(4)
    await click(copyButtons[0]!)
    expect(writeText).toHaveBeenCalledWith(code!.textContent)
    expect(copyButtons[0]!.textContent).toBe('Copied')
    // the advanced rows copy the bare endpoints
    await click(copyButtons[1]!)
    expect(writeText).toHaveBeenLastCalledWith('http://127.0.0.1:3093/mcp')
  })

  it('shows the capability rows and a stopped status when disabled', async () => {
    await renderModal({
      running: false,
      enabled: false,
      port: 3093,
      url: null,
    })

    expect(host.querySelector('.set-status-dot.running')).toBeNull()
    expect(host.textContent).toContain('Not running')
  })

  it('background and logging toggles persist through setMcpSettings', async () => {
    await renderModal({
      running: true,
      enabled: true,
      port: 3093,
      background: false,
      logging: false,
      url: 'http://127.0.0.1:3093',
    })

    // health check row sits next to the other connection URLs
    expect(host.textContent).toContain('http://127.0.0.1:3093/health')
    // logging off: the log-file row is hidden
    expect(host.textContent).not.toContain('mcp-log.txt')

    const bg = host.querySelector<HTMLButtonElement>(
      '.set-switch[aria-label="Background generation"]',
    )
    expect(bg?.getAttribute('aria-checked')).toBe('false')
    await click(bg!)
    expect(setMcpSettings).toHaveBeenLastCalledWith({
      enabled: true,
      port: 3093,
      background: true,
      logging: false,
    })

    const log = host.querySelector<HTMLButtonElement>('.set-switch[aria-label="Logging"]')
    await click(log!)
    expect(setMcpSettings).toHaveBeenLastCalledWith({
      enabled: true,
      port: 3093,
      background: false,
      logging: true,
    })
  })

  it('shows the log-file row with open and clear actions when logging is on', async () => {
    await renderModal({
      running: true,
      enabled: true,
      port: 3093,
      background: true,
      logging: true,
      url: 'http://127.0.0.1:3093',
    })

    expect(host.textContent).toContain('mcp-log.txt')
    const bg = host.querySelector<HTMLButtonElement>(
      '.set-switch[aria-label="Background generation"]',
    )
    expect(bg?.getAttribute('aria-checked')).toBe('true')
    const log = host.querySelector<HTMLButtonElement>('.set-switch[aria-label="Logging"]')
    expect(log?.getAttribute('aria-checked')).toBe('true')

    // the in-pane viewer shows the log tail; refresh/clear/open actions sit beside it
    const view = host.querySelector<HTMLPreElement>('.set-log')
    expect(view?.textContent).toContain('[mcp] listening on')
    expect(view?.textContent).toContain('[mcp] tool get_app_info ok')
    const buttons = Array.from(host.querySelectorAll<HTMLButtonElement>('.set-btn'))
    expect(buttons.some((b) => b.textContent === 'Refresh')).toBe(true)
    expect(buttons.some((b) => b.textContent === 'Open')).toBe(true)
    expect(buttons.some((b) => b.textContent === 'Clear')).toBe(true)

    // Clear empties the file and re-fetches: the viewer flips to the empty state
    await click(buttons.find((b) => b.textContent === 'Clear')!)
    expect(getMcpLogs).toHaveBeenCalled()
  })

  it('shows the empty state before any log lines exist', async () => {
    await renderModal(
      { running: true, enabled: true, port: 3093, logging: true, url: 'http://127.0.0.1:3093' },
      [],
    )
    const view = host.querySelector('.set-log')
    expect(view?.textContent).toContain('No log entries yet')
  })
})
