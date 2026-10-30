import { describe, expect, it, vi } from 'vitest'
import type { WebContents } from 'electron'

const { handlers } = vi.hoisted(() => ({
  handlers: new Map<string, (event: unknown, args: unknown) => unknown>(),
}))

vi.mock('electron', () => ({
  ipcMain: {
    handle: (channel: string, listener: (event: unknown, args: unknown) => unknown) => {
      handlers.set(channel, listener)
    },
  },
}))

import {
  APP_SKILL_CHANNELS,
  buildAppSkillManifest,
  executeAppSkillRead,
  registerAppSkillIpc,
  type AppDocumentReader,
  type AppSkillProxyDeps,
} from '../../src/main/mcp/app-skill-proxy'
import type { DesktopWindowInfo } from '../../src/main/desktop-control'

/**
 * The shell-main half of the desktop loop: the manifest follows the
 * open-window enumeration (only tabs of apps with a reader expose their tool),
 * and one prefixed read call routes to the right family reader and window.
 * The enumeration is a fake in the desktop-control shape; the readers are
 * fakes too — the point under test is the routing, not any app's read.
 */

function contentsFor(id: number): WebContents {
  return { id, isDestroyed: () => false } as unknown as WebContents
}

const WINDOWS: DesktopWindowInfo[] = [
  { id: 'home', kind: 'home', title: 'GenOffice', active: true },
  { id: 'tab-2', kind: 'docs', title: 'report.docx', filePath: '/tmp/report.docx', active: false },
  {
    id: 'tab-5',
    kind: 'sheets',
    title: 'budget.xlsx',
    filePath: '/tmp/budget.xlsx',
    active: false,
  },
  { id: 'tab-6', kind: 'pdf', title: 'scan.pdf', filePath: '/tmp/scan.pdf', active: false },
  { id: 'tab-7', kind: 'docs', title: 'notes.docx', active: true },
]

const READERS: Partial<Record<'docs' | 'sheets', AppDocumentReader>> = {
  docs: async () => ({ text: '0|heading|Report' }),
  sheets: async () => ({ sheets: [{ name: 'Sheet1' }] }),
}

function depsWith(
  windows: readonly DesktopWindowInfo[],
  readers: Partial<Record<string, AppDocumentReader>> = READERS,
): AppSkillProxyDeps {
  const byId = new Map(WINDOWS.map((w) => [w.id, contentsFor(Number(w.id.replace(/\D/g, '')))]))
  return {
    listWindows: () => [...windows],
    webContentsFor: (tabId) => byId.get(tabId),
    readers: readers as AppSkillProxyDeps['readers'],
  }
}

describe('the app-skill manifest', () => {
  it('offers the read tool only of apps with an open tab and a reader', () => {
    const manifest = buildAppSkillManifest(WINDOWS, READERS as AppSkillProxyDeps['readers'])
    expect(manifest.tools.map((t) => t.name)).toEqual(['docs:read_document', 'sheets:read_sheet'])
    expect(manifest.tools.map((t) => t.kind)).toEqual(['docs', 'sheets'])
  })

  it('lists every open document window (Home excluded) even without a tool', () => {
    const manifest = buildAppSkillManifest(WINDOWS, READERS as AppSkillProxyDeps['readers'])
    expect(manifest.windows.map((w) => w.id)).toEqual(['tab-2', 'tab-5', 'tab-6', 'tab-7'])
    // the pdf tab is context, but pdf has no read surface (the app is a viewer)
    expect(manifest.tools.find((t) => t.kind === 'pdf')).toBeUndefined()
  })

  it('drops a family whose tabs all closed (the degrade path)', () => {
    const windows = WINDOWS.filter((w) => w.kind !== 'sheets')
    const manifest = buildAppSkillManifest(windows, READERS as AppSkillProxyDeps['readers'])
    expect(manifest.tools.map((t) => t.name)).toEqual(['docs:read_document'])
    expect(manifest.windows.map((w) => w.kind)).not.toContain('sheets')
  })
})

describe('executeAppSkillRead routing', () => {
  it('routes by prefix to the family reader with the target window', async () => {
    const read: AppDocumentReader = async () => ({ text: 'blocks' })
    const seen: number[] = []
    const deps = {
      listWindows: () => [...WINDOWS],
      webContentsFor: (id: string) => {
        const wc = contentsFor(Number(id.replace(/\D/g, '')))
        seen.push(wc.id)
        return wc
      },
      readers: { docs: read },
    }
    const result = await executeAppSkillRead(deps, 'docs:read_document', { document: 'tab-2' })
    expect(result).toMatchObject({
      target: { id: 'tab-2', kind: 'docs', title: 'report.docx', path: '/tmp/report.docx' },
      result: { text: 'blocks' },
    })
    expect(seen).toEqual([2])
  })

  it('defaults to the family’s active window when no target is given', async () => {
    const result = await executeAppSkillRead(depsWith(WINDOWS), 'docs:read_document', {})
    expect(result.target.id).toBe('tab-7')
  })

  it('resolves a target by path as well as by tab id', async () => {
    const result = await executeAppSkillRead(depsWith(WINDOWS), 'docs:read_document', {
      document: '/tmp/report.docx',
    })
    expect(result.target.id).toBe('tab-2')
  })

  it('forwards and validates the sheets read payload', async () => {
    const read = vi.fn(async () => ({}))
    const deps = {
      listWindows: () => [...WINDOWS],
      webContentsFor: () => contentsFor(5),
      readers: { sheets: read },
    }
    await executeAppSkillRead(deps, 'sheets:read_sheet', {
      document: 'tab-5',
      addresses: ['A1', 'B2'],
      sheet: 'Sheet1',
    })
    expect(read).toHaveBeenCalledWith(expect.anything(), {
      addresses: ['A1', 'B2'],
      sheet: 'Sheet1',
    })
    await expect(
      executeAppSkillRead(deps, 'sheets:read_sheet', { addresses: 'A1' }),
    ).rejects.toThrow(/addresses.*array/)
    await expect(
      executeAppSkillRead(deps, 'sheets:read_sheet', {
        addresses: Array.from({ length: 101 }, (_, i) => `A${i}`),
      }),
    ).rejects.toThrow(/at most 100/)
  })

  it('guides a no-match target with what is actually open', async () => {
    await expect(
      executeAppSkillRead(depsWith(WINDOWS), 'docs:read_document', { document: 'tab-404' }),
    ).rejects.toThrow(/no open Word document matches "tab-404".*tab-2.*tab-7/s)
  })

  it('names the family when none of its documents is open (degrade, not crash)', async () => {
    const windows = WINDOWS.filter((w) => w.kind !== 'sheets')
    await expect(executeAppSkillRead(depsWith(windows), 'sheets:read_sheet', {})).rejects.toThrow(
      /no spreadsheet document is open/,
    )
  })

  it('refuses a family without a read surface (pdf) and an unknown tool', async () => {
    await expect(executeAppSkillRead(depsWith(WINDOWS), 'pdf:read_text', {})).rejects.toThrow(
      /PDF documents cannot be read/,
    )
    await expect(executeAppSkillRead(depsWith(WINDOWS), 'draw:read', {})).rejects.toThrow(
      /unknown app tool/,
    )
  })

  it('reports a window that closed between manifest and call', async () => {
    const deps = {
      listWindows: () => [...WINDOWS],
      webContentsFor: () => undefined,
      readers: { docs: async () => ({}) },
    }
    await expect(
      executeAppSkillRead(deps, 'docs:read_document', { document: 'tab-2' }),
    ).rejects.toThrow(/no longer open/)
  })
})

describe('the IPC registration', () => {
  it('registers both channels and serves them like invoke would', async () => {
    handlers.clear()
    registerAppSkillIpc(depsWith(WINDOWS))
    expect([...handlers.keys()].sort()).toEqual([
      APP_SKILL_CHANNELS.execute,
      APP_SKILL_CHANNELS.manifest,
    ])
    const manifest = (await handlers.get(APP_SKILL_CHANNELS.manifest)!({}, undefined)) as {
      tools: Array<{ name: string }>
    }
    expect(manifest.tools.map((t) => t.name)).toContain('docs:read_document')
    const executed = (await handlers.get(APP_SKILL_CHANNELS.execute)!(
      {},
      {
        tool: 'docs:read_document',
        input: { document: 'tab-2' },
      },
    )) as { target: { id: string } }
    expect(executed.target.id).toBe('tab-2')
  })

  it('validates the execute envelope before routing', async () => {
    handlers.clear()
    registerAppSkillIpc(depsWith(WINDOWS))
    const invoke = handlers.get(APP_SKILL_CHANNELS.execute)!
    await expect(invoke({}, {})).rejects.toThrow(/requires a tool name/)
    await expect(invoke({}, { tool: 'docs:read_document', input: 'nope' })).rejects.toThrow(
      /input must be an object/,
    )
  })
})
