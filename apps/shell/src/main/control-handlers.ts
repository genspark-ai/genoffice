import { existsSync } from 'node:fs'
import type { WebContents } from 'electron'
import type {
  ControlReply,
  ControlRequest,
  ControlTarget,
  RendererControlReply,
  RendererControlRequest,
} from '@genoffice/cli/control-protocol'
import type { TabKind } from '../shared/tabs-api'

export interface ControlHost {
  reveal(): void
  openDocument(path: string): boolean
  activateTab(id: string): void
  findTab(path: string): { id: string; kind: TabKind; webContents: WebContents } | undefined
  /** ms to keep asking a renderer that has not loaded its document yet */
  readyTimeoutMs?: number
}

const TARGET_KIND_BY_TAB: Partial<Record<TabKind, ControlTarget['kind']>> = {
  slides: 'slide',
  docs: 'block',
  sheets: 'range',
  pdf: 'page',
}

/**
 * Runs one CLI control request against the open tabs. Renderers register
 * `window.__genofficeControl`; the shell evaluates it through the debugger
 * channel, so no preload surface grows and an unloaded document simply
 * answers `not_ready` until its file is in.
 */
export function controlHandler(host: ControlHost): (req: ControlRequest) => Promise<ControlReply> {
  return async (req) => {
    if (!existsSync(req.path)) {
      return fail('file_not_found', `file not found: ${req.path}`)
    }
    if (req.cmd === 'open') {
      host.reveal()
      // the tab lookup resolves links and casing; the raw-path route would open a twin tab
      const existing = host.findTab(req.path)
      if (existing) host.activateTab(existing.id)
      else if (!host.openDocument(req.path)) {
        return fail('unsupported', `GenOffice cannot open ${req.path}`)
      }
      if (!req.target) return { ok: true, result: { opened: req.path } }
      const tab = existing ?? host.findTab(req.path)
      if (!tab) return fail('file_not_open_in_gui', `no editor tab is showing ${req.path}`)
      const expected = TARGET_KIND_BY_TAB[tab.kind]
      if (expected !== req.target.kind) {
        return fail('unsupported', `a ${tab.kind} tab takes a --${flagFor(expected)} target`, {
          supported: expected ? [flagFor(expected)] : [],
        })
      }
      return ask(tab.webContents, { cmd: 'goto', target: req.target }, host.readyTimeoutMs)
    }
    const tab = host.findTab(req.path)
    if (!tab) {
      return fail('file_not_open_in_gui', `GenOffice does not have ${req.path} open`, {
        suggestion: `genoffice open ${req.path}`,
      })
    }
    return ask(tab.webContents, { cmd: 'selection' }, host.readyTimeoutMs)
  }
}

function flagFor(kind: ControlTarget['kind'] | undefined): string {
  switch (kind) {
    case 'slide':
      return 'slide'
    case 'block':
      return 'block'
    case 'range':
      return 'range'
    case 'page':
      return 'page'
    default:
      return 'target'
  }
}

function fail(
  reason: Extract<ControlReply, { ok: false }>['error']['reason'],
  message: string,
  detail?: Record<string, unknown>,
): ControlReply {
  return { ok: false, error: { reason, message, ...(detail ? { detail } : {}) } }
}

async function ask(
  wc: WebContents,
  request: RendererControlRequest,
  readyTimeoutMs = 20_000,
): Promise<ControlReply> {
  const deadline = Date.now() + readyTimeoutMs
  for (;;) {
    const reply = await evaluateControl(wc, request)
    if (reply.status === 'ok') return { ok: true, result: reply.result }
    if (reply.status === 'error') return { ok: false, error: reply.error }
    if (Date.now() > deadline) {
      return fail('app_unavailable', 'the editor did not finish loading the document in time')
    }
    await new Promise((r) => setTimeout(r, 250))
  }
}

async function evaluateControl(
  wc: WebContents,
  request: RendererControlRequest,
): Promise<RendererControlReply> {
  if (wc.isDestroyed() || wc.isLoading()) return { status: 'not_ready' }
  const script = `(async () => {
    const handler = window.__genofficeControl
    if (typeof handler !== 'function') return { status: 'not_ready' }
    try {
      return await handler(${JSON.stringify(request)})
    } catch (err) {
      return { status: 'error', error: { reason: 'app_unavailable', message: String(err && err.message || err) } }
    }
  })()`
  try {
    const reply = (await wc.executeJavaScript(script, true)) as RendererControlReply | undefined
    return reply && typeof reply === 'object' && 'status' in reply ? reply : { status: 'not_ready' }
  } catch {
    return { status: 'not_ready' }
  }
}
