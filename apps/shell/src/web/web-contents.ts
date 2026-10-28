import { session } from 'electron'
import type { WebContents } from 'electron'

/**
 * A browser tab standing in for an Electron WebContents inside the web-server
 * mode. Every IPC handler keeps working because it talks to this object: pushes
 * (`send`) ride the tab's WebSocket, per-tab state maps key off `id`, cleanup
 * listeners registered on 'destroyed' fire when the browser tab disconnects,
 * and `printToPDF` is delegated to the server-side headless export pipeline
 * (the browser cannot print the live DOM, and the desktop pipeline is exactly
 * what the CLI uses, so output cannot drift).
 *
 * The class implements the WebContents surface the main-process modules
 * actually touch (audited across every app's main directory); anything else
 * throws, which keeps an un-audited call site loud instead of silently wrong.
 */

let nextId = 1_000_000_000

/** Docs-style live-DOM printing is impossible in a browser tab; the bridge
 *  installs this delegate which re-renders server-side from the saved file. */
export type WebPrintToPDF = (sourceUrl: string, options: Record<string, unknown>) => Promise<Buffer>

let printToPdfDelegate: WebPrintToPDF | null = null

export function setWebPrintToPDFDelegate(fn: WebPrintToPDF | null): void {
  printToPdfDelegate = fn
}

export class WebTabContents {
  readonly id: number
  readonly session = session.defaultSession
  /** the URL the tab is served from, query included — pending-open flows and
   *  the print delegate read the ?path= parameter off it */
  readonly url: string

  private readonly destroyedListeners = new Set<() => void>()
  private destroyed = false

  constructor(url: string) {
    this.id = nextId++
    this.url = url
  }

  /** plumb a WebContents-typed reference through main-process helpers that
   *  only touch the members implemented above */
  asWebContents(): WebContents {
    return this as unknown as WebContents
  }

  send(channel: string, ...args: unknown[]): void {
    this.deliver(channel, args)
  }

  /** installed by the bridge; must be set before any tab can push */
  deliver: (channel: string, args: unknown[]) => void = () => {
    throw new Error('web tab deliver not wired')
  }

  isDestroyed(): boolean {
    return this.destroyed
  }

  getURL(): string {
    return this.url
  }

  /** dialog parenting walks host window → BrowserWindow.fromWebContents; a
   *  browser tab has no owner window and every caller accepts undefined */
  getOwnerBrowserWindow(): undefined {
    return undefined
  }

  /** the copy-image guard and friends probe the loading URL */
  isLoading(): boolean {
    return false
  }

  /** window-opening is denied before it reaches here (the shim has no
   *  window.open routing to main), so no handler state to keep */
  setWindowOpenHandler(): { action: 'deny' } {
    return { action: 'deny' }
  }

  async printToPDF(options: Record<string, unknown> = {}): Promise<Buffer> {
    if (!printToPdfDelegate) throw new Error('PDF export is not available in the browser')
    return printToPdfDelegate(this.url, options)
  }

  async getPrintersAsync(): Promise<unknown[]> {
    return []
  }

  print(): void {
    throw new Error('Native printing is not available in the browser; use the browser print dialog')
  }

  replaceMisspelling(): void {
    // browser spellcheck UI owns corrections
  }

  on(event: string, listener: () => void): this
  on(event: 'destroyed', listener: () => void): this
  on(event: string, listener: () => void): this {
    if (event === 'destroyed') {
      if (this.destroyed) listener()
      else this.destroyedListeners.add(listener)
    }
    return this
  }

  once(event: string, listener: () => void): this {
    // self-removing wrapper: Electron once() never fires twice, and the
    // listener set must not accumulate a spent copy forever
    const wrapped = (): void => {
      this.removeListener(event, wrapped)
      listener()
    }
    return this.on(event, wrapped)
  }

  removeListener(event: string, listener: () => void): this {
    if (event === 'destroyed') this.destroyedListeners.delete(listener)
    return this
  }

  /** the browser tab went away: run every cleanup the app registered */
  destroy(): void {
    if (this.destroyed) return
    this.destroyed = true
    for (const listener of [...this.destroyedListeners]) listener()
    this.destroyedListeners.clear()
  }
}

const tabsById = new Map<number, WebTabContents>()

export function registerWebTab(tab: WebTabContents): void {
  tabsById.set(tab.id, tab)
}

export function unregisterWebTab(tab: WebTabContents): void {
  tabsById.delete(tab.id)
}

export function webTabById(id: number): WebTabContents | undefined {
  return tabsById.get(id)
}
