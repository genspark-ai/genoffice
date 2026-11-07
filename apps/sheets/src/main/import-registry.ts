/**
 * The delimited imports (a .csv/.tsv opened as a workbook) that are running
 * right now, one per tab, so the renderer can cancel the one it is waiting on.
 *
 * A 250 MB file takes a while, and the user should not have to wait it out.
 * The main process owns the work (it parses the text and writes the
 * workbook), so the cancel has to travel from the renderer through IPC to
 * the controller that is driving that work — this registry is where the two
 * meet.
 *
 * Kept out of `sheets-main.ts` because that module touches `electron.app` at
 * import time, and this logic is worth testing without Electron.
 */

/** Who a running import belongs to: a webContents id, one per tab. */
export type ImportOwner = number

export class ImportRegistry {
  /** the import running for an owner, if any */
  private readonly running = new Map<ImportOwner, AbortController>()

  /**
   * Start an import for an owner. A second import for the same owner
   * replaces the first: the first one's work is stale, and its controller
   * must not be left behind to be cancelled at the wrong moment.
   */
  start(owner: ImportOwner): AbortController {
    const controller = new AbortController()
    this.running.set(owner, controller)
    return controller
  }

  /**
   * Stop the import running for an owner, if there is one. Returns whether
   * something was actually stopped.
   */
  cancel(owner: ImportOwner): boolean {
    const controller = this.running.get(owner)
    if (!controller) return false
    controller.abort()
    return true
  }

  /**
   * The import running for an owner, if any.
   */
  controllerFor(owner: ImportOwner): AbortController | undefined {
    return this.running.get(owner)
  }

  /**
   * Forget an import that has finished. The owner check keeps a slow import
   * from clearing the entry of the next one that started meanwhile.
   */
  finish(owner: ImportOwner, controller: AbortController): void {
    if (this.running.get(owner) === controller) {
      this.running.delete(owner)
    }
  }
}
