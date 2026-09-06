/**
 * Univer rebuilds every header "unhide" arrow (one shape per hidden run, each
 * disposal scheduling its own MessageChannel task) on every skeleton change.
 * A streamed load of a filtered sheet with a thousand hidden runs therefore
 * paid a full rebuild per chunk. Coalesce the rebuilds so a burst of skeleton
 * changes ends in one.
 */
import { HeaderUnhideRenderController } from '@univerjs/sheets-ui'

const HEADER_UNHIDE_SETTLE_MS = 120

let installed = false

export function installHeaderUnhideDebounce(): void {
  if (installed) return
  installed = true
  const proto = HeaderUnhideRenderController.prototype as any
  const origUpdate = proto._update
  const pending = new WeakMap<object, ReturnType<typeof setTimeout>>()
  proto._update = function (workbook: unknown, worksheet: unknown) {
    const timer = pending.get(this)
    if (timer !== undefined) clearTimeout(timer)
    pending.set(
      this,
      setTimeout(() => {
        pending.delete(this)
        if (this._disposed) return
        origUpdate.call(this, workbook, worksheet)
      }, HEADER_UNHIDE_SETTLE_MS),
    )
  }
}
