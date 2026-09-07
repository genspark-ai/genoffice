/**
 * Slides module entry — wires every slides sub-domain into the shared
 * registry. Handlers are split across `core` (lifecycle), `elements`
 * (mutations), `state` (queries), `master` (master-view), and `files`
 * (file pickers).
 */
import { registerSlidesCoreHandlers } from './core.js'
import { registerSlidesElementHandlers } from './elements.js'
import { registerSlidesFileHandlers } from './files.js'
import { registerSlidesMasterHandlers } from './master.js'
import { registerSlidesStateHandlers } from './state.js'

export function registerSlidesHandlers(): void {
  registerSlidesCoreHandlers()
  registerSlidesElementHandlers()
  registerSlidesStateHandlers()
  registerSlidesMasterHandlers()
  registerSlidesFileHandlers()
}
