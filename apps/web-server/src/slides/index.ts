/**
 * Slides module entry — wires every slides sub-domain into the shared
 * registry. Handlers are split across `core` (lifecycle), `elements`
 * (mutations), `state` (queries), `master` (master-view), and `files`
 * (file pickers).
 */
import { registerSlidesCoreHandlers } from './core'
import { registerSlidesElementHandlers } from './elements'
import { registerSlidesFileHandlers } from './files'
import { registerSlidesMasterHandlers } from './master'
import { registerSlidesStateHandlers } from './state'

export function registerSlidesHandlers(): void {
  registerSlidesCoreHandlers()
  registerSlidesElementHandlers()
  registerSlidesStateHandlers()
  registerSlidesMasterHandlers()
  registerSlidesFileHandlers()
}
