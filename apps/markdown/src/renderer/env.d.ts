import type { DesktopApi, ProjectApi } from '../shared/ipc'

declare global {
  interface Window {
    markdownApi: DesktopApi
    projectApi: ProjectApi
  }
}
