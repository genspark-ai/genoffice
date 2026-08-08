import { defaultAiSettings } from '@genoffice/ai-provider'

/** Browser-side compatibility bridge for the Electron renderer contracts.
 * Android progressively replaces individual methods with Capacitor implementations;
 * unsupported desktop-only operations resolve safely instead of crashing the renderer.
 */
export function installRendererStub(name: string, overrides: Record<string, unknown> = {}): void {
  const target = (globalThis as any)[name] ?? {}
  const api = new Proxy({ ...target, ...overrides }, {
    get(obj, key: string | symbol) {
      if (key in obj) return (obj as any)[key]
      return (..._args: any[]) => Promise.resolve(undefined)
    },
  })
  ;(globalThis as any)[name] = api
}

export function aiDefaults() {
  return defaultAiSettings()
}

// Slides reads its renderer APIs during initial mount. Seed the minimal Android
// compatibility APIs at module evaluation time so the first render cannot race
// a useEffect-based installation in the screen component.
installRendererStub('slidesApi', {
  getAiSettings: async () => aiDefaults(),
  setAiSettings: async (settings: unknown) => localStorage.setItem('genoffice.android.ai.settings', JSON.stringify(settings)),
  aiGskStatus: async () => ({ loggedIn: false }),
  consumePendingOpen: async () => null,
  consumeNewPresentation: async () => true,
  setAutoSavePref: () => {},
  onMenuAction: () => () => {},
  onCloseCheck: () => () => {},
  reportCloseCheck: () => {},
  onCloseSaveRequest: () => () => {},
  reportCloseSaveResult: () => {},
})
installRendererStub('desktopApi', {
  getAiSettings: async () => aiDefaults(),
  aiGskStatus: async () => ({ loggedIn: false }),
})
