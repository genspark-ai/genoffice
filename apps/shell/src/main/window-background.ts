/// Native window background shown until the renderer paints its first frame.
/// Must equal `--surface` in packages/ui/src/tokens.css (light and dark) so the
/// handoff from the native window to the page is invisible.
export const SHELL_BACKGROUND_LIGHT = '#ffffff'
export const SHELL_BACKGROUND_DARK = '#1e1e1e'

export function shellWindowBackground(dark: boolean): string {
  return dark ? SHELL_BACKGROUND_DARK : SHELL_BACKGROUND_LIGHT
}
