import { RENDERER_SCHEME } from '@genoffice/electron-utils'

/**
 * Whether a getDisplayMedia request may be answered with a screen.
 *
 * The display-media handler sits on the shared default session, so an
 * ungated grant hands the first screen to any page in any app of the suite —
 * a hostile document's script in the HTML preview iframe included, one user
 * click away on Windows/Linux, where there is no system picker to confirm
 * the capture (macOS's `useSystemPicker` is what makes the ungated handler
 * acceptable there). Only the Slides renderer may ask: its scheme URL in
 * packaged builds, its dev-server origin in dev.
 */
export function isSlidesScreenRecordingSource(
  frameUrl: string | undefined,
  devUrl: string | undefined,
): boolean {
  if (!frameUrl) return false
  let url: URL
  try {
    url = new URL(frameUrl)
  } catch {
    return false
  }
  if (url.protocol === `${RENDERER_SCHEME}:` && url.hostname === 'slides') return true
  if (devUrl) {
    try {
      const dev = new URL(devUrl)
      if (url.protocol === dev.protocol && url.host === dev.host) return true
    } catch {
      /* an unusable dev URL falls back to the scheme check above */
    }
  }
  return false
}
