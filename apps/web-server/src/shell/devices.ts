/**
 * Mobile / multimodal channels.
 *
 * Neither has a backend in the standalone web server: device geometry is a
 * renderer-side concern (the client knows its own viewport), and image/table
 * understanding needs a vision model. Both report `unsupported` so callers
 * feature-detect and hide the UI instead of rendering fabricated results.
 */
import { registerHandle } from '../common/index'

const UNSUPPORTED = {
  ok: false,
  unsupported: true,
} as const

export function registerMobileHandlers(): void {
  registerHandle('mobile:get-settings', () => ({
    ...UNSUPPORTED,
    error: 'Device settings are resolved by the client, not the server.',
  }))

  registerHandle('mobile:detect', () => ({
    ...UNSUPPORTED,
    error: 'Device detection runs in the client; no server-side detection.',
  }))
}

export function registerMultimodalHandlers(): void {
  registerHandle('multimodal:analyze-image', async () => ({
    ...UNSUPPORTED,
    error: 'Image understanding needs a configured vision model.',
  }))

  registerHandle('multimodal:extract-table', async () => ({
    ...UNSUPPORTED,
    error: 'Table extraction needs a configured vision model.',
  }))
}
