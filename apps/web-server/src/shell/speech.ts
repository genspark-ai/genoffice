/**
 * Speech channels — recognize / synthesize / list voices.
 *
 * The standalone web server ships without an STT/TTS provider, so these
 * handlers report an explicit `unsupported` error rather than returning
 * fabricated transcripts or empty audio. Callers can feature-detect on the
 * error and hide the microphone UI. The Electron build keeps the real
 * provider-backed implementation.
 */
import { registerHandle } from '../common/index'

const UNSUPPORTED = {
  ok: false,
  unsupported: true,
  error: 'Speech recognition/synthesis needs a configured STT/TTS provider; not available in the web build.',
} as const

export function registerSpeechHandlers(): void {
  registerHandle('speech:recognize', async (_event: unknown, args: unknown) => {
    const { language } = (args || {}) as { audioBytes?: ArrayBuffer; language?: string }
    return { ...UNSUPPORTED, language: language ?? 'zh-CN' }
  })

  registerHandle('speech:synthesize', async (_event: unknown, args: unknown) => {
    const { voice } = (args || {}) as { text?: string; voice?: string }
    return { ...UNSUPPORTED, voice: voice ?? '' }
  })

  // Voice catalog is static metadata (no audio), so it stays real.
  registerHandle('speech:get-voices', () => [])
}
