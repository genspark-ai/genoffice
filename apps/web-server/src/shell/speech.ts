/**
 * Speech channels — recognize, synthesize, list voices. All return
 * placeholder payloads until a real STT/TTS provider is wired up.
 */
import { registerHandle } from '../common/index.js'

export function registerSpeechHandlers(): void {
  registerHandle('speech:recognize', async (_event: unknown, args: unknown) => {
    const { audioBytes, language } = args as { audioBytes: ArrayBuffer; language?: string }

    return {
      text: '这是模拟的语音识别结果',
      confidence: 0.95,
      language: language || 'zh-CN',
      words: [
        { word: '这是', start: 0, end: 0.5, confidence: 0.95 },
        { word: '模拟的', start: 0.5, end: 1.0, confidence: 0.92 },
        { word: '语音', start: 1.0, end: 1.5, confidence: 0.98 },
        { word: '识别', start: 1.5, end: 2.0, confidence: 0.91 },
        { word: '结果', start: 2.0, end: 2.5, confidence: 0.94 },
      ],
    }
  })

  registerHandle('speech:synthesize', async (_event: unknown, args: unknown) => {
    const { text, voice, speed, pitch } = args as {
      text: string
      voice?: string
      speed?: number
      pitch?: number
    }

    return {
      audioBytes: new ArrayBuffer(0),
      duration: Math.ceil(text.length * 0.3),
      format: 'mp3',
      voice: voice || 'zh-CN-female',
      speed: speed || 1.0,
      pitch: pitch || 1.0,
    }
  })

  registerHandle('speech:get-voices', () => [
    { id: 'zh-CN-female', name: '中文女声', language: 'zh-CN' },
    { id: 'zh-CN-male', name: '中文男声', language: 'zh-CN' },
    { id: 'en-US-female', name: 'English Female', language: 'en-US' },
    { id: 'en-US-male', name: 'English Male', language: 'en-US' },
    { id: 'ja-JP-female', name: '日本語女性', language: 'ja-JP' },
  ])
}
