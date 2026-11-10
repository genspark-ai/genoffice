import type { SaveRecordingOp } from '../../shared/ipc'
import { bytesToBase64 } from '../insert-actions'
import { joinSpans, slideNarrations, type NarrationMark } from './segments'
import { encodeWav } from './wav'

/** Speech needs no more: 16-bit mono WAV at this rate weighs about 2.9 MB per minute. */
const NARRATION_SAMPLE_RATE = 24000

function mono(audio: AudioBuffer): Float32Array {
  if (audio.numberOfChannels === 1) return audio.getChannelData(0)
  const mixed = new Float32Array(audio.length)
  for (let c = 0; c < audio.numberOfChannels; c++) {
    const channel = audio.getChannelData(c)
    for (let i = 0; i < mixed.length; i++) mixed[i]! += channel[i]! / audio.numberOfChannels
  }
  return mixed
}

/** One recording cut into the slides it was spoken on, each as a WAV. */
export async function recordedSlides(
  audio: Blob,
  marks: NarrationMark[],
  endMs: number,
): Promise<SaveRecordingOp['slides']> {
  const context = new AudioContext({ sampleRate: NARRATION_SAMPLE_RATE })
  try {
    const decoded = await context.decodeAudioData(await audio.arrayBuffer())
    const samples = mono(decoded)
    return slideNarrations(marks, endMs).map((slide) => ({
      slideIndex: slide.slideIndex,
      wavBase64: bytesToBase64(
        encodeWav(joinSpans(samples, decoded.sampleRate, slide.spans), decoded.sampleRate),
      ),
      durationMs: Math.round(slide.durationMs),
      advanceMs: Math.round(slide.durationMs),
    }))
  } finally {
    void context.close()
  }
}
