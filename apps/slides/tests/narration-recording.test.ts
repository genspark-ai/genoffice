/** Slide Show > Record: cutting one recording into slides, and the WAV each slide gets. */
import { describe, expect, it } from 'vitest'
import { joinSpans, slideNarrations } from '../src/renderer/narration/segments'
import { encodeWav } from '../src/renderer/narration/wav'

describe('slideNarrations', () => {
  it('gives each slide the time spoken on it, revisits added in order', () => {
    const marks = [
      { slideIndex: 0, atMs: 0 },
      { slideIndex: 1, atMs: 4000 },
      { slideIndex: 0, atMs: 6500 },
      { slideIndex: 2, atMs: 7000 },
    ]
    expect(slideNarrations(marks, 9000)).toEqual([
      {
        slideIndex: 0,
        spans: [
          { startMs: 0, endMs: 4000 },
          { startMs: 6500, endMs: 7000 },
        ],
        durationMs: 4500,
      },
      { slideIndex: 1, spans: [{ startMs: 4000, endMs: 6500 }], durationMs: 2500 },
      { slideIndex: 2, spans: [{ startMs: 7000, endMs: 9000 }], durationMs: 2000 },
    ])
  })

  it('drops a slide passed through without recorded time', () => {
    const marks = [
      { slideIndex: 0, atMs: 0 },
      { slideIndex: 1, atMs: 1000 },
      { slideIndex: 2, atMs: 1000 },
    ]
    expect(slideNarrations(marks, 1500).map((s) => s.slideIndex)).toEqual([0, 2])
  })
})

describe('slideNarrations at the end screen', () => {
  it('gives the end screen time to no slide', () => {
    const marks = [
      { slideIndex: 0, atMs: 0 },
      { slideIndex: -1, atMs: 3000 },
    ]
    expect(slideNarrations(marks, 8000)).toEqual([
      { slideIndex: 0, spans: [{ startMs: 0, endMs: 3000 }], durationMs: 3000 },
    ])
  })
})

describe('joinSpans', () => {
  it('joins the samples of each span end to end', () => {
    const samples = Float32Array.from({ length: 10 }, (_, i) => i)
    const joined = joinSpans(samples, 1000, [
      { startMs: 1, endMs: 3 },
      { startMs: 7, endMs: 20 },
    ])
    expect([...joined]).toEqual([1, 2, 7, 8, 9])
  })
})

describe('encodeWav', () => {
  it('writes a 16-bit mono PCM WAV', () => {
    const wav = encodeWav(Float32Array.from([0, 1, -1, 2]), 24000)
    const view = new DataView(wav.buffer)
    const text = (at: number, n: number) => String.fromCharCode(...wav.subarray(at, at + n))
    expect(text(0, 4)).toBe('RIFF')
    expect(text(8, 8)).toBe('WAVEfmt ')
    expect(view.getUint16(20, true)).toBe(1)
    expect(view.getUint16(22, true)).toBe(1)
    expect(view.getUint32(24, true)).toBe(24000)
    expect(view.getUint16(34, true)).toBe(16)
    expect(view.getUint32(40, true)).toBe(8)
    expect([0, 1, 2, 3].map((i) => view.getInt16(44 + i * 2, true))).toEqual([
      0, 32767, -32768, 32767,
    ])
  })
})
