/**
 * A slide entered during the recording, at `atMs` of recorded time (pauses excluded); a negative
 * index is the show's end screen, which belongs to no slide.
 */
export interface NarrationMark {
  slideIndex: number
  atMs: number
}

/** A recorded slide: the parts of the recording spoken on it, in order. */
export interface SlideNarration {
  slideIndex: number
  spans: Array<{ startMs: number; endMs: number }>
  durationMs: number
}

/**
 * Cuts one continuous recording into the slides it was spoken on. A slide visited twice gets
 * both parts, in order, as rehearsed timings add up revisits.
 */
export function slideNarrations(marks: NarrationMark[], endMs: number): SlideNarration[] {
  const bySlide = new Map<number, SlideNarration>()
  marks.forEach((mark, i) => {
    const startMs = mark.atMs
    const stopMs = Math.min(endMs, marks[i + 1]?.atMs ?? endMs)
    if (mark.slideIndex < 0 || stopMs <= startMs) return
    const slide = bySlide.get(mark.slideIndex) ?? {
      slideIndex: mark.slideIndex,
      spans: [],
      durationMs: 0,
    }
    slide.spans.push({ startMs, endMs: stopMs })
    slide.durationMs += stopMs - startMs
    bySlide.set(mark.slideIndex, slide)
  })
  return [...bySlide.values()].sort((a, b) => a.slideIndex - b.slideIndex)
}

export function joinSpans(
  samples: Float32Array,
  sampleRate: number,
  spans: SlideNarration['spans'],
): Float32Array {
  const at = (ms: number) => Math.min(samples.length, Math.round((ms * sampleRate) / 1000))
  const parts = spans.map((s) => samples.subarray(at(s.startMs), at(s.endMs)))
  const joined = new Float32Array(parts.reduce((n, p) => n + p.length, 0))
  let offset = 0
  for (const part of parts) {
    joined.set(part, offset)
    offset += part.length
  }
  return joined
}
