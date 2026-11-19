/** ST_SlideId is 256..2147483647. */
const SLIDE_ID_MIN = 256
const SLIDE_ID_MAX = 2147483647

/** Next p:sldId/@id for a presentation part: max+1, or the lowest free id once a deck sits at the ceiling. */
export function nextSlideId(presXml: string): number {
  const used = new Set<number>()
  for (const m of presXml.matchAll(/<p:sldId\s[^>]*\bid=(?:"(\d+)"|'(\d+)')/g)) {
    // Out-of-range values are invalid per ST_SlideId; ignoring them keeps a
    // single hostile attribute from forcing the ceiling scan below.
    const n = Number(m[1] ?? m[2])
    if (Number.isSafeInteger(n) && n >= SLIDE_ID_MIN && n <= SLIDE_ID_MAX) used.add(n)
  }
  // Loop-based max: spreading `used` into Math.max overflows the call stack
  // on decks with hundreds of thousands of slides.
  let max = SLIDE_ID_MIN - 1
  for (const n of used) if (n > max) max = n
  let id = max + 1
  if (id > SLIDE_ID_MAX) {
    id = SLIDE_ID_MIN
    while (used.has(id)) {
      id++
      if (id > SLIDE_ID_MAX) throw new Error('pptx: slide id space exhausted')
    }
  }
  return id
}
