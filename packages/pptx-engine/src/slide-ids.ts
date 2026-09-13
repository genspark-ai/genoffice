/** ST_SlideId is 256..2147483647. */
const SLIDE_ID_MIN = 256
const SLIDE_ID_MAX = 2147483647

/** Next p:sldId/@id for a presentation part: max+1, or the lowest free id once a deck sits at the ceiling. */
export function nextSlideId(presXml: string): number {
  const used = new Set<number>()
  for (const m of presXml.matchAll(/<p:sldId\s[^>]*\bid="(\d+)"/g)) used.add(Number(m[1]))
  let id = Math.max(SLIDE_ID_MIN - 1, ...used) + 1
  if (id > SLIDE_ID_MAX) {
    id = SLIDE_ID_MIN
    while (used.has(id)) id++
  }
  return id
}
