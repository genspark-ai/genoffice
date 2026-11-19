/**
 * Reusable Range: Blink fixes up every live Range on each DOM removal, so
 * per-word probe Ranges make later ProseMirror re-renders crawl. Do not nest.
 */
export function rangeSlot(): () => Range {
  let range: Range | undefined
  return () => (range ??= document.createRange())
}
