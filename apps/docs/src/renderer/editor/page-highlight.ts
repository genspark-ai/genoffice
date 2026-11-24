interface HighlightRegistry {
  get(name: string): Set<AbstractRange> | undefined
  set(name: string, highlight: Set<AbstractRange>): void
}

/** the named CSS highlight of the window, shared by its editors; null where highlights do not exist */
export function pageHighlight(name: string): Set<AbstractRange> | null {
  const registry = (globalThis.CSS as { highlights?: HighlightRegistry } | undefined)?.highlights
  const Highlight = (globalThis as { Highlight?: new () => Set<AbstractRange> }).Highlight
  if (!registry || !Highlight) return null
  let highlight = registry.get(name)
  if (!highlight) registry.set(name, (highlight = new Highlight()))
  return highlight
}
