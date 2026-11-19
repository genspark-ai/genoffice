/** PowerPoint show actions: <a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=<name>"/>. */
const SHOWJUMP_ACTION = 'ppaction://hlinkshowjump?jump='

export const NAMED_ACTIONS = [
  'nextslide',
  'previousslide',
  'firstslide',
  'lastslide',
  'lastslideviewed',
  'endshow',
] as const
export type NamedAction = (typeof NAMED_ACTIONS)[number]

/** The named action of an hlinkClick `action` attribute value, or null for other actions. */
export function namedActionOf(action: string | undefined | null): NamedAction | null {
  if (!action || !action.startsWith(SHOWJUMP_ACTION)) return null
  const name = action.slice(SHOWJUMP_ACTION.length)
  return (NAMED_ACTIONS as readonly string[]).includes(name) ? (name as NamedAction) : null
}

export function namedActionAttr(action: string): string {
  return SHOWJUMP_ACTION + action
}
