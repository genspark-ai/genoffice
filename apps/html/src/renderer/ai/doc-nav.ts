/**
 * In-answer citations: the model links passages as [label](htmlnav://sid/N);
 * clicking one selects that element in the source pane (and, later, the preview).
 */
export const DOC_NAV_SCHEME = 'htmlnav://'

/** htmlnav://sid/N -> N; null for anything else */
export function parseDocNavHref(href: string): number | null {
  const m = /^htmlnav:\/\/sid\/(\d+)$/.exec(href)
  return m ? Number(m[1]) : null
}
