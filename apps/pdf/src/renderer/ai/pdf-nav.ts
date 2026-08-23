/**
 * In-answer page citations: the model links pages as [p.N](pdfnav://page/N);
 * clicking one scrolls the reading view to that page.
 */

export const PDF_NAV_SCHEME = 'pdfnav://'

/** pdfnav://page/N -> N (original 1-based page number); null for anything else */
export function parsePdfNavHref(href: string): number | null {
  const m = /^pdfnav:\/\/page\/(\d+)$/.exec(href)
  return m ? Number(m[1]) : null
}
