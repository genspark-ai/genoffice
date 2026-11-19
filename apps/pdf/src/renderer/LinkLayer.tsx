import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { pdfRectToCss } from './annotations'
import type { PageGeom } from './annotations'

interface LinkItem {
  rect: [number, number, number, number]
  url?: string
  dest?: unknown
}

interface RawLinkAnnotation {
  subtype?: string
  rect?: number[]
  url?: string
  dest?: unknown
  action?: string
}

/** Max link overlays per page: a hostile PDF can carry thousands of annots. */
export const MAX_PAGE_LINKS = 500

function isFiniteRect(rect: unknown): rect is [number, number, number, number] {
  return (
    Array.isArray(rect) &&
    rect.length === 4 &&
    rect.every((n) => typeof n === 'number' && Number.isFinite(n))
  )
}

/**
 * Filter raw pdfjs annotations down to renderable links: Link subtype with a
 * finite rect and a target, capped per page. Exported for tests.
 */
export function collectPageLinks(annots: RawLinkAnnotation[]): LinkItem[] {
  const out: LinkItem[] = []
  for (const a of annots) {
    if (out.length >= MAX_PAGE_LINKS) break
    if (a.subtype !== 'Link' || (!a.url && !a.dest)) continue
    if (!isFiniteRect(a.rect)) continue
    out.push({ rect: a.rect, url: a.url, dest: a.dest })
  }
  return out
}

/** Link annot hit areas: external links open a new window (main process routes to shell.openExternal); internal dests jump pages */
export function LinkLayer({
  doc,
  pageNo,
  geom,
  scale,
  onGoToDest,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  geom: PageGeom
  scale: number
  onGoToDest: (dest: unknown) => void
}): ReactElement | null {
  const [links, setLinks] = useState<LinkItem[] | null>(null)

  useEffect(() => {
    let cancelled = false
    void (async () => {
      const page = await doc.getPage(pageNo)
      const annots = (await page.getAnnotations()) as RawLinkAnnotation[]
      if (cancelled) return
      setLinks(collectPageLinks(annots))
    })()
    return () => {
      cancelled = true
    }
  }, [doc, pageNo])

  if (!links || links.length === 0) return null

  return (
    <div className="pdf-link-layer">
      {links.map((l, i) => (
        <a
          key={i}
          className="pdf-link"
          style={pdfRectToCss(geom, l.rect, scale)}
          href={l.url ?? '#'}
          data-tip={l.url}
          target={l.url ? '_blank' : undefined}
          rel="noreferrer"
          onClick={(e) => {
            if (!l.url) {
              e.preventDefault()
              onGoToDest(l.dest)
            }
          }}
        />
      ))}
    </div>
  )
}
