
export interface PageGeometry {
  width: number
  height: number
  rotate: number
  origin: [number, number]
  userUnit: number
}

export interface PageProbeFailure {
  /** 1-based page number */
  page: number
  message: string
}

export interface PageProbeResult {
  pages: PageGeometry[]
  failures: PageProbeFailure[]
}

const FALLBACK: PageGeometry = { width: 595, height: 842, rotate: 0, origin: [0, 0], userUnit: 1 }

/**
 * Probes every page's geometry (size, rotation, origin, userUnit) in document
 * order. A page whose probe throws degrades to the previous good geometry (A4
 * before the first success) and is reported in `failures` instead of failing
 * the whole document — one broken page deep in a 450-page textbook must not
 * blank the other 449 (genoffice#2366: the open path used to destroy the
 * document over a single getPage rejection, leaving a white viewer and the
 * error only in the console).
 */
export async function probePageGeometry(doc: {
  numPages: number
  getPage: (n: number) => Promise<{
    getViewport: (o: { scale: number; rotation: number }) => { width: number; height: number }
    rotate?: number | null
    view: number[]
    userUnit?: number | null
  }>
}): Promise<PageProbeResult> {
  const pages: PageGeometry[] = []
  const failures: PageProbeFailure[] = []
  let lastGood = FALLBACK
  for (let i = 1; i <= doc.numPages; i++) {
    try {
      const page = await doc.getPage(i)
      // Unrotated size; display size is derived by geom from the total rotation
      const vp = page.getViewport({ scale: 1, rotation: 0 })
      const geo: PageGeometry = {
        width: vp.width,
        height: vp.height,
        rotate: page.rotate ?? 0,
        origin: [page.view[0] ?? 0, page.view[1] ?? 0],
        userUnit: page.userUnit ?? 1,
      }
      pages.push(geo)
      lastGood = geo
    } catch (err) {
      failures.push({ page: i, message: err instanceof Error ? err.message : String(err) })
      pages.push(lastGood)
    }
  }
  return { pages, failures }
}
