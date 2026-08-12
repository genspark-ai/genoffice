/**
 * Core image cutout (background removal) algorithm — pure functions, no DOM dependency, easy to unit test.
 * Shares its source with apps/docs and apps/slides src/renderer/cutout.ts (keep the algorithms identical; sync changes to all three).
 *
 * Approach A: edge flood fill with color tolerance (magic-wand style removal, a simplified take on PowerPoint's "Remove Background"):
 *  1. Sample a ring along the image border (corners included); greedy clustering yields ≤4 background representative colors.
 *  2. With "color distance to any representative ≤ tolerance threshold" as the passable condition, 4-connected flood fill from the edges.
 *  3. Flooded pixels get alpha 0. Only background connected to the edge is removed;
 *     enclosed regions inside shapes matching the background color (e.g. white highlights on a white shirt) are unaffected.
 *
 * TODO (approach B): brush-marked "keep/remove" regions + region growing for interactive edge refinement.
 */

/** Pixel map isomorphic to canvas ImageData (RGBA row-major), easy to construct directly in Node unit tests. */
export interface PixelImage {
  data: Uint8ClampedArray
  width: number
  height: number
}

export type RGB = readonly [number, number, number]

export interface CutoutResult {
  /** Processed RGBA pixels (a new array, input unchanged; can be fed straight to new ImageData) */
  data: Uint8ClampedArray<ArrayBuffer>
  /** Number of pixels made transparent (only counts pixels that weren't already fully transparent) */
  removedCount: number
}

/**
 * Tolerance 0..100 → color-distance threshold.
 * Distance is per-channel RMS (0..255 scale); at tolerance 100 the threshold 255 ≈ black-white distance,
 * i.e. at max nearly every color connected to the edge is judged background.
 */
export function toleranceToThreshold(tolerance: number): number {
  const t = Math.max(0, Math.min(100, tolerance))
  return (t * 255) / 100
}

/** Color distance: per-channel RMS, range 0..255 (black vs white = 255). */
export function colorDistance(a: RGB, b: RGB): number {
  const dr = a[0] - b[0]
  const dg = a[1] - b[1]
  const db = a[2] - b[2]
  return Math.sqrt((dr * dr + dg * dg + db * db) / 3)
}

/** Cluster-merge distance for sample clustering (0..255 scale) */
const CLUSTER_DIST = 30
/** Clusters below this share are treated as edge noise, not background representatives */
const MIN_CLUSTER_RATIO = 0.05
/** Alpha below this counts as "already transparent": skipped in sampling, passable for flood fill */
const ALPHA_TRANSPARENT = 16

/**
 * Sample background representative colors along the image border (corners included);
 * greedy-cluster and return in descending cluster size.
 * Fully transparent pixels don't participate; tiny clusters (edge noise) are filtered out.
 */
export function sampleBackgroundColors(img: PixelImage, maxColors = 4): RGB[] {
  const { data, width: w, height: h } = img
  if (w <= 0 || h <= 0) return []
  // sample the border ring at a stride, capped at ~512 sample points
  const perimeter = Math.max(1, 2 * (w + h) - 4)
  const step = Math.max(1, Math.floor(perimeter / 512))
  const sampleIdx: number[] = []
  for (let x = 0; x < w; x += step) {
    sampleIdx.push((0 * w + x) * 4)
    sampleIdx.push(((h - 1) * w + x) * 4)
  }
  for (let y = 0; y < h; y += step) {
    sampleIdx.push((y * w + 0) * 4)
    sampleIdx.push((y * w + w - 1) * 4)
  }

  // greedy clustering: merge into the nearest cluster (running mean), otherwise open a new one
  const clusters: Array<{ sr: number; sg: number; sb: number; n: number }> = []
  for (const i of sampleIdx) {
    if ((data[i + 3] ?? 0) < ALPHA_TRANSPARENT) continue
    const c: RGB = [data[i]!, data[i + 1]!, data[i + 2]!]
    let best: (typeof clusters)[number] | null = null
    let bestD = Infinity
    for (const cl of clusters) {
      const d = colorDistance([cl.sr / cl.n, cl.sg / cl.n, cl.sb / cl.n], c)
      if (d < bestD) {
        bestD = d
        best = cl
      }
    }
    if (best && bestD <= CLUSTER_DIST) {
      best.sr += c[0]
      best.sg += c[1]
      best.sb += c[2]
      best.n += 1
    } else {
      clusters.push({ sr: c[0], sg: c[1], sb: c[2], n: 1 })
    }
  }
  const total = clusters.reduce((s, c) => s + c.n, 0)
  if (total === 0) return []
  return clusters
    .filter((c) => c.n >= Math.max(1, total * MIN_CLUSTER_RATIO))
    .sort((a, b) => b.n - a.n)
    .slice(0, maxColors)
    .map((c) => [Math.round(c.sr / c.n), Math.round(c.sg / c.n), Math.round(c.sb / c.n)] as const)
}

/**
 * Background removal: 4-connected flood fill from the edges; pixels close to a background
 * representative (within tolerance) and connected to the edge get alpha 0.
 * Returns a new pixel array; the input is not modified.
 *
 * @param tolerance tolerance 0..100 (see toleranceToThreshold)
 * @param bgColors  background representative colors; defaults to auto-sampling from the border
 */
export function removeBackground(
  img: PixelImage,
  tolerance: number,
  bgColors: RGB[] = sampleBackgroundColors(img),
): CutoutResult {
  const { data: src, width: w, height: h } = img
  const out = new Uint8ClampedArray(src)
  if (w <= 0 || h <= 0 || bgColors.length === 0) return { data: out, removedCount: 0 }

  const thr = toleranceToThreshold(tolerance)
  // compare distances with the sum of squares (restoring the /3 of per-channel RMS), skipping sqrt
  const thr2 = thr * thr * 3

  /** Whether a pixel "looks like background": already transparent, or within tolerance of any background representative */
  const isBgLike = (i: number): boolean => {
    const p = i * 4
    if ((src[p + 3] ?? 0) < ALPHA_TRANSPARENT) return true
    const r = src[p]!
    const g = src[p + 1]!
    const b = src[p + 2]!
    for (const c of bgColors) {
      const dr = r - c[0]
      const dg = g - c[1]
      const db = b - c[2]
      if (dr * dr + dg * dg + db * db <= thr2) return true
    }
    return false
  }

  const visited = new Uint8Array(w * h)
  const stack: number[] = []
  const trySeed = (i: number) => {
    if (!visited[i] && isBgLike(i)) {
      visited[i] = 1
      stack.push(i)
    }
  }
  // seeds: all passable pixels along the border ring
  for (let x = 0; x < w; x++) {
    trySeed(x)
    trySeed((h - 1) * w + x)
  }
  for (let y = 0; y < h; y++) {
    trySeed(y * w)
    trySeed(y * w + w - 1)
  }

  let removedCount = 0
  while (stack.length > 0) {
    const i = stack.pop()!
    const p = i * 4
    if (out[p + 3]! > 0) {
      out[p + 3] = 0
      removedCount++
    }
    const x = i % w
    if (x > 0) trySeed(i - 1)
    if (x < w - 1) trySeed(i + 1)
    if (i >= w) trySeed(i - w)
    if (i < w * (h - 1)) trySeed(i + w)
  }
  return { data: out, removedCount }
}
