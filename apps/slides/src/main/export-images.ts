import { join } from 'node:path'

/**
 * File paths for a page-image export, in page order: <baseName>-01.<ext> …
 * The zero-padding width follows the total page count (3 digits for ≥100 pages)
 * so file browsers sort the pages naturally.
 */
export function buildExportImagePaths(
  dir: string,
  baseName: string,
  count: number,
  ext: 'png' | 'jpg',
): string[] {
  const pad = count >= 100 ? 3 : 2
  return Array.from({ length: Math.max(0, count) }, (_, i) =>
    join(dir, `${baseName}-${String(i + 1).padStart(pad, '0')}.${ext}`),
  )
}
