/**
 * Pictures inserted from a vector are shown, copied and cropped as the bitmap
 * drawn from it; the save writes the SVG next to that bitmap for as long as
 * the bitmap is the one drawn here, so anything that changes the bytes (crop,
 * Change Picture) leaves the vector behind. The name the vector came with is
 * kept the same way, for a picture in a table cell, which has no attribute to
 * hold alternative text.
 */
const vectors = new Map<string, { svgBase64: string; altText?: string }>()

export function rememberPictureVector(
  bitmapBase64: string,
  svgBase64: string,
  altText?: string,
): void {
  vectors.set(bitmapBase64, { svgBase64, altText })
}

export function pictureVectorOf(bitmapBase64: string): string | undefined {
  return vectors.get(bitmapBase64)?.svgBase64
}

export function pictureVectorAltOf(bitmapBase64: string): string | undefined {
  return vectors.get(bitmapBase64)?.altText
}

/**
 * The SVG a picture is drawn from on the page, and the bitmap it goes with:
 * once the bitmap is another one (crop, Change Picture) the SVG no longer is
 * that picture and the bitmap is drawn.
 */
export interface PictureVector {
  for: string
  url: string
}

export function pictureVectorAttr(
  bitmapUrl: string | null | undefined,
  vectorUrl: string | null | undefined,
): PictureVector | null {
  return bitmapUrl && vectorUrl ? { for: bitmapUrl, url: vectorUrl } : null
}

/** what the page draws for a picture: its SVG while that still is the picture, else the bitmap */
export function pictureDisplaySrc(bitmapUrl: unknown, vector: unknown): string {
  const v = vector as PictureVector | null | undefined
  return v && v.url && v.for === bitmapUrl ? v.url : String(bitmapUrl)
}
