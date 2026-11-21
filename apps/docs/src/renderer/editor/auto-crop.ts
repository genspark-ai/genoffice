/** the part of a picture to keep, in source pixels */
export interface CropBox {
  x: number
  y: number
  w: number
  h: number
}

const TOLERANCE = 12

/**
 * The box left after trimming the uniform border off a picture: rows and
 * columns that are all transparent, or all the colour of the top-left corner,
 * go. Null when nothing can be trimmed or nothing would be left.
 */
export function autoCropBox(
  data: Uint8ClampedArray,
  width: number,
  height: number,
): CropBox | null {
  if (width <= 0 || height <= 0 || data.length < width * height * 4) return null
  const [r0, g0, b0, a0] = [data[0], data[1], data[2], data[3]]
  const isBorder = (i: number): boolean => {
    const a = data[i + 3]
    if (a <= TOLERANCE) return true
    if (a0 <= TOLERANCE) return false
    return (
      Math.abs(data[i] - r0) <= TOLERANCE &&
      Math.abs(data[i + 1] - g0) <= TOLERANCE &&
      Math.abs(data[i + 2] - b0) <= TOLERANCE &&
      Math.abs(a - a0) <= TOLERANCE
    )
  }
  const rowIsBorder = (y: number): boolean => {
    for (let x = 0; x < width; x++) if (!isBorder((y * width + x) * 4)) return false
    return true
  }
  let top = 0
  while (top < height && rowIsBorder(top)) top++
  if (top === height) return null
  let bottom = height - 1
  while (bottom > top && rowIsBorder(bottom)) bottom--
  const colIsBorder = (x: number): boolean => {
    for (let y = top; y <= bottom; y++) if (!isBorder((y * width + x) * 4)) return false
    return true
  }
  let left = 0
  while (left < width && colIsBorder(left)) left++
  let right = width - 1
  while (right > left && colIsBorder(right)) right--
  const box = { x: left, y: top, w: right - left + 1, h: bottom - top + 1 }
  return box.w === width && box.h === height ? null : box
}

/** The picture with its uniform border trimmed off, as a PNG data URL; null when there is none. */
export async function autoCropDataUrl(
  dataUrl: string,
): Promise<{ dataUrl: string; box: CropBox; width: number; height: number } | null> {
  const img = new Image()
  await new Promise<void>((resolve, reject) => {
    img.onload = () => resolve()
    img.onerror = () => reject(new Error('Failed to read image'))
    img.src = dataUrl
  })
  const width = img.naturalWidth
  const height = img.naturalHeight
  if (!width || !height) return null
  const source = document.createElement('canvas')
  source.width = width
  source.height = height
  const sctx = source.getContext('2d')
  if (!sctx) return null
  sctx.drawImage(img, 0, 0)
  const box = autoCropBox(sctx.getImageData(0, 0, width, height).data, width, height)
  if (!box) return null
  const out = document.createElement('canvas')
  out.width = box.w
  out.height = box.h
  out.getContext('2d')?.drawImage(source, box.x, box.y, box.w, box.h, 0, 0, box.w, box.h)
  return { dataUrl: out.toDataURL('image/png'), box, width, height }
}
