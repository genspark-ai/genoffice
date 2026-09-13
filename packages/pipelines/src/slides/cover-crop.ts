/**
 * Center-crop fractions that make an image fill a target frame without
 * distortion (CSS object-fit: cover): the image is scaled uniformly until it
 * covers the frame and the overflow on the long axis is cropped symmetrically.
 * Returns null when no crop is needed (aspect ratios already match within 1%)
 * or when any dimension is unusable.
 */
export function coverCropFractions(
  imageW: number,
  imageH: number,
  frameW: number,
  frameH: number,
): { l: number; t: number; r: number; b: number } | null {
  if (!(imageW > 0) || !(imageH > 0) || !(frameW > 0) || !(frameH > 0)) return null
  const imageAspect = imageW / imageH
  const frameAspect = frameW / frameH
  if (Math.abs(imageAspect - frameAspect) / frameAspect < 0.01) return null
  if (imageAspect > frameAspect) {
    const crop = (1 - frameAspect / imageAspect) / 2
    return { l: crop, t: 0, r: crop, b: 0 }
  }
  const crop = (1 - imageAspect / frameAspect) / 2
  return { l: 0, t: crop, r: 0, b: crop }
}
