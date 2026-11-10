// a gradient-painted card has a transparent background-color: the swatch and the pick must see the image layer

const hasUrl = (image: string) => /\burl\(/i.test(image)
const isGradient = (image: string) => /\bgradient\(/i.test(image) && !hasUrl(image)

export function backgroundSwatch(color: string, image: string): string {
  return color || (isGradient(image) ? image : '')
}

/** declarations a colour pick writes: a gradient is replaced by the colour, a picture stays on top */
export function backgroundPickStyles(hex: string | null, image: string): Record<string, string> {
  const color = hex ?? 'transparent'
  return isGradient(image)
    ? { 'background-image': 'none', 'background-color': color }
    : { 'background-color': color }
}
