/** the fallback bitmap holds this many pixels per pixel of the size the picture is shown at */
export const SVG_FALLBACK_SCALE = 4

const utf8Base64 = (text: string): string => {
  let binary = ''
  for (const byte of new TextEncoder().encode(text)) binary += String.fromCharCode(byte)
  return btoa(binary)
}

export const svgBase64Of = (svg: string): string => utf8Base64(svg)
export const svgDataUrl = (svg: string): string => `data:image/svg+xml;base64,${utf8Base64(svg)}`

/** the root tag given the size the bitmap has, so no browser draws the SVG small and stretches it */
function sizedSvg(svg: string, widthPx: number, heightPx: number): string {
  return svg.replace(/<svg\b[^>]*>/, (root) => {
    const natural = {
      width: /\swidth="([\d.]+)(?:px)?"/.exec(root)?.[1],
      height: /\sheight="([\d.]+)(?:px)?"/.exec(root)?.[1],
    }
    let tag = root.replace(/\s(?:width|height)="[^"]*"/g, '')
    // without a viewBox the drawing would keep its own size inside the larger bitmap
    if (!/\sviewBox="/.test(tag) && natural.width && natural.height)
      tag = tag.replace(/^<svg\b/, `<svg viewBox="0 0 ${natural.width} ${natural.height}"`)
    return tag.replace(/^<svg\b/, `<svg width="${widthPx}" height="${heightPx}"`)
  })
}

/** a PNG data URL of the SVG, SVG_FALLBACK_SCALE times the size it is shown at */
export async function rasteriseSvg(svg: string, shownW: number, shownH: number): Promise<string> {
  const width = Math.max(1, Math.round(shownW * SVG_FALLBACK_SCALE))
  const height = Math.max(1, Math.round(shownH * SVG_FALLBACK_SCALE))
  const image = new Image()
  image.src = svgDataUrl(sizedSvg(svg, width, height))
  await image.decode()
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('no 2d context')
  context.drawImage(image, 0, 0, width, height)
  return canvas.toDataURL('image/png')
}
