/**
 * Paste as Picture (Word's Paste Special ▸ Picture, and the Picture entry of
 * the Paste menu): copied text, a table or a range from a spreadsheet is drawn
 * as it looks and inserted as one PNG picture, so it can no longer be edited
 * or reflowed. The clipboard HTML is laid out off screen with its own styles,
 * then painted through an SVG foreignObject onto a canvas.
 */

/** the widest the drawing gets before the content wraps (Word's text column at 6.5in) */
export const PICTURE_MAX_WIDTH_PX = 624
/** canvas pixels per CSS pixel, so the picture stays sharp when zoomed or printed */
const PICTURE_SCALE = 2
/** Chromium refuses canvases past this many pixels on a side */
const CANVAS_MAX_PX = 16_384

const XHTML_NS = 'http://www.w3.org/1999/xhtml'

/** what a picture of the clipboard carries: the PNG and the size it is drawn at */
export interface ClipboardPicture {
  dataUrl: string
  widthPx: number
  heightPx: number
}

/**
 * The clipboard content as a detached element ready to draw: the HTML's own
 * <style> rules kept (spreadsheets style their cells through classes), active
 * or remote content dropped. A remote image would taint the canvas and make
 * the whole picture unreadable, so only inline data: images stay.
 */
export function pictureContent(html: string, text: string): HTMLElement {
  const frame = document.createElement('div')
  if (html.trim()) {
    const parsed = new DOMParser().parseFromString(html, 'text/html')
    for (const style of parsed.querySelectorAll('style')) frame.append(style.cloneNode(true))
    const body = parsed.body
    for (const el of body.querySelectorAll('script, link, meta, iframe, object, embed, base')) {
      el.remove()
    }
    for (const img of body.querySelectorAll('img')) {
      if (!(img.getAttribute('src') ?? '').startsWith('data:')) img.remove()
    }
    for (const node of [...body.childNodes]) frame.append(document.importNode(node, true))
  } else {
    const pre = document.createElement('div')
    pre.style.whiteSpace = 'pre-wrap'
    pre.textContent = text
    frame.append(pre)
  }
  return frame
}

/** draws the clipboard content as a PNG; null when there is nothing to draw */
export async function clipboardPicture(
  html: string,
  text: string,
): Promise<ClipboardPicture | null> {
  if (!html.trim() && !text.trim()) return null
  const frame = pictureContent(html, text)
  frame.setAttribute('xmlns', XHTML_NS)
  // document content, not chrome: black on white like the page it lands on
  frame.style.cssText +=
    `;display:inline-block;box-sizing:border-box;max-width:${PICTURE_MAX_WIDTH_PX}px;` +
    'padding:2px;background:#ffffff;color:#000000;' +
    'font:11pt Calibri,Carlito,Aptos,Arial,sans-serif;line-height:1.2'
  const host = document.createElement('div')
  host.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none'
  host.append(frame)
  document.body.append(host)
  try {
    const box = frame.getBoundingClientRect()
    const width = Math.max(1, Math.ceil(box.width))
    const height = Math.max(1, Math.ceil(box.height))
    const scale = Math.min(PICTURE_SCALE, CANVAS_MAX_PX / width, CANVAS_MAX_PX / height)
    const xhtml = new XMLSerializer().serializeToString(frame)
    const svg =
      `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">` +
      `<foreignObject x="0" y="0" width="${width}" height="${height}">${xhtml}</foreignObject></svg>`
    const image = new Image()
    image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
    await image.decode()
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(width * scale))
    canvas.height = Math.max(1, Math.round(height * scale))
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.scale(scale, scale)
    ctx.drawImage(image, 0, 0, width, height)
    return { dataUrl: canvas.toDataURL('image/png'), widthPx: width, heightPx: height }
  } catch {
    return null
  } finally {
    host.remove()
  }
}
