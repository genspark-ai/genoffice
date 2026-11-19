import { Context } from 'konva/lib/Context'

/**
 * Chromium's default canvas `fontKerning: auto` kerns some fonts only partially
 * (Constantia measured 1.4% wider than its GPOS pairs say; Times New Roman and Arial
 * were unaffected) while PowerPoint applies the full pair table. Every Konva text
 * draw sets `font`, so hooking that attribute pins each context to full kerning.
 */
export function installCanvasKerning(): void {
  const proto = Context.prototype as unknown as {
    setAttr(attr: string, val: unknown): void
    _context: CanvasRenderingContext2D
  }
  const base = proto.setAttr
  proto.setAttr = function (this: typeof proto, attr: string, val: unknown) {
    base.call(this, attr, val)
    if (attr === 'font' && this._context && 'fontKerning' in this._context)
      this._context.fontKerning = 'normal'
  }
}
