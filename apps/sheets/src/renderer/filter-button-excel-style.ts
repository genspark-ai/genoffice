/**
 * Excel draws the AutoFilter header button as a bordered box with a down
 * arrow, and swaps in a funnel beside a smaller arrow once the column carries
 * criteria. Stock Univer paints three shrinking lines (no criteria) or a solid
 * funnel (criteria) on a borderless white tile, which reads as a list icon.
 * The button shape class is not exported, so the scene's addObjects hook
 * patches the first instance's prototype.
 */
import { Scene } from '@univerjs/engine-render'

const VIEWPORT = 16

export interface FilterButtonColors {
  readonly fill: string
  readonly border: string
  readonly glyph: string
}

export interface FilterButtonState {
  readonly hasCriteria: boolean
  readonly hovered: boolean
}

interface ThemeServiceLike {
  getColorFromTheme(key: string): string
}

interface FilterButtonShapeLike {
  _cellWidth: number
  _cellHeight: number
  _hovered: boolean
  _filterParams: { hasCriteria: boolean }
  _themeService: ThemeServiceLike
  _draw(ctx: CanvasRenderingContext2D): void
}

export function filterButtonColors(theme: ThemeServiceLike, hovered: boolean): FilterButtonColors {
  return {
    fill: theme.getColorFromTheme(hovered ? 'gray.50' : 'white'),
    border: theme.getColorFromTheme(hovered ? 'primary.600' : 'gray.300'),
    glyph: theme.getColorFromTheme('gray.900'),
  }
}

function triangle(ctx: CanvasRenderingContext2D, cx: number, top: number, half: number): void {
  ctx.beginPath()
  ctx.moveTo(cx - half, top)
  ctx.lineTo(cx + half, top)
  ctx.lineTo(cx, top + half * 1.15)
  ctx.closePath()
  ctx.fill()
}

function funnel(ctx: CanvasRenderingContext2D): void {
  ctx.beginPath()
  ctx.moveTo(2.5, 3.5)
  ctx.lineTo(9.5, 3.5)
  ctx.lineTo(6.8, 7.5)
  ctx.lineTo(6.8, 11.5)
  ctx.lineTo(5.2, 12.5)
  ctx.lineTo(5.2, 7.5)
  ctx.closePath()
  ctx.fill()
}

export function drawExcelFilterButton(
  ctx: CanvasRenderingContext2D,
  size: number,
  state: FilterButtonState,
  colors: FilterButtonColors,
): void {
  ctx.save()
  ctx.scale(size / VIEWPORT, size / VIEWPORT)
  ctx.beginPath()
  ctx.roundRect(0.5, 0.5, VIEWPORT - 1, VIEWPORT - 1, 1)
  ctx.fillStyle = colors.fill
  ctx.fill()
  ctx.lineWidth = 1
  ctx.strokeStyle = colors.border
  ctx.stroke()
  ctx.fillStyle = colors.glyph
  if (state.hasCriteria) {
    funnel(ctx)
    triangle(ctx, 11.5, 9, 2.5)
  } else {
    triangle(ctx, 8, 6, 3.5)
  }
  ctx.restore()
}

function isFilterButtonShape(object: unknown): object is FilterButtonShapeLike {
  if (typeof object !== 'object' || object === null) return false
  const shape = object as Partial<FilterButtonShapeLike>
  return (
    typeof shape._draw === 'function' &&
    typeof shape._filterParams === 'object' &&
    shape._filterParams !== null &&
    'hasCriteria' in shape._filterParams &&
    typeof shape._themeService?.getColorFromTheme === 'function'
  )
}

const PATCHED = Symbol('excel-filter-button')

/// Replaces `_draw` on the prototype shared by every filter button; returns
/// true once a prototype was patched.
export function patchFilterButtonShapes(objects: readonly unknown[]): boolean {
  for (const object of objects) {
    if (!isFilterButtonShape(object)) continue
    const proto = Object.getPrototypeOf(object) as FilterButtonShapeLike & {
      [PATCHED]?: true
    }
    if (proto[PATCHED]) return true
    proto[PATCHED] = true
    proto._draw = function excelFilterButtonDraw(
      this: FilterButtonShapeLike,
      ctx: CanvasRenderingContext2D,
    ) {
      ctx.save()
      // Same clip as stock: the tile sits in the cell's bottom-right corner
      // and must not spill into a neighbor when the cell is narrower.
      const region = new Path2D()
      region.rect(
        VIEWPORT - this._cellWidth,
        VIEWPORT - this._cellHeight,
        this._cellWidth,
        this._cellHeight,
      )
      ctx.clip(region)
      drawExcelFilterButton(
        ctx,
        VIEWPORT,
        { hasCriteria: this._filterParams.hasCriteria, hovered: this._hovered },
        filterButtonColors(this._themeService, this._hovered),
      )
      ctx.restore()
    }
    return true
  }
  return false
}

let installed = false

export function installExcelFilterButton(): void {
  if (installed) return
  installed = true
  const proto = Scene.prototype as unknown as {
    addObjects(objects: unknown[], zIndex?: number): Scene
  }
  const original = proto.addObjects
  let patched = false
  proto.addObjects = function patchedAddObjects(this: Scene, objects, zIndex) {
    if (!patched) patched = patchFilterButtonShapes(objects)
    return original.call(this, objects, zIndex)
  }
}
