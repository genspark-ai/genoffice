/**
 * Inline content walker shared by both import filters. A paragraph's children
 * (`#text`, `text:span`, `text:a`, `text:tab`, `text:line-break`, …) become a
 * flat run list in a neutral shape each converter maps onto its engine's run
 * type. The rule from the .ods work carries over: a construct that cannot be
 * represented is counted and dropped, never approximated into wrong markup.
 */

import type { OdfStyles, OdfTextProps } from './styles'
import { odfLengthToEmu } from './units'
import { XNode, attr, isTextNode, nodeChildren, nodeName, textOfNode } from './xml'

/** A run before engine mapping — every field optional except the text itself. */
export interface NeutralRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  sizePt?: number
  /** hex without '#' */
  color?: string
  font?: string
  vertAlign?: 'superscript' | 'subscript'
  caps?: 'all' | 'small'
  /** hex without '#' */
  shading?: string
  /** external hyperlink target (http/https/mailto/ftp only) */
  href?: string
}

/** an inline image reference collected while walking; the converter loads the bytes */
export interface InlineImageRef {
  /** zip path from xlink:href (may carry a leading './') or base64 payload */
  href?: string
  base64?: string
  widthEmu?: number
  heightEmu?: number
}

/** dropped-construct tally, rendered into the result's `dropped` summary */
export class Drops {
  private counts = new Map<string, number>()

  add(what: string): void {
    this.counts.set(what, (this.counts.get(what) ?? 0) + 1)
  }

  list(): Array<{ what: string; count: number }> {
    return [...this.counts.entries()]
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
      .map(([what, count]) => ({ what, count }))
  }
}

export interface WalkContext {
  styles: OdfStyles
  /** namespace-qualified style families to try when resolving a style name */
  families: readonly string[]
  drops: Drops
  images: InlineImageRef[]
}

const LINKABLE_HREF = /^(https?|mailto|ftp):/i

export function runsFromParagraph(node: XNode, ctx: WalkContext): NeutralRun[] {
  const styleName = attr(node, 'text:style-name')
  // a paragraph style's own text-properties dress its direct text (the parts
  // not wrapped in a span), so the paragraph family is tried first
  const inherited = ctx.styles.textProps(styleName, ['paragraph', ...ctx.families]) ?? undefined
  const out = runsFromChildren(nodeChildren(node), inherited, ctx)
  return mergeAdjacent(out)
}

function runsFromChildren(
  children: XNode[],
  inherited: OdfTextProps | undefined,
  ctx: WalkContext,
): NeutralRun[] {
  const out: NeutralRun[] = []
  const push = (run: NeutralRun): void => {
    if (run.text.length === 0) return
    out.push(run)
  }
  const walk = (nodes: XNode[], props: OdfTextProps | undefined, href?: string): void => {
    for (const node of nodes) {
      if (isTextNode(node)) {
        const text = textOfNode(node)
        if (text) push({ text, ...base(props), ...(href ? { href } : {}) })
        continue
      }
      const name = nodeName(node)
      switch (name) {
        case '#comment':
          break
        case 'text:span': {
          const spanProps = mergeProps(
            props,
            ctx.styles.textProps(attr(node, 'text:style-name'), ctx.families),
          )
          walk(nodeChildren(node), spanProps, href)
          break
        }
        case 'text:a': {
          const target = attr(node, 'xlink:href') ?? ''
          walk(nodeChildren(node), props, LINKABLE_HREF.test(target) ? target : href)
          break
        }
        case 'text:tab':
          push({ text: '\t', ...base(props), ...(href ? { href } : {}) })
          break
        case 'text:line-break':
          push({ text: '\n', ...base(props), ...(href ? { href } : {}) })
          break
        case 'text:s': {
          const count = Math.min(Number(attr(node, 'text:c') ?? '1') || 1, 1000)
          push({ text: ' '.repeat(count), ...base(props), ...(href ? { href } : {}) })
          break
        }
        case 'text:soft-page-break':
        case 'text:bookmark':
        case 'text:bookmark-start':
        case 'text:bookmark-end':
        case 'text:reference-mark':
        case 'text:reference-mark-start':
        case 'text:reference-mark-end':
        case 'text:sequence':
          break
        case 'text:note':
          ctx.drops.add('footnotes')
          break
        case 'office:annotation':
          ctx.drops.add('comments')
          break
        case 'draw:frame': {
          const image = imageRefOfFrame(node)
          if (image) {
            ctx.images.push(image)
            break
          }
          if (firstDescendant(node, 'draw:text-box')) {
            // an inline text frame flattens into the paragraph's flow
            ctx.drops.add('textboxes-flattened')
            walk(nodeChildren(node), props, href)
            break
          }
          ctx.drops.add('shapes')
          break
        }
        case 'draw:image':
        case 'draw:object':
        case 'math:math':
          ctx.drops.add('objects')
          break
        case 'text:page-number':
        case 'text:page-count':
        case 'text:sequence-ref':
        case 'text:bookmark-ref':
          // fields arrive with their cached value as text content
          push({ text: textOfNode(node), ...base(props), ...(href ? { href } : {}) })
          break
        default:
          // unknown container: recurse so nested spans/text still survive
          walk(nodeChildren(node), props, href)
          break
      }
    }
  }
  walk(children, inherited)
  return out
}

/** the formatting fields of a run, without its text */
function base(props: OdfTextProps | undefined): Omit<NeutralRun, 'text'> {
  if (!props) return {}
  const out: Omit<NeutralRun, 'text'> = {}
  if (props.bold) out.bold = true
  if (props.italic) out.italic = true
  if (props.underline) out.underline = true
  if (props.strike) out.strike = true
  if (props.sizePt !== undefined) out.sizePt = props.sizePt
  if (props.color) out.color = props.color
  if (props.font) out.font = props.font
  if (props.vertAlign) out.vertAlign = props.vertAlign
  if (props.caps) out.caps = props.caps
  if (props.shading) out.shading = props.shading
  return out
}

function mergeProps(
  baseProps: OdfTextProps | undefined,
  override: OdfTextProps | null,
): OdfTextProps | undefined {
  if (!override) return baseProps
  return { ...(baseProps ?? {}), ...override }
}

/** collapse runs whose formatting is identical — ODF spans split more eagerly than a reader wants */
function mergeAdjacent(runs: NeutralRun[]): NeutralRun[] {
  const out: NeutralRun[] = []
  for (const run of runs) {
    const previous = out[out.length - 1]
    if (previous && sameFormat(previous, run)) {
      previous.text += run.text
    } else {
      out.push({ ...run })
    }
  }
  return out
}

const FORMAT_KEYS: Array<keyof NeutralRun> = [
  'bold',
  'italic',
  'underline',
  'strike',
  'sizePt',
  'color',
  'font',
  'vertAlign',
  'caps',
  'shading',
  'href',
]

function sameFormat(a: NeutralRun, b: NeutralRun): boolean {
  return FORMAT_KEYS.every((key) => a[key] === b[key])
}

/** an image inside a frame: the href (zip path) or inline base64, plus the frame's box */
export function imageRefOfFrame(frame: XNode): InlineImageRef | null {
  const image = firstDescendant(frame, 'draw:image')
  if (!image) return null
  const ref: InlineImageRef = {
    href: attr(image, 'xlink:href'),
    widthEmu: lengthOf(attr(frame, 'svg:width')),
    heightEmu: lengthOf(attr(frame, 'svg:height')),
  }
  for (const child of nodeChildren(image)) {
    if (nodeName(child) === 'office:binary-data') ref.base64 = textOfNode(child).replace(/\s+/g, '')
  }
  if (!ref.href && !ref.base64) return null
  return ref
}

export function firstDescendant(node: XNode, name: string): XNode | null {
  for (const child of nodeChildren(node)) {
    if (nodeName(child) === name) return child
    const nested = firstDescendant(child, name)
    if (nested) return nested
  }
  return null
}

function lengthOf(text: string | undefined): number | undefined {
  const value = odfLengthToEmu(text)
  return value === null ? undefined : value
}
