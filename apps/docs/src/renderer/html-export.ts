/** Standalone HTML export: live editor DOM re-emitted with computed styles inlined, no app CSS needed. */

const SKIP_CLASSES = new Set([
  'ProseMirror-widget',
  'ProseMirror-separator',
  'ProseMirror-gapcursor',
  'doc-del',
  'doc-anchor-origin',
  'doc-anchor-strut',
  'doc-image-anchor-marker',
  'doc-inline-img-anchor',
  'ai-queue-anchor',
  'page-gap',
  'page-gap-cut',
  'page-gap-table-fill',
  'page-cut-overlays',
  'page-break-lead',
  'page-hf',
  'doc-sectbreak-label',
  'doc-protected-sectbreak',
  'doc-move-handle',
  'box-resize-handle',
])

const UNWRAP_CLASSES = new Set([
  'doc-comment',
  'doc-comment-resolved',
  'comment-pending',
  'doc-ins',
  'has-rpr-change',
  'doc-inactive-selection',
  'doc-comment-flash',
])

const KEEP_WIDGET_CLASSES = ['page-float-host']

const KEEP_ATTRS = ['href', 'src', 'alt', 'title', 'colspan', 'rowspan', 'dir', 'lang', 'target']

const INHERITED = [
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'font-variant',
  'font-feature-settings',
  'color',
  'line-height',
  'letter-spacing',
  'word-spacing',
  'text-align',
  'text-align-last',
  'text-indent',
  'text-transform',
  'text-shadow',
  'direction',
  'white-space',
  'word-break',
  'overflow-wrap',
  'line-break',
  'tab-size',
  'writing-mode',
  'list-style-type',
  'list-style-position',
  'border-collapse',
  'border-spacing',
  'visibility',
] as const

const INITIAL: Record<string, string> = {
  'font-weight': '400',
  'font-style': 'normal',
  'font-variant': 'normal',
  'font-feature-settings': 'normal',
  'letter-spacing': 'normal',
  'word-spacing': '0px',
  'text-align': 'start',
  'text-align-last': 'auto',
  'text-indent': '0px',
  'text-transform': 'none',
  'text-shadow': 'none',
  direction: 'ltr',
  'white-space': 'normal',
  'word-break': 'normal',
  'overflow-wrap': 'normal',
  'line-break': 'auto',
  'tab-size': '8',
  'writing-mode': 'horizontal-tb',
  'list-style-type': 'disc',
  'list-style-position': 'outside',
  'border-collapse': 'separate',
  'border-spacing': '0px',
  visibility: 'visible',
}

const BIDI_KEEP = new Set(['bidi-override', 'isolate-override', 'plaintext', 'embed'])

/** Review / content-control badges drawn by the editor, not document text. */
const CHROME_PSEUDO_CONTENT = /attr\(\s*data-(?:ppr-change-label|sdt-alias)\s*\)/

const INTERNAL_FONTS = /,\s*"(?:[A-Za-z ]+ GO|GenOffice [A-Za-z ]+)"/g

const SIDES = ['top', 'right', 'bottom', 'left'] as const

const TAG_DISPLAY: Record<string, string> = {
  div: 'block',
  p: 'block',
  h1: 'block',
  h2: 'block',
  h3: 'block',
  h4: 'block',
  h5: 'block',
  h6: 'block',
  ul: 'block',
  ol: 'block',
  li: 'list-item',
  table: 'table',
  caption: 'table-caption',
  colgroup: 'table-column-group',
  col: 'table-column',
  thead: 'table-header-group',
  tbody: 'table-row-group',
  tfoot: 'table-footer-group',
  tr: 'table-row',
  td: 'table-cell',
  th: 'table-cell',
  span: 'inline',
  a: 'inline',
  img: 'inline',
  br: 'inline',
  strong: 'inline',
  em: 'inline',
  u: 'inline',
  s: 'inline',
  sub: 'inline',
  sup: 'inline',
  code: 'inline',
  ruby: 'ruby',
  rt: 'ruby-text',
  math: 'inline',
  svg: 'inline',
}

const PASSTHROUGH_TAGS = new Set(['svg', 'math'])

const VOID_TAGS = new Set(['br', 'img', 'hr', 'col'])

const ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }
const escHtml = (s: string) => s.replace(/[&<>"]/g, (c) => ESC[c])
const escText = (s: string) => s.replace(/[&<>]/g, (c) => ESC[c])

const isZero = (v: string) => v === '0px' || v === '0' || v === ''
const isTransparent = (v: string) => v === 'rgba(0, 0, 0, 0)' || v === 'transparent' || v === ''

function hasClass(el: Element, set: Set<string> | string[]): boolean {
  for (const c of el.classList) {
    if (Array.isArray(set) ? set.includes(c) : set.has(c)) return true
  }
  return false
}

export function resolveContent(content: string, el: Element): string {
  if (!content || content === 'none' || content === 'normal') return ''
  let out = ''
  const re =
    /"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|attr\(\s*([\w-]+)\s*\)|counter\(\s*([\w-]+)[^)]*\)/g
  let m: RegExpExecArray | null
  while ((m = re.exec(content))) {
    if (m[1] !== undefined || m[2] !== undefined) {
      out += (m[1] ?? m[2]).replace(/\\([0-9a-fA-F]{1,6})\s?/g, (_, hex: string) =>
        String.fromCodePoint(parseInt(hex, 16)),
      )
    } else if (m[3]) {
      out += el.getAttribute(m[3]) ?? ''
    } else if (m[4]) {
      out += String(orderedIndex(el))
    }
  }
  return out
}

function orderedIndex(el: Element): number {
  let n = 1
  let prev = el.previousElementSibling
  while (prev && prev.classList.contains('doc-li-ordered')) {
    n++
    prev = prev.previousElementSibling
  }
  return n
}

function emitTag(el: Element, cs: CSSStyleDeclaration): string {
  const tag = el.tagName.toLowerCase()
  if (tag in TAG_DISPLAY) return tag
  const d = cs.display
  return d.startsWith('inline') || d === 'contents' ? 'span' : 'div'
}

function borderDecl(cs: CSSStyleDeclaration, side: string): string | null {
  const w = cs.getPropertyValue(`border-${side}-width`)
  const st = cs.getPropertyValue(`border-${side}-style`)
  if (isZero(w) || st === 'none' || st === 'hidden') return null
  return `border-${side}:${w} ${st} ${cs.getPropertyValue(`border-${side}-color`)}`
}

interface StyleOpts {
  parent: CSSStyleDeclaration | null
  tag: string
  authoredWidth: boolean
  authoredHeight: boolean
  skipBorders: boolean
  positioned: boolean
}

function styleDecls(cs: CSSStyleDeclaration, o: StyleOpts): string[] {
  const out: string[] = []
  for (const p of INHERITED) {
    let v = cs.getPropertyValue(p)
    if (!v) continue
    if (o.parent ? o.parent.getPropertyValue(p) === v : INITIAL[p] === v) continue
    if (p === 'font-family') v = v.replace(INTERNAL_FONTS, '')
    out.push(`${p}:${v}`)
  }
  const bidi = cs.unicodeBidi
  if (BIDI_KEEP.has(bidi) && o.parent?.unicodeBidi !== bidi) out.push(`unicode-bidi:${bidi}`)
  const display = cs.display
  const defaultDisplay = TAG_DISPLAY[o.tag] ?? 'block'
  if (o.positioned) {
    if (display !== 'block') out.push('display:block')
  } else if (display !== defaultDisplay && display !== 'none') {
    out.push(`display:${display}`)
  }
  const margins = SIDES.map((s) => cs.getPropertyValue(`margin-${s}`))
  if (margins.some((m) => !isZero(m))) out.push(`margin:${margins.join(' ')}`)
  const pads = SIDES.map((s) => cs.getPropertyValue(`padding-${s}`))
  if (pads.some((m) => !isZero(m))) out.push(`padding:${pads.join(' ')}`)
  if (!o.skipBorders) {
    for (const s of SIDES) {
      const b = borderDecl(cs, s)
      if (b) out.push(b)
    }
    const radius = cs.borderRadius
    if (radius && !isZero(radius)) out.push(`border-radius:${radius}`)
  }
  if (!isTransparent(cs.backgroundColor)) out.push(`background-color:${cs.backgroundColor}`)
  if (cs.backgroundImage && cs.backgroundImage !== 'none') {
    out.push(`background-image:${cs.backgroundImage}`)
    out.push(`background-size:${cs.backgroundSize}`)
    out.push(`background-position:${cs.backgroundPosition}`)
    out.push(`background-repeat:${cs.backgroundRepeat}`)
  }
  const deco = cs.textDecorationLine
  if (deco && deco !== 'none') {
    out.push(`text-decoration:${deco} ${cs.textDecorationStyle} ${cs.textDecorationColor}`)
  }
  if (cs.verticalAlign && cs.verticalAlign !== 'baseline')
    out.push(`vertical-align:${cs.verticalAlign}`)
  const sized =
    o.tag === 'img' ||
    o.tag === 'svg' ||
    o.tag === 'table' ||
    o.tag === 'td' ||
    o.tag === 'th' ||
    display === 'inline-block' ||
    o.positioned
  if (o.tag !== 'img') {
    if ((sized || o.authoredWidth) && cs.width && cs.width !== 'auto') out.push(`width:${cs.width}`)
    if ((o.tag === 'svg' || o.authoredHeight) && cs.height && cs.height !== 'auto') {
      out.push(`height:${cs.height}`)
    }
  }
  if (cs.minWidth && !isZero(cs.minWidth) && cs.minWidth !== 'auto')
    out.push(`min-width:${cs.minWidth}`)
  if (cs.maxWidth && cs.maxWidth !== 'none') out.push(`max-width:${cs.maxWidth}`)
  if (cs.float && cs.float !== 'none') out.push(`float:${cs.float}`)
  if (cs.clear && cs.clear !== 'none') out.push(`clear:${cs.clear}`)
  if (cs.transform && cs.transform !== 'none' && cs.transform !== 'matrix(1, 0, 0, 1, 0, 0)') {
    out.push(`transform:${cs.transform}`)
  }
  if (cs.opacity && cs.opacity !== '1') out.push(`opacity:${cs.opacity}`)
  if (cs.columnCount && cs.columnCount !== 'auto') {
    out.push(`column-count:${cs.columnCount}`)
    out.push(`column-gap:${cs.columnGap}`)
  }
  if (cs.breakBefore === 'page' || cs.breakBefore === 'always') out.push('page-break-before:always')
  if (cs.breakAfter === 'page' || cs.breakAfter === 'always') out.push('page-break-after:always')
  if (cs.boxSizing === 'border-box' && (sized || o.authoredWidth)) out.push('box-sizing:border-box')
  return out.map((d) =>
    d.replace(/(-?\d+\.\d{3,})px/g, (_, n: string) => `${parseFloat(Number(n).toFixed(2))}px`),
  )
}

export interface HtmlExportOptions {
  title: string
  lang?: string
  textWidthPx?: number
}

const SCREEN_ONLY_CLASSES = ['page-dark', 'show-marks', 'workspace-dark']

/** Screen-only ancestor classes (dark page, marks) are lifted during the synchronous walk so every theme exports the same file. */
export function buildStandaloneHtml(root: HTMLElement, opts: HtmlExportOptions): string {
  const restore: Array<() => void> = []
  for (let el: HTMLElement | null = root; el; el = el.parentElement) {
    for (const c of SCREEN_ONLY_CLASSES) {
      if (el.classList.contains(c)) {
        const target = el
        target.classList.remove(c)
        restore.push(() => target.classList.add(c))
      }
    }
  }
  try {
    return serializeDocument(root, opts)
  } finally {
    for (const r of restore) r()
  }
}

function serializeDocument(root: HTMLElement, opts: HtmlExportOptions): string {
  const win = root.ownerDocument.defaultView
  if (!win) return ''
  const rootCs = win.getComputedStyle(root)
  const rootDecls = styleDecls(rootCs, {
    parent: null,
    tag: 'div',
    authoredWidth: false,
    authoredHeight: false,
    skipBorders: true,
    positioned: false,
  }).filter((d) => !/^(margin|padding|width|max-width|min-width|box-sizing|background)/.test(d))
  const body: string[] = []
  for (const child of Array.from(root.childNodes)) serializeNode(child, rootCs, win, body)
  const width =
    opts.textWidthPx && opts.textWidthPx > 0 ? `${Math.round(opts.textWidthPx)}px` : '8.5in'
  const lang = opts.lang ? ` lang="${escHtml(opts.lang)}"` : ''
  const dir = rootCs.direction === 'rtl' ? ' dir="rtl"' : ''
  return (
    `<!DOCTYPE html>\n<html${lang}${dir}>\n<head>\n<meta charset="utf-8">\n` +
    `<meta name="viewport" content="width=device-width, initial-scale=1">\n` +
    `<title>${escText(opts.title)}</title>\n` +
    `<style>\nbody{margin:0;background:#fff}\n` +
    `p,h1,h2,h3,h4,h5,h6,ul,ol,li,table,td,th,div,span{margin:0;padding:0}\n` +
    `h1,h2,h3,h4,h5,h6{font-size:inherit;font-weight:inherit}\n` +
    `ul,ol{list-style:none}\na{color:inherit;text-decoration:inherit}\n` +
    `sub,sup{font-size:inherit;line-height:inherit;vertical-align:baseline}\n` +
    `.doc{max-width:${width};margin:0 auto;padding:48px 40px;box-sizing:content-box}\n` +
    `img{max-width:100%;height:auto}\ntable{border-collapse:collapse;border-spacing:0}\n` +
    `@media print{.doc{max-width:none;padding:0}}\n</style>\n</head>\n<body>\n` +
    `<div class="doc" style="${escHtml(rootDecls.join(';'))}">${body.join('')}</div>\n</body>\n</html>\n`
  )
}

function serializeNode(
  node: Node,
  parentCs: CSSStyleDeclaration,
  win: Window,
  out: string[],
): void {
  if (node.nodeType === Node.TEXT_NODE) {
    out.push(escText(node.nodeValue ?? ''))
    return
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return
  const el = node as HTMLElement
  if (hasClass(el, SKIP_CLASSES) && !hasClass(el, KEEP_WIDGET_CLASSES)) return
  const cs = win.getComputedStyle(el)
  if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') return
  if (hasClass(el, UNWRAP_CLASSES)) {
    // diff children against the wrapper so its inherited review styling (e.g. insertion
    // color) is dropped rather than re-inlined on nested elements
    for (const child of Array.from(el.childNodes)) serializeNode(child, cs, win, out)
    return
  }
  const rawTag = el.tagName.toLowerCase()
  if (PASSTHROUGH_TAGS.has(rawTag)) {
    const clone = el.cloneNode(true) as Element
    clone.removeAttribute('class')
    const size = [`width:${cs.width}`, `height:${cs.height}`]
    if (cs.verticalAlign !== 'baseline') size.push(`vertical-align:${cs.verticalAlign}`)
    clone.setAttribute('style', size.join(';'))
    out.push(clone.outerHTML)
    return
  }
  const tag = emitTag(el, cs)
  const positioned = cs.position === 'absolute' || cs.position === 'fixed'
  const decls = styleDecls(cs, {
    parent: parentCs,
    tag,
    authoredWidth: !!el.style.width,
    authoredHeight: !!el.style.height,
    skipBorders: el.classList.contains('has-ppr-change'),
    positioned,
  })
  if (positioned) {
    const host = el.offsetParent as HTMLElement | null
    const side =
      host && el.offsetLeft + el.offsetWidth / 2 > host.clientWidth / 2 ? 'right' : 'left'
    decls.push(`float:${side}`, 'margin:4px 8px')
  }
  const before = pseudo(el, cs, '::before', win)
  const after = pseudo(el, cs, '::after', win)
  const attrs: string[] = []
  for (const a of KEEP_ATTRS) {
    const v = el.getAttribute(a)
    if (v !== null && v !== '') attrs.push(` ${a}="${escHtml(v)}"`)
  }
  if (tag === 'a' && !el.getAttribute('href')) attrs.length = 0
  if (tag === 'img') {
    const w = Math.round(parseFloat(cs.width))
    const h = Math.round(parseFloat(cs.height))
    if (w > 0 && h > 0) attrs.push(` width="${w}" height="${h}"`)
    // crop windows and diagram fills deliberately overflow their box: keep them unclamped
    const box = el.parentElement?.clientWidth ?? 0
    if (box > 0 && w > box + 1) decls.push('max-width:none')
  }
  const style = decls.length ? ` style="${escHtml(decls.join(';'))}"` : ''
  if (VOID_TAGS.has(tag)) {
    out.push(`<${tag}${attrs.join('')}${style}>`)
    return
  }
  out.push(`<${tag}${attrs.join('')}${style}>`)
  if (before) out.push(before)
  for (const child of Array.from(el.childNodes)) serializeNode(child, cs, win, out)
  if (after) out.push(after)
  out.push(`</${tag}>`)
}

function pseudo(
  el: HTMLElement,
  elCs: CSSStyleDeclaration,
  which: '::before' | '::after',
  win: Window,
): string {
  const cs = win.getComputedStyle(el, which)
  if (!cs || cs.display === 'none' || cs.visibility === 'hidden') return ''
  if (CHROME_PSEUDO_CONTENT.test(cs.content)) return ''
  const text = resolveContent(cs.content, el)
  if (!text) return ''
  const decls = styleDecls(cs, {
    parent: elCs,
    tag: 'span',
    authoredWidth: false,
    authoredHeight: false,
    skipBorders: false,
    positioned: false,
  })
  const style = decls.length ? ` style="${escHtml(decls.join(';'))}"` : ''
  return `<span${style}>${escText(text)}</span>`
}
