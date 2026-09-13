/* global window, document */
/* Injected into the sandboxed preview frame. Plain ES2020, no imports: this
 * file is loaded as a raw string and appended as a <script>. It never mutates
 * the document except for its own overlay nodes (marked data-gx-inspector) and
 * a temporary contenteditable during text edits (a single click on rendered text
 * starts one with the caret at the click, like PowerPoint). */
;(() => {
  const SID = 'data-sid'
  const MARK = 'data-gx-inspector'
  // replaced by instrumentForPreview with the parse-map version this copy was built from
  const VERSION = Number('__GX_VERSION__')
  const post = (msg) => window.parent.postMessage({ ...msg, version: VERSION }, '*')
  const TEXT_TAGS = new Set([
    'p',
    'h1',
    'h2',
    'h3',
    'h4',
    'h5',
    'h6',
    'li',
    'td',
    'th',
    'span',
    'a',
    'blockquote',
    'figcaption',
    'label',
    'button',
    'dt',
    'dd',
    'caption',
    'summary',
    'legend',
    'em',
    'strong',
    'b',
    'i',
    'u',
    's',
    'small',
    'code',
    'cite',
    'q',
    'time',
    'div',
    'section',
    'article',
    'header',
    'footer',
    'nav',
    'aside',
    'main',
    'pre',
  ])
  const ATOMIC = new Set(['svg', 'canvas', 'iframe', 'video', 'audio', 'object', 'embed', 'math'])
  /** phrasing children a text element may carry and still edit inline as one run of rich text */
  const INLINE_TAGS = new Set([
    'a',
    'abbr',
    'b',
    'bdi',
    'bdo',
    'br',
    'cite',
    'code',
    'data',
    'del',
    'dfn',
    'em',
    'i',
    'ins',
    'kbd',
    'mark',
    'q',
    's',
    'samp',
    'small',
    'span',
    'strong',
    'sub',
    'sup',
    'time',
    'u',
    'var',
    'wbr',
  ])

  let mode = 'inspect'
  const selecting = () => mode === 'inspect'
  let selected = null
  let editing = null
  let dark = false
  let resizing = null
  let drag = null
  /** horizontal slide of a replaced element (image / video): live margin-left, committed as styles */
  let slide = null
  let pressed = null
  let swallowClick = false

  const style = document.createElement('style')
  style.setAttribute(MARK, '')
  style.textContent = `
    [${MARK}-overlay] { position: fixed; pointer-events: none; z-index: 2147483646; box-sizing: border-box; border-radius: 2px; transition: none; }
    [${MARK}-overlay="hover"] { outline: 1.5px dashed var(--gx-hover, #0f7fff); outline-offset: -1px; }
    [${MARK}-overlay="select"] { outline: 2px solid var(--gx-select, #0f7fff); outline-offset: -1px; background: rgba(15,127,255,0.05); }
    [${MARK}-overlay="dynamic"] { outline: 2px solid #f59e0b; outline-offset: -1px; }
    [${MARK}-overlay="highlight"] { background: rgba(255,213,79,0.35); }
    [${MARK}-label] { position: fixed; z-index: 2147483647; pointer-events: none; font: 11px/1 -apple-system, system-ui, sans-serif; color: #fff; background: var(--gx-select, #0f7fff); padding: 2px 6px; border-radius: 3px; white-space: nowrap; }
    [${MARK}-editing] { outline: 2px solid #2563eb !important; outline-offset: -1px; }
    [${MARK}-pin] { position: fixed; z-index: 2147483647; min-width: 18px; height: 18px; padding: 0 5px; box-sizing: border-box; border-radius: 9px; background: var(--gx-select, #0f7fff); color: #fff; font: 600 11px/18px -apple-system, system-ui, sans-serif; text-align: center; cursor: pointer; box-shadow: 0 1px 3px rgba(0,0,0,0.3); }
    [${MARK}-handle] { position: fixed; z-index: 2147483647; width: 9px; height: 9px; box-sizing: border-box; border: 1.5px solid var(--gx-select, #0f7fff); border-radius: 2px; background: #fff; pointer-events: auto; }
    [${MARK}-handle="n"], [${MARK}-handle="s"] { cursor: ns-resize; }
    [${MARK}-handle="e"], [${MARK}-handle="w"] { cursor: ew-resize; }
    [${MARK}-handle="ne"], [${MARK}-handle="sw"] { cursor: nesw-resize; }
    [${MARK}-handle="nw"], [${MARK}-handle="se"] { cursor: nwse-resize; }
    [${MARK}-grip] { position: fixed; z-index: 2147483647; width: 12px; height: 22px; box-sizing: border-box; border-radius: 4px; background: var(--gx-select, #0f7fff); color: #fff; font: 700 10px/22px -apple-system, system-ui, sans-serif; letter-spacing: -1px; text-align: center; cursor: grab; pointer-events: auto; box-shadow: 0 1px 3px rgba(0,0,0,0.3); }
    [${MARK}-drop] { position: fixed; z-index: 2147483647; pointer-events: none; background: var(--gx-select, #0f7fff); border-radius: 2px; box-shadow: 0 0 0 1px #fff; }
    [${MARK}-dragsrc] { opacity: 0.4 !important; }
    html[${MARK}-dragging], html[${MARK}-dragging] * { cursor: grabbing !important; user-select: none !important; }
    html[${MARK}-cursor="text"] *:not([${MARK}]) { cursor: text !important; }
    html[${MARK}-cursor="move"] *:not([${MARK}]) { cursor: move !important; }
    html[${MARK}-resizing], html[${MARK}-resizing] * { user-select: none !important; }
    html[${MARK}-sliding], html[${MARK}-sliding] * { cursor: move !important; user-select: none !important; }
  `
  document.documentElement.appendChild(style)

  const overlay = (kind) => {
    const el = document.createElement('div')
    el.setAttribute(MARK, '')
    el.setAttribute(`${MARK}-overlay`, kind)
    el.style.display = 'none'
    document.documentElement.appendChild(el)
    return el
  }
  const hoverBox = overlay('hover')
  const selectBox = overlay('select')
  const label = document.createElement('div')
  label.setAttribute(MARK, '')
  label.setAttribute(`${MARK}-label`, '')
  label.style.display = 'none'
  document.documentElement.appendChild(label)
  let highlightBoxes = []
  let markPins = []
  const HANDLE_DIRS = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
  const handles = HANDLE_DIRS.map((dir) => {
    const h = document.createElement('div')
    h.setAttribute(MARK, '')
    h.setAttribute(`${MARK}-handle`, dir)
    h.style.display = 'none'
    document.documentElement.appendChild(h)
    return { dir, el: h }
  })
  const grip = document.createElement('div')
  grip.setAttribute(MARK, '')
  grip.setAttribute(`${MARK}-grip`, '')
  grip.textContent = '\u22ee\u22ee'
  grip.style.display = 'none'
  document.documentElement.appendChild(grip)
  const dropBar = document.createElement('div')
  dropBar.setAttribute(MARK, '')
  dropBar.setAttribute(`${MARK}-drop`, '')
  dropBar.style.display = 'none'
  document.documentElement.appendChild(dropBar)

  const place = (box, el) => {
    if (!el || !el.isConnected) {
      box.style.display = 'none'
      return
    }
    const r = el.getBoundingClientRect()
    box.style.display = 'block'
    box.style.left = `${r.left}px`
    box.style.top = `${r.top}px`
    box.style.width = `${Math.max(r.width, 2)}px`
    box.style.height = `${Math.max(r.height, 2)}px`
  }
  const describe = (el) => {
    const cls = (el.getAttribute('class') || '').trim().split(/\s+/).filter(Boolean).slice(0, 2)
    return (
      el.tagName.toLowerCase() +
      (el.id ? `#${el.id}` : '') +
      (cls.length ? `.${cls.join('.')}` : '')
    )
  }
  const placeLabel = (el) => {
    if (!el) {
      label.style.display = 'none'
      return
    }
    const r = el.getBoundingClientRect()
    label.textContent = describe(el)
    label.style.display = 'block'
    label.style.left = `${Math.max(0, r.left)}px`
    label.style.top = `${r.top >= 18 ? r.top - 18 : r.bottom + 2}px`
  }
  const isOwn = (el) => !!(el && el.closest && el.closest(`[${MARK}]`))
  const sidOf = (el) => {
    const v = el && el.getAttribute && el.getAttribute(SID)
    return v ? Number(v) : null
  }
  /** nearest selectable target: atomic subtrees select as a whole, own overlays never */
  const targetOf = (node) => {
    let el = node && node.nodeType === 3 ? node.parentElement : node
    if (!el || isOwn(el)) return null
    for (let cur = el; cur && cur !== document.documentElement; cur = cur.parentElement) {
      if (ATOMIC.has(cur.tagName.toLowerCase())) el = cur
    }
    if (el === document.documentElement || el === document.body) return null
    return el
  }
  const rectOf = (el) => {
    const r = el.getBoundingClientRect()
    return { x: r.left, y: r.top, width: r.width, height: r.height }
  }
  const px = (v) => (v && v.endsWith('px') ? String(Math.round(parseFloat(v) * 10) / 10) : v)
  const toHex = (rgb) => {
    const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/.exec(rgb || '')
    if (!m) return ''
    if (m[4] !== undefined && Number(m[4]) === 0) return ''
    const h = (n) => Number(n).toString(16).padStart(2, '0')
    return `#${h(m[1])}${h(m[2])}${h(m[3])}`
  }
  /** effective styles the floating toolbar / style panel start from */
  const snapshot = (el) => {
    const c = window.getComputedStyle(el)
    return {
      color: toHex(c.color),
      fontSize: px(c.fontSize),
      fontWeight: c.fontWeight,
      fontStyle: c.fontStyle,
      textAlign: c.textAlign,
      background: toHex(c.backgroundColor),
      backgroundImage: c.backgroundImage,
      width: px(c.width),
      height: px(c.height),
      borderRadius: px(c.borderTopLeftRadius),
      padding: px(c.paddingTop),
      opacity: c.opacity,
      transform: el.style.transform,
      marginLeft: el.style.marginLeft,
      marginRight: el.style.marginRight,
      objectFit: c.objectFit,
      ...(el.tagName === 'IMG'
        ? {
            image: {
              src: el.getAttribute('src') || '',
              alt: el.getAttribute('alt') || '',
              naturalWidth: el.naturalWidth,
              naturalHeight: el.naturalHeight,
            },
          }
        : {}),
    }
  }
  /** the single editable text run (and its index among the direct text nodes) for the style panel's
   * text field, plus whether the element edits inline at all (same rules as beginEdit) */
  const textRunOf = (el) => {
    const inlineEditable = editMode(el) !== null
    const runs = Array.from(el.childNodes).filter((n) => n.nodeType === 3 && n.textContent.trim())
    if (runs.length !== 1) return { textRun: null, textRunIndex: -1, inlineEditable }
    return {
      textRun: runs[0].textContent,
      textRunIndex: textNodeIndex(el, runs[0]),
      inlineEditable,
    }
  }
  const fingerprint = (el) => ({
    sid: sidOf(el),
    tag: el.tagName.toLowerCase(),
    className: el.getAttribute('class') || '',
    childElementCount: el.childElementCount,
    text: (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 400),
    rect: rectOf(el),
    computed: snapshot(el),
    ...textRunOf(el),
  })
  let rectQueued = false
  /** the host draws the floating toolbar from this; sent after every scroll / resize / live style poke */
  const postRect = () => {
    if (rectQueued) return
    rectQueued = true
    window.requestAnimationFrame(() => {
      rectQueued = false
      const sid = selected && selected.isConnected ? sidOf(selected) : null
      if (sid === null) return
      post({
        type: 'gx:rect',
        sid,
        rect: rectOf(selected),
        computed: snapshot(selected),
        ...textRunOf(selected),
      })
    })
  }
  const placePin = (pin, el) => {
    if (!el || !el.isConnected) {
      pin.style.display = 'none'
      return
    }
    const r = el.getBoundingClientRect()
    pin.style.display = 'block'
    pin.style.left = `${Math.max(0, r.right - 9)}px`
    pin.style.top = `${Math.max(0, r.top - 9)}px`
  }
  // the hover box is fixed-position: without re-hit-testing under the pointer it stays put while the page scrolls under it
  let pointer = null
  /** pointer feedback like PowerPoint/WPS: an I-beam over editable text (a click there places the
   * caret), the move cursor over the rest of the selected element (a press there drags it) */
  const setCursor = (kind) => {
    if (kind) document.documentElement.setAttribute(`${MARK}-cursor`, kind)
    else document.documentElement.removeAttribute(`${MARK}-cursor`)
  }
  const cursorFor = (el, x, y) => {
    if (!el) return null
    if (editMode(el) && overText(el, x, y)) return 'text'
    return el === selected && sidOf(el) !== null ? 'move' : null
  }
  let hovered = null
  const hover = (el) => {
    place(hoverBox, el)
    if (el === hovered) return
    hovered = el
    post({ type: 'gx:hover', sid: el ? sidOf(el) : null })
  }
  const rehover = () => {
    if (!pointer || !selecting() || editing || drag || resizing || slide) return hover(null)
    hover(targetOf(document.elementFromPoint(pointer.x, pointer.y)))
  }
  const NO_RESIZE = new Set(['tr', 'tbody', 'thead', 'tfoot', 'colgroup', 'col', 'li', 'option'])
  /** replaced elements take width / height (and slide sideways) whatever their display */
  const REPLACED = new Set(['img', 'video', 'iframe', 'canvas', 'svg'])
  /** sized boxes only: inline text and table rows ignore width / height */
  const resizable = (el) => {
    if (!el || !el.isConnected || sidOf(el) === null || el === document.body) return false
    const tag = el.tagName.toLowerCase()
    if (NO_RESIZE.has(tag)) return false
    if (REPLACED.has(tag)) return true
    const d = window.getComputedStyle(el).display
    return d !== 'inline' && d !== 'contents' && d !== 'none'
  }
  const placeHandles = (el) => {
    const show = selecting() && !editing && !drag && resizable(el)
    if (!show) {
      for (const { el: h } of handles) h.style.display = 'none'
      return
    }
    const r = el.getBoundingClientRect()
    const xs = { w: r.left, e: r.right, n: r.left + r.width / 2, s: r.left + r.width / 2 }
    const ys = { n: r.top, s: r.bottom, e: r.top + r.height / 2, w: r.top + r.height / 2 }
    for (const { dir, el: h } of handles) {
      // a mid-edge handle on a tiny side would sit on top of the corner handles
      const tinySide = dir === 'n' || dir === 's' ? r.width < 30 : r.height < 14
      if (dir.length === 1 && tinySide) {
        h.style.display = 'none'
        continue
      }
      const x = dir.includes('w') ? xs.w : dir.includes('e') ? xs.e : xs.n
      const y = dir.includes('n') ? ys.n : dir.includes('s') ? ys.s : ys.e
      h.style.display = 'block'
      h.style.left = `${x - 4.5}px`
      h.style.top = `${y - 4.5}px`
    }
  }
  /** drag grip beside the left edge; below the box when the element hugs the viewport edge, inside as a last resort.
   * Replaced elements (images / video) have no text to protect: their whole body drags, so they carry no grip. */
  const placeGrip = (el) => {
    const show =
      selecting() &&
      !editing &&
      !drag &&
      !slide &&
      el &&
      el.isConnected &&
      sidOf(el) !== null &&
      el !== document.body &&
      !REPLACED.has(el.tagName.toLowerCase())
    if (!show) {
      grip.style.display = 'none'
      return
    }
    const r = el.getBoundingClientRect()
    grip.style.display = 'block'
    if (r.left >= 16) {
      grip.style.left = `${r.left - 15}px`
      grip.style.top = `${Math.max(0, r.top + r.height / 2 - 11)}px`
    } else if (r.bottom + 26 <= window.innerHeight) {
      grip.style.left = `${Math.max(0, r.left)}px`
      grip.style.top = `${r.bottom + 3}px`
    } else {
      grip.style.left = `${r.left + 3}px`
      grip.style.top = `${Math.max(0, r.top + 3)}px`
    }
  }
  const refresh = () => {
    rehover()
    place(selectBox, selected)
    placeLabel(selected)
    placeHandles(selected)
    placeGrip(selected)
    for (const { box, el } of highlightBoxes) place(box, el)
    for (const { pin, el } of markPins) placePin(pin, el)
    if (selected) postRect()
  }

  const select = (el, notify) => {
    if (editing && el !== editing.el) finishEdit(true)
    selected = el
    refresh()
    if (notify) {
      if (!el) post({ type: 'gx:select', element: null, dynamic: false })
      else post({ type: 'gx:select', element: fingerprint(el), dynamic: sidOf(el) === null })
    }
  }
  const bySid = (sid) => document.querySelector(`[${SID}="${sid}"]`)

  // ---- inline text editing (one element at a time) ----
  const textNodeIndex = (el, node) => {
    let i = 0
    for (const child of el.childNodes) {
      if (child.nodeType === 3) {
        if (child === node) return i
        i++
      }
    }
    return -1
  }
  const directTextNodes = (el) =>
    Array.from(el.childNodes).filter((n) => n.nodeType === 3 && n.textContent.trim())
  /** how the element's text edits inline: 'text' = one direct text node and no child elements
   * (committed as set_text_node, the smallest source diff); 'html' = text mixed with phrasing
   * children like <strong> / <a> / <br> (committed as the whole inner HTML); null = not inline
   * editable (no text, or block children next to several runs: those edits go through AI or source) */
  const editMode = (el) => {
    if (!el || sidOf(el) === null || !TEXT_TAGS.has(el.tagName.toLowerCase())) return null
    if (!(el.textContent || '').trim()) return null
    const children = Array.from(el.children)
    if (children.length && children.every((c) => INLINE_TAGS.has(c.tagName.toLowerCase())))
      return 'html'
    // no children, or block children (an <li> with a nested list): its single direct run still edits on its own
    return directTextNodes(el).length === 1 ? 'text' : null
  }
  /** inner HTML as the source would carry it: the preview's own data-sid markers stripped */
  const innerSource = (el) => {
    const clone = el.cloneNode(true)
    for (const n of clone.querySelectorAll(`[${SID}]`)) n.removeAttribute(SID)
    return clone.innerHTML
  }
  /** the current native selection, when it lies entirely inside `el` */
  const selectionIn = (el) => {
    const sel = window.getSelection()
    if (!sel || sel.rangeCount === 0) return null
    const r = sel.getRangeAt(0)
    return el.contains(r.startContainer) && el.contains(r.endContainer) ? r : null
  }
  /** `at` = the click that started the edit: the caret lands there, like PowerPoint (a drag-selected
   * range made before the edit is kept). Without a point the whole run is selected (toolbar / host entry). */
  const beginEdit = (el, at) => {
    if (!el || editing) return
    const mode = editMode(el)
    if (!mode) return
    // read the pre-edit selection before focus() can move it
    const kept = at ? selectionIn(el) : null
    const node = mode === 'text' ? directTextNodes(el)[0] : null
    editing = {
      el,
      mode,
      node,
      index: node ? textNodeIndex(el, node) : -1,
      before: node ? node.textContent : innerSource(el),
      rawBefore: el.innerHTML,
    }
    el.setAttribute(`${MARK}-editing`, '')
    el.setAttribute('contenteditable', 'plaintext-only')
    el.focus()
    setCursor(null)
    let range = null
    if (kept && !kept.collapsed) range = kept
    else if (at) {
      const caret = document.caretRangeFromPoint && document.caretRangeFromPoint(at.x, at.y)
      if (caret && el.contains(caret.startContainer)) range = caret
    }
    if (!range) {
      range = document.createRange()
      range.selectNodeContents(node || el)
    }
    const sel = window.getSelection()
    sel.removeAllRanges()
    sel.addRange(range)
    // the host cleared its text selection on gx:select; re-adding an unchanged range fires no selectionchange
    if (!range.collapsed) postTextSelection()
  }
  const finishEdit = (commit) => {
    if (!editing) return
    const { el, mode, node, index, before, rawBefore } = editing
    editing = null
    el.removeAttribute('contenteditable')
    el.removeAttribute(`${MARK}-editing`)
    // contenteditable may have split/merged text nodes; read the element's direct text again
    const after =
      mode === 'text'
        ? Array.from(el.childNodes)
            .filter((n) => n.nodeType === 3)
            .map((n) => n.textContent)
            .join('')
        : innerSource(el)
    if (commit && after !== before) {
      if (mode === 'text')
        post({ type: 'gx:textEditCommit', sid: sidOf(el), textNodeIndex: index, newText: after })
      else post({ type: 'gx:htmlEditCommit', sid: sidOf(el), html: after })
      // the host drops its text selection when it re-selects after the commit: a range still
      // selected in the run goes out again so bold / italic right after typing wrap it
      postTextSelection()
    } else {
      // an unchanged run is left untouched (rewriting it would collapse the selection)
      if (after !== before) {
        if (mode === 'text' && node.isConnected) node.textContent = before
        else if (mode === 'html') el.innerHTML = rawBefore
      }
      post({ type: 'gx:textEditCancel' })
    }
    refresh()
  }

  // ---- resize handles: live inline width / height, committed by the host as one set_style ----
  const KEEP_RATIO = new Set(['IMG', 'VIDEO', 'IFRAME', 'CANVAS', 'SVG', 'PICTURE'])
  for (const { dir, el: h } of handles) {
    h.addEventListener('pointerdown', (e) => {
      if (!selected || e.button !== 0) return
      e.preventDefault()
      e.stopPropagation()
      const r = selected.getBoundingClientRect()
      resizing = {
        dir,
        el: selected,
        x: e.clientX,
        y: e.clientY,
        w: r.width,
        h: r.height,
        ratio: KEEP_RATIO.has(selected.tagName) && r.height > 0 ? r.width / r.height : null,
        moved: false,
        before: { width: selected.style.width, height: selected.style.height },
      }
      h.setPointerCapture(e.pointerId)
      document.documentElement.setAttribute(`${MARK}-resizing`, '')
      post({ type: 'gx:drag', active: true })
    })
    h.addEventListener('pointermove', (e) => {
      if (!resizing || resizing.dir !== dir) return
      const { el, x, y, w, ratio } = resizing
      const sx = dir.includes('e') ? 1 : dir.includes('w') ? -1 : 0
      const sy = dir.includes('s') ? 1 : dir.includes('n') ? -1 : 0
      let nw = Math.max(8, w + sx * (e.clientX - x))
      let nh = Math.max(8, resizing.h + sy * (e.clientY - y))
      if (ratio && sx && sy) nh = nw / ratio
      resizing.moved = true
      if (sx) el.style.width = `${Math.round(nw)}px`
      if (sy || (ratio && sx)) el.style.height = `${Math.round(nh)}px`
      refresh()
    })
    const finish = (e) => {
      if (!resizing || resizing.dir !== dir) return
      const { el, moved, before } = resizing
      resizing = null
      document.documentElement.removeAttribute(`${MARK}-resizing`)
      if (h.hasPointerCapture && h.hasPointerCapture(e.pointerId))
        h.releasePointerCapture(e.pointerId)
      const styles = {}
      if (el.style.width !== before.width) styles.width = el.style.width
      if (el.style.height !== before.height) styles.height = el.style.height
      if (moved && Object.keys(styles).length) post({ type: 'gx:resize', sid: sidOf(el), styles })
      post({ type: 'gx:drag', active: false })
      refresh()
    }
    h.addEventListener('pointerup', finish)
    h.addEventListener('pointercancel', finish)
  }

  // ---- drag the selected element to a new place in the flow (dropped before / after another element) ----
  const DRAG_START_PX = 5
  // a block between two <li>s or <td>s is parsed out of its list / table: such parts only pair with
  // their own kind, anything else lands next to the whole list / table
  const PARTS = new Set([
    'li',
    'dt',
    'dd',
    'tr',
    'td',
    'th',
    'thead',
    'tbody',
    'tfoot',
    'caption',
    'colgroup',
    'col',
    'option',
    'optgroup',
  ])
  const tagOf = (el) => el.tagName.toLowerCase()
  const dropTargetAt = (x, y) => {
    if (!drag) return null
    const dragTag = tagOf(drag.el)
    for (const cand of document.elementsFromPoint(x, y)) {
      if (isOwn(cand) || cand === document.body || cand === document.documentElement) continue
      if (drag.el === cand || drag.el.contains(cand)) continue
      let el = targetOf(cand)
      if (!el) continue
      if (PARTS.has(dragTag)) {
        if (tagOf(el) !== dragTag) continue
      } else {
        while (PARTS.has(tagOf(el)) && el.parentElement && el.parentElement !== document.body) {
          el = el.parentElement
        }
        if (PARTS.has(tagOf(el))) continue
      }
      if (sidOf(el) === null || el.parentElement === document.documentElement) continue
      if (drag.el === el || drag.el.contains(el)) continue
      const r = el.getBoundingClientRect()
      const ps = window.getComputedStyle(el.parentElement)
      const d = window.getComputedStyle(el).display
      // side-by-side children (flex rows, grids, inline boxes) split left / right instead of top / bottom
      const horizontal =
        (ps.display.includes('flex') && !ps.flexDirection.startsWith('column')) ||
        ps.display.includes('grid') ||
        d.startsWith('inline') ||
        d === 'table-cell'
      const position = horizontal
        ? x < r.left + r.width / 2
          ? 'before'
          : 'after'
        : y < r.top + r.height / 2
          ? 'before'
          : 'after'
      return { el, position, horizontal, rect: r }
    }
    return null
  }
  const showDrop = (target) => {
    if (!target) {
      dropBar.style.display = 'none'
      return
    }
    const { rect: r, position, horizontal } = target
    dropBar.style.display = 'block'
    if (horizontal) {
      dropBar.style.left = `${(position === 'before' ? r.left : r.right) - 1.5}px`
      dropBar.style.top = `${r.top}px`
      dropBar.style.width = '3px'
      dropBar.style.height = `${Math.max(r.height, 4)}px`
    } else {
      dropBar.style.left = `${r.left}px`
      dropBar.style.top = `${(position === 'before' ? r.top : r.bottom) - 1.5}px`
      dropBar.style.width = `${Math.max(r.width, 4)}px`
      dropBar.style.height = '3px'
    }
  }
  const endDrag = (commit) => {
    if (!drag) return
    const { el, target } = drag
    drag = null
    el.removeAttribute(`${MARK}-dragsrc`)
    document.documentElement.removeAttribute(`${MARK}-dragging`)
    showDrop(null)
    if (commit && target) {
      post({
        type: 'gx:moveTo',
        sid: sidOf(el),
        position: target.position,
        ref_sid: sidOf(target.el),
      })
    }
    post({ type: 'gx:drag', active: false })
    refresh()
  }
  /** the point sits on rendered text of the element: a drag there selects words, it does not move the box */
  const overText = (el, x, y) => {
    const caret = document.caretRangeFromPoint && document.caretRangeFromPoint(x, y)
    const node = caret && caret.startContainer
    if (!node || node.nodeType !== 3 || !el.contains(node) || !node.textContent.trim()) return false
    const range = document.createRange()
    range.selectNodeContents(node)
    for (const r of range.getClientRects()) {
      if (x >= r.left && x <= r.right && y >= r.top && y <= r.bottom) return true
    }
    return false
  }
  document.addEventListener(
    'mousedown',
    (e) => {
      pressed = null
      if (!selecting() || editing || e.button !== 0 || !selected) return
      if (e.target === grip) {
        e.preventDefault()
        pressed = { x: e.clientX, y: e.clientY, grip: true }
        return
      }
      const el = targetOf(e.target)
      if (el !== selected || el === document.body || sidOf(el) === null) return
      if (overText(el, e.clientX, e.clientY)) return
      // a replaced element's own drag would otherwise ghost the image under our gesture
      if (REPLACED.has(el.tagName.toLowerCase())) e.preventDefault()
      pressed = { x: e.clientX, y: e.clientY, grip: false }
    },
    true,
  )
  // ---- slide a replaced element sideways inside its container: live margin-left, snapping to
  // left / center / right (written in the toolbar's own auto-margin form so its buttons light up) ----
  const SNAP_PX = 8
  const setInline = (el, styles) => {
    for (const [prop, val] of Object.entries(styles)) {
      if (val) el.style.setProperty(prop, val)
      else el.style.removeProperty(prop)
    }
  }
  const alignStyles = (align) => ({
    display: 'block',
    'margin-left': align === 'left' ? '0' : 'auto',
    'margin-right': align === 'right' ? '0' : 'auto',
  })
  const beginSlide = (el, x) => {
    const parent = el.parentElement
    if (!parent) return false
    const pc = window.getComputedStyle(parent)
    const pr = parent.getBoundingClientRect()
    const contentLeft = pr.left + parseFloat(pc.borderLeftWidth) + parseFloat(pc.paddingLeft)
    const contentW = parent.clientWidth - parseFloat(pc.paddingLeft) - parseFloat(pc.paddingRight)
    const avail = Math.max(0, contentW - el.offsetWidth)
    if (avail <= 0) return false
    // rotation / flip keep the center: the layout left is the visual center minus half the layout width
    const r = el.getBoundingClientRect()
    const start = Math.min(
      avail,
      Math.max(0, r.left + r.width / 2 - el.offsetWidth / 2 - contentLeft),
    )
    slide = {
      el,
      x,
      start,
      avail,
      moved: false,
      before: {
        display: el.style.display,
        'margin-left': el.style.marginLeft,
        'margin-right': el.style.marginRight,
      },
    }
    document.documentElement.setAttribute(`${MARK}-sliding`, '')
    hover(null)
    post({ type: 'gx:drag', active: true })
    return true
  }
  const applySlide = (styles) => setInline(slide.el, styles)
  const moveSlide = (clientX) => {
    const { start, avail, x } = slide
    const ml = Math.min(avail, Math.max(0, start + clientX - x))
    slide.moved = true
    if (ml <= SNAP_PX) applySlide(alignStyles('left'))
    else if (Math.abs(ml - avail / 2) <= SNAP_PX) applySlide(alignStyles('center'))
    else if (avail - ml <= SNAP_PX) applySlide(alignStyles('right'))
    else
      applySlide({ display: 'block', 'margin-left': `${Math.round(ml)}px`, 'margin-right': 'auto' })
    refresh()
  }
  const endSlide = (commit) => {
    if (!slide) return
    const { el, moved, before } = slide
    slide = null
    document.documentElement.removeAttribute(`${MARK}-sliding`)
    if (!commit || !moved) setInline(el, before)
    else {
      const styles = {}
      for (const prop of Object.keys(before)) {
        const now = el.style.getPropertyValue(prop)
        if (now !== before[prop]) styles[prop] = now
      }
      if (Object.keys(styles).length) post({ type: 'gx:resize', sid: sidOf(el), styles })
    }
    post({ type: 'gx:drag', active: false })
    refresh()
  }
  document.addEventListener(
    'mousemove',
    (e) => {
      if (slide) {
        e.preventDefault()
        if (!(e.buttons & 1)) return endSlide(false)
        moveSlide(e.clientX)
        return
      }
      if (drag) {
        e.preventDefault()
        // released outside the frame: the mouseup never reached us
        if (!(e.buttons & 1)) return endDrag(false)
        drag.target = dropTargetAt(e.clientX, e.clientY)
        showDrop(drag.target)
        // keep the page moving when the pointer rides an edge of the viewport
        const edge = 32
        if (e.clientY < edge) window.scrollBy(0, -12)
        else if (e.clientY > window.innerHeight - edge) window.scrollBy(0, 12)
        return
      }
      if (!pressed || !(e.buttons & 1)) return
      if (Math.hypot(e.clientX - pressed.x, e.clientY - pressed.y) < DRAG_START_PX) return
      const { x, y, grip: fromGrip } = pressed
      pressed = null
      if (!selected || !selected.isConnected) return
      // a mostly horizontal pull on an image moves it sideways; the grip and vertical pulls reorder
      if (
        !fromGrip &&
        REPLACED.has(selected.tagName.toLowerCase()) &&
        Math.abs(e.clientX - x) > Math.abs(e.clientY - y) &&
        beginSlide(selected, x)
      )
        return moveSlide(e.clientX)
      drag = { el: selected, target: null }
      selected.setAttribute(`${MARK}-dragsrc`, '')
      document.documentElement.setAttribute(`${MARK}-dragging`, '')
      hover(null)
      for (const { el: h } of handles) h.style.display = 'none'
      grip.style.display = 'none'
      post({ type: 'gx:drag', active: true })
    },
    true,
  )
  document.addEventListener(
    'mouseup',
    () => {
      pressed = null
      if (slide) {
        swallowClick = true
        endSlide(true)
        return
      }
      if (!drag) return
      swallowClick = true
      endDrag(true)
    },
    true,
  )
  // the page's own draggable images / links must not start a native drag under ours
  document.addEventListener(
    'dragstart',
    (e) => {
      if (selecting()) e.preventDefault()
    },
    true,
  )

  // ---- events (capture phase: we see them before the page's own handlers) ----
  document.addEventListener(
    'mousemove',
    (e) => {
      pointer = { x: e.clientX, y: e.clientY }
      if (!selecting() || editing || drag || resizing || slide) return setCursor(null)
      const el = targetOf(e.target)
      hover(el)
      setCursor(cursorFor(el, e.clientX, e.clientY))
    },
    true,
  )
  // on the root element only: a capture listener on document would fire for every descendant
  document.documentElement.addEventListener('mouseleave', () => {
    pointer = null
    hover(null)
    pressed = null
    if (drag) {
      drag.target = null
      showDrop(null)
    }
  })
  document.addEventListener(
    'click',
    (e) => {
      if (!selecting()) return
      if (swallowClick) {
        swallowClick = false
        e.preventDefault()
        e.stopPropagation()
        return
      }
      const a = e.target && e.target.closest && e.target.closest('a[href]')
      if (a) {
        e.preventDefault()
        if (e.metaKey || e.ctrlKey) post({ type: 'gx:navigateBlocked', href: a.href })
      }
      if (editing) {
        if (editing.el.contains(e.target)) return
        finishEdit(true)
      }
      const el = targetOf(e.target)
      if (!el) {
        // a click on the bare page (body / html) clears the selection like a click on the canvas
        if (e.target === document.body || e.target === document.documentElement) select(null, true)
        return
      }
      e.preventDefault()
      e.stopPropagation()
      select(el, true)
      // A single left click on the text itself starts editing with the caret at the click
      // (PowerPoint/WPS); the box around the text only selects, so it stays the drag surface.
      const plain = e.button === 0 && !e.shiftKey && !e.metaKey && !e.ctrlKey && !e.altKey
      if (plain && editMode(el) && overText(el, e.clientX, e.clientY))
        beginEdit(el, { x: e.clientX, y: e.clientY })
    },
    true,
  )
  // Double-click on the padding / box of a text element (its text already edits on a single
  // click, where the native double-click selects a word)
  document.addEventListener(
    'dblclick',
    (e) => {
      if (mode !== 'inspect' || editing) return
      const el = targetOf(e.target)
      if (!el || !TEXT_TAGS.has(el.tagName.toLowerCase())) return
      e.preventDefault()
      e.stopPropagation()
      select(el, true)
      beginEdit(el)
    },
    true,
  )
  document.addEventListener(
    'submit',
    (e) => {
      if (selecting()) e.preventDefault()
    },
    true,
  )
  /** a non-collapsed native range inside one direct text run of a source element, as the
   * gx:textSelect message (the host wraps it on bold / italic); null for anything else */
  const textSelection = () => {
    if (mode !== 'inspect') return null
    const sel = window.getSelection()
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null
    const r = sel.getRangeAt(0)
    if (r.startContainer !== r.endContainer || r.startContainer.nodeType !== 3) return null
    const node = r.startContainer
    const el = node.parentElement
    if (!el || isOwn(el) || sidOf(el) === null) return null
    // while editing, only a range inside the edited run counts (bold / italic wrap it after the commit)
    if (editing && el !== editing.el) return null
    return {
      type: 'gx:textSelect',
      sid: sidOf(el),
      textNodeIndex: textNodeIndex(el, node),
      start: r.startOffset,
      end: r.endOffset,
      text: node.textContent.slice(r.startOffset, r.endOffset),
    }
  }
  const postTextSelection = () => {
    const msg = textSelection()
    if (msg) post(msg)
  }
  document.addEventListener('selectionchange', postTextSelection, true)
  document.addEventListener(
    'keydown',
    (e) => {
      if (editing) {
        const mod = (e.metaKey || e.ctrlKey) && !e.altKey
        if (e.key === 'Escape') {
          // Esc = commit the text and return to the selected state (typed input is not lost, as in
          // Slides); an unchanged run goes out as a cancel and produces no history step
          e.preventDefault()
          finishEdit(true)
        } else if (e.key === 'Enter' && mod) {
          e.preventDefault()
          finishEdit(true)
        } else if (mod && (e.key.toLowerCase() === 'b' || e.key.toLowerCase() === 'i')) {
          // bold / italic on a range selected inside the edited run: commit the run first, then
          // hand the host the range again (its re-select after the commit cleared it) and the command
          const range = textSelection()
          if (!range) return
          e.preventDefault()
          finishEdit(true)
          post(range)
          post({ type: 'gx:keyCommand', command: e.key.toLowerCase() === 'b' ? 'bold' : 'italic' })
        }
        return
      }
      if ((drag || slide) && e.key === 'Escape') {
        e.preventDefault()
        if (drag) endDrag(false)
        else endSlide(false)
        return
      }
      const inField =
        e.target &&
        (/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) || e.target.isContentEditable)
      const mod = (e.metaKey || e.ctrlKey) && !e.altKey
      // the host owns history and zoom; the frame has keyboard focus, so relay the shortcuts
      if (mod && !inField) {
        const lower = e.key.toLowerCase()
        const hostCommand =
          lower === 'z'
            ? e.shiftKey
              ? 'redo'
              : 'undo'
            : lower === 'y' && !e.shiftKey
              ? 'redo'
              : { '=': 'zoomIn', '+': 'zoomIn', '-': 'zoomOut', _: 'zoomOut', 0: 'zoomReset' }[
                  e.key
                ]
        if (hostCommand) {
          e.preventDefault()
          post({ type: 'gx:keyCommand', command: hostCommand })
          return
        }
      }
      if (mode !== 'inspect') {
        if (e.key === 'Escape') post({ type: 'gx:keyCommand', command: 'escape' })
        return
      }
      const map = {
        Delete: 'delete',
        Backspace: 'delete',
        ArrowUp: 'prev',
        ArrowDown: 'next',
        ArrowLeft: 'parent',
        ArrowRight: 'child',
        Escape: 'escape',
      }
      const modKey = mod ? { k: 'askAi', b: 'bold', i: 'italic' }[e.key.toLowerCase()] : undefined
      const altKey =
        e.altKey && !mod ? { ArrowUp: 'moveUp', ArrowDown: 'moveDown' }[e.key] : undefined
      const command = modKey || altKey || map[e.key]
      if (!command || !selected || inField) return
      e.preventDefault()
      post({ type: 'gx:keyCommand', command })
    },
    true,
  )
  // Chromium reports trackpad pinch as ctrl+wheel; the frame swallows wheel events, so relay it.
  // Browse mode (present views) has no zoom control: leave the event alone there.
  document.addEventListener(
    'wheel',
    (e) => {
      if (mode !== 'inspect' || (!e.ctrlKey && !e.metaKey)) return
      e.preventDefault()
      post({ type: 'gx:zoom', delta: e.deltaY })
    },
    { passive: false, capture: true },
  )
  let scrollQueued = false
  window.addEventListener(
    'scroll',
    (e) => {
      refresh()
      if (e.target !== document || scrollQueued) return
      scrollQueued = true
      window.requestAnimationFrame(() => {
        scrollQueued = false
        post({ type: 'gx:scroll', y: window.scrollY })
      })
    },
    true,
  )
  window.addEventListener('resize', refresh)
  document.addEventListener('focusout', (e) => {
    if (editing && e.target === editing.el) finishEdit(true)
  })

  window.addEventListener('message', (e) => {
    const msg = e.data
    if (!msg || typeof msg.type !== 'string') return
    switch (msg.type) {
      case 'gx:select':
        select(msg.sid === null ? null : bySid(msg.sid), false)
        if (selected) {
          selected.scrollIntoView({ block: 'nearest' })
          postRect()
        }
        break
      case 'gx:previewStyle': {
        // live preview only: the host commits the same declarations to the source afterwards
        const el = bySid(msg.sid)
        if (!el) break
        for (const [prop, val] of Object.entries(msg.styles)) {
          if (val === null) el.style.removeProperty(prop)
          else el.style.setProperty(prop, val)
        }
        refresh()
        break
      }
      case 'gx:highlight':
        for (const { box } of highlightBoxes) box.remove()
        highlightBoxes = msg.sids
          .map((sid) => bySid(sid))
          .filter(Boolean)
          .map((el) => ({ box: overlay('highlight'), el }))
        refresh()
        break
      case 'gx:mark':
        for (const { pin } of markPins) pin.remove()
        markPins = msg.marks
          .map(({ sid, label }) => ({ el: bySid(sid), sid, label }))
          .filter(({ el }) => el)
          .map(({ el, sid, label }) => {
            const pin = document.createElement('div')
            pin.setAttribute(MARK, '')
            pin.setAttribute(`${MARK}-pin`, '')
            pin.textContent = label
            pin.addEventListener('click', (e) => {
              e.preventDefault()
              e.stopPropagation()
              post({ type: 'gx:markClick', sid })
            })
            document.documentElement.appendChild(pin)
            return { pin, el }
          })
        refresh()
        break
      case 'gx:clearHighlight':
        for (const { box } of highlightBoxes) box.remove()
        highlightBoxes = []
        break
      case 'gx:endDrag':
        if (drag) endDrag(false)
        if (slide) endSlide(false)
        pressed = null
        break
      case 'gx:setMode':
        if (editing) finishEdit(true)
        if (drag) endDrag(false)
        if (slide) endSlide(false)
        mode = msg.mode
        rehover()
        if (mode === 'browse') select(null, false)
        break
      case 'gx:beginTextEdit':
        select(bySid(msg.sid), false)
        beginEdit(selected)
        break
      case 'gx:scrollTo':
        window.scrollTo(0, msg.y)
        break
      case 'gx:theme':
        dark = !!msg.dark
        document.documentElement.style.setProperty('--gx-select', dark ? '#4a9eff' : '#0f7fff')
        document.documentElement.style.setProperty('--gx-hover', dark ? '#4a9eff' : '#0f7fff')
        break
    }
  })

  post({
    type: 'gx:ready',
    title: document.title,
    docHeight: document.documentElement.scrollHeight,
  })
})()
