/* global window, document */
/* Injected into the sandboxed preview frame. Plain ES2020, no imports: this
 * file is loaded as a raw string and appended as a <script>. It never mutates
 * the document except for its own overlay nodes (marked data-gx-inspector) and
 * a temporary contenteditable during text edits. */
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

  let mode = 'inspect'
  const selecting = () => mode === 'inspect'
  let selected = null
  let editing = null
  let dark = false

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
  /** the single editable text run (and its index among the direct text nodes), same rule as beginEdit */
  const textRunOf = (el) => {
    const runs = Array.from(el.childNodes).filter((n) => n.nodeType === 3 && n.textContent.trim())
    if (runs.length !== 1) return { textRun: null, textRunIndex: -1 }
    return { textRun: runs[0].textContent, textRunIndex: textNodeIndex(el, runs[0]) }
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
  let hovered = null
  const hover = (el) => {
    place(hoverBox, el)
    if (el === hovered) return
    hovered = el
    post({ type: 'gx:hover', sid: el ? sidOf(el) : null })
  }
  const rehover = () => {
    if (!pointer || !selecting() || editing) return hover(null)
    hover(targetOf(document.elementFromPoint(pointer.x, pointer.y)))
  }
  const refresh = () => {
    rehover()
    place(selectBox, selected)
    placeLabel(selected)
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
  const beginEdit = (el) => {
    if (!el || editing || sidOf(el) === null) return
    const textNodes = Array.from(el.childNodes).filter(
      (n) => n.nodeType === 3 && n.textContent.trim(),
    )
    // only elements whose editable text lives in a single direct text node (mixed content edits go through AI or source)
    if (textNodes.length !== 1) return
    const node = textNodes[0]
    editing = { el, node, index: textNodeIndex(el, node), before: node.textContent }
    el.setAttribute(`${MARK}-editing`, '')
    el.setAttribute('contenteditable', 'plaintext-only')
    el.focus()
    const range = document.createRange()
    range.selectNodeContents(node)
    const sel = window.getSelection()
    sel.removeAllRanges()
    sel.addRange(range)
  }
  const finishEdit = (commit) => {
    if (!editing) return
    const { el, node, index, before } = editing
    editing = null
    el.removeAttribute('contenteditable')
    el.removeAttribute(`${MARK}-editing`)
    // contenteditable may have split/merged text nodes; read the element's direct text again
    const after = Array.from(el.childNodes)
      .filter((n) => n.nodeType === 3)
      .map((n) => n.textContent)
      .join('')
    if (commit && after !== before) {
      post({ type: 'gx:textEditCommit', sid: sidOf(el), textNodeIndex: index, newText: after })
    } else {
      if (node.isConnected) node.textContent = before
      post({ type: 'gx:textEditCancel' })
    }
    refresh()
  }

  // ---- events (capture phase: we see them before the page's own handlers) ----
  document.addEventListener(
    'mousemove',
    (e) => {
      pointer = { x: e.clientX, y: e.clientY }
      if (!selecting() || editing) return
      hover(targetOf(e.target))
    },
    true,
  )
  // on the root element only: a capture listener on document would fire for every descendant
  document.documentElement.addEventListener('mouseleave', () => {
    pointer = null
    hover(null)
  })
  document.addEventListener(
    'click',
    (e) => {
      if (!selecting()) return
      if (editing) {
        if (editing.el.contains(e.target)) return
        finishEdit(true)
      }
      const a = e.target && e.target.closest && e.target.closest('a[href]')
      if (a) {
        e.preventDefault()
        if (e.metaKey || e.ctrlKey) post({ type: 'gx:navigateBlocked', href: a.href })
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
    },
    true,
  )
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
  document.addEventListener(
    'selectionchange',
    () => {
      if (mode !== 'inspect' || editing) return
      const sel = window.getSelection()
      if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return
      const r = sel.getRangeAt(0)
      if (r.startContainer !== r.endContainer || r.startContainer.nodeType !== 3) return
      const node = r.startContainer
      const el = node.parentElement
      if (!el || isOwn(el) || sidOf(el) === null) return
      post({
        type: 'gx:textSelect',
        sid: sidOf(el),
        textNodeIndex: textNodeIndex(el, node),
        start: r.startOffset,
        end: r.endOffset,
        text: node.textContent.slice(r.startOffset, r.endOffset),
      })
    },
    true,
  )
  document.addEventListener(
    'keydown',
    (e) => {
      if (editing) {
        if (e.key === 'Escape') {
          e.preventDefault()
          finishEdit(false)
        } else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
          e.preventDefault()
          finishEdit(true)
        }
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
      const command = modKey || map[e.key]
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
  window.addEventListener('scroll', refresh, true)
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
      case 'gx:setMode':
        if (editing) finishEdit(true)
        mode = msg.mode
        rehover()
        if (mode === 'browse') select(null, false)
        break
      case 'gx:beginTextEdit':
        select(bySid(msg.sid), false)
        beginEdit(selected)
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
