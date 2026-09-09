  const SKIP_TAGS = new Set([
    'SCRIPT', 'STYLE', 'LINK', 'META', 'NOSCRIPT', 'TEMPLATE', 'HEAD', 'TITLE',
  ]);
  let shotCounter = 0;

  const styleCache = new WeakMap();
  function cs(el) {
    if (!styleCache.has(el)) styleCache.set(el, getComputedStyle(el));
    return styleCache.get(el);
  }

  function isVisible(el) {
    const s = cs(el);
    if (s.display === 'none' || s.visibility === 'hidden' || parseFloat(s.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 1 && r.height > 1;
  }

  function parseColor(colorStr) {
    if (!colorStr) return null;
    const m = colorStr.match(/rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)/);
    if (!m) return null;
    return {
      r: +m[1],
      g: +m[2],
      b: +m[3],
      a: m[4] === undefined ? 1 : parseFloat(m[4]),
    };
  }

  function compositeRgb(color, backdrop) {
    const a = color.a;
    return {
      r: Math.round(color.r * a + backdrop.r * (1 - a)),
      g: Math.round(color.g * a + backdrop.g * (1 - a)),
      b: Math.round(color.b * a + backdrop.b * (1 - a)),
      a: 1,
    };
  }

  function backdropRgbOf(el) {
    const ancestors = [];
    for (let cur = el?.parentElement; cur; cur = cur.parentElement) ancestors.push(cur);
    let backdrop = { r: 255, g: 255, b: 255, a: 1 };
    for (const ancestor of ancestors.reverse()) {
      let color = parseColor(cs(ancestor).backgroundColor);
      // A gradient-painted ancestor (background-image, not backgroundColor)
      // is invisible to this walk otherwise: a translucent white overlay
      // nested inside it composites against the default white backdrop and
      // reads as near-white, so paintedBgHex's white-preserve fallback then
      // hardcodes opaque FFFFFF — white text on a white cell, content
      // effectively invisible. Approximate with the first gradient stop,
      // same fallback bgHex uses for the gradient element itself.
      if (!color || color.a < 0.05) {
        const stop = gradientBaseStop(cs(ancestor).backgroundImage);
        if (stop) color = parseColor(stop);
      }
      if (color && color.a >= 0.05) backdrop = compositeRgb(color, backdrop);
    }
    return backdrop;
  }

  // 'rgb(a)' -> 'RRGGBB' hex, or null for transparent / near-white.
  // Light documents retain white-backed tints; dark documents pass their
  // resolved ancestor paint so alpha colors keep their authored contrast.
  function toHex(colorStr, { dropWhite = false, backdrop = null } = {}) {
    const color = parseColor(colorStr);
    if (!color) return null;
    let { r, g, b, a } = color;
    if (a < 0.05) return null;
    if (a < 1) {
      ({ r, g, b } = compositeRgb(color, backdrop || { r: 255, g: 255, b: 255 }));
    }
    if (dropWhite && r > 250 && g > 250 && b > 250) return null;
    const hex = (n) => n.toString(16).padStart(2, '0');
    return (hex(r) + hex(g) + hex(b)).toUpperCase();
  }

  function hexLuminance(hex) {
    if (!hex || hex.length < 6) return 0;
    const n = parseInt(hex.slice(0, 6), 16);
    if (Number.isNaN(n)) return 0;
    const r = (n >> 16) & 255;
    const g = (n >> 8) & 255;
    const b = n & 255;
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  }

  // Representative paint of a (possibly multi-layer) gradient background:
  // the first sufficiently-opaque stop of the BOTTOM layer. CSS paints the
  // first listed layer on top, so decorative low-alpha radial accents come
  // first and the base gradient last — grabbing the first color in the whole
  // string painted a dark hero with its 0.18-alpha gold accent (near-white
  // shading under white text) instead of its dark base.
  function gradientBaseStop(bgImage) {
    if (!bgImage || !bgImage.includes('gradient')) return null;
    const lastLayerStart = bgImage.lastIndexOf('-gradient(');
    const layer = lastLayerStart >= 0 ? bgImage.slice(lastLayerStart) : bgImage;
    let first = null;
    for (const match of layer.matchAll(/rgba?\([^)]+\)|#[0-9a-fA-F]{3,8}/g)) {
      first = first || match[0];
      const color = parseColor(match[0]);
      // #hex tokens don't parse -> opaque; skip translucent overlay stops
      if (!color || color.a >= 0.5) return match[0];
    }
    return first;
  }

  function bgHex(el) {
    const solid = toHex(cs(el).backgroundColor, { dropWhite: true, backdrop: backdropRgbOf(el) });
    if (solid) return solid;
    // gradient-painted box with no solid backgroundColor: approximate with
    // a representative stop, or light foreground text lands on a white page
    const stop = gradientBaseStop(cs(el).backgroundImage);
    if (stop) return toHex(stop, { dropWhite: true });
    return null;
  }

  // True when el sits on a solid tint or a CSS background-image/gradient.
  // White cards on those surfaces must keep FFFFFF or the backdrop shows through.
  function hasTintedBackdrop(el) {
    let anc = el.parentElement;
    while (anc) {
      if (bgHex(anc)) return true;
      const s = cs(anc);
      if (s.backgroundImage && s.backgroundImage !== 'none') return true;
      if (anc === document.documentElement) break;
      anc = anc.parentElement;
    }
    return false;
  }

  // Like bgHex, but keeps authored white/near-white paint when it covers a
  // tinted or gradient backdrop (page wash, section band, etc.).
  function paintedBgHex(el) {
    const shaded = bgHex(el);
    if (shaded) return shaded;
    const raw = toHex(cs(el).backgroundColor);
    if (!raw || !hasTintedBackdrop(el)) return null;
    return 'FFFFFF';
  }


  // A compact painted pill/badge (inline-block, nowrap, own background):
  // shared signal that a container's exact visual row must be screenshot.
  function isCompactPaintedLabel(child) {
    const childStyle = cs(child);
    return (
      childStyle.display === 'inline-block' &&
      childStyle.whiteSpace === 'nowrap' &&
      Boolean(bgHex(child)) &&
      (child.innerText || '').trim().length <= 30 &&
      child.getBoundingClientRect().height <= 50
    );
  }

  function isMono(s) {
    return /mono|consolas|courier|menlo|jetbrains/i.test(s.fontFamily);
  }

  const DOC_RTL =
    getComputedStyle(document.body).direction === 'rtl' ||
    getComputedStyle(document.documentElement).direction === 'rtl';

  function isRTL(el) {
    return cs(el).direction === 'rtl';
  }

  // Resolve text-align to a physical alignment, accounting for direction.
  // In an RTL block, 'start' (the browser default) means the right side —
  // Word's w:bidi paragraphs already default to right, so return undefined
  // and let the bidi default apply; an explicit left/end must survive as
  // physical left.
  function textAlignOf(el) {
    const s = cs(el);
    const ta = s.textAlign;
    if (ta === 'center') return 'center';
    if (ta.startsWith('justify')) return 'justify';
    if (s.direction === 'rtl') {
      if (ta === 'left' || ta === 'end') return 'left';
      return undefined; // start / right == bidi paragraph default
    }
    if (ta === 'right' || ta === 'end') return 'right';
    return undefined;
  }

  function borderLeftOf(el) {
    const s = cs(el);
    const w = parseFloat(s.borderLeftWidth);
    if (w >= 2 && s.borderLeftStyle !== 'none') {
      const color = toHex(s.borderLeftColor);
      if (color) return { color, widthPx: w };
    }
    // ::before painted as a narrow vertical accent bar (▎ section headings,
    // timeline rails): visually a left border, but authored as a
    // pseudo-element; gradient rails take their first color stop.
    const before = getComputedStyle(el, '::before');
    if (before && before.content === '""') {
      const bw = parseFloat(before.width);
      const bh = parseFloat(before.height);
      const gradientStop = before.backgroundImage.match(/rgba?\([^)]+\)|#[0-9a-fA-F]{3,8}/)?.[0];
      const color = toHex(before.backgroundColor) || (gradientStop && toHex(gradientStop));
      if (color && bw >= 2 && bw <= 10 && bh >= 10) {
        return { color, widthPx: bw };
      }
    }
    return null;
  }

  function borderRightOf(el) {
    const s = cs(el);
    const w = parseFloat(s.borderRightWidth);
    if (w >= 2 && s.borderRightStyle !== 'none') {
      const color = toHex(s.borderRightColor);
      if (color) return { color, widthPx: w };
    }
    return null;
  }

  // ::before/::after painted as a full-width horizontal bar (decorative
  // section rules). Not in the DOM, so probed via computed pseudo styles.
  function pseudoBar(el, which) {
    const s = getComputedStyle(el, which);
    if (!s || s.content === 'none') return null;
    if (s.position === 'fixed') return null;
    const h = parseFloat(s.height);
    if (!(h >= 1 && h <= 8)) return null;
    const w = parseFloat(s.width);
    const elW = el.getBoundingClientRect().width;
    const fullRule = w >= elW * 0.5;
    // absolute pseudos are usually dropped (not in flow), but a thin
    // full-width bar hugging the element is a decorative underline that
    // should survive (section-title::after { position:absolute; bottom:-5px })
    if (s.position === 'absolute') {
      const top = Math.abs(parseFloat(s.top));
      const bottom = Math.abs(parseFloat(s.bottom));
      const hugsEdge = Math.min(isNaN(top) ? 99 : top, isNaN(bottom) ? 99 : bottom) <= 12;
      if (!fullRule || !hugsEdge) return null;
    }
    if (s.position !== 'absolute' && s.display !== 'block') return null;
    const shortAccent = w >= 20 && w <= 120;
    if (!(w > 0) || (!fullRule && !shortAccent)) return null;
    const color = toHex(s.backgroundColor, { dropWhite: true });
    if (!color) return null;
    return {
      type: 'hr',
      color,
      heightPx: h,
      widthFrac: Math.min(1, w / elW),
      short: shortAccent && !fullRule,
      spacingBeforePx: Math.max(0, parseFloat(s.marginTop) || 0),
      spacingAfterPx: Math.max(0, parseFloat(s.marginBottom) || 0),
    };
  }

  function pseudoMarker(el) {
    for (const which of ['::before', '::after']) {
      const s = getComputedStyle(el, which);
      if (!s || s.content === 'none' || s.content === 'normal') continue;
      const content = s.content.replace(/^['"]|['"]$/g, '');
      if (content.trim()) continue;
      const width = parseFloat(s.width);
      const height = parseFloat(s.height);
      if (
        width < 5 ||
        width > 24 ||
        height < 5 ||
        height > 24 ||
        Math.abs(width - height) > 6
      ) {
        continue;
      }
      const color =
        toHex(s.backgroundColor, { dropWhite: true }) ||
        toHex(s.borderTopColor, { dropWhite: true });
      if (!color) continue;
      return {
        text: s.transform && s.transform !== 'none' ? '\u25c6 ' : '\u25a0 ',
        color,
        sizePx: Math.max(width, height),
        bold: false,
      };
    }
    return null;
  }

  // Timeline entry decorations: an absolutely positioned pseudo dot (round
  // marker) and/or a thin vertical connector rail in the left padding gutter.
  // Both are position:absolute, so the flow rules drop them; callers
  // reproduce the rail as a paragraph left border and the dot as a colored
  // disc run.
  function timelineDecorOf(el) {
    const s = cs(el);
    if (s.position !== 'relative') return null;
    const padL = parseFloat(s.paddingLeft) || 0;
    if (padL < 12 || padL > 60) return null;
    let dot = null;
    let rail = null;
    for (const which of ['::before', '::after']) {
      const p = getComputedStyle(el, which);
      if (!p || p.content === 'none' || p.content === 'normal') continue;
      if (p.position !== 'absolute') continue;
      if (p.content.replace(/^['"]|['"]$/g, '').trim()) continue;
      const w = parseFloat(p.width);
      const h = parseFloat(p.height);
      const left = parseFloat(p.left);
      if (!(left >= -2 && left < padL)) continue;
      const color = toHex(p.backgroundColor, { dropWhite: true });
      if (!color) continue;
      if (w >= 5 && w <= 24 && h >= 5 && h <= 24 && Math.abs(w - h) <= 6) {
        dot = { color, sizePx: Math.max(w, h) };
      } else if (w >= 1 && w <= 4 && h > 30) {
        rail = { color, widthPx: Math.max(1, Math.round(w)), leftPx: left };
      }
    }
    return dot || rail ? { dot, rail, padL, railLeftPx: rail?.leftPx } : null;
  }

  function borderTopOf(el) {
    const s = cs(el);
    const w = parseFloat(s.borderTopWidth);
    if (w >= 2 && s.borderTopStyle !== 'none') {
      const color = toHex(s.borderTopColor);
      if (color) return { color, widthPx: w };
    }
    return null;
  }

  function borderSideOf(el, side) {
    const s = cs(el);
    const cap = side[0].toUpperCase() + side.slice(1);
    const width = parseFloat(s[`border${cap}Width`]);
    if (width < 1 || s[`border${cap}Style`] === 'none') return null;
    const color = toHex(s[`border${cap}Color`]);
    return color ? { color, widthPx: width, style: s[`border${cap}Style`] } : null;
  }

  function bordersOf(el) {
    const borders = {};
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const border = borderSideOf(el, side);
      if (border) borders[side] = border;
    }
    return Object.keys(borders).length ? borders : null;
  }

  function cellPaddingOf(el) {
    const s = cs(el);
    return {
      top: Math.max(0, Math.round(parseFloat(s.paddingTop) || 0)),
      bottom: Math.max(0, Math.round(parseFloat(s.paddingBottom) || 0)),
      left: Math.max(0, Math.round(parseFloat(s.paddingLeft) || 0)),
      right: Math.max(0, Math.round(parseFloat(s.paddingRight) || 0)),
    };
  }

  // Background painted by an *inline* ancestor (text highlight). Block-level
  // backgrounds are handled as paragraph/cell shading, so stop at the first
  // non-inline ancestor to avoid double-painting.
  // Marker-pen highlight: linear-gradient(transparent 60%, color 40%) on an
  // inline span. Word has no partial-height run shading, so the marker color
  // becomes a full-height run highlight.
  function markerGradientHex(el) {
    const bgImage = cs(el).backgroundImage;
    if (!bgImage || !bgImage.startsWith('linear-gradient')) return null;
    const stops = [...bgImage.matchAll(/rgba?\(([\d.\s,]+)\)/g)];
    if (!stops.length) return null;
    const last = stops[stops.length - 1][1].split(',').map(parseFloat);
    const [r, g, b, a = 1] = last;
    if (!(a >= 0.15)) return null;
    const mix = (c) => Math.round(c * a + 255 * (1 - a));
    return toHex(`rgb(${mix(r)}, ${mix(g)}, ${mix(b)})`, { dropWhite: true });
  }

  function inlineBgHex(el) {
    let cur = el;
    while (cur && cur.nodeType === Node.ELEMENT_NODE) {
      const d = cs(cur).display;
      if (!(d === 'inline' || d.startsWith('inline-') || d === 'contents')) return null;
      const bg = bgHex(cur) || markerGradientHex(cur);
      if (bg) return bg;
      cur = cur.parentElement;
    }
    return null;
  }

  function runStyleOf(el) {
    const s = cs(el);
    const style = {
      bold: parseInt(s.fontWeight, 10) >= 600,
      italic: s.fontStyle === 'italic',
      underline: s.textDecorationLine.includes('underline'),
      strike: s.textDecorationLine.includes('line-through'),
      color: toHex(s.color) || '000000',
      sizePx: parseFloat(s.fontSize),
      mono: isMono(s),
      letterSpacing: parseFloat(s.letterSpacing) || 0,
      fontFamily: s.fontFamily,
    };
    // Inline(-block) span underlined via border-bottom (recipient names,
    // signature lines): text-decoration never sees it, so map the lone
    // bottom border to a run underline. Boxed chips (any other side) skip.
    if (
      !style.underline &&
      (s.display || '').startsWith('inline') &&
      s.borderBottomStyle !== 'none' &&
      parseFloat(s.borderBottomWidth) >= 1 &&
      s.borderTopStyle === 'none' &&
      s.borderLeftStyle === 'none' &&
      s.borderRightStyle === 'none'
    ) {
      style.underline = true;
      if (s.borderBottomStyle === 'dotted' || s.borderBottomStyle === 'dashed') {
        style.underlineStyle = s.borderBottomStyle;
      }
      const underlineHex = toHex(s.borderBottomColor);
      if (underlineHex) style.underlineColor = underlineHex;
    }
    // Inline boxed chip (O/X answer labels, small tags): a same-color border
    // on every side of a short inline element is a character box — map it to a
    // run border (w:bdr). Word can't round the corners, but the frame survives.
    // Restricted to short text so a bordered inline phrase isn't boxed whole.
    if (
      (s.display || '').startsWith('inline') &&
      (el.innerText || '').trim().length <= 8 &&
      parseFloat(s.borderTopWidth) >= 1 &&
      s.borderTopStyle !== 'none'
    ) {
      const bc = toHex(s.borderTopColor);
      if (
        bc &&
        toHex(s.borderRightColor) === bc &&
        toHex(s.borderBottomColor) === bc &&
        toHex(s.borderLeftColor) === bc
      ) {
        style.charBorder = { color: bc, widthPx: parseFloat(s.borderTopWidth) };
      }
    }
    const hl = inlineBgHex(el);
    if (hl) style.highlight = hl;
    if (s.verticalAlign === 'super' || el.closest('sup')) style.superscript = true;
    else if (s.verticalAlign === 'sub' || el.closest('sub')) style.subscript = true;
    if (s.textTransform === 'uppercase') style.uppercase = true;
    if (s.direction === 'rtl') style.rtl = true;
    const a = el.closest('a[href]');
    if (a) {
      const href = a.getAttribute('href') || '';
      if (href.startsWith('#') && href.length > 1) {
        try {
          style.anchor = decodeURIComponent(href.slice(1));
        } catch {
          style.anchor = href.slice(1);
        }
      } else if (href && !/^javascript:/i.test(href)) {
        style.href = a.href;
      }
    }
    return style;
  }

