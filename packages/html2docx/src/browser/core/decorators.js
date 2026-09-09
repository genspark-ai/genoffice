  // Rendered math (KaTeX / MathJax / MathML): the internal spans are layout
  // hacks — walking them yields garbage text runs ("y=ncx"). Screenshot the
  // whole formula instead.
  function isMathContainer(el) {
    const tag = el.tagName;
    if (tag === 'MATH' || tag === 'MJX-CONTAINER') return true;
    const cls = el.classList;
    return Boolean(
      cls && (cls.contains('katex') || cls.contains('katex-display') || cls.contains('MathJax')),
    );
  }

  function pseudoBullet(el) {
    const s = getComputedStyle(el, '::before');
    if (!s || s.content === 'none' || s.content === 'normal') return false;
    const content = s.content.replace(/^['"]|['"]$/g, '').trim();
    return BULLET_GLYPHS.has(content);
  }

  // Real-element twin of the painted ::before dot below: a small empty
  // painted first child (<div class="dot"></div>) is a list marker, not a
  // block — classifier would silently drop it.
  function chipBulletRun(el) {
    const first = [...el.children].find((child) => isVisible(child));
    if (!first || first.childElementCount || (first.innerText || '').trim()) return null;
    const r = first.getBoundingClientRect();
    if (r.width < 4 || r.width > 14 || r.height < 4 || r.height > 14) return null;
    const color = toHex(cs(first).backgroundColor);
    if (!color) return null;
    const round = parseFloat(cs(first).borderRadius) >= r.width / 2 - 1;
    return {
      text: `${round ? '●' : '■'} `,
      color,
      // the dot glyph inks ~60% of its em box: scale up to match the chip
      sizePx: Math.max(8, Math.round(Math.min(r.width, r.height) / 0.6)),
    };
  }

  function pseudoBulletRun(el) {
    const s = getComputedStyle(el, '::before');
    if (!s || s.content === 'none' || s.content === 'normal') return null;
    const content = s.content.replace(/^['"]|['"]$/g, '').trim();
    if (!content) {
      // Painted dot bullet: empty ::before drawn as a small colored circle
      // or square \u2014 emit an equivalent colored glyph run.
      const dotW = parseFloat(s.width);
      const dotH = parseFloat(s.height);
      const dotColor = toHex(s.backgroundColor);
      // Composite icon: ::before circle + ::after strokes (check-in-circle
      // list markers) \u2014 a flat glyph drops the checkmark, so clip-screenshot
      // the drawn icon box instead.
      const after = getComputedStyle(el, '::after');
      const afterPaints =
        after &&
        after.content !== 'none' &&
        (parseFloat(after.borderLeftWidth) >= 1 ||
          parseFloat(after.borderBottomWidth) >= 1 ||
          Boolean(toHex(after.backgroundColor)));
      if (
        dotColor &&
        afterPaints &&
        s.position === 'absolute' &&
        dotW >= 10 &&
        dotW <= 28 &&
        dotH >= 10 &&
        dotH <= 28
      ) {
        const host = el.getBoundingClientRect();
        const x = host.left + (parseFloat(s.left) || 0);
        const y = host.top + (parseFloat(s.top) || 0);
        if (x >= 0 && y >= 0) {
          const gapPx = 6;
          return {
            text: '',
            inlineImage: true,
            shotId: `h2d-${shotCounter++}`,
            clip: {
              x: Math.round(x),
              y: Math.round(y),
              width: Math.round(dotW + gapPx),
              height: Math.round(dotH),
            },
            width: Math.round(dotW + gapPx),
            height: Math.round(dotH),
          };
        }
      }
      if (dotColor && dotW >= 4 && dotW <= 14 && dotH >= 4 && dotH <= 14) {
        const round = parseFloat(s.borderRadius) >= dotW / 2 - 1;
        return {
          text: `${round ? '\u25cf' : '\u25a0'}\u00a0`,
          color: dotColor,
          // the dot glyph inks ~60% of its em box: scale up to match the chip
          sizePx: Math.max(8, Math.round(Math.min(dotW, dotH) / 0.6)),
        };
      }
      return null;
    }
    // Numbered circle badge: content:counter(x) painted as a colored circle
    // with a light number.
    if (content.includes('counter(') && el.parentElement) {
      const badgeW = parseFloat(s.width);
      // gradient badges have a transparent backgroundColor: take the first
      // gradient stop as the representative color
      const gradientStop = s.backgroundImage.match(/rgba?\([^)]+\)|#[0-9a-f]{3,8}/i)?.[0];
      const badgeColor =
        toHex(s.backgroundColor) ||
        (gradientStop && toHex(gradientStop, { dropWhite: false }));
      // Pseudo-elements have no DOM node to screenshot, but an absolutely
      // positioned badge's rendered box is computable from the host rect +
      // its own offsets/size — clip-screenshot it so size and gradient
      // survive exactly.
      const badgeH = parseFloat(s.height) || badgeW;
      if (
        badgeColor &&
        s.position === 'absolute' &&
        badgeW >= 14 &&
        badgeW <= 72 &&
        badgeH >= 14 &&
        badgeH <= 72 &&
        parseFloat(s.borderRadius) >= badgeW / 2 - 2
      ) {
        const host = el.getBoundingClientRect();
        const x = host.left + (parseFloat(s.left) || 0);
        const y = host.top + (parseFloat(s.top) || 0);
        if (x >= 0 && y >= 0) {
          const gapPx = 8; // right whitespace baked in as the text gap
          return {
            text: '',
            inlineImage: true,
            shotId: `h2d-${shotCounter++}`,
            clip: {
              x: Math.round(x),
              y: Math.round(y),
              width: Math.round(badgeW + gapPx),
              height: Math.round(badgeH),
            },
            width: Math.round(badgeW + gapPx),
            height: Math.round(badgeH),
          };
        }
      }
      // Word has no circular chips — a negative circled digit (❶❷…) in the
      // badge color reads the same.
      if (badgeColor && badgeW >= 14 && parseFloat(s.borderRadius) >= badgeW / 2 - 2) {
        const siblings = [...el.parentElement.children].filter(
          (sibling) => sibling.tagName === el.tagName && isVisible(sibling),
        );
        const index = siblings.indexOf(el) + 1;
        const glyph =
          index >= 1 && index <= 10
            ? String.fromCodePoint(0x2775 + index)
            : `${index}.`;
        return {
          text: `${glyph} `,
          color: badgeColor,
          sizePx: Math.max(10, Math.round(parseFloat(s.fontSize) * 1.15) || 16),
        };
      }
    }
    // Authored visible text prefixes ("◆ SLIDE ") count too, not only
    // single-glyph markers.
    const literalListMarker =
      content.length <= 24 &&
      !content.includes('counter(') &&
      !content.includes('attr(') &&
      !content.includes('url(');
    if (!BULLET_GLYPHS.has(content) && !literalListMarker) return null;
    return {
      text: `${content}\u00a0`,
      color: toHex(s.color) || undefined,
      sizePx: parseFloat(s.fontSize) || undefined,
      fontFamily: BULLET_GLYPHS.has(content) ? s.fontFamily || undefined : 'Arial Unicode MS',
    };
  }

  function hrNode(el) {
    const s = cs(el);
    const r = el.getBoundingClientRect();
    const body = document.body.getBoundingClientRect();
    const border = borderSideOf(el, 'top') || borderSideOf(el, 'bottom');
    return {
      type: 'hr',
      color: border?.color || bgHex(el) || 'CCCCCC',
      heightPx: border?.widthPx || Math.max(1, Math.min(r.height, 8)),
      widthFrac: body.width > 0 ? Math.min(1, r.width / body.width) : 1,
      align:
        Math.abs(r.left + r.width / 2 - (body.left + body.width / 2)) < 12
          ? 'center'
          : 'left',
    };
  }

  function thinRuleNode(el) {
    if ((el.innerText || '').trim() || el.children.length) return null;
    const r = el.getBoundingClientRect();
    const parentWidth = el.parentElement?.getBoundingClientRect().width || r.width;
    if (r.height < 1 || r.height > 8 || r.width < parentWidth * 0.5) return null;
    const border = borderSideOf(el, 'top') || borderSideOf(el, 'bottom');
    const color = border?.color || bgHex(el);
    if (!color) return null;
    return {
      type: 'hr',
      color,
      heightPx: border?.widthPx || Math.max(1, r.height),
      widthFrac: Math.min(1, r.width / document.body.getBoundingClientRect().width),
      align: 'left',
      spacingBeforePx: Math.max(0, Math.round(parseFloat(cs(el).marginTop) || 0)),
      spacingAfterPx: Math.max(0, Math.round(parseFloat(cs(el).marginBottom) || 0)),
    };
  }

  function isOverlaidInfographic(el) {
    const rect = el.getBoundingClientRect();
    if (rect.width < 200 || rect.height < 160 || rect.width > 1000 || rect.height > 1000) {
      return false;
    }
    const children = [...el.children].filter((child) => isVisible(child));
    const layoutChild = children.find((child) => {
      const display = cs(child).display;
      return display === 'grid' || display === 'flex';
    });
    if (!layoutChild) return false;
    const layoutRect = layoutChild.getBoundingClientRect();
    return children.some((child) => {
      if (child === layoutChild || cs(child).position !== 'absolute') return false;
      const childRect = child.getBoundingClientRect();
      if (childRect.width < 40 || childRect.height < 40) return false;
      return (
        childRect.left < layoutRect.right &&
        childRect.right > layoutRect.left &&
        childRect.top < layoutRect.bottom &&
        childRect.bottom > layoutRect.top
      );
    });
  }

  function needsVisualEffectScreenshot(el) {
    return globalThis.__html2docxVisualEffects.needsScreenshot(el, {
      bgHex,
      cs,
      fullBorderOf,
    });
  }

