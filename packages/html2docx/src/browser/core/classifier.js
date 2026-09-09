  function markForScreenshotSlices(el, maxSliceHeight = 700) {
    const base = markForScreenshot(el);
    const rect = el.getBoundingClientRect();
    // Decorative absolutes bleeding past the element box (half-out hero
    // circles) get flat-cut by an element-width clip: extend to overhanging
    // descendants still touching the element band, capped so parked
    // off-screen decorations (left: -10000px) can't drag the clip away.
    let paintLeft = rect.left;
    let paintRight = rect.right;
    const maxBleed = 200;
    const elOverflow = cs(el).overflow;
    if (elOverflow !== 'hidden' && elOverflow !== 'clip') {
      for (const child of el.querySelectorAll('*')) {
        const childRect = child.getBoundingClientRect();
        if (!childRect.width || !childRect.height) continue;
        if (
          childRect.right <= rect.left - maxBleed ||
          childRect.left >= rect.right + maxBleed
        ) continue;
        if (childRect.left < paintLeft) paintLeft = Math.max(childRect.left, rect.left - maxBleed);
        if (childRect.right > paintRight) paintRight = Math.min(childRect.right, rect.right + maxBleed);
      }
    }
    const clipX = Math.max(0, Math.round(paintLeft));
    const clipWidth = Math.round(paintRight) - clipX;
    const sliceCount = Math.ceil(rect.height / maxSliceHeight);
    return Array.from({ length: sliceCount }, (_, index) => {
      const sliceHeight = Math.min(maxSliceHeight, rect.height - index * maxSliceHeight);
      return {
        ...base,
        shotId: index === 0 ? base.shotId : `h2d-${shotCounter++}`,
        width: clipWidth,
        height: Math.round(sliceHeight),
        spacingBeforePx: index === 0 ? base.spacingBeforePx : 0,
        spacingAfterPx: index === sliceCount - 1 ? base.spacingAfterPx : 0,
        clip: {
          x: clipX,
          y: Math.max(0, Math.round(rect.top + index * maxSliceHeight)),
          width: clipWidth,
          height: Math.round(sliceHeight),
        },
      };
    });
  }

  let cachedForcedPageBreakCount = null;
  function forcedPageBreakCount() {
    if (cachedForcedPageBreakCount === null) {
      cachedForcedPageBreakCount = [...document.querySelectorAll('*')].filter((element) => {
        const style = cs(element);
        return style.breakBefore === 'page' || style.pageBreakBefore === 'always';
      }).length;
    }
    return cachedForcedPageBreakCount;
  }

  // Classification tracing (H2D_TRACE_SELECTOR): record what each matching
  // element became — the recurring debugging question is "which branch
  // consumed this element", answered by tag+depth+output node types without
  // hand-inserting probes.
  function processElement(el, depth) {
    const traceSelector = globalThis.__h2dTraceSelector;
    if (!traceSelector) return processElementImpl(el, depth);
    const nodes = processElementImpl(el, depth);
    try {
      if (el.matches?.(traceSelector)) {
        (globalThis.__h2dTrace ||= []).push({
          where: 'classify',
          tag: el.tagName,
          cls: String(el.className).slice(0, 60),
          depth,
          out: nodes.map((node) => `${node.type}${node.shotId ? ':shot' : ''}`).join(',') || '(dropped)',
        });
      }
    } catch (e) { /* invalid selector: trace disabled */ }
    return nodes;
  }

  function processElementImpl(el, depth) {
    if (SKIP_TAGS.has(el.tagName)) return [];
    if (!isVisible(el)) {
      // Zero-height fill-in rules (border-top signature lines) fail the
      // height check but must still emit their line.
      const thinLine = fillInLineNode(el);
      return thinLine && !thinLine.inline ? [thinLine] : [];
    }
    const out = [];
    const s = cs(el);
    const explicitHeader = el.hasAttribute('data-docx-header');
    const explicitFooter = el.hasAttribute('data-docx-footer');
    const fixedSemanticHeader =
      s.position === 'fixed' && (el.tagName === 'HEADER' || el.getAttribute('role') === 'banner');
    const fixedSemanticFooter =
      s.position === 'fixed' && (el.tagName === 'FOOTER' || el.getAttribute('role') === 'contentinfo');
    if (depth <= 1 && (explicitHeader || explicitFooter || fixedSemanticHeader || fixedSemanticFooter)) {
      return [{
        type: explicitFooter || fixedSemanticFooter ? 'footerpart' : 'headerpart',
        children: needsBlockWalk(el)
          ? processChildren(el, depth + 1, { decorations: false })
          : paraNodesFrom(el, false),
      }];
    }

    if (isMathContainer(el)) {
      // katex-display spans the full line: shoot the inner .katex tight box
      // so the image is crisp and markForScreenshot detects center alignment
      const target =
        (el.classList?.contains('katex-display') && el.querySelector(':scope > .katex')) || el;
      const r = target.getBoundingClientRect();
      return r.width >= 4 && r.height >= 6 ? [markForScreenshot(target)] : [];
    }

    if (isIconGlyph(el)) return [markForScreenshot(el)];

    const positionedTextOverlay = [...el.children].some((child) => {
      const childPosition = cs(child).position;
      return (
        (childPosition === 'absolute' || childPosition === 'fixed') &&
        (child.innerText || '').trim().length >= 4
      );
    });
    if (s.position === 'relative' && positionedTextOverlay) {
      const rect = el.getBoundingClientRect();
      if (rect.width >= 40 && rect.width <= 1000 && rect.height >= 20 && rect.height <= 600) {
        return [markForScreenshot(el)];
      }
      // Coordinate-composed page section (cover pages): every visible child
      // is absolutely positioned, so flow classification would emit nothing
      // and silently drop the whole section.
      const visibleKids = [...el.children].filter((child) => isVisible(child));
      if (
        visibleKids.length >= 2 &&
        visibleKids.every((child) => {
          const position = cs(child).position;
          return position === 'absolute' || position === 'fixed';
        }) &&
        rect.height <= window.innerHeight * 1.25 &&
        (el.innerText || '').trim().length >= 20
      ) {
        return [{ ...markForScreenshot(el), pageComposition: rect.height >= window.innerHeight * 0.7 }];
      }
    }

    // Oversized final names/signatures are decorative page furniture. Word
    // otherwise moves the entire line to a blank second page even when the
    // HTML places it at the bottom of page one.
    if (
      el.parentElement === document.body &&
      !el.nextElementSibling &&
      parseFloat(s.fontSize) > 40
    ) {
      const r = el.getBoundingClientRect();
      const body = document.body.getBoundingClientRect();
      const bs = cs(document.body);
      const shot = markForScreenshot(el);
      return [{
        ...shot,
        type: 'floatimg',
        isolate: true,
        xPx: Math.round(r.left - body.left - parseFloat(bs.paddingLeft)),
        yPx: Math.round(r.top - body.top - parseFloat(bs.paddingTop)),
        pageWpx: Math.round(body.width - parseFloat(bs.paddingLeft) - parseFloat(bs.paddingRight)),
        pageHpx: Math.round(body.height - parseFloat(bs.paddingTop) - parseFloat(bs.paddingBottom)),
        atPageBottom: true,
      }];
    }

    // absolutely positioned decorations (page art, footers/headers) don't
    // belong in flow. Visible ones on the first page become floating images
    // anchored at their rendered page position; the rest are dropped.
    // A childless, textless painted box shoved far from its flow slot by
    // relative offsets is coordinate-placed art too — in flow it interrupts
    // the surrounding text at its static position.
    const relativeDecoration =
      s.position === 'relative' &&
      !el.children.length &&
      !(el.innerText || '').trim() &&
      (s.backgroundImage !== 'none' || paintedBgHex(el)) &&
      Math.abs(parseFloat(s.top) || 0) +
        Math.abs(parseFloat(s.left) || 0) +
        Math.abs(parseFloat(s.right) || 0) +
        Math.abs(parseFloat(s.bottom) || 0) >
        40;
    if (s.position === 'absolute' || s.position === 'fixed' || relativeDecoration) {
      if (el.hasAttribute('data-h2d-skip')) return [];
      const r = el.getBoundingClientRect();
      // effective z: nearest explicit z-index up the ancestor chain (the
      // stacking context often lives on a wrapper, not the painted box)
      const effectiveZ = (node) => {
        let cur = node;
        while (cur && cur !== document.body) {
          const z = Number.parseInt(cs(cur).zIndex, 10);
          if (Number.isFinite(z)) return z;
          cur = cur.parentElement;
        }
        return null;
      };
      const zIndex = effectiveZ(el) ?? 0;
      const coveredByPaint = [...document.querySelectorAll('*')].some((candidate) => {
        if (
          candidate === el ||
          candidate.contains(el) ||
          el.contains(candidate) ||
          !isVisible(candidate)
        ) {
          return false;
        }
        // gradient/image fills are paint too (gold banner boards etc.)
        const candidateStyle = cs(candidate);
        if (!paintedBgHex(candidate) && candidateStyle.backgroundImage === 'none') return false;
        // no explicit z counts as 0: static paint still covers negative-z
        // decorations (z:-1 backdrop art behind a painted card)
        const candidateZ = effectiveZ(candidate) ?? 0;
        if (candidateZ <= zIndex) return false;
        const cr = candidate.getBoundingClientRect();
        const overlapWidth = Math.max(0, Math.min(r.right, cr.right) - Math.max(r.left, cr.left));
        const overlapHeight = Math.max(0, Math.min(r.bottom, cr.bottom) - Math.max(r.top, cr.top));
        return overlapWidth * overlapHeight >= r.width * r.height * 0.8;
      });
      if (coveredByPaint) return [];
      const PAGE_H = 1123; // A4 @96dpi
      const thinBar = r.height <= 40 && r.width > 500; // full-width rule/footer strip
      const smallDecoration = r.width <= 140 && r.height <= 140;
      // small art at the very bottom of a roughly one-page document (poster
      // corner triangles): keep it, pinned to the page bottom
      const docBottom = Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);
      const bottomDecoration =
        smallDecoration && docBottom <= PAGE_H * 1.45 && r.top + r.height >= docBottom - 80;
      // decoration positioned inside an ancestor box (banner leaf icons,
      // section watermark numerals): page-anchoring drifts when the ancestor
      // lands elsewhere in Word, so also record coordinates relative to the
      // containing block. Such decorations are kept on ANY page — the card
      // that carries them anchors them locally; leftovers without a card
      // context are dropped by the generator instead.
      let inParent = null;
      // not offsetParent: SVG elements (a leaf icon) don't have it
      let anchorParent = null;
      for (let cur = el.parentElement; cur && cur !== document.body; cur = cur.parentElement) {
        if (cs(cur).position !== 'static') {
          anchorParent = cur;
          break;
        }
      }
      if (anchorParent && anchorParent !== document.documentElement) {
        const pr = anchorParent.getBoundingClientRect();
        if (
          pr.width > 0 && pr.height > 0 &&
          r.left >= pr.left - 8 && r.right <= pr.right + 8 &&
          r.top >= pr.top - 8 && r.bottom <= pr.bottom + 8
        ) {
          inParent = {
            xPx: Math.round(r.left - pr.left),
            yPx: Math.round(r.top - pr.top),
            parentWpx: Math.round(pr.width),
            parentHpx: Math.round(pr.height),
          };
        }
      }
      // In-card overlays draw in FRONT of text (behind-text images vanish
      // under cell shading). An under-text decoration (z<=0) that is neither
      // a see-through ghost nor clear of the parent's text would cover words
      // — those keep the legacy page-anchor path instead.
      if (inParent && zIndex <= 0) {
        const elOpacity = parseFloat(s.opacity);
        const ghost = Number.isFinite(elOpacity) && elOpacity <= 0.35;
        if (!ghost) {
          // glyph-level rects, not element boxes: a centered heading's block
          // box spans the full banner width without its TEXT reaching the
          // decoration
          const overlapsText = [...anchorParent.querySelectorAll('*')].some((cand) => {
            if (cand === el || el.contains(cand) || cand.contains(el) || !isVisible(cand)) {
              return false;
            }
            for (const child of cand.childNodes) {
              if (child.nodeType !== Node.TEXT_NODE || !child.textContent.trim()) continue;
              const range = document.createRange();
              range.selectNodeContents(child);
              for (const cr of range.getClientRects()) {
                const ow = Math.max(0, Math.min(r.right, cr.right) - Math.max(r.left, cr.left));
                const oh = Math.max(0, Math.min(r.bottom, cr.bottom) - Math.max(r.top, cr.top));
                if (ow * oh > r.width * r.height * 0.15) return true;
              }
            }
            return false;
          });
          if (overlapsText) inParent = null;
        }
      }
      if (
        (depth <= 1 || smallDecoration || thinBar || inParent) &&
        r.width >= 8 &&
        r.height >= 8 &&
        (r.width <= 500 || thinBar) &&
        r.height <= 500 &&
        r.top >= 0 &&
        (r.top + r.height <= PAGE_H + 40 || bottomDecoration || inParent)
      ) {
        // coordinates relative to the body's content box, so the generator can
        // anchor to the DOCX margin origin (page paddings differ)
        const body = document.body.getBoundingClientRect();
        const bs = cs(document.body);
        const originX = body.left + parseFloat(bs.paddingLeft);
        const originY = body.top + parseFloat(bs.paddingTop);
        const shot = markForScreenshot(el);
        return [{
          ...shot,
          type: 'floatimg',
          isolate: true,
          xPx: Math.round(r.left - originX),
          yPx: Math.round(r.top - originY),
          // content-box size, so the generator can rescale coordinates to the
          // (usually narrower) DOCX content area
          pageWpx: Math.round(body.width - parseFloat(bs.paddingLeft) - parseFloat(bs.paddingRight)),
          pageHpx: Math.round(
            Math.min(PAGE_H, body.height) - parseFloat(bs.paddingTop) - parseFloat(bs.paddingBottom),
          ),
          // pinned to the page bottom (footer bars): anchor there in DOCX too.
          // Measured against the page/document end, not body.bottom — a short
          // body (~90px) otherwise turns every top decoration into a "footer".
          atPageBottom:
            bottomDecoration ||
            r.top + r.height >= Math.min(PAGE_H, Math.max(body.bottom, docBottom)) - 8,
          // z <= 0 paints under the content layer in HTML (section pills,
          // corner art) — in front it would cover Word text instead
          behindText: zIndex <= 0,
          ...(inParent ? { inParent } : {}),
        }];
      }
      return [];
    }

    // Page-break semantics. A handful of authored breaks are deliberate
    // chapter starts and are honored. Templates that spray break-before on
    // EVERY section are format noise: honoring them yields a train of
    // mostly-empty pages (a 380px section alone, stranded section tails),
    // so noisy documents flow naturally instead — keep chains already hold
    // each heading to its body.
    if (
      (s.breakBefore === 'page' || s.pageBreakBefore === 'always') &&
      forcedPageBreakCount() <= 3
    ) {
      out.push({ type: 'pagebreak' });
    }

    // large top padding/margin (cover pages, section spacing) -> vertical spacer;
    // threshold 80px excludes ordinary page-frame padding, catches cover-style gaps
    const topGap = parseFloat(s.paddingTop) + parseFloat(s.marginTop);
    if (topGap > 80 && depth === 0) {
      out.push({ type: 'spacer', px: Math.min(topGap, 260) });
    }

    const tag = el.tagName;
    const headingMatch = tag.match(/^H([1-6])$/);
    const rect = el.getBoundingClientRect();
    // Image mosaic (grid/flex of imgs, often with row/col spans): rasterize
    // the collage whole — native processing stacks the photos vertically and
    // destroys the composition.
    {
      const visibleChildren = [...el.children].filter((child) => isVisible(child));
      const directImages = visibleChildren.filter((child) => child.tagName === 'IMG');
      if (
        (s.display === 'grid' || s.display === 'flex') &&
        directImages.length >= 3 &&
        directImages.length === visibleChildren.length &&
        rect.height <= 700 &&
        rect.height >= 80
      ) {
        return out.concat([markForScreenshot(el)]);
      }
    }
    // An empty, unpainted, borderless block with explicit height is
    // intentional whitespace (e.g. a hand-signature gap between "Sincerely,"
    // and the name) — preserve the height instead of dropping an empty
    // paragraph. Replaced elements (broken img alt fallback) and bordered
    // blanks (underline fill-in fields) have their own handling.
    if (
      !/^(IMG|SVG|CANVAS|VIDEO|IFRAME|INPUT|TEXTAREA|SELECT|BUTTON|HR)$/.test(tag) &&
      !(el.innerText || '').trim() &&
      el.children.length === 0 &&
      rect.height >= 16 &&
      !bgHex(el) &&
      !paintedBgHex(el) &&
      s.backgroundImage === 'none' &&
      !(parseFloat(s.borderTopWidth) > 0) &&
      !(parseFloat(s.borderBottomWidth) > 0) &&
      !(parseFloat(s.borderLeftWidth) > 0) &&
      !(parseFloat(s.borderRightWidth) > 0) &&
      s.boxShadow === 'none'
    ) {
      const beforePseudo = getComputedStyle(el, '::before');
      const beforeText = (beforePseudo?.content || '').replace(/^['"]|['"]$/g, '').trim();
      if (!beforeText || beforeText === 'none' || beforeText === 'normal') {
        // bindsNeighbors: an authored in-flow gap (signature space) belongs
        // WITH its neighbors — "Sincerely," + gap + name must not split
        // across a page boundary.
        return out.concat([
          { type: 'spacer', px: Math.min(Math.round(rect.height), 400), bindsNeighbors: true },
        ]);
      }
    }
    if (
      tag === 'SECTION' &&
      rect.height > 1000 &&
      // a self-painted section must stay a card: flattening drops the
      // backdrop (dark paint, or a white sheet on a tinted page) and
      // strands its text on the page background
      !paintedBgHex(el) &&
      el.querySelector('h1, h2, h3') &&
      el.querySelector('svg, canvas')
    ) {
      const paddingTopPx = Math.max(0, parseFloat(s.paddingTop) || 0);
      const paddingBottomPx = Math.max(0, parseFloat(s.paddingBottom) || 0);
      return out.concat(
        paddingTopPx > 2 ? [{ type: 'spacer', px: paddingTopPx }] : [],
        processChildren(el, depth + 1),
        paddingBottomPx > 2 ? [{ type: 'spacer', px: paddingBottomPx }] : [],
      );
    }
    const beforeStyle = getComputedStyle(el, '::before');
    const beforeContent = (beforeStyle?.content || '').replace(/^['"]|['"]$/g, '').trim();
    if (
      !(el.innerText || '').trim() &&
      el.children.length === 0 &&
      beforeContent &&
      beforeContent !== 'none' &&
      beforeContent !== 'normal' &&
      rect.width >= 4 &&
      rect.width <= 64 &&
      rect.height >= 4 &&
      rect.height <= 64
    ) {
      return out.concat([markForScreenshot(el)]);
    }

    const hasFloatedLabel = [...el.querySelectorAll('*')].some((child) => {
      const float = cs(child).cssFloat;
      return float === 'left' || float === 'right';
    });
    const compactPaintedTracks = [...el.querySelectorAll('*')].filter((child) => {
      const childRect = child.getBoundingClientRect();
      return (
        !(child.innerText || '').trim() &&
        childRect.width >= 40 &&
        childRect.height >= 3 &&
        childRect.height <= 20 &&
        Boolean(bgHex(child) || paintedBgHex(child))
      );
    });
    if (
      hasFloatedLabel &&
      compactPaintedTracks.length >= 2 &&
      rect.width <= 400 &&
      rect.height <= 400
    ) {
      return out.concat([markForScreenshot(el)]);
    }

    if (tag === 'DETAILS' && !el.open) {
      const summary = [...el.children].find((child) => child.tagName === 'SUMMARY' && isVisible(child));
      if (!summary) return out;
      const border = fullBorderOf(el);
      const nodes = paraNodesFrom(summary, false);
      for (const node of nodes) {
        node.style = {
          ...node.style,
          shading: paintedBgHex(el),
          borderTop: border || undefined,
          borderBottom: border || undefined,
          borderLeft: border || undefined,
          borderRight: border || undefined,
          indentLeftPx: Math.max(
            node.style?.indentLeftPx || 0,
            Math.round(parseFloat(s.paddingLeft) || 0),
          ),
          spacingBeforePx: Math.max(0, Math.round(parseFloat(s.marginTop) || 0)),
          spacingAfterPx: Math.max(0, Math.round(parseFloat(s.marginBottom) || 0)),
        };
      }
      return out.concat(nodes);
    }

    const pseudoLabel = tag !== 'LI' ? pseudoBulletRun(el) : null;
    if (pseudoLabel && bgHex(el)) {
      const rect = el.getBoundingClientRect();
      if (rect.width <= 500 && rect.height <= 60) {
        return out.concat([markForScreenshot(el)]);
      }
    }

    if (headingMatch) {
      const rect = el.getBoundingClientRect();
      const backgroundClip =
        s.backgroundClip ||
        s.webkitBackgroundClip ||
        s.getPropertyValue('background-clip') ||
        s.getPropertyValue('-webkit-background-clip');
      const clippedGradientText =
        s.backgroundImage.includes('gradient') &&
        backgroundClip === 'text';
      if (clippedGradientText && rect.height <= 240) {
        return out.concat([markForScreenshot(el)]);
      }
      if (
        bgHex(el) &&
        rect.height <= 120 &&
        parseFloat(s.borderTopLeftRadius) >= 4
      ) {
        return out.concat([markForScreenshot(el)]);
      }
      const runs = collectRuns(el);
      const marker = pseudoMarker(el);
      if (marker) runs.unshift(marker);
      if (runs.length) {
        const headingBefore = pseudoBar(el, '::before');
        const headingAfter = pseudoBar(el, '::after');
        const style = paraStyleOf(el);
        if (headingAfter) {
          // the bar is a decorative underline: keep it tight under the text
          // and move the heading's own bottom margin below the bar. The
          // explicit small after-spacing also suppresses the built-in
          // Heading style's default (larger) spacing.
          headingAfter.spacingBeforePx = 0;
          headingAfter.spacingAfterPx =
            (headingAfter.spacingAfterPx || 0) + Math.max(0, (style.spacingAfterPx || 0) - 4);
          style.spacingAfterPx = 4;
        }
        if (headingBefore) out.push(headingBefore);
        out.push({ type: 'heading', level: +headingMatch[1], runs, style });
        if (headingAfter) out.push(headingAfter);
      } else {
        // empty heading (author-inserted blank line) still occupies height
        const r = el.getBoundingClientRect();
        const gap = r.height + parseFloat(s.marginTop) + parseFloat(s.marginBottom);
        if (gap > 16) out.push({ type: 'spacer', px: Math.min(gap, 260) });
      }
      return out;
    }
    if (tag === 'TABLE') {
      const tableNodes = extractTable(el);
      const cap = el.caption;
      if (cap && isVisible(cap) && tableNodes.length) {
        const capRuns = collectRuns(cap);
        if (capRuns.length) {
          const capStyle = paraStyleOf(cap);
          const capPara = { type: 'para', runs: capRuns, style: capStyle };
          if (cs(cap).captionSide === 'bottom') return out.concat(tableNodes, [capPara]);
          capStyle.keepNext = true;
          return out.concat([capPara], tableNodes);
        }
      }
      return out.concat(tableNodes);
    }
    if (tag === 'UL') return out.concat(extractList(el, false, depth));
    if (tag === 'OL') return out.concat(extractList(el, true, depth));
    if (tag === 'PRE') return out.concat(extractCode(el));
    if (tag === 'HR') return out.concat([hrNode(el)]);
    if (tag === 'IMG') {
      if (!el.complete || el.naturalWidth === 0) {
        const rect = el.getBoundingClientRect();
        if (
          rect.width >= 80 &&
          rect.height >= 60 &&
          (bgHex(el) || paintedBgHex(el))
        ) {
          return out.concat([markForScreenshot(el)]);
        }
        const alt = (el.getAttribute('alt') || '').trim();
        if (!alt) return out;
        const parent = el.parentElement || el;
        return out.concat([{
          type: 'para',
          runs: [{ text: alt, ...runStyleOf(parent) }],
          style: paraStyleOf(parent),
        }]);
      }
      return out.concat([markForScreenshot(el)]);
    }
    if (tag === 'TEXTAREA' || tag === 'SELECT') {
      return out.concat([formFieldNode(el)]);
    }
    if (tag === 'IFRAME' || tag === 'VIDEO' || tag === 'AUDIO') {
      if (tag === 'VIDEO' || tag === 'AUDIO') {
        try {
          el.pause();
        } catch {
          // Some custom media elements reject scripted playback control.
        }
      }
      return out.concat([embeddedMediaNode(el)]);
    }
    if (tag === 'CANVAS' || tag === 'SVG' || tag === 'svg') {
      if (rect.height > 700) {
        return out.concat(markForScreenshotSlices(el));
      }
      return out.concat([markForScreenshot(el)]);
    }
    // Chart-library mount node (ECharts/Chart.js): the canvas/svg sits inside
    // nested wrappers, often absolutely positioned - walking children would
    // drop it via the absolute-decoration rule. Capture the mount as one image.
    {
      const chartKid = el.querySelector('canvas, svg');
      if (chartKid) {
        const r = el.getBoundingClientRect();
        const cr = chartKid.getBoundingClientRect();
        if (
          r.height > 40 &&
          cr.width >= r.width * 0.7 &&
          cr.height >= r.height * 0.5 &&
          isVisible(chartKid)
        ) {
          if (r.height > 700) {
            return out.concat(markForScreenshotSlices(el));
          }
          return out.concat([markForScreenshot(el)]);
        }
      }
    }
    if (isCheckboxLike(el)) {
      // flex/margin centering of the box inside its parent -> paragraph align
      const r = el.getBoundingClientRect();
      const pr = el.parentElement ? el.parentElement.getBoundingClientRect() : r;
      const offL = r.left - pr.left;
      const offR = pr.right - r.right;
      const style = {};
      if (offL > 8 && offR > 8 && Math.abs(offL - offR) < Math.max(8, offL * 0.4)) {
        style.align = 'center';
      } else if (offL > 12 && offL > offR * 3) {
        style.align = 'right';
      }
      return out.concat([{ ...formFieldNode(el), style }]);
    }
    if (tag === 'INPUT') {
      return out.concat([formFieldNode(el)]);
    }
    {
      const fillIn = fillInLineNode(el);
      if (fillIn && !fillIn.inline) return out.concat([fillIn]);
      const fillBox = fillInBoxNode(el);
      if (fillBox) return out.concat([fillBox]);
    }
    const visibleElementChildren = [...el.children].filter(
      (child) => !SKIP_TAGS.has(child.tagName) && isVisible(child),
    );
    const inlineIconTextRow =
      s.display === 'flex' &&
      (s.flexDirection === 'row' || s.flexDirection === '') &&
      visibleElementChildren.length === 1 &&
      isReplacedTag(visibleElementChildren[0]) &&
      visibleElementChildren[0].getBoundingClientRect().width <= 64 &&
      [...el.childNodes].some(
        (node) => node.nodeType === Node.TEXT_NODE && node.textContent.trim(),
      );
    if (inlineIconTextRow) {
      return out.concat(paraNodesFrom(el));
    }
    if (needsVisualEffectScreenshot(el)) {
      if (rect.height > 1000 && rect.width >= 400) {
        return out.concat(markForScreenshotSlices(el));
      }
      return out.concat([markForScreenshot(el)]);
    }
    const thinRule = thinRuleNode(el);
    if (thinRule) return out.concat([thinRule]);
    if (isPill(el) || isBadgeGrid(el) || isDecorativeLeaf(el)) {
      return out.concat([markForScreenshot(el)]);
    }
    // gradient-painted container: OOXML shading is flat color only, and the
    // text inside is usually styled for the gradient (white on purple) - the
    // whole block must be captured as one image. Tall gradient heroes (cover
    // banners with KPI strips) also qualify: flowing their light-on-gradient
    // text onto a white page makes it unreadable.
    if (s.backgroundImage.includes('gradient') && rect.height <= 1000) {
      // Text-heavy gradient cards (class schedules, content panels) lose
      // all editability as pixels — let them fall through to the card path,
      // which approximates the gradient with its first color stop and keeps
      // the text native. Text-light decorative fills (fee chips, banners)
      // stay screenshots for fidelity.
      const gradText = (el.innerText || '').trim();
      const gradLines = gradText ? gradText.split('\n').filter((line) => line.trim()).length : 0;
      const textHeavy = gradLines >= 4 && gradText.length >= 30 && Boolean(bgHex(el));
      if (!textHeavy) {
        if (rect.height > 700) {
          return out.concat(markForScreenshotSlices(el));
        }
        return out.concat([markForScreenshot(el)]);
      }
    }
    // block-level element rendered entirely in monospace -> code block
    if (isMono(s) && hasBlockChildren(el) === false && el.innerText.includes('\n')) {
      return out.concat(extractCode(el));
    }

    if (isOverlaidInfographic(el)) {
      // A tall section often contains one bounded infographic plus native
      // headings. Rasterizing the parent hides those headings when Word has
      // to paginate the oversized image. Let its children classify instead.
      if (rect.height > 1000) {
        // Continue through the normal container flow below.
      } else if (rect.height > 700) {
        return out.concat(markForScreenshotSlices(el));
      } else {
        return out.concat([markForScreenshot(el)]);
      }
    }

    const gridTable = gridTableRows(el);
    if (gridTable) {
      return out.concat(extractGridTable(el, gridTable, depth));
    }

    const pageColumns = pageColumnParts(el);
    if (pageColumns) {
      const leading = pageColumns.leading.flatMap((child) => processElement(child, depth + 1));
      return out.concat(leading, extractPageColumns(el, depth, pageColumns.columns));
    }

    const cssColumns = depth < 4 ? cssMultiColumnParts(el) : null;
    if (cssColumns) {
      const columnRows = extractCssColumns(el, depth, cssColumns.colCount, cssColumns.colGapPx);
      if (columnRows) return out.concat(columnRows);
    }

    const largeTableGrid = depth < 4 ? largeTableGridRows(el) : null;
    if (largeTableGrid) {
      return out.concat(
        extractGridRows(el, largeTableGrid, depth, {
          screenshotCells: true,
          noSplit: false,
          largeTableGrid: true,
        }),
      );
    }

    const compactWrappedDiagram =
      s.display === 'flex' &&
      s.flexWrap === 'wrap' &&
      el.children.length > 6 &&
      el.getBoundingClientRect().height <= 420 &&
      (el.innerText || '').trim().length <= 500;
    if (compactWrappedDiagram) {
      return out.concat([markForScreenshot(el)]);
    }

    const gridRows = depth < 4 ? gridRowClusters(el) : null;
    if (gridRows) {
      return out.concat(extractGridRows(el, gridRows, depth));
    }

    const boxRow = isBoxRow(el);
    if (boxRow) {
      const rowChildren = [...el.children].filter(
        (child) => !SKIP_TAGS.has(child.tagName) && isVisible(child),
      );
      const compactParallelTextBlocks =
        rowChildren.length >= 2 &&
        rowChildren.length <= 3 &&
        el.getBoundingClientRect().height <= 180 &&
        (parseFloat(s.columnGap) || parseFloat(s.gap) || 0) >= 20 &&
        rowChildren.every(
          (child) =>
            child.children.length >= 2 &&
            (child.innerText || '').trim().length <= 180 &&
            !paintedBgHex(child) &&
            !fullBorderOf(child),
        );
      if (compactParallelTextBlocks) {
        return out.concat([markForScreenshot(el)]);
      }
      // physical left-to-right order regardless of DOM/flex direction (RTL
      // rows arrive in right-to-left DOM order; the geometry math below
      // assumes ascending left edges)
      const kids = [...el.children]
        .filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c))
        .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
      // measured geometry: column boundaries at gap midpoints, cell alignment
      // from where the child actually sits inside its column (reproduces
      // justify-content: space-between / center layouts)
      const pr = el.getBoundingClientRect();
      const rects = kids.map((k) => k.getBoundingClientRect());
      const bounds = [pr.left];
      for (let i = 0; i < rects.length - 1; i++) {
        bounds.push((rects[i].right + rects[i + 1].left) / 2);
      }
      bounds.push(pr.right);
      const colWidths = rects.map((r, i) => Math.max(10, Math.round(bounds[i + 1] - bounds[i])));
      const alignOf = (i) => {
        const colW = bounds[i + 1] - bounds[i];
        const offL = rects[i].left - bounds[i];
        const offR = bounds[i + 1] - rects[i].right;
        // child fills its column (flex: 1 etc, or narrower only by the
        // split flex gap): box edges carry no alignment intent, leave inner
        // text-align in charge
        if (rects[i].width >= colW - 16 || offL + offR <= 28) return 'left';
        if (offL <= 8 || offL <= offR / 3) return 'left';
        if (offR <= 8 || offR <= offL / 3) return 'right';
        return 'center';
      };
      const applyAlign = (nodes, align) => {
        if (align === 'left') return nodes;
        for (const n of nodes) {
          if (n.type === 'para' || n.type === 'heading') {
            if (!n.style) n.style = {};
            if (!n.style.align) n.style.align = align;
          } else if (n.type === 'image') {
            n.align = align;
          }
        }
        return nodes;
      };
      const containerTop = borderTopOf(el);
      // Keep white row paint over tinted/gradient backdrops.
      let rowBg = paintedBgHex(el);
      // uniform box border (bordered form fields / cards) -> full cell border
      const cellBorderOf = (c) => {
        const csc = cs(c);
        const w = parseFloat(csc.borderTopWidth);
        if (w < 1 || csc.borderTopStyle === 'none') return null;
        if (parseFloat(csc.borderLeftWidth) < 1 || parseFloat(csc.borderBottomWidth) < 1) return null;
        const color = toHex(csc.borderTopColor);
        return color ? { color, widthPx: w } : null;
      };
      const cells = kids.map((c, i) => {
        const align = alignOf(i);
        // screenshot cells already contain their own paint - no cell shading
        // from the child itself, but the container bar color must still fill
        // the cell area around the image
        if (
          isReplacedTag(c) ||
          isCheckboxLike(c) ||
          isPill(c) ||
          isDecorativeLeaf(c) ||
          isBadgeGrid(c) ||
          isIconGlyph(c) ||
          needsVisualEffectScreenshot(c)
        ) {
          return {
            children: applyAlign(processElement(c, depth + 1), align),
            shading: rowBg,
            borderTop: borderTopOf(c) || containerTop,
          };
        }
        // A bordered/shaded child inside a layout row (for example a Hero
        // side panel) owns its visual height. Keep it as a nested card instead
        // of stretching its paint across the full outer table cell.
        if (isCard(c, depth + 1)) {
          return {
            children: applyAlign(processElement(c, depth + 1), align),
            borderTop: containerTop,
          };
        }
        // nested flex row (logo + name): needs the full classifier
        if (isBoxRow(c)) {
          return {
            children: applyAlign(processElement(c, depth + 1), align),
            // fall back to the container paint: a transparent cell inside a
            // shaded row must not punch a white hole in the bar
            shading: bgHex(c) || rowBg,
            borderTop: borderTopOf(c) || containerTop,
          };
        }
        // An empty underline blank (form-input div) as a direct row cell:
        // the block walk below finds no text and emits an empty cell.
        {
          const cellFillIn = fillInLineNode(c);
          if (cellFillIn && !cellFillIn.inline) {
            return {
              children: [cellFillIn],
              shading: bgHex(c) || rowBg,
              borderTop: borderTopOf(c) || containerTop,
            };
          }
        }
        return {
          children: applyAlign(
            needsBlockWalk(c)
              ? processChildren(c, depth + 1, { decorations: false })
              : paraNodesFrom(c, false),
            align,
          ),
          shading: bgHex(c) || rowBg,
          border: cellBorderOf(c),
          borderTop: borderTopOf(c) || containerTop,
        };
      });
      // align-self: center children (arrows between process steps) center in
      // their cell even when the row itself stretches.
      kids.forEach((c, i) => {
        if (cells[i] && cs(c).alignSelf === 'center') cells[i].vAlign = 'center';
      });
      const gapWidths = rects
        .slice(0, -1)
        .map((rect, i) => Math.max(0, Math.round(rects[i + 1].left - rect.right)));
      const outerGapWidths = [
        Math.max(0, Math.round(rects[0].left - pr.left)),
        Math.max(0, Math.round(pr.right - rects[rects.length - 1].right)),
      ];
      const explicitGaps =
        gapWidths.some((gap) => gap >= 2) || outerGapWidths.some((gap) => gap >= 2);
      const separateCells =
        cells.length > 1 &&
        cells.every((cell) => Boolean(cell.border)) &&
        gapWidths.every((gap) => gap >= 4);
      if (separateCells) {
        for (const cell of cells) {
          for (const child of cell.children || []) {
            if (child.type === 'para' || child.type === 'heading') {
              child.style = { ...child.style, align: 'center' };
            }
          }
        }
      }
      const hasNestedCard = cells.some((cell) =>
        (cell.children || []).some((child) => child.type === 'card'),
      );
      out.push({
        type: 'kpirow',
        cells,
        colWidths,
        // Short labels ("Q10.", "Delivery") wrap mid-word when Word's font
        // metrics run a hair wider than the browser's; give tight cells a
        // little slack so single words never split. CJK glyphs run wider in
        // Word than in Chrome, so scale the slack with the CJK char count.
        itemWidths: rects.map((rect, i) => {
          const text = (kids[i].innerText || '').trim();
          let bonus = 0;
          if (text && !text.includes('\n')) {
            const cjkCount = (text.match(/[\u3000-\u9fff\uf900-\ufaff\uff00-\uffef]/g) || []).length;
            // Word's font metrics run wider than Chrome's; without slack,
            // short labels shed their last character onto a new line
            // ("2023.08" -> "2023.0 / 8", "RESUME" -> "RESUM / E").
            if (text.length <= 20) bonus = Math.max(10, Math.round(rect.width * 0.05));
            if (/^\S+[-–]\S+$/.test(text)) {
              bonus = Math.max(bonus, 15 + Math.round(rect.width * 0.25));
            }
            if (cjkCount >= 2) {
              // CJK glyph width delta grows with font size (28px headings
              // need far more slack than 12px labels)
              const fsz = parseFloat(cs(kids[i]).fontSize) || 14;
              bonus = Math.max(bonus, Math.min(90, Math.round(cjkCount * fsz * 0.25)));
            }
            if (cs(kids[i]).whiteSpace === 'nowrap') {
              // author explicitly forbids wrapping; the docx library exposes
              // no <w:noWrap/>, so buy enough slack for Word's wider metrics
              bonus = Math.max(bonus, 12 + Math.round(rect.width * 0.15));
            }
            // Emails / URLs / phones / long contact lines: Word wraps the last
            // glyph ("…email.co / m", "(555) 456-789 / 0") when the cell is
            // sized to Chrome's metrics; a wrapped footer bar can also get
            // orphaned onto a blank page.
            const hasEmailOrUrl = /@|https?:\/\/|\.(com|net|org|io|email)\b/i.test(text);
            const hasPhone = /\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/.test(text);
            const longContact = text.length > 20 && text.length <= 64;
            if (hasEmailOrUrl || hasPhone || longContact) {
              bonus = Math.max(
                bonus,
                18 + Math.round(Math.min(rect.width, 280) * 0.14),
              );
            }
          }
          return Math.round(rect.width) + bonus;
        }),
        gapWidths,
        outerGapWidths,
        explicitGaps: explicitGaps || undefined,
        gapShading: rowBg,
        // row-level paint (accent bar / full border on the row element):
        // deep rows re-materialize as cards from these
        rowBorderLeft: borderLeftOf(el) || undefined,
        rowBorder: cellBorderOf(el) || undefined,
        rowShading: rowBg || undefined,
        separateCells: separateCells || undefined,
        compact: tag === 'NAV' || undefined,
        // Navigation labels may wrap slightly differently in Word; an exact
        // browser row height would clip them. Let Word grow this compact row.
        heightPx:
          tag === 'NAV' || hasNestedCard
            ? undefined
            : Math.round(el.getBoundingClientRect().height),
        padPx: cellPaddingOf(el),
        vAlign: separateCells || s.alignItems === 'center' ? 'center' : undefined,
        spacingBeforePx: Math.max(0, Math.round(parseFloat(s.marginTop) || 0)),
        spacingAfterPx: Math.max(0, Math.round(parseFloat(s.marginBottom) || 0)),
      });
      return out;
    }

    if (isCard(el, depth)) {
      const accent = leftAccentOf(el);
      // mark before walking children so the accent isn't picked up as a float
      if (accent) accent.el.setAttribute('data-h2d-skip', '1');
      const children = needsBlockWalk(el)
        ? processChildren(el, depth + 1, { decorations: false })
        : paraNodesFrom(el, false);
      const before = pseudoBar(el, '::before');
      const after = pseudoBar(el, '::after');
      if (before) children.unshift(before);
      if (after) children.push(after);
      if (accent) {
        // painted strip on the card's left edge -> narrow shaded cell
        const w = el.getBoundingClientRect().width;
        for (const child of children) {
          if (child.style?.indentLeftPx) {
            child.style.indentLeftPx = Math.max(0, child.style.indentLeftPx - accent.widthPx);
          }
        }
        out.push({
          type: 'kpirow',
          cells: [
            { children: [], shading: accent.color },
            { children, shading: paintedBgHex(el) },
          ],
          colWidths: [accent.widthPx, Math.max(10, Math.round(w - accent.widthPx))],
          heightPx: Math.round(el.getBoundingClientRect().height),
          padPx: cellPaddingOf(el),
          spacingBeforePx: Math.max(0, Math.round(parseFloat(s.marginTop) || 0)),
          spacingAfterPx: Math.max(0, Math.round(parseFloat(s.marginBottom) || 0)),
        });
        return out;
      }
      out.push({
        type: 'card',
        children,
        shading: paintedBgHex(el),
        border: fullBorderOf(el),
        borderLeft: borderLeftOf(el),
        borderRight: borderRightOf(el),
        borderTop: borderTopOf(el),
        radiusPx: Math.max(0, Math.round(parseFloat(s.borderTopLeftRadius) || 0)),
        widthPx:
          s.display === 'inline-block'
            ? Math.round(el.getBoundingClientRect().width)
            : undefined,
        // Keep-whole is a page-space budget, not an absolute: >80% of a page
        // can never be protected without stranding most of a page, so it
        // always splits (even over an authored break-inside:avoid); >35%
        // follows print semantics and splits unless the author opted out.
        allowSplit:
          el.getBoundingClientRect().height > 1123 * 0.8 ||
          (el.getBoundingClientRect().height > 1123 * 0.35 && cs(el).breakInside !== 'avoid') ||
          undefined,
        heightPx: Math.round(el.getBoundingClientRect().height),
        padPx: cellPaddingOf(el),
        spacingBeforePx: Math.max(0, Math.round(parseFloat(s.marginTop) || 0)),
        spacingAfterPx: Math.max(0, Math.round(parseFloat(s.marginBottom) || 0)),
      });
      return out;
    }

    // Colored top/bottom rules on flow containers (header { border-top: 5px
    // solid accent }): the container itself never becomes a paragraph or
    // card, so without synthesis the authored rule vanishes.
    const containerRule = (side) => {
      const width = parseFloat(s[side === 'top' ? 'borderTopWidth' : 'borderBottomWidth']);
      const borderStyle = s[side === 'top' ? 'borderTopStyle' : 'borderBottomStyle'];
      const color = toHex(s[side === 'top' ? 'borderTopColor' : 'borderBottomColor']);
      // 1px hairlines are the most common section separator (header rules)
      if (!(width >= 0.8) || borderStyle === 'none' || !color) return null;
      const body = document.body.getBoundingClientRect();
      return {
        type: 'hr',
        color,
        heightPx: Math.min(width, 8),
        widthFrac: body.width > 0 ? Math.min(1, rect.width / body.width) : 1,
        align: 'left',
      };
    };
    if (!needsBlockWalk(el)) {
      const before = pseudoBar(el, '::before');
      const after = pseudoBar(el, '::after');
      const nodes = paraNodesFrom(el);
      if (pseudoBullet(el)) {
        for (const node of nodes) {
          // A colored glyph run is already prepended by pseudoBulletRun and
          // kept literal (makePara preserves colored markers); adding Word
          // numbering on top would double the bullet ("• •").
          const first = (node.runs || [])[0];
          const firstIsColoredGlyph =
            first &&
            BULLET_GLYPHS.has((first.text || '').trim()) &&
            first.color &&
            first.color !== '000000';
          if (!firstIsColoredGlyph) node.style = { ...node.style, bullet: true };
        }
      }
      return out.concat(before ? [before] : [], nodes, after ? [after] : []);
    }
    {
      const before = pseudoBar(el, '::before');
      const after = pseudoBar(el, '::after');
      const children = processChildren(el, depth);
      // timeline entry (dot + vertical rail drawn as absolute pseudos in the
      // left padding gutter): dot -> colored disc run on the first line;
      // rail -> narrow shaded column (paragraph left borders are ignored by
      // Word/LibreOffice inside table cells, so reuse the accent-strip
      // pattern instead)
      const decor = timelineDecorOf(el);
      const marginBottomPx = Math.max(0, parseFloat(s.marginBottom) || 0);
      const paddingTopPx = Math.max(0, parseFloat(s.paddingTop) || 0);
      const paddingBottomPx = Math.max(0, parseFloat(s.paddingBottom) || 0);
      if (decor) {
        // Nested tables can't split across pages in Word (a whole timeline
        // block gets shoved to the next page), so the rail becomes a
        // paragraph left border. Word/LibreOffice drop paragraph borders in
        // table cells unless the paragraph is indented, hence indentLeftPx.
        const railLeft = Math.max(4, Math.round(decor.railLeftPx ?? 5));
        let firstWithRuns = true;
        for (const child of children) {
          if (child.type !== 'para' && child.type !== 'heading') continue;
          if (decor.rail) {
            child.style = {
              ...child.style,
              borderLeft: { color: decor.rail.color, widthPx: decor.rail.widthPx },
              // indent the text to the authored padding, then push the border
              // back toward the rail's authored offset so wrapped lines keep
              // a consistent gap from the line
              indentLeftPx: Math.max(child.style?.indentLeftPx || 0, Math.round(decor.padL)),
              borderLeftOutsetPx: Math.max(2, Math.round(decor.padL) - railLeft),
            };
          }
          if (firstWithRuns && decor.dot && child.runs?.length) {
            child.runs.unshift({
              text: '\u25CF ',
              color: decor.dot.color,
              sizePx: decor.dot.sizePx,
            });
            firstWithRuns = false;
          }
        }
      }
      const topRule = isCard(el, depth) ? null : containerRule('top');
      const bottomRule = isCard(el, depth) ? null : containerRule('bottom');
      return out.concat(
        topRule ? [topRule] : [],
        before ? [before] : [],
        paddingTopPx > 2 ? [{ type: 'spacer', px: paddingTopPx }] : [],
        children,
        after ? [after] : [],
        paddingBottomPx > 2 ? [{ type: 'spacer', px: paddingBottomPx }] : [],
        bottomRule ? [bottomRule] : [],
        marginBottomPx > 2 ? [{ type: 'spacer', px: marginBottomPx }] : [],
      );
    }
  }

  // A paragraph whose text starts with a literal bullet glyph (Word-exported
  // HTML fakes lists this way: Symbol-font "·" + nbsp gap) is a real list
  // item — strip the fake marker and flag it so the generator emits a native
  // Word bullet.
  const BULLET_GLYPHS = new Set(['·', '•', '●', '■', '▪', '‣', '◦', '∙', '§']);
