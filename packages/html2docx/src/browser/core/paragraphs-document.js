  function makePara(runs, style) {
    const first = runs[0];
    const marker = first && first.text.trim();
    if (marker && BULLET_GLYPHS.has(marker)) {
      if (first.color && first.color !== '000000') {
        return { type: 'para', runs, style };
      }
      let rest = runs.slice(1);
      // swallow the separator gap right after the marker
      if (rest.length && !rest[0].text.replace(/[\s\u00a0\t]/g, '')) rest = rest.slice(1);
      if (rest.length) return { type: 'para', runs: rest, style: { ...style, bullet: true } };
    }
    if (first && /^[·•●▪‣◦∙]\s+/.test(first.text)) {
      if (first.color && first.color !== '000000') {
        return { type: 'para', runs, style };
      }
      const runsCopy = runs.slice();
      runsCopy[0] = { ...first, text: first.text.replace(/^[·•●▪‣◦∙]\s+/, '') };
      return { type: 'para', runs: runsCopy, style: { ...style, bullet: true } };
    }
    return { type: 'para', runs, style };
  }

  function makeParaNodes(runs, style) {
    if (style.align !== 'justify' || !runs.some((run) => (run.text || '').includes('\n'))) {
      return [makePara(runs, style)];
    }
    const lines = [[]];
    for (const run of runs) {
      if (run.inlineImage) {
        lines[lines.length - 1].push(run);
        continue;
      }
      const parts = (run.text || '').split('\n');
      parts.forEach((part, index) => {
        if (part) lines[lines.length - 1].push({ ...run, text: part });
        if (index < parts.length - 1) lines.push([]);
      });
    }
    const nonEmptyLines = lines.filter((line) => line.length);
    if (nonEmptyLines.length <= 1) return [makePara(runs, style)];
    return nonEmptyLines.map((line, index) => {
      const lineStyle = { ...style };
      if (index < nonEmptyLines.length - 1) delete lineStyle.align;
      if (index > 0) delete lineStyle.spacingBeforePx;
      if (index < nonEmptyLines.length - 1) delete lineStyle.spacingAfterPx;
      return makePara(line, lineStyle);
    });
  }

  function paraNodesFrom(el, decorations = true) {
    // A painted display:block span inside a paragraph (scripture-line quote
    // boxes with border-left + tint) renders as its own bordered block in
    // the browser — merging it into the parent's runs loses both the
    // paragraph break and the decoration. Split the paragraph around it.
    const paintedBlocks = [...el.children].filter((child) => {
      if (SKIP_TAGS.has(child.tagName) || !isVisible(child)) return false;
      const childStyle = cs(child);
      const painted = Boolean(bgHex(child)) || Boolean(borderLeftOf(child));
      if (!painted) return false;
      if (childStyle.display === 'block') return true;
      // Inline quote box (scripture-line): border/tint + padding wrapping
      // multiple lines reads as a block quote even at display:inline —
      // short painted chips stay inline via the single-line guard.
      if (parseFloat(childStyle.paddingLeft) < 6) return false;
      if ((child.innerText || '').trim().length < 40) return false;
      const lineHeight = parseFloat(childStyle.lineHeight) || 20;
      return child.getBoundingClientRect().height > lineHeight * 1.4;
    });
    if (paintedBlocks.length && el.children.length <= 8) {
      const style = paraStyleOf(el, decorations);
      const nodes = [];
      let pending = [];
      const flushPending = () => {
        const runs = trimRuns(runsFromNodes(pending, el));
        pending = [];
        if (runs.length) nodes.push(...makeParaNodes(runs, { ...style }));
      };
      for (const node of el.childNodes) {
        if (paintedBlocks.includes(node)) {
          flushPending();
          const quoteRuns = trimRuns(runsFromNodes([...node.childNodes], node));
          if (quoteRuns.length) {
            const quoteStyle = paraStyleOf(node, decorations);
            const bl = borderLeftOf(node);
            if (bl) quoteStyle.borderLeft = bl;
            const shading = bgHex(node);
            if (shading) quoteStyle.shading = shading;
            const padLeft = parseFloat(cs(node).paddingLeft) || 0;
            const marginLeft = parseFloat(cs(node).marginLeft) || 0;
            if (padLeft + marginLeft > 2) quoteStyle.indentLeftPx = Math.round(padLeft + marginLeft);
            nodes.push(...makeParaNodes(quoteRuns, quoteStyle));
          }
        } else pending.push(node);
      }
      flushPending();
      if (nodes.length) return nodes;
    }
    const runs = collectRuns(el);
    if (!runs.length) return [];
    if (el.tagName !== 'LI') {
      const pseudoRun = pseudoBulletRun(el);
      if (pseudoRun) runs.unshift(pseudoRun);
    }
    return makeParaNodes(runs, paraStyleOf(el, decorations));
  }

  const PHRASING_TAGS = new Set([
    'STRONG', 'B', 'EM', 'I', 'U', 'S', 'SPAN', 'A', 'CODE', 'SMALL', 'SUB', 'SUP', 'MARK', 'ABBR', 'Q', 'CITE', 'KBD',
  ]);

  function isInlineNode(node) {
    if (node.nodeType === Node.TEXT_NODE) return true;
    if (node.nodeType !== Node.ELEMENT_NODE) return false;
    if (SKIP_TAGS.has(node.tagName)) return false;
    if (node.tagName === 'IMG' || node.tagName === 'CANVAS' || node.tagName === 'svg') return false;
    if (node.tagName === 'IFRAME' || node.tagName === 'VIDEO' || node.tagName === 'AUDIO') return false;
    if (node.tagName === 'TEXTAREA' || node.tagName === 'SELECT') return false;
    // checkbox inputs flow inline with their label; other inputs are blocks
    if (node.tagName === 'INPUT') return isCheckboxLike(node);
    // Compact circular number badges (30×30 inline-flex chips) must go through
    // processElement for screenshot — otherwise they flatten to white bars.
    // Wider tag pills stay inline so wrap/flow is preserved.
    if (isCompactCircleBadge(node)) return false;
    const d = cs(node).display;
    if (
      d === 'inline-block' &&
      node.getBoundingClientRect().width >= 120 &&
      fullBorderOf(node) &&
      (parseFloat(cs(node).paddingLeft) >= 2 || parseFloat(cs(node).paddingTop) >= 2)
    ) {
      return false;
    }
    if (d === 'inline' || d.startsWith('inline-') || d === 'contents') return true;
    // Chrome reports display:block for phrasing elements that are actually
    // laid out inline (flex containers blockify children; some elements
    // compute block with no matching rule at all). Trust the geometry: if
    // the element shares a line with adjacent text, it flows inline.
    if (PHRASING_TAGS.has(node.tagName) && node.children.length === 0) {
      const sameLine = (textNode, textEdge) => {
        try {
          const range = document.createRange();
          range.selectNodeContents(textNode);
          const rects = range.getClientRects();
          const box = textEdge === 'last' ? rects[rects.length - 1] : rects[0];
          const r = node.getBoundingClientRect();
          if (!box) return false;
          if (textEdge === 'last') return r.top < box.bottom - 2 && r.left >= box.right - 4;
          return box.top < r.bottom - 2 && box.left >= r.right - 4;
        } catch (e) {
          return false;
        }
      };
      const prev = node.previousSibling;
      if (prev && prev.nodeType === Node.TEXT_NODE && prev.textContent.trim() && sameLine(prev, 'last')) {
        return true;
      }
      const next = node.nextSibling;
      if (next && next.nodeType === Node.TEXT_NODE && next.textContent.trim() && sameLine(next, 'first')) {
        return true;
      }
    }
    return false;
  }

  function attachBookmark(nodes, id) {
    if (!id) return false;
    for (const node of nodes) {
      if (node.type === 'para' || node.type === 'heading') {
        node.bookmarks = [...new Set([...(node.bookmarks || []), id])];
        return true;
      }
      if (node.type === 'list' && node.items?.length) {
        const item = node.items[0];
        item.bookmarks = [...new Set([...(item.bookmarks || []), id])];
        return true;
      }
      if (node.children && attachBookmark(node.children, id)) return true;
      for (const cell of node.cells || []) {
        if (cell.children && attachBookmark(cell.children, id)) return true;
      }
    }
    return false;
  }

  // Walk a block container: consecutive inline siblings merge into one paragraph,
  // block-level children are classified recursively.
  function processChildren(el, depth, opts = {}) {
    const decorations = opts.decorations !== false;
    const out = [];
    const containerStyle = cs(el);
    const singleColumnGrid =
      containerStyle.display === 'grid' &&
      containerStyle.gridTemplateColumns.split(' ').filter(Boolean).length === 1;
    const columnFlex =
      containerStyle.display === 'flex' &&
      (containerStyle.flexDirection === 'column' || containerStyle.flexDirection === 'column-reverse');
    const stackGapPx =
      singleColumnGrid || columnFlex
        ? Math.min(
            120,
            Math.max(
              0,
              parseFloat(containerStyle.rowGap) || parseFloat(containerStyle.gap) || 0,
            ),
          )
        : 0;
    let emittedGroup = false;
    let inlineBuffer = [];
    const appendGroup = (nodes) => {
      if (!nodes.length) return;
      if (emittedGroup && stackGapPx > 0) out.push({ type: 'spacer', px: stackGapPx });
      out.push(...nodes);
      emittedGroup = true;
    };
    const flush = () => {
      if (!inlineBuffer.length) return;
      const runs = trimRuns(runsFromNodes(inlineBuffer, el));
      if (runs.length) appendGroup(makeParaNodes(runs, paraStyleOf(el, decorations)));
      inlineBuffer = [];
    };
    // CSS `order` / *-reverse reorders flex/grid items visually without
    // touching the DOM: walking childNodes would emit e.g. the photo/content
    // stack in authored order while the browser shows the reverse. Only
    // reorder when the author opted in and no bare text is in play.
    let iterNodes = el.childNodes;
    if (
      (containerStyle.display === 'grid' || containerStyle.display === 'flex') &&
      ([...el.children].some((c) => (parseInt(cs(c).order, 10) || 0) !== 0) ||
        containerStyle.flexDirection.endsWith('-reverse')) &&
      [...el.childNodes].every(
        (n) => n.nodeType !== Node.TEXT_NODE || !n.textContent.trim(),
      )
    ) {
      iterNodes = [...el.children].sort((a, b) => {
        const ra = a.getBoundingClientRect();
        const rb = b.getBoundingClientRect();
        return ra.top - rb.top || (DOC_RTL ? rb.left - ra.left : ra.left - rb.left);
      });
    }
    for (const node of iterNodes) {
      if (isInlineNode(node)) {
        // An inline-block that starts on a fresh line below the buffered
        // content is a visual block (inline-block heading + inline-block
        // paragraph stacked by wrapping): break the paragraph between them.
        if (
          inlineBuffer.length &&
          node.nodeType === Node.ELEMENT_NODE &&
          cs(node).display.startsWith('inline-')
        ) {
          const r = node.getBoundingClientRect();
          // skip buffered whitespace-only text nodes to find the last element
          let prevEl = null;
          for (let i = inlineBuffer.length - 1; i >= 0; i--) {
            const b = inlineBuffer[i];
            if (b.nodeType === Node.TEXT_NODE && !b.textContent.trim()) continue;
            if (b.nodeType === Node.ELEMENT_NODE) prevEl = b;
            break;
          }
          const pr = prevEl && prevEl.getBoundingClientRect();
          if (
            pr &&
            r.height > 0 &&
            pr.height > 0 &&
            r.top >= pr.bottom - 2 &&
            r.width >= el.getBoundingClientRect().width * 0.4
          ) {
            flush();
            appendGroup(processElement(node, depth));
            continue;
          }
        }
        inlineBuffer.push(node);
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        flush();
        const children = processElement(node, depth);
        attachBookmark(children, node.id);
        appendGroup(children);
      }
    }
    flush();
    return out;
  }

  return globalThis.__html2docxPages.build({
    cs,
    docRTL: DOC_RTL,
    isVisible,
    markForScreenshot,
    nextShotId: () => `h2d-${shotCounter++}`,
    processChildren,
    processElement,
    toHex,
  });