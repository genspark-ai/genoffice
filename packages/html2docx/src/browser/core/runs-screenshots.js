  // Regional_Indicator pairs (flag emoji \uD83C\uDDEC\uD83C\uDDE7) are not Extended_Pictographic, so
  // match them explicitly or they stay in the body-font run and render as the
  // bare letters ("GB").
  const EMOJI_RE = /\p{Extended_Pictographic}|\p{Regional_Indicator}/u;
  const EMOJI_SEQUENCE_RE =
    /(\p{Regional_Indicator}{2}|\p{Extended_Pictographic}(?:\uFE0E|\uFE0F)?(?:\p{Emoji_Modifier})?(?:\u200D\p{Extended_Pictographic}(?:\uFE0E|\uFE0F)?(?:\p{Emoji_Modifier})?)*)/gu;

  // Arrows and misc technical symbols (↔ → ⇄ …): CJK body fonts often lack
  // the glyph in Word and render a box. Split them into symbol-font runs.
  const SYMBOL_RE = /[←-⇿■-◿⤀-⥿⬀-⯿]/;
  const SYMBOL_SEQUENCE_RE = /([←-⇿■-◿⤀-⥿⬀-⯿]+)/g;

  function splitEmojiRuns(runs) {
    return runs
      .flatMap((run) => {
        // Arrows first: U+2194-21FF are Extended_Pictographic, but in prose
        // they are text glyphs — the emoji font would paint them as blue
        // emoji chips.
        if (!run.text || !SYMBOL_RE.test(run.text)) return [run];
        return run.text
          .split(SYMBOL_SEQUENCE_RE)
          .filter(Boolean)
          // No VS15 injection: the symbol-font assignment already forces text
          // presentation, and the selector pollutes copied/searched text.
          .map((text) =>
            SYMBOL_RE.test(text) ? { ...run, text, symbol: true } : { ...run, text },
          );
      })
      .flatMap((run) => {
        if (run.symbol || !run.text || !EMOJI_RE.test(run.text)) return [run];
        return run.text
          .split(EMOJI_SEQUENCE_RE)
          .filter(Boolean)
          .map((text) => ({ ...run, text, emoji: isColorEmoji(text) }));
      });
  }

  // Only default-emoji-presentation glyphs (📘 🇬🇧 ⭐), a VS16-forced one
  // (★️), or a flag pair get the unset-font color-emoji path. Text-presentation
  // dingbats (★ ☆ ● ■ U+2605…) stay in the body font so they keep the run's
  // color instead of turning into a colored emoji star.
  const COLOR_EMOJI_RE = /\p{Emoji_Presentation}|\p{Regional_Indicator}/u;
  function isColorEmoji(text) {
    return COLOR_EMOJI_RE.test(text) || /️/.test(text);
  }

  // Collect styled text runs from a list of DOM nodes (text + inline elements).
  // baseEl provides the coordinate frame for tab-stop measurement.
  function runsFromNodes(nodes, baseEl) {
    const runs = [];
    const baseRect = baseEl ? baseEl.getBoundingClientRect() : null;
    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) {
        // A long run of spaces/nbsp is a column separator (Word-exported HTML
        // spreads columns and right-aligns dates this way). Emit a tab marker
        // and measure where the following content actually starts so the
        // generator can place an equivalent tab stop.
        const raw = node.textContent;
        // newline-preserving white-space (pre-wrap message bodies): authored
        // \n are real line/paragraph breaks, not collapsible spaces
        if (raw.includes('\n') && node.parentElement) {
          const ws = cs(node.parentElement).whiteSpace || '';
          if (ws.startsWith('pre') || ws === 'break-spaces') {
            const style = runStyleOf(node.parentElement);
            raw.split('\n').forEach((line, i) => {
              if (i) runs.push({ text: '\n', ...style });
              const text = line.replace(/\s+/g, ' ');
              if (text.trim()) runs.push({ text, ...style });
            });
            return;
          }
        }
        const parts = raw.split(/(\s{6,})/);
        let offset = 0;
        for (const part of parts) {
          if (part) {
            let isSeparator = false;
            let tabFrac = null;
            if (/^\s{6,}$/.test(part) && baseRect && baseRect.width > 0) {
              // Only a *rendered* wide gap is a column separator. Raw source
              // may contain long space runs that the browser collapses to a
              // single space (Word-exported HTML) — those must stay spaces.
              try {
                const range = document.createRange();
                range.setStart(node, offset);
                range.setEnd(node, offset + part.length);
                const rect = range.getBoundingClientRect();
                if (rect && rect.width > 14) {
                  isSeparator = true;
                  tabFrac = (rect.right - baseRect.left) / baseRect.width;
                }
              } catch (e) { /* fall through as plain space */ }
            }
            if (isSeparator && tabFrac != null && tabFrac >= 0.6) {
              // Gap pushing content to the far right: right-aligned tab (dates).
              runs.push({ text: '\t', tabFrac, ...runStyleOf(node.parentElement) });
            } else if (isSeparator) {
              // Mid-line column gap: fixed wide space (tab stops in Word are
              // too fragile to reproduce arbitrary column positions).
              runs.push({ text: '\u00a0\u00a0\u00a0 ', ...runStyleOf(node.parentElement) });
            } else {
              const text = part.replace(/\s+/g, ' ');
              if (text) runs.push({ text, ...runStyleOf(node.parentElement) });
            }
          }
          offset += part.length;
        }
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      if (SKIP_TAGS.has(node.tagName)) return;
      if (node.tagName === 'BR') {
        // BR has a 0x0 rect, so this must come before the visibility check
        runs.push({ text: '\n', ...runStyleOf(node.parentElement) });
        return;
      }
      // Empty underline blanks can report ~1px height and fail isVisible.
      {
        const fillInEarly = fillInLineNode(node);
        if (fillInEarly?.inline) {
          const style = cs(node);
          const fontSize = parseFloat(style.fontSize) || 14;
          const spaceWidth = Math.max(3, fontSize * 0.45);
          const count = Math.max(3, Math.round(fillInEarly.width / spaceWidth));
          const bottom = fillInEarly.borders?.bottom;
          runs.push({
            text: '\u00a0'.repeat(count),
            underline: true,
            underlineStyle: bottom?.style || 'solid',
            color: bottom?.color || '000000',
            sizePx: fontSize,
            preserveWhitespace: true,
          });
          return;
        }
      }
      if (!isVisible(node)) return;
      // Preceding content glued to this element in the source (no whitespace
      // text node): a block child renders on its own LINE in the browser
      // (August 16<span class="h3-sub" style="display:block">Arrival…), and
      // an inline sibling separated by CSS margin needs a space — without
      // either, the words fuse in Word ("August 16Arrival").
      if (
        runs.length &&
        runs[runs.length - 1].text &&
        !/\s$/.test(runs[runs.length - 1].text)
      ) {
        if (cs(node).display === 'block') {
          runs.push({ text: '\n', ...runStyleOf(node.parentElement || node) });
        } else {
          const prevSibling = node.previousSibling;
          if (
            prevSibling &&
            prevSibling.nodeType === Node.TEXT_NODE &&
            /\S$/.test(prevSibling.textContent)
          ) {
            try {
              const lastGlyph = document.createRange();
              lastGlyph.setStart(prevSibling, prevSibling.textContent.length - 1);
              lastGlyph.setEnd(prevSibling, prevSibling.textContent.length);
              const prevRight = lastGlyph.getBoundingClientRect().right;
              const glyphs = document.createRange();
              glyphs.selectNodeContents(node);
              const nodeLeft = glyphs.getBoundingClientRect().left;
              if (nodeLeft - prevRight >= 3) {
                runs.push({ text: ' ', ...runStyleOf(node.parentElement || node) });
              }
            } catch (e) {
              /* no gap detected, nothing to add */
            }
          }
        }
      }
      if (isMathContainer(node)) {
        const r = node.getBoundingClientRect();
        if (r.width >= 3 && r.height >= 6) {
          const shot = markForScreenshot(node);
          runs.push({
            text: '',
            inlineImage: true,
            shotId: shot.shotId,
            width: shot.width,
            height: shot.height,
          });
        }
        return;
      }
      if (node.tagName === 'SVG' || node.tagName === 'svg' || node.tagName === 'IMG') {
        const r = node.getBoundingClientRect();
        if (r.width <= 64 && r.height <= 64) {
          const shot = markForScreenshot(node);
          runs.push({
            text: '',
            inlineImage: true,
            shotId: shot.shotId,
            width: shot.width,
            height: shot.height,
          });
          const style = cs(node);
          const parentStyle = node.parentElement ? cs(node.parentElement) : null;
          const gapPx = Math.max(
            parseFloat(style.marginRight) || 0,
            parseFloat(parentStyle?.columnGap) || parseFloat(parentStyle?.gap) || 0,
          );
          if (gapPx >= 4) {
            const spaceWidth = Math.max(3, (parseFloat(parentStyle?.fontSize) || 14) * 0.3);
            const count = Math.max(1, Math.round((gapPx - spaceWidth) / spaceWidth));
            runs.push({
              text: '\u00a0'.repeat(count),
              ...runStyleOf(node.parentElement || node),
              preserveWhitespace: true,
            });
          }
        }
        return;
      }
      // Empty painted chip (legend color swatch, status dot): no text, no
      // children, just a small background/border box \u2014 dropping it loses the
      // color-to-meaning key, so inline-embed it like a small icon.
      {
        const chipStyle = cs(node);
        const r = node.getBoundingClientRect();
        if (
          !node.childElementCount &&
          !(node.innerText || '').trim() &&
          r.width >= 4 && r.width <= 28 &&
          r.height >= 4 && r.height <= 28 &&
          (bgHex(node) ||
            (chipStyle.backgroundImage || '').includes('gradient') ||
            (parseFloat(chipStyle.borderTopWidth) > 0 && parseFloat(chipStyle.borderRadius) > 0))
        ) {
          const shot = markForScreenshot(node);
          runs.push({
            text: '',
            inlineImage: true,
            shotId: shot.shotId,
            width: shot.width,
            height: shot.height,
          });
          const parentStyle = node.parentElement ? cs(node.parentElement) : null;
          const gapPx = Math.max(
            parseFloat(chipStyle.marginRight) || 0,
            parseFloat(parentStyle?.columnGap) || parseFloat(parentStyle?.gap) || 0,
          );
          if (gapPx >= 4) {
            const spaceWidth = Math.max(3, (parseFloat(parentStyle?.fontSize) || 14) * 0.3);
            const count = Math.max(1, Math.round((gapPx - spaceWidth) / spaceWidth));
            runs.push({
              text: '\u00a0'.repeat(count),
              ...runStyleOf(node.parentElement || node),
              preserveWhitespace: true,
            });
          }
          return;
        }
      }
      if (isCheckboxLike(node)) {
        runs.push(checkboxNode(node));
        return;
      }
      const nodeStyle = cs(node);
      const nodeRect = node.getBoundingClientRect();
      // A circular painted badge (numbered day markers) has no text-run
      // equivalent — the round paint is the point. Plain rectangular chips
      // stay text so they remain searchable.
      const circularBadge =
        (parseFloat(nodeStyle.borderRadius) || 0) >= nodeRect.height / 2 - 1 &&
        Math.abs(nodeRect.width - nodeRect.height) <= nodeRect.height * 0.4 &&
        !/\s/.test((node.innerText || '').trim());
      const compactPaintedLabel =
        (['inline-block', 'inline-flex', 'inline-grid'].includes(nodeStyle.display) ||
          // a circular badge centering its digit with flex/grid still sits
          // inline in its heading row
          (circularBadge && ['flex', 'grid'].includes(nodeStyle.display))) &&
        (nodeStyle.whiteSpace === 'nowrap' || circularBadge) &&
        // inlineBgHex rejects non-inline displays; a flex badge paints its
        // own background directly
        Boolean(inlineBgHex(node) || (circularBadge && bgHex(node))) &&
        (node.innerText || '').trim().length <= 30 &&
        nodeRect.width >= 12 &&
        nodeRect.width <= 220 &&
        nodeRect.height >= 12 &&
        nodeRect.height <= 50;
      if (compactPaintedLabel) {
        const shot = markForScreenshot(node);
        // Label crushed by a too-narrow column (letters stacked vertically):
        // ask the screenshot pass to relax it to its natural one-line size,
        // and record that relaxed size (measured under nowrap) so the docx
        // shows the image at its true aspect.
        const crushed = node.scrollWidth > node.clientWidth + 2;
        if (crushed) {
          shot.unwrap = true;
          const previousWhiteSpace = node.style.whiteSpace;
          node.style.whiteSpace = 'nowrap';
          shot.width = Math.min(220, node.scrollWidth + 2);
          shot.height = node.clientHeight || shot.height;
          node.style.whiteSpace = previousWhiteSpace || '';
        }
        // A circular badge taller than its line's text inflates the Word row
        // height and pushes near-page-height cards past the page. Display it
        // at the surrounding text's line scale; the PNG keeps full detail.
        let displayWidth = shot.width;
        let displayHeight = shot.height;
        if (circularBadge && node.parentElement) {
          const parentFont = parseFloat(cs(node.parentElement).fontSize) || 16;
          const cap = Math.round(parentFont * 1.35);
          if (displayHeight > cap) {
            displayWidth = Math.round((shot.width * cap) / shot.height);
            displayHeight = cap;
          }
        }
        runs.push({
          text: '',
          inlineImage: true,
          shotId: shot.shotId,
          width: displayWidth,
          height: displayHeight,
          unwrap: shot.unwrap,
        });
        return;
      }
      if (
        !(node.innerText || '').trim() &&
        node.children.length === 0 &&
        (nodeStyle.display === 'inline' || nodeStyle.display === 'inline-block') &&
        nodeRect.width >= 4 &&
        nodeRect.width <= 32 &&
        nodeRect.height >= 4 &&
        nodeRect.height <= 32
      ) {
        const swatch = bgHex(node);
        if (swatch) {
          runs.push({
            text: '\u25a0 ',
            color: swatch,
            sizePx: Math.max(nodeRect.width, nodeRect.height),
          });
          return;
        }
      }
      const highlightedInline =
        nodeStyle.display.startsWith('inline-') &&
        Boolean(inlineBgHex(node)) &&
        Boolean((node.innerText || '').trim());
      const paddedSpaces = (px) => {
        if (px < 2) return '';
        const spaceWidth = Math.max(3, (parseFloat(nodeStyle.fontSize) || 14) * 0.3);
        return '\u00a0'.repeat(Math.max(1, Math.round(px / spaceWidth)));
      };
      if (highlightedInline) {
        const leftPad = paddedSpaces(parseFloat(nodeStyle.paddingLeft) || 0);
        if (leftPad) runs.push({ text: leftPad, ...runStyleOf(node), preserveWhitespace: true });
      }
      const inlinePseudoRun = pseudoBulletRun(node);
      if (inlinePseudoRun) runs.push(inlinePseudoRun);
      for (const child of node.childNodes) walk(child);
      // Symmetric to the block-BEFORE rule above: a block child also ends
      // its own line, so following content (an eyebrow <span
      // class="section-num" style="display:block"> glued before heading
      // text) must not fuse onto it. trimRuns drops a trailing break.
      if (
        nodeStyle.display === 'block' &&
        runs.length &&
        runs[runs.length - 1].text &&
        !/\n$/.test(runs[runs.length - 1].text)
      ) {
        runs.push({ text: '\n', ...runStyleOf(node.parentElement || node) });
      }
      if (highlightedInline) {
        const rightPad = paddedSpaces(parseFloat(nodeStyle.paddingRight) || 0);
        if (rightPad) runs.push({ text: rightPad, ...runStyleOf(node), preserveWhitespace: true });
      }
      // Inline chip followed directly by text with no space character in the
      // source (<span class="q-mark">Q</span>How...): the visual gap comes
      // from min-width/margin/padding. Keep a space or the words glue.
      const next = node.nextSibling;
      if (
        next &&
        next.nodeType === Node.TEXT_NODE &&
        /^\S/.test(next.textContent) &&
        runs.length &&
        runs[runs.length - 1].text &&
        !/\s$/.test(runs[runs.length - 1].text)
      ) {
        try {
          const glyphs = document.createRange();
          glyphs.selectNodeContents(node);
          const glyphRight = glyphs.getBoundingClientRect().right;
          const first = document.createRange();
          first.setStart(next, 0);
          first.setEnd(next, 1);
          const nextLeft = first.getBoundingClientRect().left;
          if (nextLeft - glyphRight >= 3) {
            runs.push({ text: ' ', ...runStyleOf(node.parentElement || node) });
          }
        } catch (e) {
          /* no gap detected, nothing to add */
        }
      }
    }
    for (const node of nodes) walk(node);
    // Latin fragments with neutral punctuation inside an RTL context get
    // bidi-reordered by Word ("Motorola 68010, supplier A" -> "Motorola
    // supplier ,68010 A"): bracket them with LRM so the whole fragment
    // stays one LTR segment. Runs carrying RTL script keep native ordering.
    const isolated = runs.map((run) => {
      if (!run.rtl || !run.text || run.inlineImage) return run;
      if (/[֐-ࣿיִ-ﻼ]/.test(run.text)) return run;
      if (!/[A-Za-z0-9]/.test(run.text)) return run;
      if (!/[()[\]{}+*/,.:%~=-]/.test(run.text)) return run;
      return { ...run, text: `‎${run.text}‎` };
    });
    return splitEmojiRuns(isolated);
  }

  // Measured gap between the element's left edge and its first visible glyph
  // (covers both CSS padding-left and leading nbsp used as manual inset).
  function textInsetLeft(el, rect) {
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      const m = node.textContent.match(/[^\s\u00a0]/);
      if (!m) continue;
      try {
        const range = document.createRange();
        range.setStart(node, m.index);
        range.setEnd(node, m.index + 1);
        const glyph = range.getBoundingClientRect();
        if (glyph.width > 0) return glyph.left - rect.left;
      } catch (e) {
        return 0;
      }
      return 0;
    }
    return 0;
  }

  function trimRuns(runs) {
    if (runs.length) {
      if (!runs[0].preserveWhitespace) runs[0].text = runs[0].text.replace(/^\s+/, '');
      if (!runs[runs.length - 1].preserveWhitespace) {
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      }
    }
    return runs.filter((r) => r.inlineImage || r.text.length > 0);
  }

  function collectRuns(el) {
    return trimRuns(runsFromNodes([...el.childNodes], el));
  }

  function hasBlockChildren(el) {
    let prevRect = null;
    for (const child of el.children) {
      if (SKIP_TAGS.has(child.tagName)) continue;
      const d = cs(child).display;
      if (d !== 'inline' && !d.startsWith('inline-')) return true;
      // Two wide inline-blocks that render on separate lines (an inline-block
      // heading too wide to share its line with the following inline-block
      // paragraph) are visually blocks; merging them glues the texts. The
      // width guard keeps small wrapping chips/badges on the inline path.
      if (d.startsWith('inline-')) {
        const r = child.getBoundingClientRect();
        const wide = r.width >= el.getBoundingClientRect().width * 0.4;
        if (prevRect && wide && r.height > 0 && r.top >= prevRect.bottom - 2) return true;
        if (r.height > 0 && wide) prevRect = r;
        else if (r.height > 0) prevRect = null;
      }
    }
    return false;
  }

  // Replaced elements (img/svg/canvas) are excluded from inline runs, so a
  // wrapper holding only an inline <img> must still walk children as blocks
  // or the image silently disappears.
  function isReplacedTag(el) {
    const t = el.tagName;
    if (
      t === 'IMG' ||
      t === 'CANVAS' ||
      t === 'SVG' ||
      t === 'svg' ||
      t === 'IFRAME' ||
      t === 'VIDEO' ||
      t === 'AUDIO'
    ) {
      return true;
    }
    if (t === 'TEXTAREA' || t === 'SELECT') return true;
    if (t === 'INPUT') return !isCheckboxLike(el);
    return false;
  }

  function hasReplacedDescendant(el) {
    if (isReplacedTag(el)) return true;
    for (const d of el.querySelectorAll(
      'img, svg, canvas, iframe, video, audio, textarea, select, input:not([type=checkbox]):not([type=radio])',
    )) {
      if (isVisible(d)) return true;
    }
    return false;
  }

  // needs full block walk (vs a single merged paragraph)
  function needsBlockWalk(el) {
    return hasBlockChildren(el) || hasReplacedDescendant(el);
  }

  // decorations=false: only text-level styling (align/spacing/line-height).
  // Used for paragraphs inside cards/cells where the container already owns
  // the border and shading (avoids double-drawn vertical bars).
  function paraStyleOf(el, decorations = true) {
    const s = cs(el);
    const style = {};
    style.cjk = /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/.test(
      el.textContent || '',
    );
    if (isRTL(el)) style.rtl = true;
    const align = textAlignOf(el);
    if (align) style.align = align;
    const mb = parseFloat(s.marginBottom);
    if (mb > 2) style.spacingAfterPx = Math.min(mb, 40);
    const mt = parseFloat(s.marginTop);
    if (mt > 2) style.spacingBeforePx = Math.min(mt, 40);
    if (
      style.spacingBeforePx &&
      el.parentElement === document.body &&
      !el.nextElementSibling
    ) {
      // Large decorative signatures/flourishes at the end of a one-page
      // document should use the remaining page space, not force a blank page.
      style.spacingBeforePx = Math.min(style.spacingBeforePx, 2);
    }
    const ti = parseFloat(s.textIndent);
    if (ti > 2) style.firstLineIndentPx = Math.min(ti, 100);
    const lh = parseFloat(s.lineHeight);
    const fsz = parseFloat(s.fontSize);
    if (lh && fsz && lh / fsz > 1.25) style.lineRatio = Math.min(lh / fsz, 2.2);
    if (style.cjk && lh && fsz) style.exactLineHeightPx = lh;
    // Heading-like paragraph (bold, short, followed by content): keep it on
    // the same page as what it introduces - FAQ questions, styled pseudo-
    // headings etc. are divs/ps in HTML but behave like headings in print.
    const headText = (el.innerText || '').trim();
    if (
      parseFloat(s.fontWeight) >= 600 &&
      headText.length > 0 &&
      headText.length <= 120 &&
      !headText.includes('\n') &&
      el.nextElementSibling
    ) {
      style.keepNext = true;
      // A short bold heading-like line must also not split across pages
      // mid-word ("EDUCATIO" / "N").
      style.keepLines = true;
    }
    // Short text rendered on ONE browser line: record the measured width so
    // the generator can size narrow cells (or shrink display banners)
    // instead of letting Word wrap "Sub Total :" or "INVOICE" mid-word.
    if (fsz >= 10 && headText.length > 0 && headText.length <= 40 && !headText.includes('\n')) {
      try {
        const range = document.createRange();
        range.selectNodeContents(el);
        const textRect = range.getBoundingClientRect();
        if (textRect.width > 0 && textRect.height <= fsz * 2.1) {
          // Word's substituted font may run much wider than the webfont.
          const wordWidth = window.__h2dWordFontLineWidth
            ? window.__h2dWordFontLineWidth(el)
            : 0;
          style.oneLineWidthPx = Math.round(Math.max(textRect.width, wordWidth));
        }
      } catch (e) {
        /* leave unset */
      }
    }
    if (!decorations) {
      // decorations=false suppresses the container-owned paint, but a child
      // with its own distinct background (label chips, "OBJECTION 1 OF 6"
      // bars inside cards) must keep it.
      const ownBg = bgHex(el);
      if (ownBg && (!el.parentElement || ownBg !== bgHex(el.parentElement))) {
        style.shading = ownBg;
      }
      return style;
    }
    // manual left inset (decoration: containers that own their inset skip it)
    const padL = parseFloat(s.paddingLeft) + parseFloat(s.marginLeft);
    if (padL >= 12) style.indentLeftPx = Math.min(padL, 200);
    let shading = paintedBgHex(el);
    // authored vertical padding on unshaded wrappers (e.g. a card body div)
    // reads as spacing; shaded bars handle their padding via exact height
    if (!shading) {
      const pt = parseFloat(s.paddingTop) || 0;
      const pb = parseFloat(s.paddingBottom) || 0;
      if (pt > 2) style.spacingBeforePx = Math.min((style.spacingBeforePx || 0) + pt, 40);
      if (pb > 2) style.spacingAfterPx = Math.min((style.spacingAfterPx || 0) + pb, 40);
    }
    if (shading) {
      const textColor = toHex(s.color) || '000000';
      const shadeLum = hexLuminance(shading);
      const textLum = hexLuminance(textColor);
      // Near-white chip on a colored parent + light text would become an
      // invisible full-width white bar (benefit-num / step circles). Skip.
      if (shadeLum > 230 && textLum > 200) {
        shading = null;
      }
    }
    if (shading) {
      style.shading = shading;
      // Word's single spacing for CJK fonts has much taller leading than the
      // browser's, which fattens shaded bars. For single-line shaded
      // paragraphs, pin the rendered height as an exact line height.
      // Elements with a left accent border are note cards, not header bars -
      // the color bar table would drop the accent, keep them as paragraphs.
      const rect = el.getBoundingClientRect();
      const lineH = lh || fsz * 1.35;
      const radius = parseFloat(s.borderTopLeftRadius) || 0;
      // Compact circular badges are not full-width header bars.
      const compactBadge = rect.width > 0 && rect.width <= 72 && radius >= rect.height * 0.4;
      if (
        rect.height > 0 &&
        lineH &&
        rect.height < lineH * 1.9 &&
        !borderLeftOf(el) &&
        !compactBadge
      ) {
        style.exactLineHeightPx = rect.height;
        // single-line bar: safe to render as a fixed-height color bar table.
        // (exactLineHeightPx alone also gets set for CJK line control and
        // must NOT trigger the bar path - it would clip multi-line text.)
        style.colorBar = true;
        // Text inset inside the bar (padding-left and/or leading nbsp runs,
        // which trimRuns strips from the text itself).
        const inset = textInsetLeft(el, rect);
        if (inset > 1) style.barInsetLeftPx = Math.min(inset, 60);
      }
    }
    const bl = borderLeftOf(el);
    if (bl) {
      style.borderLeft = bl;
      const paddingLeft = parseFloat(s.paddingLeft) || 0;
      if (paddingLeft > 0) style.borderLeftSpacePx = Math.min(paddingLeft, 40);
    }
    const br = borderRightOf(el);
    if (br) {
      style.borderRight = br;
      const paddingRight = parseFloat(s.paddingRight) || 0;
      if (paddingRight > 0) style.borderRightSpacePx = Math.min(paddingRight, 40);
    }
    const bt = borderSideOf(el, 'top');
    if (bt) style.borderTop = bt;
    const bw = parseFloat(s.borderBottomWidth);
    if (bw >= 1 && s.borderBottomStyle !== 'none') {
      const color = toHex(s.borderBottomColor);
      if (color) {
        style.borderBottom = { color, widthPx: bw };
        // dotted/dashed TOC separators must not harden into solid rules
        if (s.borderBottomStyle === 'dotted' || s.borderBottomStyle === 'dashed') {
          style.borderBottom.style = s.borderBottomStyle;
        }
      }
    }
    return style;
  }

  // Lowest visible pixel row (relative to the element top) of the element's
  // own paint and descendants: text lines, replaced elements, painted boxes.
  function visibleContentBottom(el, rect) {
    let bottom = rect.top;
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const textNode = walker.currentNode;
      if (!textNode.textContent.trim()) continue;
      const range = document.createRange();
      range.selectNodeContents(textNode);
      for (const lineRect of range.getClientRects()) {
        if (lineRect.height > 0 && lineRect.width > 0) bottom = Math.max(bottom, lineRect.bottom);
      }
    }
    for (const child of el.querySelectorAll('*')) {
      if (!isVisible(child)) continue;
      const childRect = child.getBoundingClientRect();
      if (childRect.height <= 0 || childRect.width <= 0) continue;
      const childStyle = cs(child);
      const painted =
        child.matches('img, svg, canvas, video, iframe, input, textarea, select, button, hr') ||
        childStyle.backgroundImage !== 'none' ||
        childStyle.boxShadow !== 'none' ||
        Boolean(toHex(childStyle.backgroundColor, { dropWhite: false })) ||
        parseFloat(childStyle.borderTopWidth) > 0 ||
        parseFloat(childStyle.borderBottomWidth) > 0;
      if (painted) bottom = Math.max(bottom, Math.min(childRect.bottom, rect.bottom));
    }
    return bottom - rect.top;
  }

  function markForScreenshot(el) {
    const id = `h2d-${shotCounter++}`;
    el.setAttribute('data-h2d-id', id);
    const r = el.getBoundingClientRect();
    const s = cs(el);
    const body = document.body.getBoundingClientRect();
    const offLeft = r.left - body.left;
    const offRight = body.right - r.right;
    const align =
      offLeft > 8 && offRight > 8 && Math.abs(offLeft - offRight) < Math.max(12, r.width * 0.08)
        ? 'center'
        : offLeft > offRight * 3
          ? 'right'
          : 'left';
    // A transparent wrapper much taller than its visible content (e.g. a
    // photo container stretched by the column) would embed a page-height
    // mostly-empty image and strand a phantom page. Crop the invisible tail;
    // painted elements keep their full box (the paint IS content).
    let height = Math.round(r.height);
    let clip;
    const paintedSelf =
      s.backgroundImage !== 'none' ||
      s.boxShadow !== 'none' ||
      Boolean(toHex(s.backgroundColor, { dropWhite: false })) ||
      parseFloat(s.borderTopWidth) > 0 ||
      parseFloat(s.borderBottomWidth) > 0;
    if (r.height >= 600 && !paintedSelf) {
      const contentBottom = visibleContentBottom(el, r);
      if (contentBottom > 0 && r.height - contentBottom >= 150) {
        height = Math.round(Math.min(r.height, contentBottom + 12));
        clip = {
          x: Math.max(0, Math.round(r.left)),
          y: Math.max(0, Math.round(r.top)),
          width: Math.round(r.width),
          height,
        };
      }
    }
    return {
      type: 'image',
      shotId: id,
      width: Math.round(r.width),
      height,
      ...(clip ? { clip } : {}),
      align,
      spacingBeforePx: Math.max(0, Math.round(parseFloat(s.marginTop) || 0)),
      spacingAfterPx: Math.max(0, Math.round(parseFloat(s.marginBottom) || 0)),
    };
  }

