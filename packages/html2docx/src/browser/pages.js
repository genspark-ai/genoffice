(function installPageExtractor(global) {
  function build(deps) {
    const {
      cs,
      docRTL,
      isVisible,
      markForScreenshot,
      nextShotId,
      processChildren,
      processElement,
      toHex,
    } = deps;
    const rootBodyStyle = cs(document.body);
    const absoluteRootChildren = [...document.body.children].filter(
      (child) => isVisible(child) && cs(child).position === 'absolute',
    );
    const bodyHeight = document.body.getBoundingClientRect().height;
    const visibleBodyChildren = [...document.body.children].filter(isVisible);
    const authoredPageRoots = visibleBodyChildren.filter(
      (child) => child.hasAttribute('data-docx-page') || child.hasAttribute('data-page'),
    );
    const inferredPageRoots = visibleBodyChildren.filter((child) => {
      if (authoredPageRoots.includes(child)) return false;
      if (!isVisible(child)) return false;
      const style = cs(child);
      const rect = child.getBoundingClientRect();
      return (
        (style.breakAfter === 'page' || style.pageBreakAfter === 'always') &&
        rect.width >= 300 &&
        rect.height >= 400
      );
    });
    const multiPageComposition = authoredPageRoots.length >= 2;
    const mixedPageComposition = !multiPageComposition && inferredPageRoots.length >= 2;
    const renderedPageHeight = window.innerHeight;
    // Fixed-composition slide pages: page-sized relative sections whose
    // direct children are absolutely positioned (coordinate layouts authored
    // without data-docx-page). Flow classification drops the positioned
    // content entirely, so rasterize each page like authored pages.
    const slidePageRoots = visibleBodyChildren.filter((child) => {
      if (authoredPageRoots.includes(child)) return false;
      const style = cs(child);
      if (style.position !== 'relative' && style.position !== 'absolute') return false;
      const rect = child.getBoundingClientRect();
      if (Math.abs(rect.height - renderedPageHeight) > renderedPageHeight * 0.08) return false;
      if (rect.width < window.innerWidth * 0.8) return false;
      const absoluteKids = [...child.children].filter(
        (kid) => isVisible(kid) && cs(kid).position === 'absolute',
      );
      return absoluteKids.length >= 2 && (child.innerText || '').trim().length >= 20;
    });
    const slidePageComposition =
      !multiPageComposition &&
      !mixedPageComposition &&
      slidePageRoots.length >= 2 &&
      slidePageRoots.length >= visibleBodyChildren.length - 1;
    const multiPageAbsoluteComposition =
      !multiPageComposition &&
      !mixedPageComposition &&
      bodyHeight > renderedPageHeight &&
      bodyHeight <= renderedPageHeight * 10 &&
      absoluteRootChildren.some(
        (child) => child.getBoundingClientRect().top >= renderedPageHeight - 2,
      );
    const activePageView = [...document.body.children].find(
      (child) =>
        isVisible(child) &&
        child.classList.contains('page-view') &&
        child.classList.contains('active-page'),
    );
    const fixedNavigation = [...document.body.children].some((child) => {
      if (!isVisible(child)) return false;
      const style = cs(child);
      return (
        (child.tagName === 'NAV' || child.tagName === 'HEADER') &&
        (style.position === 'fixed' || style.position === 'sticky')
      );
    });
    const webPageComposition =
      !multiPageComposition &&
      !mixedPageComposition &&
      !multiPageAbsoluteComposition &&
      Boolean(activePageView) &&
      fixedNavigation &&
      bodyHeight > renderedPageHeight &&
      activePageView.querySelector('img') &&
      activePageView.querySelector('.container, .row, [class*="grid"]');
    const pageComposition =
      (rootBodyStyle.position === 'relative' || rootBodyStyle.position === 'absolute') &&
      rootBodyStyle.overflow === 'hidden' &&
      bodyHeight >= 700 &&
      bodyHeight <= 1600 &&
      absoluteRootChildren.length >= 3 &&
      (document.body.innerText || '').trim().length >= 20;
    const singleVisualPageRoot = visibleBodyChildren.length === 1 ? visibleBodyChildren[0] : null;
    const singleVisualPageRect = singleVisualPageRoot?.getBoundingClientRect();
    const singleVisualPageStyle = singleVisualPageRoot ? cs(singleVisualPageRoot) : null;
    const singleVisualPagePaint =
      singleVisualPageStyle &&
      (singleVisualPageStyle.backgroundImage !== 'none' ||
        Boolean(toHex(singleVisualPageStyle.backgroundColor, { dropWhite: true })));
    const singleVisualPageLayout = singleVisualPageRoot
      ? [...singleVisualPageRoot.querySelectorAll('*')].some((element) => {
          const style = cs(element);
          return (
            (style.display === 'grid' && style.gridTemplateColumns.split(' ').length >= 2) ||
            (style.display === 'flex' &&
              (style.flexDirection === 'row' || style.flexDirection === ''))
          );
        })
      : false;
    const flowingVisualComposition =
      !multiPageComposition &&
      !mixedPageComposition &&
      !multiPageAbsoluteComposition &&
      !webPageComposition &&
      !pageComposition &&
      Boolean(singleVisualPageRoot) &&
      singleVisualPageStyle.position === 'relative' &&
      (singleVisualPageStyle.overflow === 'hidden' ||
        singleVisualPageStyle.overflow === 'clip') &&
      singleVisualPagePaint &&
      singleVisualPageLayout &&
      singleVisualPageRect.width >= window.innerWidth * 0.7 &&
      singleVisualPageRect.height > renderedPageHeight * 1.05 &&
      singleVisualPageRect.height <= renderedPageHeight * 3 &&
      (singleVisualPageRoot.innerText || '').trim().length >= 20;
    const singlePageVisualComposition =
      !multiPageComposition &&
      !mixedPageComposition &&
      !multiPageAbsoluteComposition &&
      !webPageComposition &&
      !pageComposition &&
      !flowingVisualComposition &&
      Boolean(singleVisualPageRoot) &&
      (singleVisualPageStyle.position === 'relative' ||
        singleVisualPageStyle.position === 'absolute') &&
      (singleVisualPageStyle.overflow === 'hidden' ||
        singleVisualPageStyle.overflow === 'clip') &&
      singleVisualPagePaint &&
      singleVisualPageRect.width >= window.innerWidth * 0.7 &&
      singleVisualPageRect.height >= renderedPageHeight * 0.7 &&
      singleVisualPageRect.height <= renderedPageHeight * 1.35 &&
      (singleVisualPageRoot.innerText || '').trim().length >= 20;
    const bodySinglePageVisualComposition =
      !multiPageComposition &&
      !mixedPageComposition &&
      !multiPageAbsoluteComposition &&
      !webPageComposition &&
      !pageComposition &&
      !singlePageVisualComposition &&
      rootBodyStyle.position === 'relative' &&
      rootBodyStyle.backgroundImage !== 'none' &&
      rootBodyStyle.maxWidth !== 'none' &&
      visibleBodyChildren.length >= 2 &&
      visibleBodyChildren.length <= 5 &&
      bodyHeight >= renderedPageHeight * 0.7 &&
      bodyHeight <= renderedPageHeight * 1.35 &&
      (document.body.innerText || '').trim().length >= 20;
    const bodyRect = document.body.getBoundingClientRect();
    // Word (Mac) insets the first line ~26px below the page top even at zero
    // margins, so a slice scaled to the full content height loses its bottom
    // rows to the page edge. Budget slices below the page height so the
    // width-driven scale governs every slice (uniform scale across pages)
    // and Word never clips a pixel row. Matches the 45px renderer headroom:
    // 1 - 45/1123.
    const sliceBudget = Math.round(renderedPageHeight * 0.96);
    // Bands (y ranges relative to body top) a slice boundary must not cross:
    // text lines and replaced elements always; painted card boxes when small
    // enough to move to the next page whole.
    const collectCutBands = () => {
      const lineBands = [];
      const boxBands = [];
      const push = (bands, rect) => {
        if (rect.height > 0 && rect.width > 0) {
          bands.push({ top: rect.top - bodyRect.top, bottom: rect.bottom - bodyRect.top });
        }
      };
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const textNode = walker.currentNode;
        if (!textNode.textContent.trim()) continue;
        const range = document.createRange();
        range.selectNodeContents(textNode);
        for (const rect of range.getClientRects()) push(lineBands, rect);
      }
      for (const element of document.body.querySelectorAll('*')) {
        if (!isVisible(element)) continue;
        const rect = element.getBoundingClientRect();
        if (rect.height <= 0 || rect.width <= 0) continue;
        if (
          element.matches('img, svg, canvas, video, iframe, input, textarea, select, button')
        ) {
          push(lineBands, rect);
          continue;
        }
        if (rect.height < 40 || rect.height > sliceBudget * 0.9) continue;
        const style = cs(element);
        // dropWhite:false — an explicit white card on a tinted page is still
        // a visual card whose rows must not be split across pages.
        const painted =
          style.backgroundImage !== 'none' ||
          style.boxShadow !== 'none' ||
          Boolean(toHex(style.backgroundColor, { dropWhite: false })) ||
          parseFloat(style.borderTopWidth) > 0 ||
          parseFloat(style.borderBottomWidth) > 0;
        if (painted) push(boxBands, rect);
      }
      return { lineBands, boxBands };
    };
    const chooseCut = (ideal, minCut, bands) => {
      const clearance = 3;
      for (let y = ideal; y >= minCut; y -= 2) {
        const blocked = bands.some(
          (band) => band.top < y + clearance && band.bottom > y - clearance,
        );
        if (!blocked) return y;
      }
      return null;
    };
    const sliceBoundaries = () => {
      // A composition only slightly taller than one page reads as a
      // single-page design (e.g. a resume): render one full slice and let
      // the renderer's page scale absorb the few percent, instead of
      // stranding a couple of lines on a second page.
      if (bodyHeight <= sliceBudget * 1.12) return [Math.round(bodyHeight)];
      const { lineBands, boxBands } = collectCutBands();
      const allBands = [...lineBands, ...boxBands];
      const boundaries = [];
      let cursor = 0;
      while (bodyHeight - cursor > sliceBudget) {
        const ideal = cursor + sliceBudget;
        const minCut = cursor + Math.round(sliceBudget * 0.6);
        // Whole cards move to the next page when a gap exists; otherwise cut
        // between text lines inside the card; otherwise hard-cut.
        const cut =
          chooseCut(ideal, minCut, allBands) ??
          chooseCut(ideal, minCut, lineBands) ??
          ideal;
        boundaries.push(cut);
        cursor = cut;
      }
      boundaries.push(Math.round(bodyHeight));
      // Trailing sliver: absorb a short final slice into the previous page
      // when the merged slice still fits with a small page-scale shrink.
      if (boundaries.length >= 2) {
        const previousStart = boundaries.length >= 3 ? boundaries[boundaries.length - 3] : 0;
        const lastStart = boundaries[boundaries.length - 2];
        if (
          bodyHeight - lastStart <= sliceBudget * 0.25 &&
          bodyHeight - previousStart <= sliceBudget * 1.12
        ) {
          boundaries.splice(boundaries.length - 2, 1);
        }
      }
      return boundaries;
    };
    const screenshotPageSlices = ({ authoredPageHeights = false } = {}) => {
      // Absolutely-positioned multi-page canvases are authored in exact
      // viewport-height pages: gap-aware cuts would shift authored page
      // boundaries, so slice at the authored pitch and let the renderer
      // headroom shrink each page to fit.
      const boundaries = authoredPageHeights
        ? Array.from(
            { length: Math.ceil(bodyHeight / renderedPageHeight) },
            (_, index) => Math.min(bodyHeight, (index + 1) * renderedPageHeight),
          )
        : sliceBoundaries();
      // Decorative absolutes bleeding past the body's box (half-out hero
      // circles) get flat-cut by a body-width clip. Extend the horizontal
      // extent to overhanging descendants that still touch the body band —
      // capped so parked off-screen decorations (left/right: -10000px)
      // never drag the clip away.
      let paintLeft = bodyRect.left;
      let paintRight = bodyRect.right;
      const maxBleed = 200;
      for (const d of document.body.querySelectorAll('*')) {
        const dr = d.getBoundingClientRect();
        if (!dr.width || !dr.height) continue;
        if (dr.right <= bodyRect.left - maxBleed || dr.left >= bodyRect.right + maxBleed) continue;
        if (dr.left < paintLeft) paintLeft = Math.max(dr.left, bodyRect.left - maxBleed);
        if (dr.right > paintRight) paintRight = Math.min(dr.right, bodyRect.right + maxBleed);
      }
      paintLeft = Math.max(0, Math.round(paintLeft));
      const paintWidth = Math.round(paintRight - paintLeft);
      let previousBoundary = 0;
      return boundaries.map((boundary) => {
        const sliceHeight = Math.round(boundary - previousBoundary);
        const sliceTop = previousBoundary;
        previousBoundary = boundary;
        return {
          type: 'image',
          shotId: nextShotId(),
          width: paintWidth,
          height: sliceHeight,
          align: 'center',
          pageComposition: true,
          capturesDocumentText: true,
          clip: {
            x: paintLeft,
            y: Math.max(0, Math.round(bodyRect.top + sliceTop)),
            width: paintWidth,
            height: sliceHeight,
          },
        };
      });
    };
    const buildMixedPageIr = () => {
      const output = [];
      const pageRoots = new Set(inferredPageRoots);
      let pendingPageBreak = false;
      let previousPageRootFillsPage = false;
      const append = (nodes) => {
        for (const node of nodes) {
          if (
            node.type === 'pagebreak' &&
            (!output.length || output[output.length - 1].type === 'pagebreak')
          ) {
            continue;
          }
          output.push(node);
        }
      };
      for (const child of visibleBodyChildren) {
        const style = cs(child);
        const isPageRoot = pageRoots.has(child);
        const breaksBefore =
          style.breakBefore === 'page' || style.pageBreakBefore === 'always';
        const naturalPageAdvance = pendingPageBreak && previousPageRootFillsPage;
        if (
          !naturalPageAdvance &&
          (pendingPageBreak || (isPageRoot && breaksBefore)) &&
          output.length
        ) {
          append([{ type: 'pagebreak' }]);
        }
        const nodes = isPageRoot
          ? [{ ...markForScreenshot(child), pageComposition: true }]
          : processElement(child, 0);
        if (naturalPageAdvance && nodes[0]?.type === 'pagebreak') nodes.shift();
        append(nodes);
        pendingPageBreak =
          style.breakAfter === 'page' || style.pageBreakAfter === 'always';
        previousPageRootFillsPage =
          isPageRoot &&
          child.getBoundingClientRect().height >= renderedPageHeight - 12;
      }
      return output;
    };
    const ir = multiPageComposition
      ? authoredPageRoots.flatMap((page, index) => [
          ...(index > 0 ? [{ type: 'pagebreak' }] : []),
          { ...markForScreenshot(page), pageComposition: true },
        ])
      : slidePageComposition
        ? slidePageRoots.flatMap((page, index) => [
            ...(index > 0 ? [{ type: 'pagebreak' }] : []),
            { ...markForScreenshot(page), pageComposition: true },
          ])
      : mixedPageComposition
        ? buildMixedPageIr()
        : webPageComposition
          ? screenshotPageSlices()
          : multiPageAbsoluteComposition
            ? screenshotPageSlices({ authoredPageHeights: true })
            : flowingVisualComposition
              ? screenshotPageSlices()
            : singlePageVisualComposition
              ? [{ ...markForScreenshot(singleVisualPageRoot), pageComposition: true }]
              : bodySinglePageVisualComposition
                ? [{ ...markForScreenshot(document.body), pageComposition: true }]
              : pageComposition
                ? [{ ...markForScreenshot(document.body), pageComposition: true }]
                : processChildren(document.body, 0);
    while (ir.length && ir[0].type === 'pagebreak') ir.shift();
    for (let index = ir.length - 1; index >= 0; index--) {
      const node = ir[index];
      if (
        node.type !== 'image' ||
        node.clip ||
        node.pageComposition ||
        node.height <= 1000
      ) {
        continue;
      }
      const element = document.querySelector(`[data-h2d-id="${node.shotId}"]`);
      if (!element) continue;
      const rect = element.getBoundingClientRect();
      const maxSliceHeight = 700;
      const sliceCount = Math.ceil(rect.height / maxSliceHeight);
      const slices = Array.from({ length: sliceCount }, (_, sliceIndex) => {
        const sliceHeight = Math.min(
          maxSliceHeight,
          rect.height - sliceIndex * maxSliceHeight,
        );
        return {
          ...node,
          shotId: sliceIndex === 0 ? node.shotId : nextShotId(),
          height: Math.round(sliceHeight),
          spacingBeforePx: sliceIndex === 0 ? node.spacingBeforePx : 0,
          spacingAfterPx:
            sliceIndex === sliceCount - 1 ? node.spacingAfterPx : 0,
          clip: {
            x: Math.max(0, Math.round(rect.left)),
            y: Math.max(0, Math.round(rect.top + sliceIndex * maxSliceHeight)),
            width: Math.round(rect.width),
            height: Math.round(sliceHeight),
          },
        };
      });
      ir.splice(index, 1, ...slices);
    }

    // A centered "paper sheet" holding all the content on a tinted desk
    // backdrop (body bg + one max-width/box-shadow child): the desk color is
    // screen chrome, the sheet IS the page — take its paint instead.
    const paperSheet = (() => {
      const kids = [...document.body.children].filter(
        (child) =>
          isVisible(child) &&
          ((child.innerText || '').trim() || child.querySelector('img,svg,canvas,table')),
      );
      if (kids.length !== 1) return null;
      const el = kids[0];
      const st = cs(el);
      if (!toHex(st.backgroundColor, { dropWhite: false })) return null;
      const rect = el.getBoundingClientRect();
      const centered = Math.abs(rect.left - (window.innerWidth - rect.right)) <= 24;
      const coversContent = rect.height >= document.body.scrollHeight * 0.8;
      const sheetLike = st.boxShadow !== 'none' || parseFloat(st.maxWidth) > 0;
      return centered && coversContent && sheetLike ? el : null;
    })();
    const bodyBackground = toHex(cs(document.body).backgroundColor, {
      dropWhite: false,
    });
    // The desk-vs-sheet call only holds for neutral backdrops: a saturated
    // color or gradient behind the sheet is poster design, not screen chrome.
    const backdropDesigned = (() => {
      if (cs(document.body).backgroundImage.includes('gradient')) return true;
      if (!bodyBackground) return false;
      const channels = [0, 2, 4].map((i) => parseInt(bodyBackground.slice(i, i + 2), 16));
      return Math.max(...channels) - Math.min(...channels) > 24;
    })();
    const deskSheet = backdropDesigned ? null : paperSheet;
    const pageBg = deskSheet
      ? toHex(cs(deskSheet).backgroundColor, { dropWhite: true })
      : bodyBackground
        ? toHex(cs(document.body).backgroundColor, { dropWhite: true })
        : toHex(cs(document.documentElement).backgroundColor, { dropWhite: true });
    const bodyHasPaint =
      !deskSheet &&
      (cs(document.body).backgroundImage !== 'none' || Boolean(bodyBackground));
    const bgSource = bodyHasPaint ? document.body : document.documentElement;
    const bgStyle = cs(bgSource);
    const pagePseudoPaints = (which) => {
      const pseudo = getComputedStyle(bgSource, which);
      if (!pseudo || pseudo.content === 'none' || pseudo.content === 'normal') return false;
      const hasBorder = ['Top', 'Right', 'Bottom', 'Left'].some(
        (side) =>
          parseFloat(pseudo[`border${side}Width`]) >= 1 &&
          pseudo[`border${side}Style`] !== 'none',
      );
      return (
        hasBorder ||
        pseudo.backgroundImage !== 'none' ||
        Boolean(toHex(pseudo.backgroundColor, { dropWhite: true }))
      );
    };
    // a flat solid body color needs no backdrop screenshot: w:background
    // renders the exact hex, while the header-image route's resampled tone
    // visibly differs from same-color shading fills on content blocks
    const backdropHasArt =
      bgStyle.backgroundImage !== 'none' ||
      pagePseudoPaints('::before') ||
      pagePseudoPaints('::after') ||
      Boolean(paperSheet && backdropDesigned);
    if (
      !pageComposition &&
      !multiPageComposition &&
      !mixedPageComposition &&
      !multiPageAbsoluteComposition &&
      !webPageComposition &&
      !flowingVisualComposition &&
      !singlePageVisualComposition &&
      !bodySinglePageVisualComposition &&
      backdropHasArt
    ) {
      const backdrop = document.createElement('div');
      const width = window.innerWidth;
      const height = window.innerHeight;
      // park off-screen on the side the document ignores for overflow:
      // in RTL, left-side overflow extends the scrollable area and shifts
      // the capture surface origin, blanking every later element screenshot
      const rtlDoc =
        cs(document.documentElement).direction === 'rtl' ||
        cs(document.body).direction === 'rtl';
      Object.assign(backdrop.style, {
        position: 'absolute',
        [rtlDoc ? 'right' : 'left']: '-10000px',
        top: '0',
        width: `${width}px`,
        height: `${height}px`,
        backgroundColor: pageBg ? `#${pageBg}` : bgStyle.backgroundColor,
        backgroundImage: bgStyle.backgroundImage,
        backgroundSize: bgStyle.backgroundSize,
        backgroundPosition: bgStyle.backgroundPosition,
        backgroundRepeat: bgStyle.backgroundRepeat,
        pointerEvents: 'none',
      });
      const appendPagePseudo = (which) => {
        const pseudo = getComputedStyle(bgSource, which);
        if (!pseudo || pseudo.content === 'none' || pseudo.content === 'normal') return;
        const hasBorder = ['Top', 'Right', 'Bottom', 'Left'].some(
          (side) =>
            parseFloat(pseudo[`border${side}Width`]) >= 1 &&
            pseudo[`border${side}Style`] !== 'none',
        );
        const hasPaint =
          pseudo.backgroundImage !== 'none' ||
          Boolean(toHex(pseudo.backgroundColor, { dropWhite: true }));
        if (!hasBorder && !hasPaint) return;
        const overlay = document.createElement('div');
        Object.assign(overlay.style, {
          position: 'absolute',
          top: pseudo.top === 'auto' ? '0' : pseudo.top,
          right: pseudo.right === 'auto' ? '0' : pseudo.right,
          bottom: pseudo.bottom === 'auto' ? '0' : pseudo.bottom,
          left: pseudo.left === 'auto' ? '0' : pseudo.left,
          boxSizing: 'border-box',
          backgroundColor: pseudo.backgroundColor,
          backgroundImage: pseudo.backgroundImage,
          backgroundSize: pseudo.backgroundSize,
          backgroundPosition: pseudo.backgroundPosition,
          backgroundRepeat: pseudo.backgroundRepeat,
          borderTop: pseudo.borderTop,
          borderRight: pseudo.borderRight,
          borderBottom: pseudo.borderBottom,
          borderLeft: pseudo.borderLeft,
          opacity: pseudo.opacity,
        });
        backdrop.appendChild(overlay);
      };
      appendPagePseudo('::before');
      appendPagePseudo('::after');
      // Designed backdrop with a content sheet on top (poster on gradient):
      // paint the sheet's box into the backdrop so page chrome = backdrop +
      // sheet, and the flowing text keeps its readable card behind it.
      if (paperSheet && backdropDesigned) {
        const sheetRect = paperSheet.getBoundingClientRect();
        const sheetStyle = cs(paperSheet);
        const sheet = document.createElement('div');
        Object.assign(sheet.style, {
          position: 'absolute',
          left: `${Math.round(sheetRect.left)}px`,
          top: `${Math.round(Math.max(0, sheetRect.top))}px`,
          width: `${Math.round(sheetRect.width)}px`,
          height: `${Math.round(Math.max(100, height - Math.max(0, sheetRect.top)))}px`,
          backgroundColor: sheetStyle.backgroundColor,
          borderRadius: sheetStyle.borderRadius,
          boxShadow: sheetStyle.boxShadow,
        });
        backdrop.appendChild(sheet);
      }
      document.body.appendChild(backdrop);
      const shot = markForScreenshot(backdrop);
      ir.unshift({
        type: 'pagebg',
        color: pageBg,
        shotId: shot.shotId,
        width,
        height,
      });
    } else if (pageBg) {
      ir.unshift({ type: 'pagebg', color: pageBg });
    }
    for (let index = ir.length - 1; index > 0; index--) {
      if (ir[index].type === 'pagebreak' && ir[index - 1].type === 'spacer') {
        ir.splice(index - 1, 1);
      }
    }
    // Section boundaries stack spacers (padding-bottom + divider walk +
    // padding-top emitted by different branches): the real CSS gap is the
    // two paddings, so keep the two largest of a run and drop the rest.
    for (let index = ir.length - 1; index >= 0; index--) {
      if (ir[index].type !== 'spacer' || ir[index].bindsNeighbors) continue;
      let first = index;
      while (first > 0 && ir[first - 1].type === 'spacer' && !ir[first - 1].bindsNeighbors) first--;
      if (first === index) continue;
      const run = ir.slice(first, index + 1).map((node) => node.px || 0).sort((a, b) => b - a);
      const merged = Math.min(260, run[0] + (run[1] || 0));
      ir.splice(first, index - first + 1, { type: 'spacer', px: merged });
      index = first;
    }
    const lastContent = ir[ir.length - 1];
    if (lastContent?.spacingAfterPx != null) lastContent.spacingAfterPx = 0;
    if (lastContent?.style?.spacingAfterPx != null) lastContent.style.spacingAfterPx = 0;

    const bodyStyle = cs(document.body);
    const cssLengthPx = (value) => {
      const match = String(value || '').trim().match(/^([\d.]+)(px|pt|in|cm|mm)?$/i);
      if (!match) return null;
      const amount = parseFloat(match[1]);
      return (
        amount *
        ({ px: 1, pt: 96 / 72, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 }[
          (match[2] || 'px').toLowerCase()
        ] || 1)
      );
    };
    const namedPageSizes = {
      a3: [1123, 1587],
      a4: [794, 1123],
      a5: [559, 794],
      letter: [816, 1056],
      legal: [816, 1344],
      tabloid: [1056, 1632],
      ledger: [1632, 1056],
    };
    let pageSizePx = null;
    for (const sheet of [...document.styleSheets]) {
      let rules;
      try {
        rules = [...sheet.cssRules];
      } catch {
        continue;
      }
      for (const rule of rules) {
        if (rule.type !== CSSRule.PAGE_RULE) continue;
        const parts = (rule.style.getPropertyValue('size') || '').trim().toLowerCase().split(/\s+/);
        const orientation = parts.find((part) => part === 'landscape' || part === 'portrait');
        const named = parts.find((part) => namedPageSizes[part]);
        if (named) {
          pageSizePx = [...namedPageSizes[named]];
        } else {
          const lengths = parts.map(cssLengthPx).filter((value) => value != null);
          if (lengths.length >= 2) pageSizePx = [lengths[0], lengths[1]];
          else if (lengths.length === 1) pageSizePx = [lengths[0], lengths[0] * Math.SQRT2];
        }
        if (pageSizePx && orientation === 'landscape' && pageSizePx[0] < pageSizePx[1]) {
          pageSizePx.reverse();
        }
        if (pageSizePx && orientation === 'portrait' && pageSizePx[0] > pageSizePx[1]) {
          pageSizePx.reverse();
        }
        break;
      }
      if (pageSizePx) break;
    }
    if (!pageSizePx && (multiPageComposition || mixedPageComposition)) {
      const pageRect = (authoredPageRoots[0] || inferredPageRoots[0]).getBoundingClientRect();
      pageSizePx = [pageRect.width, pageRect.height];
    } else if (!pageSizePx && multiPageAbsoluteComposition) {
      pageSizePx = [bodyRect.width, renderedPageHeight];
    } else if (!pageSizePx && pageComposition) {
      pageSizePx = [bodyRect.width, bodyRect.height];
    }
    pageSizePx ||= [794, 1123];
    const composed =
      pageComposition ||
      multiPageComposition ||
      slidePageComposition ||
      mixedPageComposition ||
      multiPageAbsoluteComposition ||
      webPageComposition ||
      flowingVisualComposition ||
      singlePageVisualComposition ||
      bodySinglePageVisualComposition;
    // A natively-rendered flow document slightly taller than one page (e.g. a
    // resume that strands two lines on page 2) reads as a one-page design.
    // Word lays such docs out ~10% denser than the browser proxy (narrower
    // Arial + width scale), so browser ratios up to ~1.22 are recoverable
    // with a small whole-document squeeze applied to geometry and fonts.
    // ratio against the page height at the ACTUAL viewport width — the
    // viewport height is not resized when the authored width differs, so
    // innerHeight is the wrong denominator there (880-wide doc in a
    // 1024-proportioned window read as 0.95 pages instead of 1.11)
    const flowCanvasWidth = Math.min(
      window.innerWidth,
      Math.max(Math.round(pageSizePx[0]), Math.ceil(bodyRect.width)),
    );
    const flowPageHeight = flowCanvasWidth * (pageSizePx[1] / pageSizePx[0]);
    const flowRatio = bodyHeight / flowPageHeight;
    const squeezableFlow = !composed && !authoredPageRoots.length && !inferredPageRoots.length;
    (globalThis.__h2dTrace ||= []).push({
      where: 'pageFitSqueeze',
      flowRatio: Math.round(flowRatio * 1000) / 1000,
      bodyHeight: Math.round(bodyHeight),
      renderedPageHeight,
      composed: Boolean(composed),
      authoredRoots: authoredPageRoots.length,
      inferredRoots: inferredPageRoots.length,
    });
    // Word lays text-heavy docs ~10% denser than the browser proxy (narrower
    // Arial), which the 0.9 fudge credits. Table/fixed-height-heavy docs get
    // no such densening, so past ~1.05 the fudge under-squeezes and the last
    // block still strands (invoice stamp boxes) — taper the credit away.
    const pageFitSqueeze =
      squeezableFlow && flowRatio > 1.05 && flowRatio <= 1.22
        ? // table/fixed-height-heavy docs (invoices) get none of the Arial
          // densening the 0.9 fudge below assumes — squeeze at face value
          // plus slack so the trailing block doesn't strand by one line
          // 0.90/0.82 (was 0.92/0.84): Word (Mac) lays these table-heavy
          // docs up to ~9% taller than the browser proxy — the old floor
          // left a marginal one-pager's footer stranding on page 2
          Math.max(0.8, Math.min(1, 0.88 / flowRatio))
        : squeezableFlow && flowRatio > 1.02 && flowRatio <= 1.05
          ? Math.max(0.88, Math.min(1, 0.98 / (flowRatio * 0.9)))
          : // A doc filling exactly one browser page (invoices, forms) is a
            // one-page design; Word's layout runs a few percent taller (table
            // leading, min row heights), so squeeze preemptively.
            squeezableFlow && flowRatio > 0.96 && flowRatio <= 1.02
            ? 0.94
            : 1;
    // A viewport wider than the authored canvas (1024 reload of a
    // letter-width design) centers the body with auto margins — a viewport
    // artifact, not an authored page margin. Measure against the canvas and
    // strip the centering offset, or the page gets ~3cm fake margins and the
    // content shrinks by the canvas/viewport ratio.
    const canvasWidth = Math.min(
      window.innerWidth,
      Math.max(Math.round(pageSizePx[0]), Math.ceil(bodyRect.width)),
    );
    const centerOffset = Math.max(0, (window.innerWidth - canvasWidth) / 2);
    // flowed documents get a small margin floor: an authored padding:0 body
    // puts Word text hard against the physical page edge, which no print
    // layout intends; composed pages keep exact geometry. The floor widens
    // the margins beyond the measured geometry, so the reported canvas must
    // widen by the same amount or the px->page scale over-sizes the content
    // and Word clips it at the right margin.
    const rawLeftPx = composed
      ? Math.max(0, Math.round(bodyRect.left - centerOffset))
      : Math.max(
          0,
          Math.round(bodyRect.left - centerOffset + (parseFloat(bodyStyle.paddingLeft) || 0)),
        );
    const rawRightPx = Math.max(
      0,
      composed
        ? Math.round(canvasWidth - (bodyRect.right - centerOffset))
        : Math.round(
            canvasWidth -
              (bodyRect.right - centerOffset) +
              (parseFloat(bodyStyle.paddingRight) || 0),
          ),
    );
    const leftPx = composed ? rawLeftPx : Math.max(24, rawLeftPx);
    const rightPx = composed ? rawRightPx : Math.max(24, rawRightPx);
    const metaContent = (name) =>
      document.querySelector(`meta[name="${name}" i]`)?.content?.trim() || undefined;
    ir.unshift({
      type: 'docsettings',
      viewportWidthPx: canvasWidth + (leftPx - rawLeftPx) + (rightPx - rawRightPx),
      rtl: docRTL || undefined,
      lang: document.documentElement.lang || undefined,
      meta: {
        title: (document.title || '').trim() || undefined,
        author: metaContent('author'),
        description: metaContent('description'),
        keywords: metaContent('keywords'),
      },
      composed: composed || undefined,
      pageFitSqueeze: pageFitSqueeze < 1 ? pageFitSqueeze : undefined,
      pageSizePx: { width: Math.round(pageSizePx[0]), height: Math.round(pageSizePx[1]) },
      marginsPx: {
        top: composed
          ? Math.max(0, Math.round(bodyRect.top))
          : Math.max(24, Math.round(bodyRect.top + (parseFloat(bodyStyle.paddingTop) || 0))),
        bottom: composed
          ? 0
          : Math.max(24, Math.round(parseFloat(bodyStyle.paddingBottom) || 0)),
        left: leftPx,
        right: rightPx,
      },
    });
    return ir;
  }

  global.__html2docxPages = { build };
})(globalThis);
