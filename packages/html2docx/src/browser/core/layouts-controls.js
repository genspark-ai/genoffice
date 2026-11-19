  function isBoxRow(el) {
    const s = cs(el);
    const kids = [...el.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    // bare text between element children (badge + question text + hint):
    // column extraction iterates elements only and silently drops the text
    for (const node of el.childNodes) {
      if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) return false;
    }
    const navRow = el.tagName === 'NAV' && s.display === 'flex';
    const singleGridRow = s.display === 'grid' && s.gridTemplateColumns.split(' ').length >= 2;
    // Narrow separators (▶ arrows between process steps) are not columns —
    // exclude them from the column-count cap or a 4-step flow with 3 arrows
    // (7 children) falls back to a vertical card stack.
    const columnCount = kids.filter((k) => k.getBoundingClientRect().width > 40).length;
    if (kids.length < 2 || columnCount > (navRow || singleGridRow ? 12 : 6)) return false;
    const flexRow = s.display === 'flex' && (s.flexDirection === 'row' || s.flexDirection === '');
    const gridCols = s.display === 'grid' && (s.gridTemplateColumns.split(' ').length >= 2);
    // CSS table layout (display: table + table-cell children) is a
    // side-by-side row exactly like a flex row
    const cssTableRow =
      s.display === 'table' && kids.every((k) => cs(k).display === 'table-cell');
    if (!flexRow && !gridCols && !cssTableRow) return false;
    // Page-level column layouts (sidebar + main) are far too tall for one
    // table row — Word/LO can't paginate it. Those flatten to sequential flow.
    // A row of similar-width parallel cards (pricing tiers, lead + stacked
    // cells) is not a page layout: page-level two-column designs are consumed
    // by the earlier pageColumns branch and sidebars fail the width-similarity
    // check, so allow up to the ~one-page cell budget gridRowClusters uses.
    const kidHeights = kids.map((k) => k.getBoundingClientRect().height);
    const kidWidths = kids.map((k) => k.getBoundingClientRect().width).filter((w) => w > 40);
    const parallelCards =
      kidWidths.length >= 2 && Math.max(...kidWidths) <= Math.min(...kidWidths) * 1.8;
    if (Math.max(...kidHeights) > (parallelCards ? 950 : 500)) return false;
    // children laid out side by side: every pair overlaps vertically
    // (tops alone fail for align-items: center/end rows of unequal heights)
    const rects = kids.map((k) => k.getBoundingClientRect());
    // RTL flex rows progress right-to-left in DOM order
    const ltr = rects.length < 2 || rects[1].left >= rects[0].left;
    for (let i = 0; i < rects.length - 1; i++) {
      const a = rects[i];
      const b = rects[i + 1];
      if (Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) < Math.min(a.height, b.height) * 0.5) {
        return false;
      }
      // must progress consistently in one horizontal direction
      if (ltr && b.left < a.right - 2) return false;
      if (!ltr && b.right > a.left + 2) return false;
    }
    return true;
  }

  // Multi-row grid of cards (2-4 columns, 2+ rows): children clustered into
  // rows by rendered Y position -> borderless layout table, one table row per
  // visual row, so the side-by-side arrangement survives and Word can still
  // paginate between rows.
  function gridRowClusters(el) {
    const s = cs(el);
    const gridCols = s.display === 'grid' && s.gridTemplateColumns.split(' ').length >= 2;
    const wrapFlex = s.display === 'flex' && s.flexWrap === 'wrap';
    if (!gridCols && !wrapFlex) return null;
    const kids = [...el.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    if (kids.length < 3) return null;
    // grid-area placement makes DOM order diverge from visual order —
    // cluster by rendered position, not source position
    kids.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
    const rows = [];
    let cur = null;
    let curTop = -1e9;
    for (const k of kids) {
      const t = k.getBoundingClientRect().top;
      if (!cur || t - curTop > 20) {
        cur = [];
        rows.push(cur);
        curTop = t;
      }
      cur.push(k);
    }
    if (rows.length < 2) return null;
    const nCols = Math.max(...rows.map((r) => r.length));
    if (nCols < 2 || nCols > 13) return null;
    // A table row can't paginate mid-cell, so cells must fit one page. Word's
    // usable content height is ~970px — cards up to ~950px still render as a
    // single side-by-side row (a 2×2 teacher-card grid has ~780px cells).
    if (kids.some((k) => k.getBoundingClientRect().height > 950)) return null;
    if (kids.some((k) => k.getBoundingClientRect().height > 700)) {
      // Page-tall rows are only safe for grids of self-painted cards; a
      // painted section's internal text layout (dark manifesto chapters)
      // must stay with the card/section branches that keep its backdrop.
      const paintedKids = kids.filter(
        (k) => bgHex(k) || paintedBgHex(k) || fullBorderOf(k),
      ).length;
      if (paintedKids * 2 < kids.length) return null;
    }
    // physical left-to-right cell order (RTL grids arrive right-to-left)
    for (const row of rows) {
      row.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
    }
    return rows;
  }

  function largeTableGridRows(el) {
    const s = cs(el);
    const gridCols = s.display === 'grid' && s.gridTemplateColumns.split(' ').length >= 2;
    const wrapFlex = s.display === 'flex' && s.flexWrap === 'wrap';
    if (!gridCols && !wrapFlex) return null;
    if (pageColumnParts(el)) return null;
    const kids = [...el.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    if (kids.length < 3 || kids.length > 8) return null;
    const tableKids = kids.filter((kid) => kid.querySelector('table'));
    if (tableKids.length < 2 || tableKids.length * 2 < kids.length) return null;
    if (!kids.some((kid) => kid.getBoundingClientRect().height > 700)) return null;
    if (el.getBoundingClientRect().height > 2200) return null;
    // same as gridRowClusters: visual order, not DOM order
    kids.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
    const rows = [];
    let current = null;
    let currentTop = -1e9;
    for (const kid of kids) {
      const top = kid.getBoundingClientRect().top;
      if (!current || top - currentTop > 20) {
        current = [];
        rows.push(current);
        currentTop = top;
      }
      current.push(kid);
    }
    if (rows.length < 2 || rows.length > 4) return null;
    const columnCount = Math.max(...rows.map((row) => row.length));
    if (columnCount < 2 || columnCount > 4) return null;
    for (const row of rows) {
      row.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
    }
    return rows;
  }

  function extractGridRows(
    el,
    gridRows,
    depth,
    { screenshotCells = false, noSplit = true, largeTableGrid = false } = {},
  ) {
    // Column tracks are clustered from cell left edges across ALL rows, not
    // just row 0: rows can occupy different columns (timeline layouts where
    // cards alternate left/right and empty divs are invisible), and taking
    // row 0's cells as the column set mis-registers every other row.
    const containerRect = el.getBoundingClientRect();
    const trackLefts = [];
    for (const c of gridRows.flat()) {
      const left = c.getBoundingClientRect().left;
      if (!trackLefts.some((x) => Math.abs(x - left) < 12)) trackLefts.push(left);
    }
    trackLefts.sort((a, b) => a - b);
    const trackIndex = (x) => {
      let best = 0;
      for (let i = 1; i < trackLefts.length; i++) {
        if (Math.abs(trackLefts[i] - x) < Math.abs(trackLefts[best] - x)) best = i;
      }
      return best;
    };
    let colWidths = trackLefts.map((left, i) =>
      Math.round((i + 1 < trackLefts.length ? trackLefts[i + 1] : containerRect.right) - left),
    );
    const containerBg = bgHex(el);
    const cellFor = (c) => {
      if (screenshotCells) return { children: [markForScreenshot(c)] };
      // screenshot cells carry their own paint - no cell shading on top
      if (
        isReplacedTag(c) ||
        isPill(c) ||
        isDecorativeLeaf(c) ||
        isBadgeGrid(c) ||
        needsVisualEffectScreenshot(c)
      ) {
        return { children: processElement(c, depth + 1) };
      }
      // a cell that is itself a flex row or a visual card needs the full
      // classifier: rows collapse into a stack and cards lose their
      // border/background when only their children are walked
      return {
        children: isBoxRow(c) || isCard(c, depth + 1)
          ? processElement(c, depth + 1)
          : needsBlockWalk(c)
            ? processChildren(c, depth + 1, { decorations: false })
            : paraNodesFrom(c, false),
        // on a tinted/gradient backdrop, white card paint must survive
        shading: paintedBgHex(c),
      };
    };
    const rows = gridRows.map((cells) => {
      const rowCells = [];
      let pos = 0;
      for (const c of cells) {
        const rect = c.getBoundingClientRect();
        const start = Math.max(pos, trackIndex(rect.left));
        let span = 1;
        while (start + span < trackLefts.length && trackLefts[start + span] < rect.right - 12) {
          span += 1;
        }
        // empty filler for grid columns this row does not occupy
        if (start > pos) {
          rowCells.push({ children: [], gridStart: pos, colspan: start - pos });
        }
        rowCells.push({ ...cellFor(c), gridStart: start, colspan: span });
        pos = start + span;
      }
      if (pos < trackLefts.length) {
        rowCells.push({ children: [], gridStart: pos, colspan: trackLefts.length - pos });
      }
      return {
        // planner-style rows of mostly-empty cells must keep their drawn height
        heightPx:
          cells.reduce((n, c) => n + (c.innerText || '').trim().length, 0) < 30
            ? Math.min(420, Math.max(...cells.map((c) => Math.round(c.getBoundingClientRect().height))))
            : undefined,
        cells: rowCells,
      };
    });
    const s = cs(el);
    const rowGapPx = Math.min(parseFloat(s.rowGap) || 0, 60);
    const colGapPx = Math.min(parseFloat(s.columnGap) || 0, 60);
    // Authored grid gaps become explicit gap tracks/rows shaded with the
    // container's background: as interior cell margins the gap would be
    // filled by the cell's own (e.g. white card) shading and visually vanish.
    if (colGapPx >= 8 && trackLefts.length >= 2) {
      const totalTracks = trackLefts.length * 2 - 1;
      for (const row of rows) {
        const remapped = [];
        let pos = 0;
        for (const cell of row.cells) {
          const start = cell.gridStart * 2;
          if (start > pos) {
            remapped.push({ children: [], shading: containerBg, gridStart: pos, colspan: start - pos });
          }
          const span = cell.colspan * 2 - 1;
          remapped.push({
            ...cell,
            gridStart: start,
            colspan: span,
            // fillers and transparent cells show the container tint in HTML
            shading: cell.shading ?? containerBg,
          });
          pos = start + span;
        }
        if (pos < totalTracks) {
          remapped.push({ children: [], shading: containerBg, gridStart: pos, colspan: totalTracks - pos });
        }
        row.cells = remapped;
      }
      const expanded = [];
      colWidths.forEach((w, i) => {
        if (i < colWidths.length - 1) {
          expanded.push(Math.max(10, w - colGapPx), colGapPx);
        } else {
          expanded.push(w);
        }
      });
      colWidths = expanded;
    }
    if (rowGapPx >= 8) {
      const totalTracks = colWidths.length;
      const spaced = [];
      rows.forEach((row, i) => {
        spaced.push(row);
        if (i < rows.length - 1) {
          spaced.push({
            gapRow: true,
            heightPx: rowGapPx,
            cells: [{ children: [], shading: containerBg, gridStart: 0, colspan: totalTracks }],
          });
        }
      });
      return [{
        type: 'table',
        layout: true,
        noSplit,
        largeTableGrid,
        rowGapPx: 0,
        rows: spaced,
        colWidths,
      }];
    }
    return [{ type: 'table', layout: true, noSplit, largeTableGrid, rowGapPx, rows, colWidths }];
  }

  // Page-level side-by-side columns (sidebar + main). Reproduced as a
  // borderless multi-row table: content blocks from all columns are sliced
  // into rows by their rendered Y position, so side-by-side placement
  // survives while Word can still paginate between rows.
  function pageColumnParts(el) {
    const s = cs(el);
    const kids = [...el.children].filter(
      (c) =>
        !SKIP_TAGS.has(c.tagName) &&
        isVisible(c) &&
        cs(c).position !== 'absolute' &&
        cs(c).position !== 'fixed',
    );
    if (kids.length < 2 || kids.length > 4) return null;
    const flexRow = s.display === 'flex' && (s.flexDirection === 'row' || s.flexDirection === '');
    const gridCols = s.display === 'grid' && (s.gridTemplateColumns.split(' ').length >= 2);
    if (!flexRow && !gridCols) return null;
    const containerRect = el.getBoundingClientRect();
    const columns = kids.filter((k) => k.getBoundingClientRect().height > 500);
    if (columns.length < 2 || columns.length > 3) return null;
    const tops = columns.map((k) => Math.round(k.getBoundingClientRect().top));
    if (Math.max(...tops) - Math.min(...tops) >= 60) return null;
    const ordered = [...columns].sort(
      (a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left,
    );
    for (let i = 0; i < ordered.length - 1; i++) {
      if (ordered[i + 1].getBoundingClientRect().left < ordered[i].getBoundingClientRect().right - 2) {
        return null;
      }
    }
    const leading = kids.filter((k) => !columns.includes(k));
    if (
      leading.some((k) => {
        const r = k.getBoundingClientRect();
        return r.width < containerRect.width * 0.8 || r.bottom > Math.min(...tops) + 2;
      })
    ) {
      return null;
    }
    return { leading, columns: ordered };
  }

  // Flatten a column into addressable blocks with their rendered Y position.
  function columnBlocks(col, depth, out, colIdx) {
    let kids = [...col.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    // unwrap single-child wrappers
    while (kids.length === 1 && kids[0].children.length && !directText(col)) {
      col = kids[0];
      kids = [...col.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    }
    if (!kids.length) {
      const nodes = paraNodesFrom(col);
      if (nodes.length) {
        const r = col.getBoundingClientRect();
        out.push({ col: colIdx, top: r.top, bottom: r.bottom, nodes });
      }
      return;
    }
    for (const child of kids) {
      const r = child.getBoundingClientRect();
      // very tall composite blocks get split further so row slicing stays
      // fine-grained — but never into tables/grids/lists, whose internal
      // structure must reach the classifier whole
      const atomic =
        ['TABLE', 'UL', 'OL'].includes(child.tagName) ||
        cs(child).display === 'grid' ||
        child.querySelector('table');
      // 320 (not a full page): slicing needs block granularity well under a
      // page or two columns never share a legal cut line (479px job entries
      // left no common gap below page 1's limit and degraded the layout).
      if (r.height > 320 && !atomic && hasBlockChildren(child) && !bgHex(child) && !isBadgeGrid(child)) {
        columnBlocks(child, depth, out, colIdx);
        continue;
      }
      const nodes = processElement(child, depth + 1);
      if (!nodes.length) continue;
      // A tall bullet list is the classic slice blocker (a near-page atomic
      // block forces the sequential-flow fallback): split it per top-level
      // li so cuts can land between items. Ordered lists stay atomic —
      // splitting them would restart their numbering.
      if (
        r.height > 320 &&
        (child.tagName === 'UL' || child.tagName === 'OL') &&
        nodes.length === 1 &&
        nodes[0].type === 'list' &&
        nodes[0].items.every((item) => !item.ordered && item.liIndex != null)
      ) {
        const lis = [...child.children].filter(
          (li) => li.tagName === 'LI' && isVisible(li),
        );
        const groups = new Map();
        for (const item of nodes[0].items) {
          if (!groups.has(item.liIndex)) groups.set(item.liIndex, []);
          groups.get(item.liIndex).push(item);
        }
        const liBlocks = [];
        for (const [liIdx, items] of groups) {
          const li = lis[liIdx];
          if (!li) { liBlocks.length = 0; break; }
          const liRect = li.getBoundingClientRect();
          liBlocks.push({
            col: colIdx,
            top: liRect.top,
            bottom: liRect.bottom,
            nodes: [{ ...nodes[0], items }],
          });
        }
        if (liBlocks.length >= 2) {
          out.push(...liBlocks);
          continue;
        }
      }
      out.push({ col: colIdx, top: r.top, bottom: r.bottom, nodes });
    }
  }

  function directText(el) {
    for (const n of el.childNodes) {
      if (n.nodeType === Node.TEXT_NODE && n.textContent.trim()) return true;
    }
    return false;
  }

  // One single-row borderless table, one cell per column, each cell holding
  // the column's full content. The row is allowed to break across pages
  // (Word/LO default), which paginates naturally while keeping the exact
  // side-by-side geometry — far more faithful than slicing content into
  // pseudo-rows by Y position.
  function compactColumnNodes(nodes) {
    for (const node of nodes) {
      if (node.style) {
        if (node.style.spacingBeforePx) {
          node.style.spacingBeforePx = Math.max(1, Math.round(node.style.spacingBeforePx * 0.5));
        }
        if (node.style.spacingAfterPx) {
          node.style.spacingAfterPx = Math.max(1, Math.round(node.style.spacingAfterPx * 0.5));
        }
      }
      if (node.spacingBeforePx) {
        node.spacingBeforePx = Math.max(1, Math.round(node.spacingBeforePx * 0.5));
      }
      if (node.spacingAfterPx) {
        node.spacingAfterPx = Math.max(1, Math.round(node.spacingAfterPx * 0.5));
      }
      if (node.type === 'spacer') node.px = Math.max(2, Math.round(node.px * 0.5));
      if (node.children) compactColumnNodes(node.children);
      if (node.cells) {
        for (const cell of node.cells) compactColumnNodes(cell.children || []);
      }
      if (node.rows) {
        for (const row of node.rows) {
          for (const cell of row.cells) compactColumnNodes(cell.children || []);
        }
      }
    }
    return nodes;
  }

  // A page column's own 1px edge border is the visual divider between
  // columns — carry it onto the cell (borderLeftOf demands >=2px).
  function columnEdgeBorder(csc, side) {
    const widthPx = parseFloat(csc[`border${side}Width`]);
    if (!(widthPx >= 1) || csc[`border${side}Style`] === 'none') return null;
    const color = toHex(csc[`border${side}Color`]);
    return color ? { color, widthPx } : null;
  }

  // Word never splits a table row that contains a nested table, so a
  // taller-than-page single-row column layout strands page 1 blank and dumps
  // everything onto page 2+. Slice the columns into multiple rows at Y
  // positions no block straddles: each slice fits a page and Word paginates
  // between the rows.
  function slicedPageColumnRows(el, depth, cols, colWidths) {
    const blocks = [];
    cols.forEach((col, colIdx) => columnBlocks(col, depth, blocks, colIdx));
    if (!blocks.length) return null;
    return buildColumnRows(blocks, cols.length, colWidths, (ci, sliceIdx) => {
      const csc = cs(cols[ci]);
      return {
        shading: bgHex(cols[ci]),
        borderLeft: columnEdgeBorder(csc, 'Left'),
        borderRight: columnEdgeBorder(csc, 'Right'),
        gapAfterPx:
          ci < cols.length - 1
            ? Math.max(
                0,
                Math.round(
                  cols[ci + 1].getBoundingClientRect().left -
                    cols[ci].getBoundingClientRect().right,
                ),
              )
            : 0,
        padPx: {
          top: sliceIdx === 0 ? Math.round(parseFloat(csc.paddingTop)) : 0,
          left: Math.round(parseFloat(csc.paddingLeft)),
          right: Math.round(parseFloat(csc.paddingRight)),
        },
      };
    });
  }

  // Slice pre-collected column blocks into page-height rows so Word can
  // paginate between them while keeping side-by-side placement. cellStyle
  // supplies per-column paint/padding; children come from the sliced blocks.
  function buildColumnRows(
    blocks,
    colCount,
    colWidths,
    cellStyle,
    { allowSingleRow = false, ignoreHeaderOffset = false } = {},
  ) {
    const PAGE_H = 1123; // A4 @96dpi
    const topMin = Math.min(...blocks.map((b) => b.top));
    const bottomMax = Math.max(...blocks.map((b) => b.bottom));
    // Content above the columns (title header) eats into Word's first page:
    // the first slice must be short enough to fit the remainder, or the
    // whole first row slides to page 2 and page 1 is stranded again. For a
    // column block nested deep in a multi-page document the body-top distance
    // is meaningless (Word repositions the block per page), so skip it.
    const body = document.body.getBoundingClientRect();
    const offsetAbove = ignoreHeaderOffset
      ? 0
      : Math.max(0, topMin - body.top - (parseFloat(cs(document.body).paddingTop) || 0));
    const cuts = [];
    const straddlerAt = (y) => blocks.find((b) => b.top < y - 4 && b.bottom > y + 4);
    const bumpPast = (y) => {
      for (let guard = 0; guard < 200; guard++) {
        const straddler = straddlerAt(y);
        if (!straddler) break;
        y = straddler.bottom + 6;
      }
      return y;
    };
    const bumpAbove = (y) => {
      for (let guard = 0; guard < 200; guard++) {
        const straddler = straddlerAt(y);
        if (!straddler) break;
        y = straddler.top - 6;
      }
      return y;
    };
    // Word lays out the pre-column header taller than the browser (built-in
    // heading spacing, substituted font metrics), so the browser-measured
    // offset underestimates what page 1 has left. A row that misses the
    // real remainder slides whole to page 2 and strands page 1 — budget
    // conservatively; a shorter first row costs nothing.
    const page1Limit = topMin + PAGE_H - offsetAbove * 1.2 - 160;
    let target = topMin + Math.max(250, PAGE_H * 0.75 - offsetAbove * 1.2);
    while (target < bottomMax - 150) {
      let y = bumpPast(target);
      if (!cuts.length && y > page1Limit) {
        // An atomic straddler pushed the first cut below page 1's remaining
        // space. Cutting above it keeps a short-but-valid first row instead
        // of degrading the whole two-column layout to sequential flow.
        const up = bumpAbove(Math.min(target, page1Limit));
        if (up - topMin >= 150) y = up;
      }
      if (y >= bottomMax - 100) break;
      cuts.push(y);
      target = y + PAGE_H * 0.8;
    }
    // Column content that fits one page needs no cut: emit a single row when
    // the caller has no sequential-flow fallback (CSS multi-column), else
    // signal null so the tall-sidebar path can fall back.
    if (!cuts.length && !allowSingleRow) return null;
    // A first slice taller than page 1's remaining space strands page 1
    // anyway (atomic blocks forced the cut too far down) — signal the
    // caller to fall back to sequential flow.
    if (cuts.length && cuts[0] - topMin > PAGE_H - offsetAbove - 80) return null;
    const boundaries = [topMin - 1, ...cuts, bottomMax + 1];
    const rows = [];
    for (let s = 0; s + 1 < boundaries.length; s++) {
      const cells = Array.from({ length: colCount }, (_, ci) => {
        const cellBlocks = blocks
          .filter((b) => b.col === ci && b.top >= boundaries[s] && b.top < boundaries[s + 1])
          .sort((a, b) => a.top - b.top);
        return {
          ...cellStyle(ci, s),
          children: compactColumnNodes(cellBlocks.flatMap((b) => b.nodes)),
        };
      });
      rows.push({ cells });
    }
    return [{ type: 'table', layout: true, pageColumns: true, rows, colWidths }];
  }

  // CSS multi-column (`column-count` / `columns`): content flows into N
  // physical columns with no per-column element. flatten the direct children,
  // bucket each into a column by its rendered center, then slice into
  // page-height rows like the flex/grid column path.
  function cssMultiColumnParts(el) {
    const s = cs(el);
    if (s.display !== 'block' && s.display !== 'flow-root') return null;
    let colCount = parseInt(s.columnCount, 10);
    const hasColWidth = s.columnWidth && s.columnWidth !== 'auto';
    if ((!Number.isFinite(colCount) || colCount < 2) && !hasColWidth) return null;
    const kids = [...el.children].filter(
      (c) =>
        !SKIP_TAGS.has(c.tagName) &&
        isVisible(c) &&
        cs(c).position !== 'absolute' &&
        cs(c).position !== 'fixed',
    );
    // need enough items to actually populate more than one column
    if (kids.length < 4) return null;
    const lefts = kids.map((k) => Math.round(k.getBoundingClientRect().left));
    const buckets = [];
    for (const x of lefts) if (!buckets.some((b) => Math.abs(b - x) < 20)) buckets.push(x);
    if (buckets.length < 2) return null;
    if (!Number.isFinite(colCount) || colCount < 2) colCount = buckets.length;
    if (colCount < 2 || colCount > 4) return null;
    const gap = parseFloat(s.columnGap);
    const rect = el.getBoundingClientRect();
    const colGapPx = Number.isFinite(gap) ? Math.round(gap) : Math.round(rect.width * 0.04);
    return { colCount: Math.min(colCount, buckets.length), colGapPx };
  }

  function extractCssColumns(el, depth, colCount, colGapPx) {
    const rect = el.getBoundingClientRect();
    const s = cs(el);
    const padL = parseFloat(s.paddingLeft) || 0;
    const padR = parseFloat(s.paddingRight) || 0;
    const contentLeft = rect.left + padL;
    const contentWidth = Math.max(1, rect.width - padL - padR);
    const colPitch = (contentWidth + colGapPx) / colCount;
    const kids = [...el.children].filter(
      (c) =>
        !SKIP_TAGS.has(c.tagName) &&
        isVisible(c) &&
        cs(c).position !== 'absolute' &&
        cs(c).position !== 'fixed',
    );
    const blocks = [];
    for (const child of kids) {
      const r = child.getBoundingClientRect();
      let colIdx = Math.floor((r.left + r.width / 2 - contentLeft) / colPitch);
      colIdx = Math.max(0, Math.min(colCount - 1, colIdx));
      const nodes = processElement(child, depth + 1);
      if (nodes.length) blocks.push({ col: colIdx, top: r.top, bottom: r.bottom, nodes });
    }
    if (blocks.length < 2 || new Set(blocks.map((b) => b.col)).size < 2) return null;
    const colWidthPx = Math.round((contentWidth - colGapPx * (colCount - 1)) / colCount);
    const colWidths = Array.from({ length: colCount }, () => colWidthPx);
    const containerShading = bgHex(el);
    return buildColumnRows(
      blocks,
      colCount,
      colWidths,
      (ci, sliceIdx) => ({
        shading: containerShading,
        gapAfterPx: ci < colCount - 1 ? colGapPx : 0,
        padPx: { top: sliceIdx === 0 ? Math.round(parseFloat(s.paddingTop) || 0) : 0, left: 0, right: 0 },
      }),
      { allowSingleRow: true, ignoreHeaderOffset: true },
    );
  }

  function extractPageColumns(el, depth, cols) {
    const colWidths = cols.map((c) => Math.round(c.getBoundingClientRect().width));
    const anyPositionedComposition = cols.some((col) =>
      [...col.querySelectorAll('*')].some((descendant) => {
        if (!isVisible(descendant)) return false;
        const position = cs(descendant).position;
        return position === 'absolute' || position === 'fixed';
      }),
    );
    const tallestPx = Math.max(...cols.map((c) => c.getBoundingClientRect().height));
    if (!anyPositionedComposition && tallestPx > 1123 * 1.15) {
      const sliced = slicedPageColumnRows(el, depth, cols, colWidths);
      if (sliced) return sliced;
      const flowCells = cols.map((col) =>
        compactColumnNodes(
          needsBlockWalk(col) ? processChildren(col, depth + 1) : paraNodesFrom(col),
        ),
      );
      const hasTableish = (nodes) =>
        (nodes || []).some(
          (n) =>
            ['table', 'card', 'kpirow'].includes(n.type) ||
            hasTableish(n.children) ||
            (n.rows || []).some((row) => (row.cells || []).some((c) => hasTableish(c.children))),
        );
      // Nested tables make a table row unsplittable in Word: a single
      // multi-page row would strand page 1 blank — sequential flow reads
      // better. Pure paragraph/list columns split across pages natively,
      // so they keep the two-column geometry as one row.
      if (flowCells.some(hasTableish)) return flowCells.flat();
      return [{
        type: 'table',
        layout: true,
        pageColumns: true,
        rows: [{
          cells: cols.map((col, colIndex) => {
            const csc = cs(col);
            return {
              shading: bgHex(col),
              borderLeft: columnEdgeBorder(csc, 'Left'),
              borderRight: columnEdgeBorder(csc, 'Right'),
              gapAfterPx:
                colIndex < cols.length - 1
                  ? Math.max(
                      0,
                      Math.round(
                        cols[colIndex + 1].getBoundingClientRect().left -
                          col.getBoundingClientRect().right,
                      ),
                    )
                  : 0,
              padPx: {
                top: Math.round(parseFloat(csc.paddingTop)),
                left: Math.round(parseFloat(csc.paddingLeft)),
                right: Math.round(parseFloat(csc.paddingRight)),
              },
              children: flowCells[colIndex],
            };
          }),
        }],
        colWidths,
      }];
    }
    const cells = cols.map((col, colIndex) => {
      const csc = cs(col);
      const hasPositionedComposition = [...col.querySelectorAll('*')].some((descendant) => {
        if (!isVisible(descendant)) return false;
        const position = cs(descendant).position;
        return position === 'absolute' || position === 'fixed';
      });
      const colRect = col.getBoundingClientRect();
      const children = compactColumnNodes(
        hasPositionedComposition && colRect.height <= 1800
          ? [markForScreenshot(col)]
          : needsBlockWalk(col)
            ? processChildren(col, depth + 1)
            : paraNodesFrom(col),
      );
      let topEdgeFill;
      if (children[0]?.type === 'image' && (children[0].spacingBeforePx || 0) <= 2) {
        const firstPaintedChild = [...col.children].find(
          (child) => !SKIP_TAGS.has(child.tagName) && isVisible(child),
        );
        if (firstPaintedChild) {
          const firstRect = firstPaintedChild.getBoundingClientRect();
          const offLeft = firstRect.left - colRect.left;
          const offRight = colRect.right - firstRect.right;
          children[0].bleedLeftPx =
            offLeft <= 2 ? Math.max(0, parseFloat(csc.paddingLeft) || 0) : 0;
          children[0].bleedRightPx =
            offRight <= 2 ? Math.max(0, parseFloat(csc.paddingRight) || 0) : 0;
          children[0].align =
            offLeft > 8 &&
            offRight > 8 &&
            Math.abs(offLeft - offRight) < Math.max(12, firstRect.width * 0.08)
              ? 'center'
              : offLeft > offRight * 3
                ? 'right'
                : 'left';
        }
        const firstStyle = firstPaintedChild ? cs(firstPaintedChild) : null;
        const gradientColor = firstStyle?.backgroundImage.match(
          /rgba?\([^)]+\)|#[0-9a-f]{3,8}/i,
        )?.[0];
        topEdgeFill =
          (gradientColor && toHex(gradientColor, { dropWhite: false })) ||
          (firstPaintedChild && bgHex(firstPaintedChild));
        children[0].topEdgeFill = topEdgeFill;
      }
      return {
        // column-level paint (sidebar background) becomes cell shading
        shading: bgHex(col),
        borderLeft: columnEdgeBorder(csc, 'Left'),
        borderRight: columnEdgeBorder(csc, 'Right'),
        gapAfterPx:
          colIndex < cols.length - 1
            ? Math.max(
                0,
                Math.round(
                  cols[colIndex + 1].getBoundingClientRect().left -
                    col.getBoundingClientRect().right,
                ),
              )
            : 0,
        padPx: {
          top: Math.round(parseFloat(csc.paddingTop)),
          left: Math.round(parseFloat(csc.paddingLeft)),
          right: Math.round(parseFloat(csc.paddingRight)),
        },
        children,
      };
    });
    return [{ type: 'table', layout: true, pageColumns: true, rows: [{ cells }], colWidths }];
  }

  // Grid/flex of small uniformly-painted badges (skill icons, tag clouds with
  // shaped backgrounds): text alone can't reproduce circles/pills laid out in
  // a grid, so capture the whole container as one image.
  function isBadgeGrid(el) {
    const s = cs(el);
    const kids = [...el.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    if (kids.length < 2) return false;
    const r = el.getBoundingClientRect();
    if (r.height < 8 || r.height > 320) return false;
    if ((el.innerText || '').trim().length > 150) return false;
    const flexOrGrid = s.display === 'grid' || s.display === 'flex';
    const pillKids = kids.filter((k) => isPill(k));
    // Plain block tag clouds: optional label + >=3 pill chips
    // (e.g. <div>Book tags</div><span class="tag-pill">…).
    const tagCloud =
      !flexOrGrid &&
      pillKids.length >= 3 &&
      pillKids.length >= kids.length - 1 &&
      r.height <= 160;
    if (!flexOrGrid && !tagCloud) return false;
    const checkKids = tagCloud ? pillKids : kids;
    // badges hold a short label at most; multi-line award/skill cards with
    // real sentences must stay editable text (CJK chars weighted double)
    const weighted = (t) =>
      t.length + (t.match(/[\u3000-\u9fff\uf900-\ufaff]/g) || []).length;
    return checkKids.every((k) => {
      const kr = k.getBoundingClientRect();
      return (
        kr.height <= 130 &&
        Boolean(bgHex(k) || paintedBgHex(k)) &&
        weighted((k.innerText || '').trim()) <= 30 &&
        // badges/chips are single-line; a label+value pair is a KPI stat
        // box that must stay an editable kpirow cell
        !(k.innerText || '').trim().includes('\n')
      );
    });
  }

  // Empty bordered square drawn as a form checkbox (styled div or <input>).
  function isCheckboxLike(el) {
    if (el.tagName === 'INPUT') {
      return el.type === 'checkbox' || el.type === 'radio';
    }
    if ((el.innerText || '').trim() || el.children.length) return false;
    const r = el.getBoundingClientRect();
    if (r.width < 8 || r.width > 30 || r.height < 8 || r.height > 30) return false;
    if (Math.abs(r.width - r.height) > 6) return false;
    const s = cs(el);
    return parseFloat(s.borderTopWidth) >= 1 && s.borderTopStyle !== 'none';
  }

  function checkboxNode(el) {
    const r = el.getBoundingClientRect();
    const checked = el.tagName === 'INPUT' && el.checked;
    const radio =
      (el.tagName === 'INPUT' && el.type === 'radio') ||
      // styled circle: border-radius >= 40% of the box
      parseFloat(cs(el).borderTopLeftRadius) >= r.height * 0.4;
    return checkboxRun({
      checked,
      radio,
      contentControl: el.tagName === 'INPUT' && !radio ? 'checkbox' : undefined,
      heightPx: r.height,
      color: toHex(cs(el).borderTopColor) || '333333',
    });
  }

  function checkboxRun(marker) {
    const glyph = marker.radio
      ? marker.checked
        ? '\u25c9'
        : '\u25cb'
      : marker.checked
        ? '\u2611'
        : '\u2610';
    const sizePx = Math.max(12, Math.min(22, Math.round(marker.heightPx * 1.1)));
    return {
      text: glyph,
      color: marker.color || '333333',
      sizePx,
      contentControl: marker.contentControl,
      checked: marker.checked,
    };
  }

  function pseudoCheckbox(el) {
    for (const which of ['::before', '::after']) {
      const s = getComputedStyle(el, which);
      if (!s || s.content === 'none' || s.content === 'normal') continue;
      const content = s.content.replace(/^['"]|['"]$/g, '');
      if (content.trim()) continue;
      const width = parseFloat(s.width);
      const height = parseFloat(s.height);
      if (
        width < 8 ||
        width > 30 ||
        height < 8 ||
        height > 30 ||
        Math.abs(width - height) > 6
      ) {
        continue;
      }
      const borderWidth = parseFloat(s.borderTopWidth);
      if (borderWidth < 1 || s.borderTopStyle === 'none') continue;
      const itemStyle = cs(el);
      const left = Number.isFinite(parseFloat(s.left)) ? parseFloat(s.left) : 0;
      const gap = Math.max(4, (parseFloat(itemStyle.paddingLeft) || 0) - left - width);
      return {
        widthPx: width,
        heightPx: height,
        gapPx: gap,
        checked: false,
        radio: parseFloat(s.borderTopLeftRadius) >= height * 0.4,
        color: toHex(s.borderTopColor) || '999999',
      };
    }
    return null;
  }

  // Small rounded-corner painted box with a short text (badge / pill / logo
  // circle): rectangles can't reproduce the rounding, screenshot it.
  function isPill(el) {
    const r = el.getBoundingClientRect();
    if (r.height < 14 || r.height > 90 || r.width > 420) return false;
    if ((el.innerText || '').trim().length > 60) return false;
    if (parseFloat(cs(el).borderTopLeftRadius) < Math.min(12, r.height * 0.25)) {
      return false;
    }
    // Solid tint, or translucent/near-white paint kept only because a tinted
    // parent sits behind it (e.g. rgba(255,255,255,.2) number chips on colored
    // headers). bgHex alone drops those and they wrongly become white bars.
    return Boolean(bgHex(el) || paintedBgHex(el));
  }

  // Step/number circles only — not wider tag pills that must stay inline.
  function isCompactCircleBadge(el) {
    if (!isPill(el)) return false;
    const r = el.getBoundingClientRect();
    const radius = parseFloat(cs(el).borderTopLeftRadius) || 0;
    return (
      r.width <= 48 &&
      Math.abs(r.width - r.height) <= 8 &&
      radius >= r.height * 0.4
    );
  }

  // Textless painted leaf (progress bar, color swatch, decorative strip,
  // fill-in box): nothing to map to OOXML text, screenshot it. Paint may
  // live on the element itself or on a nested fill div. Height cap is
  // generous only because the element is guaranteed textless.
  function isDecorativeLeaf(el) {
    // a lone emoji/pictograph still counts as textless decoration
    if ((el.innerText || '').trim().length > 4) return false;
    const r = el.getBoundingClientRect();
    if (r.height < 3 || r.height > 400 || r.width < 16) return false;
    if (bgHex(el)) return true;
    if (cs(el).backgroundImage !== 'none') return true;
    // bordered empty box (fill-in area) counts even without background.
    // 1px form answer boxes fail borderTopOf (>=2px), so also accept fullBorder.
    if (
      r.height >= 28 &&
      !isCheckboxLike(el) &&
      (borderTopOf(el) || fullBorderOf(el))
    ) {
      return true;
    }
    for (const d of el.querySelectorAll('*')) {
      if (isVisible(d) && bgHex(d)) return true;
    }
    return false;
  }

  // Icon-font glyph element (<i class="bi bi-telephone">): the glyph lives in
  // a ::before pseudo element, so text extraction sees nothing at all.
  function isIconGlyph(el) {
    if ((el.innerText || '').trim()) return false;
    if (el.children.length) return false;
    const r = el.getBoundingClientRect();
    if (r.width < 6 || r.height < 6 || r.width > 48 || r.height > 48) return false;
    const hasContent = (pseudo) => {
      const content = getComputedStyle(el, pseudo).content;
      return content && content !== 'none' && content !== 'normal' && content !== '""';
    };
    return hasContent('::before') || hasContent('::after');
  }

  // Absolutely-positioned full-height painted strip hugging the card's left
  // edge (decorative accent). Reproduced as a shaded table cell, not a float.
  function leftAccentOf(el) {
    const pr = el.getBoundingClientRect();
    for (const c of el.children) {
      if (cs(c).position !== 'absolute') continue;
      const bg = bgHex(c);
      if (!bg) continue;
      const r = c.getBoundingClientRect();
      if (
        Math.abs(r.left - pr.left) < 3 &&
        r.height >= pr.height * 0.85 &&
        r.width >= 8 &&
        r.width < pr.width * 0.4
      ) {
        return { el: c, color: bg, widthPx: Math.round(r.width) };
      }
    }
    return null;
  }

  // Uniform border on all four sides (outlined card / framed box)
  function fullBorderOf(el) {
    const s = cs(el);
    const w = parseFloat(s.borderTopWidth);
    if (!(w >= 1) || s.borderTopStyle === 'none') return null;
    if (
      parseFloat(s.borderLeftWidth) < 1 ||
      parseFloat(s.borderRightWidth) < 1 ||
      parseFloat(s.borderBottomWidth) < 1
    ) {
      return null;
    }
    // A uniform frame paints every side the same color. A bright accent on one
    // edge over near-background hairlines (gold top rule + dark 1px sides) is a
    // top rule, not a box — let borderTopOf claim just that edge instead of
    // drawing the accent color around all four sides.
    const color = toHex(s.borderTopColor);
    const sidesMatch =
      toHex(s.borderRightColor) === color &&
      toHex(s.borderBottomColor) === color &&
      toHex(s.borderLeftColor) === color;
    if (!sidesMatch) return null;
    return color ? { color, widthPx: w } : null;
  }

  // Visual card: own background, strong left border, or a full outline —
  // smaller than page frame
  function isCard(el, depth) {
    if (depth >= 4) return false;
    const r = el.getBoundingClientRect();
    const shading = paintedBgHex(el);
    const bl = borderLeftOf(el);
    const br = borderRightOf(el);
    const fullBorder = fullBorderOf(el);
    // Very tall unframed wrappers are page containers, but an explicit
    // four-sided border is authored content and must survive across pages.
    // Exception: a full-bleed SECTION with its own solid backdrop among
    // sibling sections (dark manifesto chapters) — without the card shading
    // its light text lands unreadable on the page background. The card row
    // splits across pages, so height is no obstacle.
    if (r.height > 1000 && !fullBorder) {
      const siblingCountAt = (node) =>
        node.parentElement
          ? [...node.parentElement.children].filter((c) => c !== node && isVisible(c)).length
          : 0;
      // shading (paintedBgHex above) includes a white sheet on a tinted
      // page — dropping it dumps the sheet's content onto the backdrop.
      // The painted panel may sit inside a transparent section wrapper
      // (<section><div style="background:…">): the "among sibling sections"
      // signal then lives on the wrapper, so look up through unpainted ones.
      let host = el;
      let siblings = siblingCountAt(host);
      while (
        siblings < 2 &&
        host.parentElement &&
        host.parentElement !== document.body &&
        !paintedBgHex(host.parentElement)
      ) {
        host = host.parentElement;
        siblings = siblingCountAt(host);
      }
      if (!(shading && siblings >= 2)) return false;
    }
    const s = cs(el);
    const hasPadding = parseFloat(s.paddingLeft) >= 6 || parseFloat(s.paddingTop) >= 6;
    return Boolean((shading || bl || br || fullBorder) && hasPadding);
  }

  function formFieldNode(el) {
    return globalThis.__html2docxFormsMedia.formFieldNode(el, {
      bgHex,
      bordersOf,
      cellPaddingOf,
      cs,
      runStyleOf,
    });
  }

  function fillInLineNode(el) {
    return globalThis.__html2docxFormsMedia.fillInLineNode(el, { bordersOf, cs });
  }

  function fillInBoxNode(el) {
    return globalThis.__html2docxFormsMedia.fillInBoxNode(el, {
      bordersOf,
      cs,
      toHex,
      paintedBgHex,
      runStyleOf,
    });
  }

  function embeddedMediaNode(el) {
    return globalThis.__html2docxFormsMedia.embeddedMediaNode(el, { markForScreenshot });
  }

