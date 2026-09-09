(function installHtmlTableExtractor(global) {
  function extractTable(el, deps) {
    const {
      bgHex,
      isCompactPaintedLabel,
      bordersOf,
      cellPaddingOf,
      collectRuns,
      cs,
      hasBlockChildren,
      isRTL,
      isVisible,
      markForScreenshot,
      processChildren,
      textAlignOf,
      toHex,
    } = deps;
    // Gradient-painted cells (th { background: linear-gradient(...) }) have a
    // transparent backgroundColor; approximate with the first gradient stop so
    // white header text doesn't land on a white cell.
    const gradHex = (node) => {
      const bgImage = cs(node).backgroundImage;
      if (!bgImage || !bgImage.includes('gradient')) return null;
      const m = bgImage.match(/rgba?\([^)]*\)|#[0-9a-fA-F]{3,8}/);
      return m ? toHex(m[0]) : null;
    };
    const cellShading = (cell, tr) => {
      const own = bgHex(cell) || gradHex(cell);
      if (own) return own;
      const behind = bgHex(tr) || gradHex(tr);
      if (!behind) return null;
      // A cell's own opaque paint (even white, which bgHex drops) covers the
      // row color — without this the row bleeds through and the cell's dark
      // text lands on dark row shading.
      return toHex(cs(cell).backgroundColor) || behind;
    };
    const rows = [];
    const boundarySamples = [];
    const activeSpans = [];
    let gridColumnCount = 0;
    const tableRect = el.getBoundingClientRect();
    const tableRows = [...el.querySelectorAll(
      ':scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr, :scope > tr',
    )];
    for (let rowIndex = 0; rowIndex < tableRows.length; rowIndex++) {
      const tr = tableRows[rowIndex];
      if (!isVisible(tr)) continue;
      const cells = [];
      const rowBorders = bordersOf(tr);
      const occupied = new Set();
      for (const span of activeSpans) {
        if (span.endRow <= rowIndex) continue;
        for (let col = span.gridStart; col < span.gridStart + span.colspan; col++) {
          occupied.add(col);
        }
        cells.push({
          ...span.cell,
          runs: [],
          children: undefined,
          gridStart: span.gridStart,
          verticalMerge: 'continue',
        });
      }
      const directCells = [...tr.children].filter(
        (cell) => cell.tagName === 'TD' || cell.tagName === 'TH',
      );
      const exactVisualRow = directCells.some(
        (cell) => [...cell.children].filter(isCompactPaintedLabel).length >= 2,
      );
      if (exactVisualRow && !activeSpans.some((span) => span.endRow > rowIndex)) {
        const colspan = directCells.reduce((sum, cell) => sum + (cell.colSpan || 1), 0);
        cells.push({
          colspan,
          rowspan: 1,
          gridStart: 0,
          children: [markForScreenshot(tr)],
          borders: {},
          padPx: { top: 0, right: 0, bottom: 0, left: 0 },
        });
        const rect = tr.getBoundingClientRect();
        (boundarySamples[0] ||= []).push(rect.left - tableRect.left);
        (boundarySamples[colspan] ||= []).push(rect.right - tableRect.left);
        gridColumnCount = Math.max(gridColumnCount, colspan);
        rows.push({ cells, heightPx: Math.round(rect.height) });
        continue;
      }
      let nextColumn = 0;
      for (const cell of tr.children) {
        if (cell.tagName !== 'TD' && cell.tagName !== 'TH') continue;
        const colspan = cell.colSpan || 1;
        while (
          [...Array(colspan).keys()].some((offset) => occupied.has(nextColumn + offset))
        ) {
          nextColumn++;
        }
        const rowspan = cell.rowSpan || 1;
        const entry = {
          colspan,
          rowspan,
          gridStart: nextColumn,
          verticalMerge: rowspan > 1 ? 'restart' : undefined,
          shading: cellShading(cell, tr),
          bold: cell.tagName === 'TH' || undefined,
          borders: bordersOf(cell) || rowBorders,
          padPx: cellPaddingOf(cell),
          rtl: isRTL(cell) || undefined,
        };
        const textAlign = textAlignOf(cell);
        if (textAlign === 'center' || textAlign === 'right') entry.align = textAlign;
        const verticalAlign = cs(cell).verticalAlign;
        if (verticalAlign === 'middle') entry.vAlign = 'center';
        else if (verticalAlign === 'bottom') entry.vAlign = 'bottom';
        // decorations: false - the cell entry already owns border/shading;
        // paragraphs merged from the cell's inline children must not inherit
        // the td's border as a paragraph border (drew stray lines inside cells)
        if (hasBlockChildren(cell)) {
          entry.children = processChildren(cell, 2, { decorations: false });
        }
        else {
          entry.runs = collectRuns(cell);
          // CJK cell text: with no explicit line height Word applies the CJK
          // font's own leading (Yu Mincho ≈1.7x) — uncapped, invoice rows
          // balloon and push a one-page layout onto a second page. Pin the
          // browser's line box; skip tight leadings that EXACT would clip.
          const lineHeightPx = parseFloat(cs(cell).lineHeight);
          const fontSizePx = parseFloat(cs(cell).fontSize);
          const hasCJK = (entry.runs || []).some((run) =>
            /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/.test(run.text || ''),
          );
          if (hasCJK && lineHeightPx && fontSizePx && lineHeightPx / fontSizePx >= 1.3) {
            entry.style = { ...entry.style, cjk: true, exactLineHeightPx: lineHeightPx };
          }
        }
        // Empty cell showing placeholder text via ::before content
        // ("Click to edit..." on contenteditable template cells): keep the
        // rendered placeholder so the grid doesn't collapse into blank space.
        if (!entry.children && !(entry.runs || []).length) {
          const pseudo = getComputedStyle(cell, '::before');
          const m = (pseudo.content || '').match(/^"(.*)"$/);
          if (m && m[1].trim()) {
            const run = { text: m[1], color: toHex(pseudo.color) || '999999' };
            if (pseudo.fontStyle === 'italic') run.italic = true;
            const sz = parseFloat(pseudo.fontSize);
            if (sz) run.sizePx = sz;
            entry.runs = [run];
          }
        }
        // Short cell content ("Wave 1" pills, "11." numbers): record the
        // forced single-line width (nowrap min-content, works even when the
        // browser already crushed the column) so the generator can guarantee
        // the column never wraps it letter-by-letter into a fake vertical
        // stack.
        {
          const cellText = (cell.innerText || '').trim();
          if (cellText && cellText.length <= 40 && !cellText.includes('\n')) {
            // nowrap + Range: the true single-line CONTENT width. Plain
            // scrollWidth reports the stretched layout width for uncrushed
            // cells, which would inflate every column's floor equally.
            // inner width BEFORE forcing nowrap — nowrap reflows the auto
            // table layout and stretches the cell to fit the text
            const pad = cellPaddingOf(cell);
            const innerWidth =
              cell.getBoundingClientRect().width - (pad.left || 0) - (pad.right || 0);
            const previousWhiteSpace = cell.style.whiteSpace;
            cell.style.whiteSpace = 'nowrap';
            try {
              const range = document.createRange();
              range.selectNodeContents(cell);
              let textWidth = range.getBoundingClientRect().width;
              // Only short label-like content that the browser itself renders
              // on one line gets a keep-on-one-line floor. A cell the browser
              // already wrapped must be allowed to wrap in Word too — floors
              // from wrapped cells starve the remaining columns (a 5-column
              // table's last column collapsed to 600 twips).
              if (textWidth > 0 && textWidth <= 160 && textWidth <= innerWidth + 2) {
                // Word's substituted font may run much wider than the webfont.
                const wordWidth = window.__h2dWordFontLineWidth
                  ? window.__h2dWordFontLineWidth(cell)
                  : 0;
                entry.oneLineWidthPx = Math.round(Math.max(textWidth, wordWidth));
              }
            } catch (e) {
              /* leave unset */
            }
            cell.style.whiteSpace = previousWhiteSpace || '';
          }
        }
        cells.push(entry);
        const rect = cell.getBoundingClientRect();
        const startPx = isRTL(el) ? tableRect.right - rect.right : rect.left - tableRect.left;
        const endPx = isRTL(el) ? tableRect.right - rect.left : rect.right - tableRect.left;
        (boundarySamples[nextColumn] ||= []).push(startPx);
        (boundarySamples[nextColumn + colspan] ||= []).push(endPx);
        if (rowspan > 1) {
          activeSpans.push({
            gridStart: nextColumn,
            colspan,
            endRow: rowIndex + rowspan,
            cell: {
              colspan,
              shading: entry.shading,
              borders: entry.borders,
              padPx: entry.padPx,
              rtl: entry.rtl,
              align: entry.align,
              vAlign: entry.vAlign,
            },
          });
          for (let col = nextColumn; col < nextColumn + colspan; col++) occupied.add(col);
        }
        nextColumn += colspan;
      }
      if (!cells.length) continue;
      cells.sort((a, b) => a.gridStart - b.gridStart);
      gridColumnCount = Math.max(
        gridColumnCount,
        ...cells.map((cell) => cell.gridStart + (cell.colspan || 1)),
      );
      const renderedHeight = Math.round(tr.getBoundingClientRect().height);
      const rowText = (tr.innerText || '').trim();
      // Only a real header row repeats across pages: flagging plain first
      // rows duplicates a data row at every page break.
      const isHeaderRow =
        tr.parentElement?.tagName === 'THEAD' ||
        (directCells.length > 0 && directCells.every((cell) => cell.tagName === 'TH'));
      rows.push({
        cells,
        header: isHeaderRow || undefined,
        rawHeightPx: renderedHeight,
        // Compact text rows should use Word's natural height. Keeping the
        // browser's measured minimum double-counts small font-metric/padding
        // differences and makes ordinary data tables visibly too loose.
        heightPx:
          renderedHeight >= 80 || (renderedHeight >= 30 && !rowText)
            ? renderedHeight
            : undefined,
      });
    }
    if (!rows.length) return [];
    // Uniform form grid (fill-in sheets sized for handwriting): the author
    // gave every body row the same generous height. Preserving it only on
    // the EMPTY rows (rule above) makes filled rows collapse to text height
    // and the grid uneven — carry the shared height to all of them.
    {
      const body = rows.slice(1);
      const heights = body.map((row) => row.rawHeightPx);
      if (
        body.length >= 4 &&
        Math.min(...heights) >= 32 &&
        Math.max(...heights) - Math.min(...heights) <= 4
      ) {
        for (const row of body) row.heightPx = row.rawHeightPx;
      }
      for (const row of rows) delete row.rawHeightPx;
    }
    const boundaries = Array(gridColumnCount + 1).fill(null);
    boundaries[0] = 0;
    boundaries[gridColumnCount] = tableRect.width;
    for (let index = 1; index < gridColumnCount; index++) {
      const samples = boundarySamples[index];
      if (samples?.length) {
        const sorted = [...samples].sort((a, b) => a - b);
        boundaries[index] = sorted[Math.floor(sorted.length / 2)];
      }
    }
    let previousKnown = 0;
    for (let index = 1; index <= gridColumnCount; index++) {
      if (boundaries[index] == null) continue;
      const span = index - previousKnown;
      const start = boundaries[previousKnown];
      const end = boundaries[index];
      for (let fill = 1; fill < span; fill++) {
        boundaries[previousKnown + fill] = start + ((end - start) * fill) / span;
      }
      previousKnown = index;
    }
    const colWidths = boundaries
      .slice(0, -1)
      .map((start, index) => Math.max(1, Math.round(boundaries[index + 1] - start)));
    return [{
      type: 'table',
      rows,
      colWidths,
      outerBorder: bordersOf(el),
      shading: bgHex(el),
      preserveRowHeights: true,
      // rendered height drives the keep-together call: a table that fits on
      // one page should jump to the next page whole instead of splitting
      heightPx: Math.round(tableRect.height),
      rtl: isRTL(el) || undefined,
    }];
  }

  global.__html2docxTable = { extractTable };
})(globalThis);
