  function extractTable(el) {
    const rect = el.getBoundingClientRect();
    const hasCompactPaintedLabel = [...el.querySelectorAll('td *, th *')].some(isCompactPaintedLabel);
    if (hasCompactPaintedLabel && rect.height <= 900) return markForScreenshot(el);
    return globalThis.__html2docxTable.extractTable(el, {
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
    });
  }

  // A CSS-grid table whose direct children are grid rows with the same number
  // of columns (for example a Likert survey built from divs instead of table).
  function gridTableRows(el) {
    const rows = [...el.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    if (rows.length < 2) return null;
    const cssTable =
      cs(el).display === 'table' &&
      rows.every(
        (row) =>
          cs(row).display === 'table-row' &&
          [...row.children]
            .filter((cell) => !SKIP_TAGS.has(cell.tagName) && isVisible(cell))
            .every((cell) => cs(cell).display === 'table-cell'),
      );
    const counts = rows.map((row) => {
      const s = cs(row);
      if (s.display !== 'grid' && !(cssTable && s.display === 'table-row')) return 0;
      return [...row.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c)).length;
    });
    const base = Math.max(...counts);
    if (base < 2 || base > 12) return null;
    // CSS tables tolerate short rows (a lone header cell): the browser leaves
    // the missing columns empty, so extractGridTable pads them with fillers.
    // Grid rows keep the strict equal-count rule.
    if (counts.some((count) => count !== base && !(cssTable && count >= 1))) return null;
    if (counts.filter((count) => count === base).length < 2) return null;
    const widths = rows.map((row) => Math.round(row.getBoundingClientRect().width));
    if (Math.max(...widths) - Math.min(...widths) > 8) return null;
    // Cells must sit in the same columns in every row: alternating layouts
    // (timeline cards flipping left/right around a center rail) are grids of
    // rows too, but mapping their cells positionally scrambles the columns.
    const leftsOf = (row) =>
      [...row.children]
        .filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c))
        .map((c) => Math.round(c.getBoundingClientRect().left))
        .sort((a, b) => a - b);
    const firstLefts = leftsOf(rows[counts.indexOf(base)]);
    for (const row of rows) {
      const lefts = leftsOf(row);
      if (lefts.length === base) {
        if (lefts.some((left, i) => Math.abs(left - firstLefts[i]) > 12)) return null;
      } else if (
        lefts.some((left) => !firstLefts.some((columnLeft) => Math.abs(left - columnLeft) <= 12))
      ) {
        return null;
      }
    }
    // Avoid classifying arbitrary stacks of small grids as a data table.
    if (!cssTable && rows.length < 4 && !fullBorderOf(el)) return null;
    return rows;
  }

  function extractGridTable(el, gridRows, depth) {
    const cellsOf = (row) =>
      [...row.children].filter((c) => !SKIP_TAGS.has(c.tagName) && isVisible(c));
    const baseRow = gridRows.reduce(
      (best, row) => (cellsOf(row).length > cellsOf(best).length ? row : best),
      gridRows[0],
    );
    const firstCells = cellsOf(baseRow);
    const colWidths = firstCells.map((c) => Math.round(c.getBoundingClientRect().width));
    const colLefts = firstCells.map((c) => Math.round(c.getBoundingClientRect().left));
    const rows = gridRows.map((row, rowIdx) => {
      const domCells = cellsOf(row);
      let cells = domCells.map((cell) => {
          const textAlign = textAlignOf(cell);
          const entry = {
            shading: bgHex(cell) || bgHex(row),
            borders: bordersOf(cell),
            padPx: cellPaddingOf(cell),
            vAlign: cs(row).alignItems === 'center' ? 'center' : undefined,
            rtl: isRTL(cell) || undefined,
          };
          if (textAlign === 'center' || textAlign === 'right') entry.align = textAlign;
          entry.children = needsBlockWalk(cell)
            ? processChildren(cell, depth + 1, { decorations: false })
            : paraNodesFrom(cell, false);
          return entry;
        });
      // Short row (lone header cell in a CSS table): seat each cell in its
      // geometric column and fill the rest with empty cells — unshaded, so
      // they inherit the table background, same as the browser leaves them.
      if (cells.length && cells.length < colWidths.length) {
        const placed = new Array(colWidths.length).fill(null);
        domCells.forEach((cell, i) => {
          const left = Math.round(cell.getBoundingClientRect().left);
          let col = colLefts.findIndex((columnLeft) => Math.abs(columnLeft - left) <= 12);
          if (col < 0 || placed[col]) col = placed.findIndex((slot) => !slot);
          placed[col] = cells[i];
        });
        cells = placed.map((entry) => entry || { children: [] });
      }
      const renderedHeight = Math.round(row.getBoundingClientRect().height);
      return {
        cells,
        // Let text-driven rows use Word's natural compact height. Preserve
        // only genuinely tall visual rows (planners / empty writing areas).
        heightPx: renderedHeight >= 80 ? renderedHeight : undefined,
        // Only a short strip reads as a repeatable header; a tall first row
        // is content (card stacks) and would reprint on every page.
        header: rowIdx === 0 && renderedHeight < 60,
      };
    });
    return [{
      type: 'table',
      gridTable: true,
      rows,
      colWidths,
      outerBorder: bordersOf(el),
      shading: bgHex(el),
      padPx: cellPaddingOf(el),
      heightPx: Math.round(el.getBoundingClientRect().height),
      rtl: isRTL(el) || undefined,
    }];
  }

  function extractCode(el) {
    const shading = bgHex(el) || 'F5F5F5';
    const lines = el.innerText.replace(/\n$/, '').split('\n');
    const { color, sizePx } = runStyleOf(el);
    return [{ type: 'code', lines, shading, color, sizePx }];
  }

  function extractList(el, ordered, depth = 1) {
    const listItems = [...el.children].filter((li) => li.tagName === 'LI' && isVisible(li));
    // list-style:none -> the author wants plain paragraphs, not numbering.
    // Judge by the items, not the UL: frameworks (pico.css) re-style the LI
    // marker even when the UL itself says none.
    const plain = listItems.length
      ? listItems.every(
          (li) => cs(li).listStyleType === 'none' || cs(li).display !== 'list-item',
        )
      : cs(el).listStyleType === 'none';
    const needsExactPseudoMarker = (li) =>
      /[☑☐❋]/u.test(pseudoBulletRun(li)?.text || '');
    const hasCompactPaintedLabels = listItems.some((li) =>
      [...li.querySelectorAll(':scope > *')].some(isCompactPaintedLabel),
    );
    if (hasCompactPaintedLabels) return [markForScreenshot(el)];
    // CSS multi-column list (columns: 2 on a TOC): reproduce as a layout
    // table with one cell per rendered column. Numbering continues across
    // columns, so synthesize literal number prefixes instead of Word lists.
    const columnCount = parseInt(cs(el).columnCount, 10);
    if (
      depth <= 2 &&
      Number.isFinite(columnCount) &&
      columnCount >= 2 &&
      listItems.length >= columnCount &&
      listItems.every((li) => !hasBlockChildren(li))
    ) {
      const lefts = listItems.map((li) => Math.round(li.getBoundingClientRect().left));
      const uniqueLefts = [...new Set(lefts)].sort((a, b) => a - b);
      if (uniqueLefts.length >= 2) {
        const groups = uniqueLefts.map(() => []);
        listItems.forEach((li, index) => {
          groups[uniqueLefts.indexOf(lefts[index])].push({ li, index });
        });
        const cells = groups.map((group) => ({
          children: group.flatMap(({ li, index }) => {
            if (plain && needsExactPseudoMarker(li)) return [markForScreenshot(li)];
            const runs = collectRuns(li);
            if (!runs.length) return [];
            const pseudoRun = pseudoBulletRun(li);
            let prefixed = Boolean(pseudoRun);
            if (pseudoRun) runs.unshift(pseudoRun);
            else if (ordered) {
              runs.unshift({ text: `${index + 1}. `, ...runStyleOf(li) });
              prefixed = true;
            } else if (!plain) {
              runs.unshift({ text: '\u2022\u00a0', ...runStyleOf(li) });
              prefixed = true;
            }
            const style = { ...paraStyleOf(el, false), ...paraStyleOf(li, false) };
            delete style.indentLeftPx;
            if (prefixed) {
              const em = parseFloat(cs(li).fontSize) || 16;
              style.indentHangingPx = Math.round(em * (ordered ? 1.1 : 0.8));
            }
            return [{ type: 'para', runs, style }];
          }),
        }));
        const width = el.getBoundingClientRect().width || 1;
        return [{
          type: 'table',
          layout: true,
          rows: [{ cells: cells.map((cell, i) => ({ ...cell, gridStart: i, colspan: 1 })) }],
          colWidths: uniqueLefts.map((left, i) =>
            Math.round((i + 1 < uniqueLefts.length ? uniqueLefts[i + 1] : el.getBoundingClientRect().right) - left) || Math.round(width / uniqueLefts.length),
          ),
        }];
      }
    }
    // Layout list: display:flex items (numbered chip + content column) need
    // the full classifier. Flattening them to runs drops the chip and
    // inlines nested callout boxes into the text. Items too tall for a table
    // row (isBoxRow rejects them) still get a block walk so nested dialogue
    // bubbles / sub-lists survive.
    // CSS-counter numbering ("counter(toc-counter)") never resolves through
    // getComputedStyle; synthesize it from the item's position.
    const counterPrefix = (li, index) => {
      const pseudoStyle = getComputedStyle(li, '::before');
      const content = pseudoStyle.content || '';
      const counterStart = content.indexOf('counter(');
      if (counterStart < 0) return null;
      const counterEnd = content.indexOf(')', counterStart);
      const quotedText = (part) =>
        [...part.matchAll(/"([^"]*)"|'([^']*)'/g)]
          .map((match) => match[1] ?? match[2] ?? '')
          .join('');
      const prefix = quotedText(content.slice(0, counterStart));
      const suffix = quotedText(content.slice(counterEnd + 1));
      let value = `${prefix}${index + 1}${suffix}`;
      if (pseudoStyle.textTransform === 'uppercase') value = value.toUpperCase();
      if (pseudoStyle.textTransform === 'lowercase') value = value.toLowerCase();
      return value.endsWith(' ') ? value : `${value} `;
    };
    if (plain && listItems.some((li) => isBoxRow(li) || hasBlockChildren(li))) {
      return listItems.flatMap((li, index) => {
        if (isBoxRow(li)) return processElement(li, depth + 1);
        // A counter badge glyph (❶) already carries the number — adding the
        // synthesized plain counter text would double-number the item.
        const bulletRun = pseudoBulletRun(li) || chipBulletRun(li);
        const prefix = bulletRun ? null : counterPrefix(li, index);
        if (hasBlockChildren(li)) {
          const nodes = processChildren(li, depth + 1);
          const first = nodes.find((n) => n.type === 'para');
          if (first && bulletRun) first.runs = [bulletRun, ...first.runs];
          if (first && prefix) first.runs = [{ text: prefix, ...runStyleOf(li) }, ...first.runs];
          return nodes;
        }
        const runs = collectRuns(li);
        if (!runs.length) return [];
        if (prefix) runs.unshift({ text: prefix, ...runStyleOf(li) });
        if (bulletRun) runs.unshift(bulletRun);
        const style = { ...paraStyleOf(el), ...paraStyleOf(li) };
        return [{ type: 'para', runs, style }];
      });
    }
    const nativeMarkerOf = (li) => {
      const itemStyle = cs(li);
      if (itemStyle.display !== 'list-item' || itemStyle.listStyleType === 'none') return null;
      const glyphs = {
        disc: '\u2022',
        circle: '\u25e6',
        square: '\u25aa',
      };
      const glyph = glyphs[itemStyle.listStyleType];
      if (!glyph) return null;
      const runStyle = runStyleOf(li);
      return {
        run: {
          text: `${glyph}\u00a0`,
          color: runStyle.color,
          sizePx: runStyle.sizePx,
          fontFamily: runStyle.fontFamily,
        },
        widthPx: Math.max(10, Math.round((runStyle.sizePx || 14) * 1.05)),
      };
    };
    const markers = listItems.map((li) => {
      const first = [...li.children].find((child) => isVisible(child));
      const nativeMarker = nativeMarkerOf(li);
      if (first && isCheckboxLike(first)) {
        const rect = first.getBoundingClientRect();
        const itemStyle = cs(li);
        return {
          element: first,
          widthPx: rect.width,
          gapPx: parseFloat(itemStyle.columnGap) || parseFloat(itemStyle.gap) || 8,
          nativeMarker,
        };
      }
      const marker = pseudoCheckbox(li);
      return marker ? { ...marker, nativeMarker } : null;
    });
    const checklist =
      plain &&
      listItems.length > 1 &&
      markers.every(Boolean);
    if (checklist) {
      const listRect = el.getBoundingClientRect();
      const firstColWidth = Math.round(
        Math.max(
          ...markers.map(
            (marker) => marker.widthPx + marker.gapPx + (marker.nativeMarker?.widthPx || 0),
          ),
        ),
      );
      const rows = listItems.map((li, index) => {
        const children = [...li.children].filter((child) => isVisible(child));
        const marker = markers[index];
        const label = marker.element ? children.find((child) => child !== marker.element) : null;
        const rowBorders = bordersOf(li);
        const pad = cellPaddingOf(li);
        const itemStyle = cs(li);
        const itemParaStyle = paraStyleOf(li);
        delete itemParaStyle.spacingBeforePx;
        delete itemParaStyle.spacingAfterPx;
        // The marker table already consumes the list item's left padding.
        // Keeping it on the label paragraph would indent the text twice.
        delete itemParaStyle.indentLeftPx;
        const markerRuns = [];
        if (marker.nativeMarker) markerRuns.push(marker.nativeMarker.run);
        markerRuns.push(marker.element ? checkboxNode(marker.element) : checkboxRun(marker));
        return {
          heightPx: Math.round(
            li.getBoundingClientRect().height + Math.max(0, parseFloat(itemStyle.marginBottom) || 0),
          ),
          cells: [
            {
              runs: markerRuns,
              borders: rowBorders || {},
              padPx: { top: pad.top, bottom: pad.bottom, left: 0, right: marker.gapPx },
            },
            {
              runs: label ? collectRuns(label) : collectRuns(li),
              style: itemParaStyle,
              borders: rowBorders || {},
              padPx: { top: pad.top, bottom: pad.bottom, left: 0, right: 0 },
            },
          ],
          last: index === listItems.length - 1,
        };
      });
      const listStyle = cs(el);
      return [{
        type: 'table',
        checklist: true,
        noHeader: true,
        rows,
        colWidths: [firstColWidth, Math.max(1, Math.round(listRect.width) - firstColWidth)],
        spacingBeforePx: Math.max(0, Math.round(parseFloat(listStyle.marginTop) || 0)),
        spacingAfterPx: Math.max(0, Math.round(parseFloat(listStyle.marginBottom) || 0)),
        rtl: isRTL(el) || undefined,
      }];
    }
    const items = [];
    for (const li of listItems) {
      items.push({ runs: collectRuns(li), style: paraStyleOf(li) });
    }
    if (!items.length) return [];
    if (plain) {
      return items
        .filter((it) => it.runs.length)
        .map((it, index) => {
          const li = listItems[index];
          if (needsExactPseudoMarker(li)) return markForScreenshot(li);
          const style = { ...paraStyleOf(el), ...it.style };
          const itemStyle = cs(li);
          const paddingTop = parseFloat(itemStyle.paddingTop) || 0;
          const paddingBottom = parseFloat(itemStyle.paddingBottom) || 0;
          if (paddingTop > 0) style.spacingBeforePx = (style.spacingBeforePx || 0) + paddingTop;
          if (paddingBottom > 0) style.spacingAfterPx = (style.spacingAfterPx || 0) + paddingBottom;
          const bulletRun = pseudoBulletRun(li) || chipBulletRun(li);
          if (bulletRun) it.runs.unshift(bulletRun);
          else {
            // list-style:none + ::before counter() (self-numbered rule
            // lists): the marker lives in the pseudo-element, not the runs
            const prefix = counterPrefix(li, index);
            if (prefix) {
              const pseudoWeight = parseInt(getComputedStyle(li, '::before').fontWeight, 10);
              it.runs.unshift({
                text: prefix,
                ...runStyleOf(li),
                ...(pseudoWeight >= 600 ? { bold: true } : {}),
              });
              const em = parseFloat(itemStyle.fontSize) || 16;
              style.indentHangingPx = Math.round(em * 1.1);
            }
          }
          return { type: 'para', runs: it.runs, style };
        });
    }
    // <ol start="951">: Word numbering always restarts at 1 (startOverride
    // is fixed), so lists with an authored start keep literal number
    // prefixes instead — the TOC must match the 951-975 headings it links.
    const startAt = ordered ? parseInt(el.getAttribute('start'), 10) : NaN;
    if (Number.isFinite(startAt) && startAt !== 1 && listItems.every((li) => !hasBlockChildren(li))) {
      return listItems.flatMap((li, index) => {
        const runs = collectRuns(li);
        if (!runs.length) return [];
        runs.unshift({ text: `${startAt + index}. `, ...runStyleOf(li) });
        const style = { ...paraStyleOf(el, false), ...paraStyleOf(li, false) };
        const em = parseFloat(cs(li).fontSize) || 16;
        style.indentHangingPx = Math.round(em * 1.6);
        return [{ type: 'para', runs, style }];
      });
    }
    const nestedItems = [];
    // liIndex = ordinal of the top-level li each item came from, so column
    // slicing can split a tall bullet list at li boundaries.
    const appendListItems = (listEl, level, parentIndentPx, rootIndex) => {
      const listStyle = cs(listEl);
      const indentLeftPx = parentIndentPx + Math.max(0, parseFloat(listStyle.paddingLeft) || 0);
      const childItems = [...listEl.children].filter(
        (child) => child.tagName === 'LI' && isVisible(child),
      );
      for (let liOrdinal = 0; liOrdinal < childItems.length; liOrdinal++) {
        const li = childItems[liOrdinal];
        const liIndex = level === 0 ? liOrdinal : rootIndex;
        // Walk li children in DOM order, splitting content around nested
        // lists: text after a nested list must stay after it (a trailing
        // "to perform..." clause hoisted above its sub-items inverts the
        // reading order).
        const segments = [];
        let current = [];
        for (const node of li.childNodes) {
          const isList =
            node.nodeType === Node.ELEMENT_NODE &&
            (node.tagName === 'UL' || node.tagName === 'OL');
          if (isList) {
            if (isVisible(node)) {
              segments.push({ content: current });
              current = [];
              segments.push({ list: node });
            }
          } else current.push(node);
        }
        segments.push({ content: current });
        const style = paraStyleOf(li);
        if (segments.some((seg) => seg.list)) delete style.spacingAfterPx;
        let leading = true;
        for (const seg of segments) {
          if (seg.list) {
            appendListItems(seg.list, level + 1, indentLeftPx, liIndex);
            continue;
          }
          const runs = trimRuns(runsFromNodes(seg.content, li));
          if (!runs.length) continue;
          // A justified li with an explicit <br> keeps the break as \n in one
          // paragraph — Word stretches the line BEFORE a manual break under
          // justify (CSS never does), blowing huge gaps into short label
          // lines. Split at the breaks like makeParaNodes: only the final
          // fragment keeps justify (Word never stretches a last line).
          const fragments = [runs];
          if (style.align === 'justify' && runs.some((run) => (run.text || '').includes('\n'))) {
            fragments.length = 0;
            let currentLine = [];
            for (const run of runs) {
              const parts = (run.text || '').split('\n');
              parts.forEach((part, partIndex) => {
                if (part) currentLine.push({ ...run, text: part });
                if (partIndex < parts.length - 1 && currentLine.length) {
                  fragments.push(currentLine);
                  currentLine = [];
                }
              });
            }
            if (currentLine.length) fragments.push(currentLine);
            if (!fragments.length) fragments.push(runs);
          }
          fragments.forEach((fragmentRuns, fragmentIndex) => {
            const fragmentStyle =
              fragments.length > 1 && fragmentIndex < fragments.length - 1
                ? { ...style, align: undefined }
                : style;
            if (leading) {
              nestedItems.push({
                runs: fragmentRuns,
                style: fragmentStyle,
                level,
                // ::before counter() numbering counts as ordered even on a UL
                ordered:
                  listEl.tagName === 'OL' ||
                  (getComputedStyle(li, '::before').content || '').includes('counter('),
                markerType: cs(li).listStyleType || listStyle.listStyleType,
                indentLeftPx,
                liIndex,
              });
              leading = false;
            } else {
              // continuation text of the same li: no marker, text-aligned indent
              nestedItems.push({
                runs: fragmentRuns,
                style: fragmentStyle,
                level,
                continuation: true,
                indentLeftPx,
                liIndex,
              });
            }
          });
        }
      }
    };
    appendListItems(el, 0, 0, 0);
    return [{
      type: 'list',
      ordered,
      items: nestedItems,
      rtl: isRTL(el) || undefined,
    }];
  }

  // A horizontal row of >=2 similar boxes (flex row / grid columns) -> kpirow
