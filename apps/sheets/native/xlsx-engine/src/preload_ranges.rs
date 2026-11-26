use super::{
    CHUNK_ROW_COUNT, CellRange, ChunkData, MAX_RANGE_CELLS, MergedRange, SheetIndex, SheetMetadata,
};

pub(super) fn chunk_bounds(chunk: &ChunkData) -> Option<CellRange> {
    let mut bounds = chunk.cells.iter().fold(None, |bounds, cell| {
        let area = point(cell.row, cell.column);
        Some(bounds.map_or_else(|| area.clone(), |bounds| union(&bounds, &area)))
    });
    if let (Some(first), Some(last)) = (chunk.rows.first(), chunk.rows.last()) {
        let bounds = bounds.get_or_insert_with(|| point(first.row, 0));
        bounds.start_row = bounds.start_row.min(first.row);
        bounds.end_row = bounds.end_row.max(last.row);
    }
    bounds
}

fn point(row: usize, column: usize) -> CellRange {
    CellRange {
        start_row: row,
        end_row: row,
        start_column: column,
        end_column: column,
    }
}

fn union(a: &CellRange, b: &CellRange) -> CellRange {
    CellRange {
        start_row: a.start_row.min(b.start_row),
        end_row: a.end_row.max(b.end_row),
        start_column: a.start_column.min(b.start_column),
        end_column: a.end_column.max(b.end_column),
    }
}

pub(super) fn plan(index: &SheetIndex, sheet: &SheetMetadata) -> Vec<CellRange> {
    let mut bands = index.chunk_bounds.clone();
    let mut include = |area: MergedRange| {
        let end_row = area.end_row.min(sheet.row_count - 1);
        let end_column = area.end_column.min(sheet.column_count - 1);
        if area.start_row > end_row || area.start_column > end_column {
            return;
        }
        for band in (area.start_row / CHUNK_ROW_COUNT)..=(end_row / CHUNK_ROW_COUNT) {
            let range = CellRange {
                start_row: area.start_row.max(band * CHUNK_ROW_COUNT),
                end_row: end_row.min((band + 1) * CHUNK_ROW_COUNT - 1),
                start_column: area.start_column,
                end_column,
            };
            bands
                .entry(band)
                .and_modify(|bounds| *bounds = union(bounds, &range))
                .or_insert(range);
        }
    };
    if let Some(decorations) = &index.decorations {
        for &merge in &decorations.merges {
            include(merge);
        }
        for link in &decorations.hyperlinks {
            include(MergedRange {
                start_row: link.row,
                end_row: link.row,
                start_column: link.column,
                end_column: link.column,
            });
        }
    }
    for area in sheet
        .tables
        .iter()
        .map(|table| table.range)
        .chain(sheet.pivot_ranges.iter().copied())
    {
        include(area);
    }

    let mut ranges = bounded_reads(bands.into_values().filter_map(|mut band| {
        band.end_row = band.end_row.min(sheet.row_count - 1);
        band.end_column = band.end_column.min(sheet.column_count - 1);
        (band.start_row <= band.end_row && band.start_column <= band.end_column).then_some(band)
    }));
    if ranges.is_empty() {
        // Even an empty sheet needs one read for its sheet-wide metadata.
        ranges.push(point(0, 0));
    }
    ranges
}

fn bounded_reads(bands: impl Iterator<Item = CellRange>) -> Vec<CellRange> {
    let mut ranges: Vec<CellRange> = Vec::new();
    for band in bands {
        let rows_per_read = (MAX_RANGE_CELLS / (band.end_column - band.start_column + 1)).max(1);
        for start_row in (band.start_row..=band.end_row).step_by(rows_per_read) {
            let range = CellRange {
                start_row,
                end_row: band.end_row.min(start_row + rows_per_read - 1),
                start_column: band.start_column,
                end_column: band.end_column,
            };
            if let Some(last) = ranges
                .last_mut()
                .filter(|last| last.end_row + 1 == start_row)
            {
                let combined = union(last, &range);
                let rows = combined.end_row - combined.start_row + 1;
                let columns = combined.end_column - combined.start_column + 1;
                if rows.saturating_mul(columns) <= MAX_RANGE_CELLS {
                    *last = combined;
                    continue;
                }
            }
            ranges.push(range);
        }
    }
    ranges
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn coalesces_adjacent_dense_chunks_and_preserves_gaps_and_the_read_limit() {
        let bands = (0..100).map(|band| CellRange {
            start_row: band * CHUNK_ROW_COUNT,
            end_row: (band + 1) * CHUNK_ROW_COUNT - 1,
            start_column: 0,
            end_column: 1,
        });
        let wide = CellRange {
            start_row: 100_000,
            end_row: 100_255,
            start_column: 0,
            end_column: 999,
        };
        let reads = bounded_reads(bands.chain([wide]));
        let row_spans: Vec<_> = reads
            .iter()
            .map(|range| (range.start_row, range.end_row))
            .collect();
        assert_eq!(
            row_spans,
            [
                (0, 25_599),
                (100_000, 100_099),
                (100_100, 100_199),
                (100_200, 100_255)
            ]
        );
    }
}
