//! Reads a table description from stdin and appends its EMF (enhanced
//! metafile) rendering to the Windows clipboard, keeping the text and HTML
//! formats the grid already wrote. CAD paste dialogs (AutoCAD, ZWCAD and
//! their PASTESPEC) list the formats the clipboard actually carries, so the
//! EMF is what makes a "Picture (Enhanced Metafile)" entry appear for a
//! copied spreadsheet range (genoffice#1985, stage 1 — no OLE server, no
//! embedded-object activation).
//!
//! The helper never empties the clipboard: OpenClipboard → SetClipboardData
//! only adds the CF_ENHMETAFILE format, so Univer's text/plain and text/html
//! survive and regular paste targets behave exactly as before.
//!
//! Non-Windows builds are stubs: they parse and validate the payload, then
//! report the platform as unsupported, so CI compiles and tests the shared
//! input handling everywhere. Output is always one JSON line on stdout.

use serde::Deserialize;
use std::io::{Read, Write};

#[cfg(windows)]
use std::collections::{HashMap, HashSet};

const MAX_EDGE_PIXELS: i32 = 100_000;

#[derive(Deserialize)]
#[serde(deny_unknown_fields, rename_all = "camelCase")]
struct EmfRequest {
    col_widths: Vec<f64>,
    row_heights: Vec<f64>,
    cells: Vec<Vec<EmfCell>>,
    merges: Vec<EmfMerge>,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields)]
// The stub build only validates the payload without rendering it, so the
// fields stay construction-only there.
#[cfg_attr(not(windows), allow(dead_code))]
struct EmfCell {
    v: String,
    /// l = left (text), r = right (numbers), c = centered (booleans)
    a: EmfAlign,
}

#[derive(Deserialize)]
#[serde(rename_all = "lowercase")]
enum EmfAlign {
    L,
    R,
    C,
}

#[derive(Deserialize)]
#[serde(deny_unknown_fields, rename_all = "camelCase")]
struct EmfMerge {
    /// Merge origin relative to the copied range, plus its span in cells.
    r: usize,
    c: usize,
    rs: usize,
    cs: usize,
}

/// Payload checks that must hold before any GDI work: the geometry arrays
/// must match the cell matrix, merges must land inside it, and the metafile
/// must stay within a sane canvas so a runaway selection cannot allocate a
/// huge bitmap.
fn validate(request: &EmfRequest) -> Result<(), String> {
    if request.cells.is_empty() || request.col_widths.is_empty() {
        return Err("empty selection".into());
    }
    let rows = request.cells.len();
    let columns = request.col_widths.len();
    for (row, cells) in request.cells.iter().enumerate() {
        if cells.len() != columns {
            return Err(format!(
                "row {row} has {} cells, expected {columns}",
                cells.len()
            ));
        }
    }
    if request.row_heights.len() != rows {
        return Err(format!(
            "{} row heights for {rows} rows",
            request.row_heights.len()
        ));
    }
    if request
        .col_widths
        .iter()
        .any(|w| !w.is_finite() || *w < 0.0)
        || request
            .row_heights
            .iter()
            .any(|h| !h.is_finite() || *h < 0.0)
    {
        return Err("negative or non-finite geometry".into());
    }
    let total_w = request.col_widths.iter().sum::<f64>();
    let total_h = request.row_heights.iter().sum::<f64>();
    if total_w > f64::from(MAX_EDGE_PIXELS) || total_h > f64::from(MAX_EDGE_PIXELS) {
        return Err("selection too large for a metafile".into());
    }
    for merge in &request.merges {
        if merge.r >= rows
            || merge.c >= columns
            || merge.rs == 0
            || merge.cs == 0
            || merge.r + merge.rs > rows
            || merge.c + merge.cs > columns
        {
            return Err("merge outside the selection".into());
        }
    }
    Ok(())
}

#[cfg(windows)]
struct TableMetrics {
    width_px: i32,
    height_px: i32,
    column_edges: Vec<i32>,
    row_edges: Vec<i32>,
}

#[cfg(windows)]
fn table_metrics(request: &EmfRequest) -> Result<TableMetrics, String> {
    let to_px = |values: &[f64]| -> Result<Vec<i32>, String> {
        values
            .iter()
            .map(|value| {
                i32::try_from(value.round() as i64).map_err(|_| "geometry out of range".to_string())
            })
            .collect()
    };
    let columns = to_px(&request.col_widths)?;
    let rows = to_px(&request.row_heights)?;
    let width_px = columns.iter().sum::<i32>();
    let height_px = rows.iter().sum::<i32>();
    if width_px <= 0 || height_px <= 0 {
        return Err("selection has no extent".into());
    }
    let mut column_edges = Vec::with_capacity(columns.len() + 1);
    let mut edge = 0;
    column_edges.push(edge);
    for width in &columns {
        edge += width;
        column_edges.push(edge);
    }
    let mut row_edges = Vec::with_capacity(rows.len() + 1);
    edge = 0;
    row_edges.push(edge);
    for height in &rows {
        edge += height;
        row_edges.push(edge);
    }
    Ok(TableMetrics {
        width_px,
        height_px,
        column_edges,
        row_edges,
    })
}

#[cfg(windows)]
/// Merge anchors keyed by their top-left cell, and every cell the merges
/// cover (grid lines and stray text stay off those).
fn merge_layout(
    request: &EmfRequest,
) -> (
    HashMap<(usize, usize), (usize, usize, usize, usize)>,
    HashSet<(usize, usize)>,
) {
    let anchors: HashMap<(usize, usize), (usize, usize, usize, usize)> = request
        .merges
        .iter()
        .map(|merge| ((merge.r, merge.c), (merge.r, merge.c, merge.rs, merge.cs)))
        .collect();
    let covered: HashSet<(usize, usize)> = request
        .merges
        .iter()
        .flat_map(|merge| {
            let (r, c, rs, cs) = (merge.r, merge.c, merge.rs, merge.cs);
            (r..r + rs).flat_map(move |row| (c..c + cs).map(move |column| (row, column)))
        })
        .collect();
    (anchors, covered)
}

#[cfg(windows)]
fn render_and_append(request: &EmfRequest) -> Result<(i32, i32), String> {
    use windows_sys::Win32::Graphics::Gdi::DeleteEnhMetaFile;

    // The clipboard the copy produced. The helper starts right after that
    // copy, so this is the sequence number the copy wrote; if it moves while
    // the metafile is rendered, the clipboard belongs to someone else now and
    // the picture must not be appended to it.
    let sequence = clipboard_sequence();
    let metrics = table_metrics(request)?;
    let metafile = render_metafile(request, &metrics)?;
    if clipboard_sequence() != sequence {
        // Another copy (ours or another application's) owns the clipboard now:
        // drop the picture rather than putting a stale one beside it.
        unsafe { DeleteEnhMetaFile(metafile) };
        return Ok((metrics.width_px, metrics.height_px));
    }
    match append_to_clipboard(metafile, sequence) {
        Ok(()) => Ok((metrics.width_px, metrics.height_px)),
        Err(error) => Err(error),
    }
}

/// The clipboard's sequence number: it moves whenever the clipboard is
/// written, so a helper can tell whether the clipboard it was started for is
/// still the one it is about to append to.
#[cfg(windows)]
fn clipboard_sequence() -> u32 {
    use windows_sys::Win32::System::DataExchange::GetClipboardSequenceNumber;
    unsafe { GetClipboardSequenceNumber() }
}

#[cfg(windows)]
fn render_metafile(
    request: &EmfRequest,
    metrics: &TableMetrics,
) -> Result<*mut core::ffi::c_void, String> {
    use windows_sys::Win32::Foundation::RECT;
    use windows_sys::Win32::Graphics::Gdi::{
        CloseEnhMetaFile, CreateEnhMetaFileW, DeleteEnhMetaFile, GetDC, GetDeviceCaps, LOGPIXELSX,
        LOGPIXELSY, ReleaseDC,
    };

    let hdc_ref = unsafe { GetDC(std::ptr::null_mut()) };
    if hdc_ref.is_null() {
        return Err("no reference device context".into());
    }
    // The frame travels with the metafile in 0.01mm units, computed from the
    // reference DC's resolution, so CAD scales the picture at a sane size.
    let (dpi_x, dpi_y) = unsafe {
        (
            GetDeviceCaps(hdc_ref, LOGPIXELSX as i32).max(1),
            GetDeviceCaps(hdc_ref, LOGPIXELSY as i32).max(1),
        )
    };
    let frame_w = i64::from(metrics.width_px) * 2540 / i64::from(dpi_x);
    let frame_h = i64::from(metrics.height_px) * 2540 / i64::from(dpi_y);
    let frame = RECT {
        left: 0,
        top: 0,
        right: frame_w.max(1) as i32,
        bottom: frame_h.max(1) as i32,
    };
    let hdc = unsafe { CreateEnhMetaFileW(hdc_ref, std::ptr::null(), &frame, std::ptr::null()) };
    // the reference DC is only needed at creation time; release it once the
    // metafile context exists
    unsafe { ReleaseDC(std::ptr::null_mut(), hdc_ref) };
    if hdc.is_null() {
        return Err("no metafile device context".into());
    }
    let painted = unsafe { paint_table(hdc, request, metrics) };
    let handle = unsafe { CloseEnhMetaFile(hdc) };
    if handle.is_null() {
        return Err("metafile close failed".into());
    }
    if let Err(error) = painted {
        unsafe { DeleteEnhMetaFile(handle) };
        return Err(error);
    }
    Ok(handle)
}

#[cfg(windows)]
unsafe fn paint_table(
    hdc: *mut core::ffi::c_void,
    request: &EmfRequest,
    metrics: &TableMetrics,
) -> Result<(), String> {
    use windows_sys::Win32::Graphics::Gdi::{
        CreatePen, CreateSolidBrush, DeleteObject, PS_SOLID, Rectangle, SelectObject,
    };

    const GRID_COLOR: u32 = 0x00C0C0C0; // COLORREF is 0x00BBGGRR

    unsafe {
        // White ground first: a transparent metafile background renders black in
        // several CAD viewers.
        let ground = CreateSolidBrush(0x00FFFFFF);
        let old_brush = SelectObject(hdc, ground);
        let pen = CreatePen(PS_SOLID, 1, GRID_COLOR);
        let old_pen = SelectObject(hdc, pen);
        Rectangle(hdc, 0, 0, metrics.width_px, metrics.height_px);

        let (anchors, covered) = merge_layout(request);
        // The grey pen stays selected for the interior grid lines: restoring it
        // here would leave the default (black) pen for them, and the table would
        // come out with a grey frame and black inner lines.
        draw_grid_lines(hdc, metrics, &anchors);
        // The text keeps the default pen, as it did before.
        DeleteObject(SelectObject(hdc, old_pen));
        draw_cell_text(hdc, request, metrics, &anchors, &covered);
        SelectObject(hdc, old_brush);
        DeleteObject(ground);
    }
    Ok(())
}

#[cfg(windows)]
/// Interior grid lines only; the painted frame already draws the outer edge.
unsafe fn draw_grid_lines(
    hdc: *mut core::ffi::c_void,
    metrics: &TableMetrics,
    anchors: &HashMap<(usize, usize), (usize, usize, usize, usize)>,
) {
    use windows_sys::Win32::Graphics::Gdi::{LineTo, MoveToEx};

    unsafe {
        for (column, x) in metrics.column_edges.iter().enumerate().skip(1) {
            if column == metrics.column_edges.len() - 1 {
                break;
            }
            for row in 0..metrics.row_edges.len() - 1 {
                // Skip the inside of a merge: the anchor cell (row, column) maps
                // to a merge whose span covers this edge.
                if merge_covers_column(anchors, row, column) {
                    continue;
                }
                MoveToEx(hdc, *x, metrics.row_edges[row], std::ptr::null_mut());
                LineTo(hdc, *x, metrics.row_edges[row + 1]);
            }
        }
        for (row, y) in metrics.row_edges.iter().enumerate().skip(1) {
            if row == metrics.row_edges.len() - 1 {
                break;
            }
            for column in 0..metrics.column_edges.len() - 1 {
                if merge_covers_row(anchors, row, column) {
                    continue;
                }
                MoveToEx(hdc, metrics.column_edges[column], *y, std::ptr::null_mut());
                LineTo(hdc, metrics.column_edges[column + 1], *y);
            }
        }
    }
}

#[cfg(windows)]
/// A vertical edge at `column` between `row`-adjacent cells sits inside a
/// merge when some anchor spans across it.
fn merge_covers_column(
    anchors: &HashMap<(usize, usize), (usize, usize, usize, usize)>,
    row: usize,
    column: usize,
) -> bool {
    anchors
        .values()
        .any(|(r, c, rs, cs)| row >= *r && row < r + rs && column > *c && column < c + cs)
}

#[cfg(windows)]
fn merge_covers_row(
    anchors: &HashMap<(usize, usize), (usize, usize, usize, usize)>,
    row: usize,
    column: usize,
) -> bool {
    anchors
        .values()
        .any(|(r, c, rs, cs)| column >= *c && column < c + cs && row > *r && row < r + rs)
}

#[cfg(windows)]
/// Text paints once per merge anchor, clipped to the whole merged rectangle;
/// covered cells paint nothing.
unsafe fn draw_cell_text(
    hdc: *mut core::ffi::c_void,
    request: &EmfRequest,
    metrics: &TableMetrics,
    anchors: &HashMap<(usize, usize), (usize, usize, usize, usize)>,
    covered: &HashSet<(usize, usize)>,
) {
    use windows_sys::Win32::Foundation::SIZE;
    use windows_sys::Win32::Graphics::Gdi::{
        CLEARTYPE_QUALITY, CreateFontW, CreateRectRgn, DEFAULT_CHARSET, DEFAULT_PITCH,
        DeleteObject, FW_NORMAL, GetTextExtentPoint32W, SelectClipRgn, SelectObject, SetBkMode,
        SetTextColor, TRANSPARENT, TextOutW,
    };

    const FONT_HEIGHT: i32 = -12; // Segoe UI at 9pt
    const TEXT_INSET: i32 = 3;

    unsafe {
        let mut face: Vec<u16> = "Segoe UI".encode_utf16().collect();
        face.push(0);
        let font = CreateFontW(
            FONT_HEIGHT,
            0,
            0,
            0,
            FW_NORMAL as i32,
            0,
            0,
            0,
            u32::from(DEFAULT_CHARSET),
            0,
            0,
            u32::from(CLEARTYPE_QUALITY),
            u32::from(DEFAULT_PITCH),
            face.as_ptr(),
        );
        if font.is_null() {
            return;
        }
        let old_font = SelectObject(hdc, font);
        SetBkMode(hdc, TRANSPARENT as i32);
        SetTextColor(hdc, 0x00000000);

        for (row, cells) in request.cells.iter().enumerate() {
            for (column, cell) in cells.iter().enumerate() {
                if cell.v.is_empty() {
                    continue;
                }
                let is_anchor = anchors.contains_key(&(row, column));
                if !is_anchor && covered.contains(&(row, column)) {
                    continue;
                }
                let (r0, c0, rs, cs) = anchors
                    .get(&(row, column))
                    .copied()
                    .unwrap_or((row, column, 1, 1));
                let x0 = metrics.column_edges[c0];
                let y0 = metrics.row_edges[r0];
                let x1 = metrics.column_edges[(c0 + cs).min(metrics.column_edges.len() - 1)];
                let y1 = metrics.row_edges[(r0 + rs).min(metrics.row_edges.len() - 1)];
                let text: Vec<u16> = cell.v.encode_utf16().collect();
                let mut extent = SIZE { cx: 0, cy: 0 };
                if GetTextExtentPoint32W(hdc, text.as_ptr(), text.len() as i32, &mut extent) == 0 {
                    continue;
                }
                let anchor_x = match cell.a {
                    EmfAlign::R => x1 - TEXT_INSET - extent.cx,
                    EmfAlign::C => x0 + (x1 - x0 - extent.cx) / 2,
                    EmfAlign::L => x0 + TEXT_INSET,
                };
                let anchor_y = y0 + (y1 - y0 - extent.cy) / 2;
                let clip = CreateRectRgn(x0, y0, x1, y1);
                SelectClipRgn(hdc, clip);
                TextOutW(hdc, anchor_x, anchor_y, text.as_ptr(), text.len() as i32);
                SelectClipRgn(hdc, std::ptr::null_mut());
                DeleteObject(clip);
            }
        }

        SelectObject(hdc, old_font);
        DeleteObject(font);
    }
}

#[cfg(windows)]
fn append_to_clipboard(metafile: *mut core::ffi::c_void, sequence: u32) -> Result<(), String> {
    use windows_sys::Win32::Graphics::Gdi::DeleteEnhMetaFile;
    use windows_sys::Win32::System::DataExchange::{
        CloseClipboard, OpenClipboard, SetClipboardData,
    };
    use windows_sys::Win32::System::Ole::CF_ENHMETAFILE;
    use windows_sys::Win32::System::Threading::Sleep;

    for _ in 0..20 {
        if unsafe { OpenClipboard(std::ptr::null_mut()) } != 0 {
            // Re-checked in the success branch, not only before the loop: a
            // busy clipboard is usually another copy being written, so the
            // open may have waited out exactly the write this baseline exists
            // to detect — appending then would put a stale picture beside the
            // newer content.
            if clipboard_sequence() != sequence {
                unsafe { CloseClipboard() };
                unsafe { DeleteEnhMetaFile(metafile) };
                return Ok(());
            }
            if unsafe { SetClipboardData(CF_ENHMETAFILE as u32, metafile) }.is_null() {
                unsafe { CloseClipboard() };
                unsafe { DeleteEnhMetaFile(metafile) };
                return Err("SetClipboardData failed".into());
            }
            // The clipboard owns the handle on success; do not delete it.
            unsafe { CloseClipboard() };
            return Ok(());
        }
        unsafe { Sleep(50) };
    }
    unsafe { DeleteEnhMetaFile(metafile) };
    Err("clipboard stayed busy".into())
}

fn write_line(value: serde_json::Value) {
    let mut stdout = std::io::stdout().lock();
    let _ = writeln!(stdout, "{value}");
}

fn finish(result: Result<(i32, i32), String>) {
    match result {
        Ok((width, height)) => write_line(serde_json::json!({
            "ok": true,
            "width": width,
            "height": height,
        })),
        Err(error) => write_line(serde_json::json!({ "ok": false, "error": error })),
    }
}

fn main() {
    let mut input = String::new();
    if std::io::stdin().read_to_string(&mut input).is_err() {
        finish(Err("unreadable stdin".into()));
        return;
    }
    let request: EmfRequest = match serde_json::from_str(input.trim()) {
        Ok(request) => request,
        Err(error) => {
            finish(Err(format!("invalid payload: {error}")));
            return;
        }
    };
    if let Err(error) = validate(&request) {
        finish(Err(error));
        return;
    }
    #[cfg(windows)]
    {
        finish(render_and_append(&request));
    }
    #[cfg(not(windows))]
    {
        let _ = &request;
        finish(Err("unsupported platform".into()));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> EmfRequest {
        EmfRequest {
            col_widths: vec![80.0, 60.0],
            row_heights: vec![20.0, 20.0],
            cells: vec![
                vec![
                    EmfCell {
                        v: "Name".into(),
                        a: EmfAlign::L,
                    },
                    EmfCell {
                        v: "42".into(),
                        a: EmfAlign::R,
                    },
                ],
                vec![
                    EmfCell {
                        v: "Ada".into(),
                        a: EmfAlign::L,
                    },
                    EmfCell {
                        v: "".into(),
                        a: EmfAlign::L,
                    },
                ],
            ],
            merges: vec![],
        }
    }

    #[test]
    fn accepts_a_well_formed_table() {
        assert_eq!(validate(&sample()), Ok(()));
    }

    #[test]
    fn rejects_row_length_mismatches() {
        let mut request = sample();
        request.cells[0].pop();
        assert!(validate(&request).is_err());
    }

    #[test]
    fn rejects_row_height_mismatches() {
        let mut request = sample();
        request.row_heights = vec![20.0];
        assert!(validate(&request).is_err());
    }

    #[test]
    fn rejects_empty_selections() {
        let mut request = sample();
        request.cells.clear();
        assert!(validate(&request).is_err());
    }

    #[test]
    fn rejects_merges_outside_the_selection() {
        let mut request = sample();
        request.merges = vec![EmfMerge {
            r: 1,
            c: 0,
            rs: 2,
            cs: 1,
        }];
        assert!(validate(&request).is_err());
    }

    #[test]
    fn rejects_runaway_canvases() {
        let mut request = sample();
        request.col_widths = vec![f64::from(MAX_EDGE_PIXELS) + 1.0, 60.0];
        assert!(validate(&request).is_err());
    }

    #[test]
    fn rejects_non_finite_geometry() {
        let mut request = sample();
        request.row_heights[0] = f64::NAN;
        assert!(validate(&request).is_err());
    }

    #[test]
    fn parses_camel_case_payloads() {
        let payload = serde_json::json!({
            "colWidths": [80, 60],
            "rowHeights": [20, 20],
            "cells": [[{"v": "A", "a": "l"}, {"v": "1", "a": "r"}], [{"v": "", "a": "l"}, {"v": "", "a": "l"}]],
            "merges": [{"r": 0, "c": 0, "rs": 1, "cs": 2}],
        });
        let request: EmfRequest = serde_json::from_value(payload).unwrap();
        assert_eq!(validate(&request), Ok(()));
        assert_eq!(request.merges[0].cs, 2);
    }

    #[test]
    fn rejects_unknown_fields() {
        let payload = serde_json::json!({
            "colWidths": [80, 60],
            "rowHeights": [20, 20],
            "cells": [[{"v": "A", "a": "l"}, {"v": "", "a": "l"}]],
            "merges": [],
            "surprise": true,
        });
        assert!(serde_json::from_value::<EmfRequest>(payload).is_err());
    }
}
