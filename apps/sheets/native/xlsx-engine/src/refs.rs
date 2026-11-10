//! A1 reference parsing: single addresses, `A1:B2` ranges and merge /
//! defined-name area references.

use super::*;

pub(crate) fn parse_area_reference(reference: &str) -> Option<MergedRange> {
    let cleaned = reference.replace('$', "");
    match cleaned.split_once(':') {
        Some((start, end)) => {
            let (start_row, start_column) = parse_address(start).ok()?;
            let (end_row, end_column) = parse_address(end).ok()?;
            (end_row >= start_row && end_column >= start_column).then_some(MergedRange {
                start_row,
                start_column,
                end_row,
                end_column,
            })
        }
        None => {
            let (row, column) = parse_address(&cleaned).ok()?;
            Some(MergedRange {
                start_row: row,
                start_column: column,
                end_row: row,
                end_column: column,
            })
        }
    }
}

pub(crate) fn parse_merge_reference(reference: &str) -> Option<MergedRange> {
    let (start, end) = reference.split_once(':')?;
    let (start_row, start_column) = parse_address(&start.replace('$', "")).ok()?;
    let (end_row, end_column) = parse_address(&end.replace('$', "")).ok()?;
    if end_row < start_row || end_column < start_column {
        return None;
    }
    Some(MergedRange {
        start_row,
        start_column,
        end_row,
        end_column,
    })
}

/// "A3:C20" (or a single "A3") → MergedRange; None on anything unparsable.
pub(crate) fn parse_range_reference(reference: &str) -> Option<MergedRange> {
    let cleaned = reference.replace('$', "");
    let mut parts = cleaned.split(':');
    let (start_row, start_column) = parse_address(parts.next()?).ok()?;
    let (end_row, end_column) = match parts.next() {
        Some(end) => parse_address(end).ok()?,
        None => (start_row, start_column),
    };
    Some(MergedRange {
        start_row: start_row.min(end_row),
        start_column: start_column.min(end_column),
        end_row: start_row.max(end_row),
        end_column: start_column.max(end_column),
    })
}

pub(crate) fn parse_address(address: &str) -> Result<(usize, usize), SidecarError> {
    let mut column = 0usize;
    let mut split = 0usize;
    for (index, byte) in address.bytes().enumerate() {
        if byte.is_ascii_alphabetic() {
            column = column
                .checked_mul(26)
                .and_then(|value| {
                    value.checked_add((byte.to_ascii_uppercase() - b'A' + 1) as usize)
                })
                .ok_or_else(|| SidecarError::Workbook("Cell column overflows.".into()))?;
            split = index + 1;
        } else {
            break;
        }
    }
    if split == 0 || split == address.len() {
        return Err(SidecarError::Workbook(format!(
            "Invalid cell address {address}."
        )));
    }
    let row = address[split..]
        .parse::<usize>()
        .map_err(|_| SidecarError::Workbook(format!("Invalid cell address {address}.")))?;
    if row == 0 || column == 0 {
        return Err(SidecarError::Workbook(format!(
            "Invalid cell address {address}."
        )));
    }
    Ok((row - 1, column - 1))
}
