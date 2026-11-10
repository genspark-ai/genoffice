//! `x14:sparklineGroups` extension reader.

use super::*;

pub(crate) const SPARKLINE_EXT_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

pub(crate) const MAX_SPARKLINE_GROUPS: usize = 100;

pub(crate) const MAX_SPARKLINE_CELLS: usize = 500;

pub(crate) fn sparkline_type(value: Option<&str>) -> &'static str {
    match value {
        Some("column") => "column",
        Some("stacked") => "stacked",
        _ => "line",
    }
}

/// ARGB or RGB hex attribute → `#RRGGBB` (alpha dropped, no theme resolution).
pub(crate) fn argb_to_hex(value: &str) -> Option<String> {
    let hex = value.trim();
    let hex = match hex.len() {
        8 => &hex[2..],
        6 => hex,
        _ => return None,
    };
    hex.bytes()
        .all(|byte| byte.is_ascii_hexdigit())
        .then(|| format!("#{}", hex.to_ascii_uppercase()))
}

pub(crate) fn read_sheet_sparklines(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
) -> Result<Vec<SparklineGroupInfo>, SidecarError> {
    let entry = zip_entry(archive, worksheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(entry));
    parse_sparkline_groups(&mut reader)
}

pub(crate) fn parse_sparkline_groups<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<Vec<SparklineGroupInfo>, SidecarError> {
    enum Target {
        None,
        Formula,
        Sqref,
    }
    let mut groups = Vec::new();
    let mut buffer = Vec::new();
    let mut in_ext = false;
    let mut group: Option<SparklineGroupInfo> = None;
    let mut in_sparkline = false;
    let mut target = Target::None;
    let mut source_ref = String::new();
    let mut host_cell = String::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) if element.local_name().as_ref() == b"ext" => {
                in_ext = attribute_value(reader, &element, b"uri")?
                    .is_some_and(|uri| uri.eq_ignore_ascii_case(SPARKLINE_EXT_URI));
            }
            Event::End(element) if element.local_name().as_ref() == b"ext" => {
                in_ext = false;
                group = None;
                in_sparkline = false;
                target = Target::None;
            }
            Event::Start(element) | Event::Empty(element)
                if in_ext && element.local_name().as_ref() == b"sparklineGroup" =>
            {
                group = Some(SparklineGroupInfo {
                    kind: sparkline_type(attribute_value(reader, &element, b"type")?.as_deref())
                        .into(),
                    color: None,
                    negative_color: None,
                    cells: Vec::new(),
                });
            }
            Event::Start(element) | Event::Empty(element)
                if in_ext && element.local_name().as_ref() == b"colorSeries" =>
            {
                if let Some(group) = &mut group {
                    group.color = attribute_value(reader, &element, b"rgb")?
                        .as_deref()
                        .and_then(argb_to_hex);
                }
            }
            Event::Start(element) | Event::Empty(element)
                if in_ext && element.local_name().as_ref() == b"colorNegative" =>
            {
                if let Some(group) = &mut group {
                    group.negative_color = attribute_value(reader, &element, b"rgb")?
                        .as_deref()
                        .and_then(argb_to_hex);
                }
            }
            Event::Start(element) if in_ext && element.local_name().as_ref() == b"sparkline" => {
                in_sparkline = true;
                source_ref.clear();
                host_cell.clear();
            }
            Event::Start(element) if in_sparkline && element.local_name().as_ref() == b"f" => {
                target = Target::Formula;
            }
            Event::Start(element) if in_sparkline && element.local_name().as_ref() == b"sqref" => {
                target = Target::Sqref;
            }
            Event::Text(text) if in_sparkline => {
                let value = decode_text(&text)?;
                match target {
                    Target::Formula => source_ref.push_str(&value),
                    Target::Sqref => host_cell.push_str(&value),
                    Target::None => {}
                }
            }
            Event::CData(text) if in_sparkline => {
                let value = decode_cdata(&text)?;
                match target {
                    Target::Formula => source_ref.push_str(&value),
                    Target::Sqref => host_cell.push_str(&value),
                    Target::None => {}
                }
            }
            Event::GeneralRef(reference) if in_sparkline => {
                let value = general_ref_text(&reference)?;
                match target {
                    Target::Formula => source_ref.push_str(&value),
                    Target::Sqref => host_cell.push_str(&value),
                    Target::None => {}
                }
            }
            Event::End(element)
                if in_sparkline
                    && (element.local_name().as_ref() == b"f"
                        || element.local_name().as_ref() == b"sqref") =>
            {
                target = Target::None;
            }
            Event::End(element)
                if in_sparkline && element.local_name().as_ref() == b"sparkline" =>
            {
                in_sparkline = false;
                target = Target::None;
                let cell = host_cell.split_whitespace().next().unwrap_or("");
                let formula = source_ref.trim();
                if let Some(group) = &mut group {
                    if !cell.is_empty()
                        && !formula.is_empty()
                        && group.cells.len() < MAX_SPARKLINE_CELLS
                    {
                        group.cells.push(SparklineCellInfo {
                            cell: cell.to_owned(),
                            source_ref: formula.to_owned(),
                        });
                    }
                }
            }
            Event::End(element) if in_ext && element.local_name().as_ref() == b"sparklineGroup" => {
                if let Some(group) = group.take() {
                    if !group.cells.is_empty() && groups.len() < MAX_SPARKLINE_GROUPS {
                        groups.push(group);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(groups)
}
