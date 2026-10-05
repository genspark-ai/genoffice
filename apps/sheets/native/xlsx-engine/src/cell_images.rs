//! Excel rich-value and WPS DISPIMG pictures displayed inside their host cells.

use super::*;

pub(crate) const MAX_CELL_IMAGES: usize = 500;

/// WPS image names resolve through cellimages.xml and its relationships. Like
/// optional Excel richData, malformed or missing picture parts are ignored.
pub(crate) fn read_wps_cell_images(archive: &mut ZipArchive<File>) -> HashMap<String, String> {
    read_wps_cell_images_inner(archive).unwrap_or_default()
}

fn read_wps_cell_images_inner(
    archive: &mut ZipArchive<File>,
) -> Result<HashMap<String, String>, SidecarError> {
    const PART: &str = "xl/cellimages.xml";
    let Some(xml) = visuals::read_optional_xml(archive, PART)? else {
        return Ok(HashMap::new());
    };
    let document = visuals::parse_document(&xml, PART)?;
    let relationships = visuals::read_relationships(archive, PART)?;
    let mut images = HashMap::new();
    for image in document
        .descendants()
        .filter(|node| node.has_tag_name("cellImage"))
    {
        let name = image
            .descendants()
            .find(|node| node.has_tag_name("cNvPr"))
            .and_then(|node| node.attribute("name"));
        let embed = image
            .descendants()
            .find(|node| node.has_tag_name("blip"))
            .and_then(|node| {
                node.attributes()
                    .find(|attribute| attribute.name() == "embed")
            })
            .map(|attribute| attribute.value());
        let (Some(name), Some(relationship)) = (name, embed.and_then(|id| relationships.get(id)))
        else {
            continue;
        };
        if !relationship.relationship_type.ends_with("/image") {
            continue;
        }
        let Ok(media_path) = visuals::resolve_part_target(PART, &relationship.target) else {
            continue;
        };
        if zip_entry(archive, &media_path).is_ok() {
            images.insert(name.to_owned(), media_path);
        }
    }
    Ok(images)
}

/// Only a DISPIMG formula identifies a picture; ordinary text mentioning the
/// function stays text, and missing images retain their cached cell value.
fn dispimg_name(formula: &str) -> Option<&str> {
    let (function, arguments) = formula.trim().trim_start_matches('=').split_once('(')?;
    if !function.trim().eq_ignore_ascii_case("DISPIMG")
        && !function.trim().eq_ignore_ascii_case("_xlfn.DISPIMG")
    {
        return None;
    }
    let (name, rest) = arguments.trim_start().strip_prefix('"')?.split_once('"')?;
    (rest.trim_start().starts_with(',') && rest.trim_end().ends_with(')')).then_some(name)
}

/// One record per cell, even when several WPS formulas reuse one image name.
/// Empty cells cannot capture a later cell's formula.
pub(crate) fn read_sheet_cell_images(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
    images_by_vm: &HashMap<u32, String>,
    images_by_name: &HashMap<String, String>,
    image_count: &mut usize,
) -> Result<Vec<CellImageInfo>, SidecarError> {
    if images_by_vm.is_empty() && images_by_name.is_empty() {
        return Ok(Vec::new());
    }
    let entry = zip_entry(archive, worksheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(entry));
    let mut buffer = Vec::new();
    let mut current_row = 0usize;
    let mut first_row = true;
    let mut next_column = 0usize;
    let mut cell_images = Vec::new();
    let mut current_cell = None;
    let mut in_formula = false;
    let mut formula = String::new();
    let mut push_image = |row, column, media_path: &String| {
        if cell_images.len() < MAX_CELL_IMAGES {
            cell_images.push(CellImageInfo {
                id: format!("cell-image-{image_count}"),
                row,
                column,
                media_path: media_path.clone(),
            });
            *image_count += 1;
        }
    };
    loop {
        let event = reader.read_event_into(&mut buffer)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"row" =>
            {
                current_row = attribute_value(&reader, &element, b"r")?
                    .and_then(|value| value.parse::<usize>().ok())
                    .filter(|value| *value > 0)
                    .map(|value| value - 1)
                    .unwrap_or(if first_row { 0 } else { current_row + 1 });
                first_row = false;
                next_column = 0;
            }
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"c" =>
            {
                // A corrupt address is the optional one in worse clothes, so
                // it lands where an omitted one does rather than failing the
                // whole open over a cell image.
                let (row, column) = match attribute_value(&reader, &element, b"r")? {
                    Some(address) => parse_address(&address).unwrap_or((current_row, next_column)),
                    None => (current_row, next_column),
                };
                next_column = column + 1;
                let media_path = attribute_value(&reader, &element, b"vm")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .and_then(|vm| images_by_vm.get(&vm));
                if let Some(media_path) = media_path {
                    push_image(row, column, media_path);
                }
                current_cell = (!empty && media_path.is_none()).then_some((row, column));
            }
            Event::Start(element) if element.local_name().as_ref() == b"f" => {
                in_formula = current_cell.is_some() && !images_by_name.is_empty();
                formula.clear();
            }
            Event::Text(text) if in_formula => formula.push_str(&decode_text(&text)?),
            Event::CData(text) if in_formula => formula.push_str(&decode_cdata(&text)?),
            Event::GeneralRef(reference) if in_formula => {
                formula.push_str(&general_ref_text(&reference)?)
            }
            Event::End(element) if element.local_name().as_ref() == b"f" => {
                if let Some(media_path) =
                    dispimg_name(&formula).and_then(|name| images_by_name.get(name))
                {
                    if let Some((row, column)) = current_cell.take() {
                        push_image(row, column, media_path);
                    }
                }
                in_formula = false;
            }
            Event::End(element) if element.local_name().as_ref() == b"c" => {
                current_cell = None;
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(cell_images)
}
