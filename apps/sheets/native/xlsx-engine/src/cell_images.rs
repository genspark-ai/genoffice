//! Attribute-only scan for `<c vm=>` cells that resolve to richData image
//! parts ("Place in Cell" pictures).

use super::*;

pub(crate) const MAX_CELL_IMAGES: usize = 500;

/// Attribute-only scan for `<c vm=>` cells whose value metadata resolves to
/// an image rich value. Runs at open time and only when the workbook carries
/// richData image parts at all.
pub(crate) fn read_sheet_cell_images(
    archive: &mut ZipArchive<File>,
    worksheet_path: &str,
    images_by_vm: &HashMap<u32, String>,
    image_count: &mut usize,
) -> Result<Vec<CellImageInfo>, SidecarError> {
    if images_by_vm.is_empty() {
        return Ok(Vec::new());
    }
    let entry = zip_entry(archive, worksheet_path)?;
    let mut reader = Reader::from_reader(BufReader::new(entry));
    let mut buffer = Vec::new();
    let mut current_row = 0usize;
    let mut first_row = true;
    let mut next_column = 0usize;
    let mut cell_images = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer)? {
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
                let (row, column) = match attribute_value(&reader, &element, b"r")? {
                    Some(address) => parse_address(&address)?,
                    None => (current_row, next_column),
                };
                next_column = column + 1;
                let media_path = attribute_value(&reader, &element, b"vm")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .and_then(|vm| images_by_vm.get(&vm));
                if let Some(media_path) = media_path {
                    if cell_images.len() < MAX_CELL_IMAGES {
                        cell_images.push(CellImageInfo {
                            id: format!("cell-image-{image_count}"),
                            row,
                            column,
                            media_path: media_path.clone(),
                        });
                        *image_count += 1;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(cell_images)
}
