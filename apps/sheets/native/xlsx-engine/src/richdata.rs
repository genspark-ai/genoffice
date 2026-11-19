//! Excel in-cell rich-value pictures ("place picture in cell"). The cell
//! carries a `vm` index and a cached #VALUE! error; the picture itself lives
//! behind xl/metadata.xml → xl/richData/richValueRel.xml → xl/media/*.

use std::collections::HashMap;
use std::fs::File;

use roxmltree::Node;
use zip::ZipArchive;

use crate::SidecarError;
use crate::visuals::{parse_document, read_optional_xml, read_relationships, resolve_part_target};

/// `vm` attribute value (1-based valueMetadata record index) → media part
/// path, for records whose rich value is a `_localImage`. Absent or
/// malformed richData parts degrade to "no in-cell pictures".
pub fn read_rich_value_images(archive: &mut ZipArchive<File>) -> HashMap<u32, String> {
    read_rich_value_images_inner(archive).unwrap_or_default()
}

fn read_rich_value_images_inner(
    archive: &mut ZipArchive<File>,
) -> Result<HashMap<u32, String>, SidecarError> {
    let workbook_rels = read_relationships(archive, "xl/workbook.xml")?;
    let part_by_type = |suffix: &str| {
        workbook_rels
            .values()
            .find(|relationship| relationship.relationship_type.ends_with(suffix))
            .map(|relationship| resolve_part_target("xl/workbook.xml", &relationship.target))
            .transpose()
    };
    let (Some(metadata_path), Some(rel_part_path), Some(value_path), Some(structure_path)) = (
        part_by_type("/sheetMetadata")?,
        part_by_type("/richValueRel")?,
        part_by_type("/rdRichValue")?,
        part_by_type("/rdRichValueStructure")?,
    ) else {
        return Ok(HashMap::new());
    };

    // Structure index → position of the local-image key among its <k> keys.
    let Some(structure_xml) = read_optional_xml(archive, &structure_path)? else {
        return Ok(HashMap::new());
    };
    let structure_doc = parse_document(&structure_xml, &structure_path)?;
    let structures: Vec<Option<usize>> = structure_doc
        .descendants()
        .filter(|node| node.has_tag_name("s"))
        .map(|node| {
            if node.attribute("t") != Some("_localImage") {
                return None;
            }
            node.children()
                .filter(|child| child.has_tag_name("k"))
                .position(|key| key.attribute("n") == Some("_rvRel:LocalImageIdentifier"))
        })
        .collect();

    // Rich value index → richValueRel entry index (image rich values only).
    let Some(value_xml) = read_optional_xml(archive, &value_path)? else {
        return Ok(HashMap::new());
    };
    let value_doc = parse_document(&value_xml, &value_path)?;
    let rich_values: Vec<Option<usize>> = value_doc
        .descendants()
        .filter(|node| node.has_tag_name("rv"))
        .map(|node| {
            let structure = node.attribute("s")?.parse::<usize>().ok()?;
            let key_position = (*structures.get(structure)?)?;
            let value = node
                .children()
                .filter(|child| child.has_tag_name("v"))
                .nth(key_position)?;
            value.text()?.trim().parse::<usize>().ok()
        })
        .collect();

    // richValueRel entry index → media part path (through the part's rels).
    let Some(rel_xml) = read_optional_xml(archive, &rel_part_path)? else {
        return Ok(HashMap::new());
    };
    let rel_doc = parse_document(&rel_xml, &rel_part_path)?;
    let part_rels = read_relationships(archive, &rel_part_path)?;
    let media_paths: Vec<Option<String>> = rel_doc
        .descendants()
        .filter(|node| node.has_tag_name("rel"))
        .map(|node| {
            let id = relationship_id(node)?;
            let relationship = part_rels.get(&id)?;
            resolve_part_target(&rel_part_path, &relationship.target).ok()
        })
        .collect();

    // valueMetadata record index (vm is that plus one) → media part path.
    let Some(metadata_xml) = read_optional_xml(archive, &metadata_path)? else {
        return Ok(HashMap::new());
    };
    let metadata_doc = parse_document(&metadata_xml, &metadata_path)?;
    let type_names: Vec<String> = metadata_doc
        .descendants()
        .filter(|node| node.has_tag_name("metadataType"))
        .map(|node| node.attribute("name").unwrap_or_default().to_owned())
        .collect();
    // futureMetadata bk records carry the rich value index (xlrd:rvb i=);
    // a record without one falls back to the identity mapping.
    let future_rich_value: Vec<Option<usize>> = metadata_doc
        .descendants()
        .find(|node| {
            node.has_tag_name("futureMetadata") && node.attribute("name") == Some("XLRICHVALUE")
        })
        .map(|future| {
            future
                .descendants()
                .filter(|node| node.has_tag_name("bk"))
                .map(|record| {
                    record
                        .descendants()
                        .find(|node| node.has_tag_name("rvb"))
                        .and_then(|rvb| rvb.attribute("i"))
                        .and_then(|index| index.parse::<usize>().ok())
                })
                .collect()
        })
        .unwrap_or_default();

    let mut by_vm = HashMap::new();
    let Some(value_metadata) = metadata_doc
        .descendants()
        .find(|node| node.has_tag_name("valueMetadata"))
    else {
        return Ok(by_vm);
    };
    for (index, record) in value_metadata
        .children()
        .filter(|node| node.has_tag_name("bk"))
        .enumerate()
    {
        // rc t= is a 1-based metadataTypes index; a bk may carry one rc per
        // metadata type, so find the XLRICHVALUE one wherever it sits.
        let Some(future_index) = record
            .children()
            .filter(|node| node.has_tag_name("rc"))
            .find(|rc| {
                rc.attribute("t")
                    .and_then(|value| value.parse::<usize>().ok())
                    .and_then(|value| value.checked_sub(1))
                    .and_then(|value| type_names.get(value))
                    .is_some_and(|name| name == "XLRICHVALUE")
            })
            .and_then(|rc| rc.attribute("v"))
            .and_then(|value| value.parse::<usize>().ok())
        else {
            continue;
        };
        let rich_value_index = future_rich_value
            .get(future_index)
            .copied()
            .flatten()
            .unwrap_or(future_index);
        let Some(rel_index) = rich_values.get(rich_value_index).copied().flatten() else {
            continue;
        };
        let Some(media_path) = media_paths.get(rel_index).cloned().flatten() else {
            continue;
        };
        by_vm.insert(index as u32 + 1, media_path);
    }
    Ok(by_vm)
}

fn relationship_id(node: Node<'_, '_>) -> Option<String> {
    node.attributes()
        .find(|attribute| attribute.name() == "id")
        .map(|attribute| attribute.value().to_owned())
}
