//! `drawingN.xml` reader: anchors, pictures, shapes (preset and custom
//! geometry), text bodies and group transforms.

use super::*;

/// xdr:txBody paragraphs with per-run styling; a:br becomes a newline run.
pub(crate) fn parse_text_paragraphs(
    body: Node<'_, '_>,
    colors: &ColorContext,
) -> Vec<ShapeParagraph> {
    body.children()
        .filter(|node| node.has_tag_name("p"))
        .map(|paragraph| {
            let align = direct_child(paragraph, "pPr")
                .and_then(|node| node.attribute("algn"))
                .map(ToOwned::to_owned);
            let mut runs = Vec::new();
            for child in paragraph.children() {
                if child.has_tag_name("br") {
                    runs.push(ShapeRun {
                        text: "\n".into(),
                        color: None,
                        bold: false,
                        italic: false,
                        underline: false,
                        size: None,
                    });
                    continue;
                }
                if !child.has_tag_name("r") {
                    continue;
                }
                let text = direct_child(child, "t")
                    .and_then(|node| node.text())
                    .unwrap_or_default()
                    .to_owned();
                let properties = direct_child(child, "rPr");
                let flag = |name: &str| {
                    properties
                        .and_then(|rpr| rpr.attribute(name))
                        .is_some_and(|value| value == "1" || value == "true")
                };
                runs.push(ShapeRun {
                    text,
                    color: properties.and_then(|rpr| drawing_fill_color(rpr, colors)),
                    bold: flag("b"),
                    italic: flag("i"),
                    underline: properties
                        .and_then(|rpr| rpr.attribute("u"))
                        .is_some_and(|value| value != "none"),
                    size: properties
                        .and_then(|rpr| rpr.attribute("sz"))
                        .and_then(|value| value.parse::<f64>().ok())
                        .map(|value| value / 100.0),
                });
            }
            ShapeParagraph { align, runs }
        })
        .collect()
}

pub(crate) fn read_drawing(
    archive: &mut ZipArchive<File>,
    drawing_path: &str,
    sheet_id: &str,
    id_offset: usize,
    colors: &ColorContext,
    ole_shape_ids: &HashSet<u32>,
) -> Result<Vec<VisualObject>, SidecarError> {
    let xml = read_xml(archive, drawing_path)?;
    let document = parse_document(&xml, drawing_path)?;
    let relationships = read_relationships(archive, drawing_path)?;
    let mut visuals = Vec::new();
    for (index, anchor_node) in document
        .descendants()
        .filter(|node| {
            node.has_tag_name("twoCellAnchor")
                || node.has_tag_name("oneCellAnchor")
                || node.has_tag_name("absoluteAnchor")
        })
        .enumerate()
    {
        let Some(anchor) = parse_anchor(anchor_node) else {
            continue;
        };
        // Excel does not render objects flagged hidden (cNvPr hidden="1").
        // OLE compat fallback shapes are the exception: they are kept so the
        // caller can put the OLE visual in their z-order slot (borrowing the
        // anchor when the object has none), then dropped.
        let nv = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("cNvPr"));
        let hidden = nv.is_some_and(hidden_attribute);
        let is_ole_placeholder = nv
            .and_then(|node| node.attribute("id"))
            .and_then(|value| value.parse::<u32>().ok())
            .is_some_and(|id| ole_shape_ids.contains(&id));
        if hidden && !is_ole_placeholder {
            continue;
        }
        let visual_id = format!("visual-{}", id_offset + index + 1);
        if let Some(group) = anchor_node
            .children()
            .find(|node| node.has_tag_name("grpSp"))
        {
            // The group box spans the whole anchor; its EMU size comes from
            // the group's own xfrm ext (Excel keeps the two equal).
            let xfrm = group_xfrm(group);
            let width = xfrm_value(xfrm, "ext", "cx").unwrap_or(0.0);
            let height = xfrm_value(xfrm, "ext", "cy").unwrap_or(0.0);
            if width > 0.0 && height > 0.0 {
                let mut counter = 0;
                expand_group(
                    group,
                    &anchor,
                    (0.0, 0.0, width, height),
                    &visual_id,
                    &mut counter,
                    sheet_id,
                    colors,
                    drawing_path,
                    &relationships,
                    &mut visuals,
                )?;
            }
            continue;
        }
        if let Some(chart_node) = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("chart"))
        {
            let Some(id) = relationship_id(chart_node) else {
                continue;
            };
            let Some(relationship) = relationships.get(&id) else {
                continue;
            };
            let chart_path = resolve_part_target(drawing_path, &relationship.target)?;
            visuals.push(VisualObject {
                id: visual_id,
                sheet_id: sheet_id.to_owned(),
                kind: "chart".into(),
                anchor,
                chart: Some(read_chart(archive, &chart_path, colors)?),
                chart_path: Some(chart_path.clone()),
                media_path: None,
                media_type: None,
                opacity: None,
                crop: None,
                fill_media_path: None,
                fill_media_type: None,
                name: drawing_name(anchor_node),
                shape_type: None,
                custom_path: None,
                fill_color: None,
                fill_gradient: None,
                line_color: None,
                line_width: None,
                line_dash: None,
                line_cap: None,
                flip_h: false,
                flip_v: false,
                text_color: None,
                text_anchor: None,
                paragraphs: None,
                text: None,
                prog_id: None,
                rotation: None,
                frame_width: None,
                frame_height: None,
                nv_id: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: Some(index),
            });
            continue;
        }
        // Only an xdr:pic is a picture — an xdr:sp with a:blipFill keeps its
        // geometry and outline and falls through to the shape branch.
        if let Some(pic_node) = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("pic"))
        {
            let Some(blip_node) = pic_node
                .descendants()
                .find(|node| node.has_tag_name("blip"))
            else {
                continue;
            };
            let Some(id) = relationship_id(blip_node) else {
                continue;
            };
            let Some(relationship) = relationships.get(&id) else {
                continue;
            };
            // Excel draws nothing for a picture whose anchor collapses to a
            // point (from == to with equal offsets, or a 0x0 ext).
            if anchor_is_zero_extent(&anchor) {
                continue;
            }
            let media_path = resolve_part_target(drawing_path, &relationship.target)?;
            visuals.push(VisualObject {
                id: visual_id,
                sheet_id: sheet_id.to_owned(),
                kind: "image".into(),
                anchor,
                chart: None,
                chart_path: None,
                media_type: media_type_for_path(&media_path).map(ToOwned::to_owned),
                media_path: Some(media_path),
                opacity: blip_opacity(blip_node),
                crop: blip_crop(pic_node),
                fill_media_path: None,
                fill_media_type: None,
                name: drawing_name(anchor_node),
                shape_type: None,
                custom_path: None,
                fill_color: None,
                fill_gradient: None,
                line_color: None,
                line_width: None,
                line_dash: None,
                line_cap: None,
                flip_h: false,
                flip_v: false,
                text_color: None,
                text_anchor: None,
                paragraphs: None,
                text: None,
                prog_id: None,
                rotation: None,
                frame_width: None,
                frame_height: None,
                nv_id: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: Some(index),
            });
            continue;
        }
        if let Some(shape_node) = anchor_node
            .descendants()
            .find(|node| node.has_tag_name("sp") || node.has_tag_name("cxnSp"))
        {
            let mut shape = shape_visual(
                shape_node,
                anchor,
                visual_id,
                sheet_id,
                drawing_name(anchor_node),
                colors,
                drawing_path,
                &relationships,
                Some(index),
            );
            // nv_id pairs only the hidden OLE compat fallback with its
            // <oleObject>; a visible shape that happens to share a cNvPr id
            // with an OLE shapeId must stay a regular shape.
            if !(hidden && is_ole_placeholder) {
                shape.nv_id = None;
            }
            visuals.push(shape);
        }
    }
    // Chart children of expanded groups carry only their part path.
    for visual in &mut visuals {
        if visual.kind == "chart" && visual.chart.is_none() {
            if let Some(chart_path) = visual.chart_path.clone() {
                visual.chart = Some(read_chart(archive, &chart_path, colors)?);
            }
        }
    }
    Ok(visuals)
}

/// One `sp`/`cxnSp` node → a shape visual placed at `anchor`. Shared by the
/// direct anchor branch and grpSp expansion (children pass their own name
/// and no drawing_index — anchor edits would rewrite the whole group).
#[allow(clippy::too_many_arguments)]
pub(crate) fn shape_visual(
    shape_node: Node<'_, '_>,
    anchor: DrawingAnchor,
    visual_id: String,
    sheet_id: &str,
    name: Option<String>,
    colors: &ColorContext,
    drawing_path: &str,
    relationships: &HashMap<String, Relationship>,
    drawing_index: Option<usize>,
) -> VisualObject {
    {
        {
            let shape_type = shape_node
                .descendants()
                .find(|node| node.has_tag_name("prstGeom"))
                .and_then(|node| node.attribute("prst"))
                .map(ToOwned::to_owned);
            let custom_path = if shape_type.is_none() {
                parse_custom_geometry(shape_node)
            } else {
                None
            };
            let shape_sppr = shape_node.children().find(|node| node.has_tag_name("spPr"));
            // An explicit <a:noFill/> directly under spPr means transparent —
            // it must not fall through to the xdr:style fillRef theme color.
            let has_no_fill = shape_sppr
                .is_some_and(|sppr| sppr.children().any(|node| node.has_tag_name("noFill")));
            let fill_color = if has_no_fill {
                Some("none".into())
            } else {
                shape_sppr.and_then(|sppr| drawing_fill_color(sppr, colors))
            };
            let fill_blip = shape_sppr
                .and_then(|sppr| sppr.children().find(|node| node.has_tag_name("blipFill")))
                .and_then(|fill| fill.descendants().find(|node| node.has_tag_name("blip")));
            let fill_media_path = fill_blip
                .and_then(relationship_id)
                .and_then(|id| relationships.get(&id))
                .and_then(|relationship| {
                    resolve_part_target(drawing_path, &relationship.target).ok()
                });
            let fill_media_opacity = fill_media_path
                .as_ref()
                .and(fill_blip)
                .and_then(blip_opacity);
            let fill_media_type = fill_media_path
                .as_deref()
                .and_then(media_type_for_path)
                .map(ToOwned::to_owned);
            let xfrm = shape_node
                .descendants()
                .find(|node| node.has_tag_name("xfrm"));
            let rotation = xfrm
                .and_then(|node| node.attribute("rot"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|value| value / 60_000.0);
            // Only rotated shapes need the true frame — unrotated anchors
            // already are the frame.
            let frame_extent = |attribute: &str| {
                if rotation.is_none() {
                    return None;
                }
                xfrm_value(xfrm, "ext", attribute).filter(|value| *value > 0.0)
            };
            let frame_width = frame_extent("cx");
            let frame_height = frame_extent("cy");
            let flipped = |attribute: &str| {
                xfrm.and_then(|node| node.attribute(attribute))
                    .is_some_and(|value| value == "1" || value == "true")
            };
            let body = shape_node
                .descendants()
                .find(|node| node.has_tag_name("txBody"));
            let paragraphs = body
                .map(|node| parse_text_paragraphs(node, colors))
                .filter(|list| !list.is_empty());
            let text = paragraphs.as_ref().map(|list| {
                list.iter()
                    .map(|paragraph| {
                        paragraph
                            .runs
                            .iter()
                            .map(|run| run.text.as_str())
                            .collect::<String>()
                    })
                    .collect::<Vec<_>>()
                    .join("\n")
            });
            let text_anchor = body
                .and_then(|node| direct_child(node, "bodyPr"))
                .and_then(|node| node.attribute("anchor"))
                .map(ToOwned::to_owned);
            // xdr:style theme references are the fallback when spPr carries
            // no explicit fill/line (Excel's default for inserted shapes).
            let style_node = shape_node
                .children()
                .find(|node| node.has_tag_name("style"));
            let style_color = |name: &str| {
                let reference = style_node?
                    .children()
                    .find(|node| node.has_tag_name(name))?;
                let scheme = reference
                    .children()
                    .find(|node| node.has_tag_name("schemeClr"))?;
                let base = scheme_color_rgb(scheme.attribute("val")?, colors)?;
                Some(apply_color_modifiers(scheme, base))
            };
            let fill_gradient = if fill_color.is_none() && fill_media_path.is_none() {
                style_fill_gradient(style_node, colors)
            } else {
                None
            };
            // An explicit spPr blipFill replaces the style fillRef entirely:
            // the frame stays transparent while the image loads.
            let fill_color = fill_color.or_else(|| {
                if fill_media_path.is_some() {
                    Some("none".into())
                } else {
                    style_color("fillRef")
                }
            });
            let line_node = shape_node
                .children()
                .find(|node| node.has_tag_name("spPr"))
                .and_then(|sppr| sppr.children().find(|node| node.has_tag_name("ln")));
            let line_color = line_node
                .and_then(|ln| {
                    if ln.children().any(|node| node.has_tag_name("noFill")) {
                        return Some("none".into());
                    }
                    drawing_fill_color(ln, colors)
                })
                .or_else(|| style_color("lnRef"));
            let line_width = line_node
                .and_then(|ln| ln.attribute("w"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|emu| emu / 12_700.0);
            let line_dash = line_node
                .and_then(|ln| ln.children().find(|node| node.has_tag_name("prstDash")))
                .and_then(|node| node.attribute("val"))
                .map(ToOwned::to_owned);
            let line_cap = line_node
                .and_then(|ln| ln.attribute("cap"))
                .map(ToOwned::to_owned);
            let text_color = style_color("fontRef");
            let nv_id = shape_node
                .descendants()
                .find(|node| node.has_tag_name("cNvPr"))
                .and_then(|node| node.attribute("id"))
                .and_then(|value| value.parse::<u32>().ok());
            VisualObject {
                id: visual_id,
                sheet_id: sheet_id.to_owned(),
                kind: "shape".into(),
                anchor,
                chart: None,
                chart_path: None,
                media_path: None,
                media_type: None,
                opacity: fill_media_opacity,
                crop: None,
                fill_media_path,
                fill_media_type,
                name,
                shape_type,
                custom_path,
                fill_color,
                fill_gradient,
                line_color,
                line_width,
                line_dash,
                line_cap,
                flip_h: flipped("flipH"),
                flip_v: flipped("flipV"),
                text_color,
                text_anchor,
                paragraphs,
                text,
                prog_id: None,
                rotation,
                frame_width,
                frame_height,
                nv_id,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index,
            }
        }
    }
}

pub(crate) fn format_path_number(value: f64) -> String {
    if value.fract() == 0.0 && value.abs() < 1e15 {
        format!("{}", value as i64)
    } else {
        format!("{value:.2}")
    }
}

/// Parse a shape's a:custGeom pathLst into one SVG path string. Multiple
/// `<a:path>` entries scale into the first path's coordinate space. Returns
/// None when a command has no SVG mapping (arcTo) — the caller falls back
/// to the placeholder frame.
pub(crate) fn parse_custom_geometry(shape_node: Node<'_, '_>) -> Option<CustomPath> {
    let geometry = shape_node
        .descendants()
        .find(|node| node.has_tag_name("custGeom"))?;
    let path_list = geometry
        .children()
        .find(|node| node.has_tag_name("pathLst"))?;
    let mut base_width = 0.0_f64;
    let mut base_height = 0.0_f64;
    let mut d = String::new();
    let mut fill_d = String::new();
    let mut stroke_only = true;
    for path in path_list
        .children()
        .filter(|node| node.has_tag_name("path"))
    {
        let dimension = |attribute: &str| {
            path.attribute(attribute)
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| *value > 0.0)
                .unwrap_or(0.0)
        };
        let width = dimension("w");
        let height = dimension("h");
        if d.is_empty() {
            base_width = width;
            base_height = height;
        }
        let path_fills = path.attribute("fill") != Some("none");
        if path_fills {
            stroke_only = false;
        }
        let segment_start = d.len();
        let scale_x = if width > 0.0 && base_width > 0.0 {
            base_width / width
        } else {
            1.0
        };
        let scale_y = if height > 0.0 && base_height > 0.0 {
            base_height / height
        } else {
            1.0
        };
        for command in path.children().filter(|node| node.is_element()) {
            let points: Vec<(f64, f64)> = command
                .children()
                .filter(|node| node.has_tag_name("pt"))
                .filter_map(|point| {
                    Some((
                        point.attribute("x")?.parse::<f64>().ok()? * scale_x,
                        point.attribute("y")?.parse::<f64>().ok()? * scale_y,
                    ))
                })
                .collect();
            let (letter, expected) = match command.tag_name().name() {
                "moveTo" => ("M", 1),
                "lnTo" => ("L", 1),
                "cubicBezTo" => ("C", 3),
                "quadBezTo" => ("Q", 2),
                "close" => ("Z", 0),
                _ => return None,
            };
            if points.len() < expected {
                return None;
            }
            if !d.is_empty() {
                d.push(' ');
            }
            d.push_str(letter);
            for (x, y) in points.iter().take(expected) {
                d.push(' ');
                d.push_str(&format_path_number(*x));
                d.push(' ');
                d.push_str(&format_path_number(*y));
            }
        }
        if path_fills && d.len() > segment_start {
            if !fill_d.is_empty() {
                fill_d.push(' ');
            }
            fill_d.push_str(d[segment_start..].trim_start());
        }
    }
    if d.is_empty() {
        return None;
    }
    let fill_d = if stroke_only || fill_d == d {
        None
    } else {
        Some(fill_d)
    };
    Some(CustomPath {
        width: base_width.max(1.0),
        height: base_height.max(1.0),
        d,
        stroke_only,
        fill_d,
    })
}

pub(crate) fn hidden_attribute(node: Node<'_, '_>) -> bool {
    node.attribute("hidden")
        .is_some_and(|value| value == "1" || value == "true")
}

pub(crate) fn group_xfrm<'a>(group: Node<'a, 'a>) -> Option<Node<'a, 'a>> {
    group
        .children()
        .find(|node| node.has_tag_name("grpSpPr"))?
        .children()
        .find(|node| node.has_tag_name("xfrm"))
}

pub(crate) fn xfrm_value(xfrm: Option<Node<'_, '_>>, tag: &str, attribute: &str) -> Option<f64> {
    xfrm?
        .children()
        .find(|node| node.has_tag_name(tag))?
        .attribute(attribute)?
        .parse::<f64>()
        .ok()
}

/// Flatten a grpSp into per-child visuals. `group_box` is the group's frame
/// as (x, y, width, height) in EMU relative to the anchor's `from` marker;
/// children map from the group's chOff/chExt space onto that box. Child
/// anchors encode the box as offsets within the from cell, which the
/// renderer's marker walk resolves across real row/column sizes.
#[allow(clippy::too_many_arguments)]
pub(crate) fn expand_group(
    group: Node<'_, '_>,
    anchor: &DrawingAnchor,
    group_box: (f64, f64, f64, f64),
    visual_id: &str,
    counter: &mut usize,
    sheet_id: &str,
    colors: &ColorContext,
    drawing_path: &str,
    relationships: &HashMap<String, Relationship>,
    visuals: &mut Vec<VisualObject>,
) -> Result<(), SidecarError> {
    let (box_x, box_y, box_width, box_height) = group_box;
    let xfrm = group_xfrm(group);
    let ch_off_x = xfrm_value(xfrm, "chOff", "x").unwrap_or(0.0);
    let ch_off_y = xfrm_value(xfrm, "chOff", "y").unwrap_or(0.0);
    let ch_ext_x = xfrm_value(xfrm, "chExt", "cx").filter(|value| *value > 0.0);
    let ch_ext_y = xfrm_value(xfrm, "chExt", "cy").filter(|value| *value > 0.0);
    let scale_x = box_width / ch_ext_x.unwrap_or(box_width);
    let scale_y = box_height / ch_ext_y.unwrap_or(box_height);
    for child in group.children() {
        let is_group = child.has_tag_name("grpSp");
        let is_shape = child.has_tag_name("sp") || child.has_tag_name("cxnSp");
        let is_picture = child.has_tag_name("pic");
        let is_frame = child.has_tag_name("graphicFrame");
        if !is_group && !is_shape && !is_picture && !is_frame {
            continue;
        }
        if child
            .descendants()
            .find(|node| node.has_tag_name("cNvPr"))
            .is_some_and(hidden_attribute)
        {
            continue;
        }
        // The first xfrm under the child is its own (spPr or grpSpPr).
        let child_xfrm = child.descendants().find(|node| node.has_tag_name("xfrm"));
        let (Some(off_x), Some(off_y), Some(ext_x), Some(ext_y)) = (
            xfrm_value(child_xfrm, "off", "x"),
            xfrm_value(child_xfrm, "off", "y"),
            xfrm_value(child_xfrm, "ext", "cx"),
            xfrm_value(child_xfrm, "ext", "cy"),
        ) else {
            continue;
        };
        let child_box = (
            box_x + (off_x - ch_off_x) * scale_x,
            box_y + (off_y - ch_off_y) * scale_y,
            ext_x * scale_x,
            ext_y * scale_y,
        );
        if is_group {
            expand_group(
                child,
                anchor,
                child_box,
                visual_id,
                counter,
                sheet_id,
                colors,
                drawing_path,
                relationships,
                visuals,
            )?;
            continue;
        }
        let child_anchor = DrawingAnchor {
            from_row: anchor.from_row,
            from_column: anchor.from_column,
            from_row_offset: anchor.from_row_offset + child_box.1.round() as i64,
            from_column_offset: anchor.from_column_offset + child_box.0.round() as i64,
            to_row: anchor.from_row,
            to_column: anchor.from_column,
            to_row_offset: anchor.from_row_offset + (child_box.1 + child_box.3).round() as i64,
            to_column_offset: anchor.from_column_offset
                + (child_box.0 + child_box.2).round() as i64,
            explicit_to: false,
        };
        *counter += 1;
        let child_id = format!("{visual_id}-{counter}");
        let child_name = child
            .descendants()
            .find(|node| node.has_tag_name("cNvPr"))
            .and_then(|node| node.attribute("name"))
            .map(ToOwned::to_owned);
        if is_shape {
            visuals.push(shape_visual(
                child,
                child_anchor,
                child_id,
                sheet_id,
                child_name,
                colors,
                drawing_path,
                relationships,
                None,
            ));
            continue;
        }
        if is_frame {
            // Chart data is backfilled by read_drawing (reading the part
            // needs the archive, which this expansion deliberately avoids).
            let Some(chart_path) = child
                .descendants()
                .find(|node| node.has_tag_name("chart"))
                .and_then(relationship_id)
                .and_then(|id| relationships.get(&id))
                .map(|relationship| resolve_part_target(drawing_path, &relationship.target))
                .transpose()?
            else {
                continue;
            };
            visuals.push(VisualObject {
                id: child_id,
                sheet_id: sheet_id.to_owned(),
                kind: "chart".into(),
                anchor: child_anchor,
                chart: None,
                chart_path: Some(chart_path),
                media_path: None,
                media_type: None,
                opacity: None,
                crop: None,
                fill_media_path: None,
                fill_media_type: None,
                name: child_name,
                shape_type: None,
                custom_path: None,
                fill_color: None,
                fill_gradient: None,
                line_color: None,
                line_width: None,
                line_dash: None,
                line_cap: None,
                flip_h: false,
                flip_v: false,
                text_color: None,
                text_anchor: None,
                paragraphs: None,
                text: None,
                prog_id: None,
                rotation: None,
                frame_width: None,
                frame_height: None,
                nv_id: None,
                drawing_path: Some(drawing_path.to_owned()),
                drawing_index: None,
            });
            continue;
        }
        let Some(blip_node) = child.descendants().find(|node| node.has_tag_name("blip")) else {
            continue;
        };
        let Some(id) = relationship_id(blip_node) else {
            continue;
        };
        let Some(relationship) = relationships.get(&id) else {
            continue;
        };
        if anchor_is_zero_extent(&child_anchor) {
            continue;
        }
        let media_path = resolve_part_target(drawing_path, &relationship.target)?;
        visuals.push(VisualObject {
            id: child_id,
            sheet_id: sheet_id.to_owned(),
            kind: "image".into(),
            anchor: child_anchor,
            chart: None,
            chart_path: None,
            media_type: media_type_for_path(&media_path).map(ToOwned::to_owned),
            media_path: Some(media_path),
            opacity: blip_opacity(blip_node),
            crop: blip_crop(child),
            fill_media_path: None,
            fill_media_type: None,
            name: child_name,
            shape_type: None,
            custom_path: None,
            fill_color: None,
            fill_gradient: None,
            line_color: None,
            line_width: None,
            line_dash: None,
            line_cap: None,
            flip_h: false,
            flip_v: false,
            text_color: None,
            text_anchor: None,
            paragraphs: None,
            text: None,
            prog_id: None,
            rotation: None,
            frame_width: None,
            frame_height: None,
            nv_id: None,
            drawing_path: Some(drawing_path.to_owned()),
            drawing_index: None,
        });
    }
    Ok(())
}

/// Both markers coincide — a zero span on both axes. (The no-ext fallbacks
/// in parse_anchor always produce a real span, so they never match.)
pub(crate) fn anchor_is_zero_extent(anchor: &DrawingAnchor) -> bool {
    anchor.from_row == anchor.to_row
        && anchor.from_row_offset == anchor.to_row_offset
        && anchor.from_column == anchor.to_column
        && anchor.from_column_offset == anchor.to_column_offset
}

/// a:blip/a:alphaModFix amt (per-100000) → 0..1; None when absent or opaque.
pub(crate) fn blip_opacity(blip_node: Node<'_, '_>) -> Option<f64> {
    let amt = blip_node
        .children()
        .find(|node| node.has_tag_name("alphaModFix"))?
        .attribute("amt")?
        .parse::<f64>()
        .ok()?;
    let opacity = (amt / 100_000.0).clamp(0.0, 1.0);
    (opacity < 1.0).then_some(opacity)
}

/// a:srcRect l/t/r/b (per-100000 of the source size) under `scope` → 0..1
/// crop fractions; None when absent, all-zero, or nothing would remain.
pub(crate) fn blip_crop(scope: Node<'_, '_>) -> Option<CropRect> {
    let rect = scope
        .descendants()
        .find(|node| node.has_tag_name("srcRect"))?;
    let side = |name: &str| {
        rect.attribute(name)
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 100_000.0)
            .unwrap_or(0.0)
    };
    let sides = [side("l"), side("t"), side("r"), side("b")];
    if sides
        .iter()
        .any(|value| !value.is_finite() || value.abs() > 1.0)
        || sides.iter().all(|value| *value == 0.0)
        || sides[0] + sides[2] >= 1.0
        || sides[1] + sides[3] >= 1.0
    {
        return None;
    }
    Some(CropRect {
        left: sides[0],
        top: sides[1],
        right: sides[2],
        bottom: sides[3],
    })
}

pub(crate) fn parse_anchor(anchor: Node<'_, '_>) -> Option<DrawingAnchor> {
    let Some(from) = direct_child(anchor, "from") else {
        // absoluteAnchor: xdr:pos + xdr:ext in sheet EMU. Encode both corners
        // as offsets from cell (0,0) — the renderer's marker walk carries
        // offsets across real row/column sizes.
        let pos = direct_child(anchor, "pos")?;
        let ext = direct_child(anchor, "ext")?;
        let coordinate =
            |node: Node<'_, '_>, attribute: &str| node.attribute(attribute)?.parse::<i64>().ok();
        let x = coordinate(pos, "x")?;
        let y = coordinate(pos, "y")?;
        let cx = coordinate(ext, "cx")?;
        let cy = coordinate(ext, "cy")?;
        return Some(DrawingAnchor {
            from_row: 0,
            from_column: 0,
            from_row_offset: y,
            from_column_offset: x,
            to_row: 0,
            to_column: 0,
            to_row_offset: y + cy,
            to_column_offset: x + cx,
            explicit_to: false,
        });
    };
    let from_row = marker_value(from, "row")?;
    let from_column = marker_value(from, "col")?;
    let from_row_offset = marker_signed_value(from, "rowOff").unwrap_or(0);
    let from_column_offset = marker_signed_value(from, "colOff").unwrap_or(0);
    if let Some(to) = direct_child(anchor, "to") {
        return Some(DrawingAnchor {
            from_row,
            from_column,
            from_row_offset,
            from_column_offset,
            to_row: marker_value(to, "row").unwrap_or(from_row + 20),
            to_column: marker_value(to, "col").unwrap_or(from_column + 8),
            to_row_offset: marker_signed_value(to, "rowOff").unwrap_or(0),
            to_column_offset: marker_signed_value(to, "colOff").unwrap_or(0),
            explicit_to: true,
        });
    }
    // oneCellAnchor: the size lives in xdr:ext (EMU). Encode it as offsets
    // within the from cell — the renderer's marker walk handles offsets past
    // the cell edge, so no new anchor fields are needed. (Previously `to`
    // fell back to `from` itself: a zero-size box.)
    let ext = direct_child(anchor, "ext");
    let extent = |attribute: &str| {
        ext.and_then(|node| node.attribute(attribute))
            .and_then(|value| value.parse::<i64>().ok())
    };
    match (extent("cx"), extent("cy")) {
        (Some(cx), Some(cy)) => Some(DrawingAnchor {
            from_row,
            from_column,
            from_row_offset,
            from_column_offset,
            to_row: from_row,
            to_column: from_column,
            to_row_offset: from_row_offset + cy,
            to_column_offset: from_column_offset + cx,
            explicit_to: false,
        }),
        _ => Some(DrawingAnchor {
            from_row,
            from_column,
            from_row_offset,
            from_column_offset,
            to_row: from_row + 20,
            to_column: from_column + 8,
            to_row_offset: 0,
            to_column_offset: 0,
            explicit_to: false,
        }),
    }
}
