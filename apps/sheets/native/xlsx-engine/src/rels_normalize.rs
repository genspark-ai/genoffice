//! Normalizes OPC relationship targets before the IronCalc import.
//!
//! IronCalc 0.9.1 resolves a worksheet's comments and table relationships by
//! assuming a relative target that begins with `../` (`load_sheet_rels` in
//! import/worksheets.rs), while OPC (ECMA-376 Part 2) also allows absolute
//! targets: a leading `/` means "from the package root". openpyxl writes the
//! comments and VML drawing relationships that way, so the importer asks the
//! archive for a bogus part name and the whole load fails with the opaque
//! `Zip Error: specified file not found in archive` (genoffice#2306).
//!
//! IronCalc is an external crate, so the repair stays on our side of the
//! import boundary: a package that uses absolute targets is copied in memory
//! with those targets rewritten to the equivalent relative form and imported
//! from those bytes. A package whose targets are all relative keeps the plain
//! path import — no copy, no behavior change, just the small `.rels` parts
//! read to prove the fixup is unnecessary.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Write};
use std::path::Path;

use ironcalc::base::Model;
use ironcalc::import::{load_from_xlsx, load_from_xlsx_bytes};
use roxmltree::Document;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::SidecarError;
use crate::archive::{canonical_entry_name, resolve_relationship_target};
use crate::xml_util::{MAX_METADATA_PART_BYTES, copy_entry_bounded};

/// The fixed namespace of a `.rels` part. The normalized copy states it
/// explicitly because the part is rebuilt from its parsed relationships.
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";

/// Imports the workbook at `path` into an IronCalc model, repairing absolute
/// relationship targets first when the package uses them.
pub(crate) fn load_model(path: &Path) -> Result<Model<'static>, SidecarError> {
    let Some(normalized) = normalized_package(path)? else {
        let path_text = path.to_str().ok_or_else(|| {
            SidecarError::InvalidRequest("Workbook path is not valid UTF-8.".into())
        })?;
        return load_from_xlsx(path_text, "en", "UTC", "en").map_err(import_error);
    };
    let name = path
        .file_stem()
        .map(|stem| stem.to_string_lossy().into_owned())
        .unwrap_or_default();
    let workbook = load_from_xlsx_bytes(&normalized, &name, "en", "UTC").map_err(import_error)?;
    Model::from_workbook(workbook, "en").map_err(import_error)
}

fn import_error(error: impl std::fmt::Display) -> SidecarError {
    SidecarError::Workbook(format!("Formula engine import failed: {error}"))
}

/// One relationship part of the package.
struct RelationshipPart {
    /// Entry index in the source archive.
    index: usize,
    /// Entry name as stored in the archive (used for the rewritten entry).
    name: String,
    /// The part whose relationships the file holds, e.g.
    /// `xl/worksheets/sheet1.xml` for `xl/worksheets/_rels/sheet1.xml.rels`.
    /// Empty for the package-root relationships (`_rels/.rels`).
    source: String,
}

/// The package with absolute relationship targets rewritten, or `None` when
/// every target is already relative and the original file can be imported
/// as is.
fn normalized_package(path: &Path) -> Result<Option<Vec<u8>>, SidecarError> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file)?;
    let parts = relationship_parts(&mut archive)?;
    let mut needs_normalization = false;
    for part in &parts {
        let xml = read_part(&mut archive, part.index)?;
        if has_absolute_target(&xml) {
            needs_normalization = true;
            break;
        }
    }
    if !needs_normalization {
        return Ok(None);
    }
    // Every part is rewritten in one pass, not just the ones that carry an
    // absolute target: the copy is handed to the importer wholesale, so this
    // is the one place that can report a dangling relationship part by name.
    let mut rewrites: HashMap<usize, String> = HashMap::new();
    for part in &parts {
        let xml = read_part(&mut archive, part.index)?;
        if let Some(rewritten) = rewrite_part(&xml, part, &mut archive)? {
            rewrites.insert(part.index, rewritten);
        }
    }
    if rewrites.is_empty() {
        // Detection promised at least one absolute target; if resolution
        // rejected it there is nothing to hand over differently.
        return Ok(None);
    }
    write_package(&mut archive, &rewrites).map(Some)
}

fn relationship_parts(
    archive: &mut ZipArchive<File>,
) -> Result<Vec<RelationshipPart>, SidecarError> {
    let mut parts = Vec::new();
    for index in 0..archive.len() {
        let name = archive.by_index_raw(index)?.name().to_string();
        if name.ends_with('/') {
            continue;
        }
        if let Some(source) = relationship_source_part(&name) {
            parts.push(RelationshipPart {
                index,
                name,
                source,
            });
        }
    }
    Ok(parts)
}

/// Maps `xl/worksheets/_rels/sheet1.xml.rels` to `xl/worksheets/sheet1.xml`,
/// and the package-root `_rels/.rels` to the empty (package root) part.
fn relationship_source_part(name: &str) -> Option<String> {
    let canonical = canonical_entry_name(name)?;
    let (directory, file_name) = canonical.rsplit_once('/')?;
    let source_directory = directory
        .strip_suffix("/_rels")
        .or_else(|| (directory == "_rels").then_some(""))?;
    let source_file = file_name.strip_suffix(".rels")?;
    match (source_directory.is_empty(), source_file.is_empty()) {
        // `_rels/.rels` holds the package root's relationships.
        (true, true) => Some(String::new()),
        (true, false) => Some(source_file.to_owned()),
        (false, true) => None,
        (false, false) => Some(format!("{source_directory}/{source_file}")),
    }
}

fn read_part(archive: &mut ZipArchive<File>, index: usize) -> Result<String, SidecarError> {
    let mut entry = archive.by_index(index)?;
    let name = entry.name().to_string();
    let declared = entry.size();
    if declared > MAX_METADATA_PART_BYTES {
        return Err(SidecarError::Workbook(format!(
            "{name} is {declared} bytes uncompressed, above the {MAX_METADATA_PART_BYTES} byte metadata part limit."
        )));
    }
    let mut bytes = Vec::new();
    copy_entry_bounded(&mut entry, declared, MAX_METADATA_PART_BYTES, &mut bytes)?;
    String::from_utf8(bytes)
        .map_err(|_| SidecarError::Workbook(format!("{name} is not valid UTF-8.")))
}

/// True when the part holds a relationship (other than an external one) whose
/// target starts with `/`. Malformed XML is not diagnosed here: the importer
/// reports it exactly as before.
fn has_absolute_target(xml: &str) -> bool {
    let Ok(document) = Document::parse(xml) else {
        return false;
    };
    document
        .root_element()
        .descendants()
        .filter(|node| node.is_element() && node.has_tag_name("Relationship"))
        .any(|node| !is_external(node) && node.attribute("Target").is_some_and(is_absolute_target))
}

fn is_external(relationship: roxmltree::Node<'_, '_>) -> bool {
    relationship
        .attribute("TargetMode")
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
}

fn is_absolute_target(target: &str) -> bool {
    let target = target.trim();
    // `//host/share` is a network path, not an absolute OPC part name.
    target.starts_with('/') && !target.starts_with("//")
}

/// Rebuilds one relationship part with absolute targets made relative.
/// Returns `None` when the part has nothing to rewrite.
fn rewrite_part(
    xml: &str,
    part: &RelationshipPart,
    archive: &mut ZipArchive<File>,
) -> Result<Option<String>, SidecarError> {
    let document = Document::parse(xml).map_err(|error| {
        SidecarError::Workbook(format!("Invalid XML in {}: {error}", part.name))
    })?;
    let root = document.root_element();
    let mut changed = false;
    let mut out = String::with_capacity(xml.len() + 64);
    out.push('<');
    out.push_str(root.tag_name().name());
    out.push_str(" xmlns=\"");
    out.push_str(RELATIONSHIPS_NAMESPACE);
    out.push('"');
    push_attributes(&mut out, root);
    out.push('>');
    for child in root.children().filter(|node| node.is_element()) {
        out.push('<');
        out.push_str(child.tag_name().name());
        let relationship = child.has_tag_name("Relationship");
        let external = relationship && is_external(child);
        for attribute in child.attributes() {
            let rewritten = if relationship && !external && attribute.name() == "Target" {
                normalize_target(&part.source, attribute.value(), archive)?
            } else {
                None
            };
            out.push(' ');
            out.push_str(attribute.name());
            out.push_str("=\"");
            push_escaped(&mut out, rewritten.as_deref().unwrap_or(attribute.value()));
            out.push('"');
            changed |= rewritten.is_some();
        }
        out.push_str("/>");
    }
    out.push_str("</");
    out.push_str(root.tag_name().name());
    out.push('>');
    Ok(changed.then_some(out))
}

fn push_attributes(out: &mut String, node: roxmltree::Node<'_, '_>) {
    for attribute in node.attributes() {
        out.push(' ');
        out.push_str(attribute.name());
        out.push_str("=\"");
        push_escaped(out, attribute.value());
        out.push('"');
    }
}

fn push_escaped(out: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(character),
        }
    }
}

/// Resolves one relationship target and returns its relative spelling when
/// the target is absolute. Relative targets are returned untouched (but still
/// resolved, so a dangling part is reported by name rather than as the zip
/// error this module exists to replace).
fn normalize_target(
    source_part: &str,
    target: &str,
    archive: &mut ZipArchive<File>,
) -> Result<Option<String>, SidecarError> {
    let absolute = is_absolute_target(target);
    let resolved = match resolve_relationship_target(source_part, target) {
        Ok(resolved) => resolved,
        Err(error) if absolute => {
            return Err(SidecarError::Workbook(format!(
                "{} relationship target {target} could not be resolved: {error}",
                source_label(source_part)
            )));
        }
        // Relative targets are not all package part names: a hyperlink can be
        // a location-only fragment or a URL written without TargetMode. Those
        // stay the importer's business, exactly as before.
        Err(_) => return Ok(None),
    };
    if crate::zip_entry(archive, &resolved).is_err() {
        return Err(SidecarError::Workbook(format!(
            "workbook relationship target not found: {} -> {resolved} (resolved from {target})",
            source_label(source_part)
        )));
    }
    if !absolute {
        return Ok(None);
    }
    let mut rewritten = relative_target(source_part, &resolved);
    if let Some((_, fragment)) = target.trim().split_once('#') {
        if !fragment.is_empty() {
            rewritten.push('#');
            rewritten.push_str(fragment);
        }
    }
    Ok(Some(rewritten))
}

fn source_label(source_part: &str) -> &str {
    if source_part.is_empty() {
        "package root"
    } else {
        source_part
    }
}

/// The path from the directory of `source_part` to the package part `target`:
/// `xl/comments/comment1.xml` from `xl/worksheets/sheet1.xml` is
/// `../comments/comment1.xml`.
fn relative_target(source_part: &str, target: &str) -> String {
    let source = source_part.replace('\\', "/");
    let mut base: Vec<&str> = source
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect();
    base.pop();
    let target: Vec<&str> = target
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect();
    let mut common = 0;
    while common < base.len() && common < target.len() && base[common] == target[common] {
        common += 1;
    }
    let mut relative = String::new();
    for _ in common..base.len() {
        relative.push_str("../");
    }
    relative.push_str(&target[common..].join("/"));
    relative
}

/// Writes the source package with the planned relationship parts replaced.
/// Every other entry is copied as its original compressed bytes, so media and
/// worksheet parts are not inflated and recompressed for a fixup that touches
/// relations only.
fn write_package(
    archive: &mut ZipArchive<File>,
    rewrites: &HashMap<usize, String>,
) -> Result<Vec<u8>, SidecarError> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for index in 0..archive.len() {
        match rewrites.get(&index) {
            Some(xml) => {
                let name = archive.by_index_raw(index)?.name().to_string();
                writer.start_file(name, options)?;
                writer.write_all(xml.as_bytes())?;
            }
            None => writer.raw_copy_file(archive.by_index(index)?)?,
        }
    }
    let cursor = writer.finish()?;
    Ok(cursor.into_inner())
}

#[cfg(test)]
mod tests {
    use std::io::Read;

    use super::*;

    #[test]
    fn relationship_parts_map_to_their_source_part() {
        assert_eq!(
            relationship_source_part("xl/worksheets/_rels/sheet1.xml.rels").as_deref(),
            Some("xl/worksheets/sheet1.xml")
        );
        assert_eq!(
            relationship_source_part("/xl/_rels/workbook.xml.rels").as_deref(),
            Some("xl/workbook.xml")
        );
        assert_eq!(relationship_source_part("_rels/.rels").as_deref(), Some(""));
        assert_eq!(relationship_source_part("xl/worksheets/sheet1.xml"), None);
        assert_eq!(relationship_source_part("xl/media/image1.png"), None);
    }

    #[test]
    fn relative_targets_climb_to_the_common_ancestor() {
        assert_eq!(
            relative_target("xl/worksheets/sheet1.xml", "xl/comments/comment1.xml"),
            "../comments/comment1.xml"
        );
        assert_eq!(
            relative_target(
                "xl/worksheets/sheet1.xml",
                "xl/drawings/commentsDrawing1.vml"
            ),
            "../drawings/commentsDrawing1.vml"
        );
        // excel's legacy comments part sits beside the workbook part
        assert_eq!(
            relative_target("xl/worksheets/sheet1.xml", "xl/comments1.xml"),
            "../comments1.xml"
        );
        // workbook rels resolve from xl/, not from a top-level part
        assert_eq!(
            relative_target("xl/workbook.xml", "xl/worksheets/sheet1.xml"),
            "worksheets/sheet1.xml"
        );
        // package-root rels resolve from the root
        assert_eq!(relative_target("", "xl/workbook.xml"), "xl/workbook.xml");
        assert_eq!(
            relative_target("xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"),
            "sheet1.xml"
        );
    }

    fn relative_only_package() -> (tempfile::TempDir, std::path::PathBuf) {
        crate::tests::open_fixture(&[
            (
                "_rels/.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/worksheets/sheet1.xml", "<worksheet/>"),
        ])
    }

    fn absolute_target_package(include_comments: bool) -> (tempfile::TempDir, std::path::PathBuf) {
        let mut entries = vec![
            (
                "_rels/.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            ("xl/workbook.xml", "<workbook/>"),
            ("xl/worksheets/sheet1.xml", "<worksheet/>"),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="/xl/comments/comment1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="/xl/drawings/commentsDrawing1.vml"/></Relationships>"#,
            ),
            ("xl/drawings/commentsDrawing1.vml", "<xml/>"),
        ];
        if include_comments {
            entries.push(("xl/comments/comment1.xml", "<comments/>"));
        }
        crate::tests::open_fixture(&entries)
    }

    fn entry_text(archive: &mut ZipArchive<Cursor<Vec<u8>>>, name: &str) -> String {
        let mut entry = archive.by_name(name).unwrap();
        let mut text = String::new();
        entry.read_to_string(&mut text).unwrap();
        text
    }

    #[test]
    fn relative_only_packages_are_not_rewritten() {
        let (_dir, path) = relative_only_package();
        assert!(normalized_package(&path).unwrap().is_none());
    }

    #[test]
    fn absolute_targets_are_rewritten_in_the_normalized_package() {
        let (_dir, path) = absolute_target_package(true);
        let bytes = normalized_package(&path).unwrap().unwrap();
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let rels = entry_text(&mut archive, "xl/worksheets/_rels/sheet1.xml.rels");
        assert!(
            rels.contains(r#"Target="../comments/comment1.xml""#),
            "{rels}"
        );
        assert!(
            rels.contains(r#"Target="../drawings/commentsDrawing1.vml""#),
            "{rels}"
        );
        // relative targets are not touched
        let workbook_rels = entry_text(&mut archive, "xl/_rels/workbook.xml.rels");
        assert!(
            workbook_rels.contains(r#"Target="worksheets/sheet1.xml""#),
            "{workbook_rels}"
        );
        // entries without a rewrite are copied verbatim
        assert_eq!(
            entry_text(&mut archive, "xl/worksheets/sheet1.xml"),
            "<worksheet/>"
        );
    }

    #[test]
    fn a_missing_absolute_target_names_the_part() {
        let (_dir, path) = absolute_target_package(false);
        let error = normalized_package(&path).unwrap_err();
        let message = error.to_string();
        assert!(
            message.contains("workbook relationship target not found:"),
            "{message}"
        );
        assert!(
            message.contains("xl/worksheets/sheet1.xml -> xl/comments/comment1.xml"),
            "{message}"
        );
        assert!(
            message.contains("(resolved from /xl/comments/comment1.xml)"),
            "{message}"
        );
    }
}
