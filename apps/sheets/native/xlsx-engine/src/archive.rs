//! Session-less archive commands backing the streaming save channel: the
//! TypeScript gateway generates patched entry contents, this module reads
//! individual entries and reassembles the package. Untouched entries are
//! copied as raw compressed bytes, so preservation holds by construction.

use std::collections::HashSet;
use std::fs::{self, File};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::SidecarError;

const MAX_ENTRY_COUNT: usize = 10_000;
/// Cap on a single decompressed entry handed to the patching layer. The
/// archive itself has no size limit — only entries the gateway edits must
/// fit in memory.
// Densely styled worksheets can exceed 256 MiB as XML while remaining
// ordinary workbooks on disk (the 88k-row suppliers fixture is ~307 MiB).
// Keep a finite anti-bomb / memory bound, but allow that real-world case.
// 500 MiB, not 512: the host patches entries as JavaScript strings, and
// V8's maximum string length is 536,870,888 bytes — an entry between that
// and 512 MiB would pass a 512 MiB cap and then fail to materialize.
const MAX_EXTRACTED_ENTRY_BYTES: u64 = 500 * 1024 * 1024;

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ArchiveEntry {
    pub name: String,
    pub crc32: u32,
    pub compressed_size: u64,
    pub uncompressed_size: u64,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExtractedEntry {
    pub name: String,
    pub path: PathBuf,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EntryContent {
    pub name: String,
    pub content_path: PathBuf,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SaveArchiveResult {
    pub before_entries: Vec<ArchiveEntry>,
    pub after_entries: Vec<ArchiveEntry>,
}

pub fn archive_manifest(path: &Path) -> Result<Vec<ArchiveEntry>, SidecarError> {
    let mut archive = open_validated(path)?;
    manifest_of(&mut archive)
}

/// Streams each entry through the XML parser and reports which ones contain
/// `needle` in their decoded text nodes. Constant memory: entity references
/// are decoded (openpyxl writes CJK as `&#NNNN;`), consecutive text events
/// accumulate until the enclosing element closes, and oversized runs keep a
/// sliding tail. Lets the gateway skip rewriting entries too large to patch
/// when they cannot reference the edited sheet.
pub fn scan_entries_for_text(
    path: &Path,
    names: &[String],
    needle: &str,
) -> Result<Vec<String>, SidecarError> {
    if needle.is_empty() {
        return Err(SidecarError::InvalidRequest("Scan needle must not be empty.".into()));
    }
    let mut archive = open_validated(path)?;
    let mut matches = Vec::new();
    for name in names {
        let entry = archive.by_name(name).map_err(|_| {
            SidecarError::Workbook(format!("Workbook is missing {name}."))
        })?;
        if entry_text_contains(entry, needle)? {
            matches.push(name.clone());
        }
    }
    Ok(matches)
}

fn entry_text_contains(
    entry: impl std::io::Read,
    needle: &str,
) -> Result<bool, SidecarError> {
    use quick_xml::events::Event;
    // Text runs are searched when the element closes; the tail keeps needle
    // matches alive across the accumulation cap.
    const RUN_CAP: usize = 1 << 20;
    let mut reader = quick_xml::Reader::from_reader(std::io::BufReader::new(entry));
    let mut buf = Vec::new();
    let mut run = String::new();
    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|error| SidecarError::Workbook(error.to_string()))?;
        match event {
            Event::Text(text) => {
                let decoded = text
                    .decode()
                    .map_err(|error| SidecarError::Workbook(error.to_string()))?;
                run.push_str(&decoded);
            }
            Event::GeneralRef(reference) => {
                run.push_str(&crate::general_ref_text(&reference)?);
            }
            Event::Eof => return Ok(run.contains(needle)),
            _ => {
                if run.contains(needle) {
                    return Ok(true);
                }
                run.clear();
            }
        }
        if run.len() > RUN_CAP {
            if run.contains(needle) {
                return Ok(true);
            }
            let keep = needle.len().saturating_mul(4).min(run.len());
            let mut cut = run.len() - keep;
            while !run.is_char_boundary(cut) {
                cut += 1;
            }
            run.drain(..cut);
        }
        buf.clear();
    }
}

pub fn read_entries_to_dir(
    path: &Path,
    names: &[String],
    output_dir: &Path,
) -> Result<Vec<ExtractedEntry>, SidecarError> {
    if !output_dir.is_dir() {
        return Err(SidecarError::InvalidRequest(
            "Entry output directory does not exist.".into(),
        ));
    }
    let mut archive = open_validated(path)?;
    let mut extracted = Vec::with_capacity(names.len());
    for (index, name) in names.iter().enumerate() {
        let mut entry = archive.by_name(name).map_err(|_| {
            SidecarError::Workbook(format!("Workbook is missing {name}."))
        })?;
        if entry.size() > MAX_EXTRACTED_ENTRY_BYTES {
            return Err(SidecarError::Workbook(format!(
                "Entry {name} is {} bytes uncompressed, above the {MAX_EXTRACTED_ENTRY_BYTES} byte patch limit.",
                entry.size()
            )));
        }
        let output_path = output_dir.join(format!("entry-{index}.bin"));
        let mut output = BufWriter::new(File::create(&output_path)?);
        std::io::copy(&mut entry, &mut output)?;
        output.flush()?;
        extracted.push(ExtractedEntry {
            name: name.clone(),
            path: output_path,
        });
    }
    Ok(extracted)
}

pub fn save_archive(
    source_path: &Path,
    target_path: &Path,
    replacements: &[EntryContent],
    removals: &[String],
    additions: &[EntryContent],
) -> Result<SaveArchiveResult, SidecarError> {
    if source_path.canonicalize()? == target_path.canonicalize().unwrap_or_default() {
        return Err(SidecarError::InvalidRequest(
            "Save target must differ from the source archive.".into(),
        ));
    }
    let mut archive = open_validated(source_path)?;
    let before_entries = manifest_of(&mut archive)?;
    let source_names: HashSet<&str> =
        before_entries.iter().map(|entry| entry.name.as_str()).collect();
    validate_edit_sets(&source_names, replacements, removals, additions)?;

    let replacement_by_name: std::collections::HashMap<&str, &EntryContent> = replacements
        .iter()
        .map(|item| (item.name.as_str(), item))
        .collect();
    let removal_set: HashSet<&str> = removals.iter().map(String::as_str).collect();

    let target_file = File::create(target_path)?;
    let mut writer = ZipWriter::new(target_file);
    for index in 0..archive.len() {
        let entry = archive.by_index_raw(index)?;
        let name = entry.name().to_owned();
        if removal_set.contains(name.as_str()) {
            continue;
        }
        match replacement_by_name.get(name.as_str()) {
            Some(replacement) => {
                drop(entry);
                write_entry(&mut writer, &name, &replacement.content_path)?;
            }
            None => writer.raw_copy_file(entry)?,
        }
    }
    for addition in additions {
        write_entry(&mut writer, &addition.name, &addition.content_path)?;
    }
    let target_file = writer.finish()?;
    target_file.sync_all()?;

    let mut saved = ZipArchive::new(File::open(target_path)?)?;
    Ok(SaveArchiveResult {
        before_entries,
        after_entries: manifest_of(&mut saved)?,
    })
}

fn write_entry(
    writer: &mut ZipWriter<File>,
    name: &str,
    content_path: &Path,
) -> Result<(), SidecarError> {
    let content = fs::read(content_path)?;
    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .large_file(content.len() as u64 >= u64::from(u32::MAX));
    writer.start_file(name, options)?;
    writer.write_all(&content)?;
    Ok(())
}

fn validate_edit_sets(
    source_names: &HashSet<&str>,
    replacements: &[EntryContent],
    removals: &[String],
    additions: &[EntryContent],
) -> Result<(), SidecarError> {
    let mut seen: HashSet<&str> = HashSet::new();
    let edited = replacements
        .iter()
        .map(|item| item.name.as_str())
        .chain(removals.iter().map(String::as_str))
        .chain(additions.iter().map(|item| item.name.as_str()));
    for name in edited {
        if !seen.insert(name) {
            return Err(SidecarError::InvalidRequest(format!(
                "Entry {name} appears in more than one edit set."
            )));
        }
        if !is_safe_entry_name(name) {
            return Err(SidecarError::InvalidRequest(format!(
                "Entry {name} has an unsafe archive path."
            )));
        }
    }
    for replacement in replacements {
        if !source_names.contains(replacement.name.as_str()) {
            return Err(SidecarError::InvalidRequest(format!(
                "Cannot replace missing entry {}.",
                replacement.name
            )));
        }
    }
    for removal in removals {
        if !source_names.contains(removal.as_str()) {
            return Err(SidecarError::InvalidRequest(format!(
                "Cannot remove missing entry {removal}."
            )));
        }
    }
    for addition in additions {
        if source_names.contains(addition.name.as_str()) {
            return Err(SidecarError::InvalidRequest(format!(
                "Cannot add entry {} — it already exists.",
                addition.name
            )));
        }
    }
    if source_names.len() + additions.len() > MAX_ENTRY_COUNT {
        return Err(SidecarError::InvalidRequest(
            "Saving would exceed the archive entry limit.".into(),
        ));
    }
    Ok(())
}

fn is_safe_entry_name(name: &str) -> bool {
    !name.is_empty()
        && !name.starts_with('/')
        && !name.ends_with('/')
        && !name.split('/').any(|segment| segment.is_empty() || segment == "..")
}

fn open_validated(path: &Path) -> Result<ZipArchive<File>, SidecarError> {
    let mut archive = ZipArchive::new(File::open(path)?)?;
    if archive.len() > MAX_ENTRY_COUNT {
        return Err(SidecarError::Workbook(
            "Workbook contains too many ZIP entries.".into(),
        ));
    }
    for index in 0..archive.len() {
        if archive.by_index_raw(index)?.enclosed_name().is_none() {
            return Err(SidecarError::Workbook(
                "Workbook contains an unsafe ZIP path.".into(),
            ));
        }
    }
    Ok(archive)
}

fn manifest_of(archive: &mut ZipArchive<File>) -> Result<Vec<ArchiveEntry>, SidecarError> {
    let mut entries = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let entry = archive.by_index_raw(index)?;
        if entry.is_dir() {
            continue;
        }
        entries.push(ArchiveEntry {
            name: entry.name().to_owned(),
            crc32: entry.crc32(),
            compressed_size: entry.compressed_size(),
            uncompressed_size: entry.size(),
        });
    }
    entries.sort_by(|left, right| left.name.cmp(&right.name));
    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;

    fn build_fixture(dir: &Path) -> PathBuf {
        let path = dir.join("fixture.zip");
        let mut writer = ZipWriter::new(File::create(&path).unwrap());
        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, content) in [
            ("keep/a.xml", "<a>untouched</a>"),
            ("replace/b.xml", "<b>old</b>"),
            ("remove/c.xml", "<c>drop me</c>"),
        ] {
            writer.start_file(name, stored).unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
        path
    }

    fn write_content(dir: &Path, name: &str, content: &str) -> EntryContent {
        let content_path = dir.join(name);
        fs::write(&content_path, content).unwrap();
        EntryContent {
            name: String::new(),
            content_path,
        }
    }

    #[test]
    fn saves_with_raw_copy_replacement_removal_and_addition() {
        let dir = tempdir();
        let source = build_fixture(&dir);
        let target = dir.join("saved.zip");
        let mut replacement = write_content(&dir, "b.new.xml", "<b>new</b>");
        replacement.name = "replace/b.xml".into();
        let mut addition = write_content(&dir, "d.new.xml", "<d>added</d>");
        addition.name = "added/d.xml".into();

        let result = save_archive(
            &source,
            &target,
            &[replacement],
            &["remove/c.xml".into()],
            &[addition],
        )
        .unwrap();

        let before: Vec<&str> = result.before_entries.iter().map(|e| e.name.as_str()).collect();
        assert_eq!(before, ["keep/a.xml", "remove/c.xml", "replace/b.xml"]);
        let after: Vec<&str> = result.after_entries.iter().map(|e| e.name.as_str()).collect();
        assert_eq!(after, ["added/d.xml", "keep/a.xml", "replace/b.xml"]);

        let untouched_before = result
            .before_entries
            .iter()
            .find(|e| e.name == "keep/a.xml")
            .unwrap();
        let untouched_after = result
            .after_entries
            .iter()
            .find(|e| e.name == "keep/a.xml")
            .unwrap();
        assert_eq!(untouched_before, untouched_after);

        let mut saved = ZipArchive::new(File::open(&target).unwrap()).unwrap();
        assert_eq!(read_entry(&mut saved, "replace/b.xml"), "<b>new</b>");
        assert_eq!(read_entry(&mut saved, "added/d.xml"), "<d>added</d>");
        assert_eq!(read_entry(&mut saved, "keep/a.xml"), "<a>untouched</a>");
        assert!(saved.by_name("remove/c.xml").is_err());
    }

    #[test]
    fn rejects_edits_on_missing_or_conflicting_entries() {
        let dir = tempdir();
        let source = build_fixture(&dir);
        let target = dir.join("saved.zip");
        let mut missing = write_content(&dir, "x.xml", "<x/>");
        missing.name = "missing/x.xml".into();
        let error = save_archive(&source, &target, &[missing], &[], &[]).unwrap_err();
        assert!(error.to_string().contains("missing entry"));

        let mut duplicate = write_content(&dir, "dup.xml", "<b/>");
        duplicate.name = "replace/b.xml".into();
        let error =
            save_archive(&source, &target, &[duplicate], &["replace/b.xml".into()], &[])
                .unwrap_err();
        assert!(error.to_string().contains("more than one edit set"));

        let mut existing = write_content(&dir, "a.xml", "<a/>");
        existing.name = "keep/a.xml".into();
        let error = save_archive(&source, &target, &[], &[], &[existing]).unwrap_err();
        assert!(error.to_string().contains("already exists"));

        let mut unsafe_name = write_content(&dir, "evil.xml", "<e/>");
        unsafe_name.name = "../evil.xml".into();
        let error = save_archive(&source, &target, &[], &[], &[unsafe_name]).unwrap_err();
        assert!(error.to_string().contains("unsafe archive path"));
    }

    #[test]
    fn scans_decoded_text_across_entity_references() {
        let dir = tempdir();
        let path = dir.join("scan.zip");
        let mut writer = ZipWriter::new(File::create(&path).unwrap());
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        // openpyxl-style numeric entities: CJK chars encoded as &#37319;&#36141; etc.
        writer.start_file("refs.xml", options).unwrap();
        writer
            .write_all("<sheet><f>SUM('&#37319;&#36141;'!A1:A5)</f></sheet>".as_bytes())
            .unwrap();
        writer.start_file("plain.xml", options).unwrap();
        writer.write_all(b"<sheet><v>12</v></sheet>").unwrap();
        writer.finish().unwrap();

        let matches = scan_entries_for_text(
            &path,
            &["refs.xml".into(), "plain.xml".into()],
            "\u{91c7}\u{8d2d}",
        )
        .unwrap();
        assert_eq!(matches, ["refs.xml"]);

        let ascii = scan_entries_for_text(&path, &["plain.xml".into()], "12").unwrap();
        assert_eq!(ascii, ["plain.xml"]);
    }

    #[test]
    fn extracts_requested_entries_and_reports_manifest() {
        let dir = tempdir();
        let source = build_fixture(&dir);
        let manifest = archive_manifest(&source).unwrap();
        assert_eq!(manifest.len(), 3);
        assert!(manifest.iter().all(|entry| entry.uncompressed_size > 0));

        let output_dir = dir.join("out");
        fs::create_dir(&output_dir).unwrap();
        let extracted =
            read_entries_to_dir(&source, &["replace/b.xml".into()], &output_dir).unwrap();
        assert_eq!(extracted.len(), 1);
        assert_eq!(fs::read_to_string(&extracted[0].path).unwrap(), "<b>old</b>");

        let error = read_entries_to_dir(&source, &["missing.xml".into()], &output_dir).unwrap_err();
        assert!(error.to_string().contains("missing"));
    }

    fn read_entry(archive: &mut ZipArchive<File>, name: &str) -> String {
        let mut entry = archive.by_name(name).unwrap();
        let mut content = String::new();
        entry.read_to_string(&mut content).unwrap();
        content
    }

    fn tempdir() -> PathBuf {
        let dir = std::env::temp_dir().join(format!(
            "xlsx-archive-test-{}-{}",
            std::process::id(),
            COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
        ));
        fs::create_dir_all(&dir).unwrap();
        dir
    }

    static COUNTER: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(0);
}
