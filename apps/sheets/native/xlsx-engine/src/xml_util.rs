//! Zip and quick-xml helpers shared by every part reader: tolerant entry
//! lookup, attribute access and text decoding.

use super::*;

/// Entry lookup tolerant of non-conformant producers: '\' separators,
/// leading '/', and case drift in entry names. Excel opens such packages
/// (tdf131575: .NET-written `xl\workbook.xml` with `sharedstrings.xml`).
pub(crate) fn zip_entry<'a>(
    archive: &'a mut ZipArchive<File>,
    name: &str,
) -> zip::result::ZipResult<zip::read::ZipFile<'a, File>> {
    let resolved = if archive.index_for_name(name).is_some() {
        None
    } else {
        let normalize = |value: &str| {
            value
                .trim_start_matches(['/', '\\'])
                .replace('\\', "/")
                .to_ascii_lowercase()
        };
        let wanted = normalize(name);
        archive
            .file_names()
            .find(|candidate| normalize(candidate) == wanted)
            .map(ToOwned::to_owned)
    };
    archive.by_name(resolved.as_deref().unwrap_or(name))
}

pub(crate) fn read_zip_string(
    archive: &mut ZipArchive<File>,
    path: &str,
) -> Result<String, SidecarError> {
    let mut entry = zip_entry(archive, path)?;
    let mut value = String::new();
    entry.read_to_string(&mut value)?;
    Ok(value)
}

pub(crate) fn attribute_value<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, SidecarError> {
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| SidecarError::Workbook(error.to_string()))?;
        if attribute.key.local_name().as_ref() == name {
            return Ok(Some(
                attribute
                    .decode_and_unescape_value(reader.decoder())?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

/// quick-xml 0.38 emits entity references (`&#26679;`, `&amp;`) as separate
/// GeneralRef events, not as part of Text. Exporters like openpyxl encode all
/// non-ASCII text as numeric character refs, so dropping these loses CJK text.
pub(crate) fn general_ref_text(
    reference: &quick_xml::events::BytesRef<'_>,
) -> Result<String, SidecarError> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| SidecarError::Workbook(error.to_string()))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| SidecarError::Workbook(error.to_string()))?;
    Ok(match name.as_ref() {
        "amp" => "&".into(),
        "lt" => "<".into(),
        "gt" => ">".into(),
        "apos" => "'".into(),
        "quot" => "\"".into(),
        other => format!("&{other};"),
    })
}

/// XML 1.0 §2.11 line-ending normalization that quick-xml leaves to the
/// caller: CRLF pairs (and stray CRs) in cell text become LF. A CR surviving
/// to the renderer doubles every line break in the document model.
pub(crate) fn normalize_line_endings(text: &mut String) {
    if text.contains('\r') {
        *text = text.replace("\r\n", "\n").replace('\r', "\n");
    }
}

pub(crate) fn decode_text(text: &quick_xml::events::BytesText<'_>) -> Result<String, SidecarError> {
    let decoded = text
        .decode()
        .map_err(|error| SidecarError::Workbook(error.to_string()))?;
    quick_xml::escape::unescape(&decoded)
        .map(|value| value.into_owned())
        .map_err(|error| SidecarError::Workbook(error.to_string()))
}

/// CDATA content is literal — decoded per the document encoding, never
/// entity-unescaped.
pub(crate) fn decode_cdata(
    text: &quick_xml::events::BytesCData<'_>,
) -> Result<String, SidecarError> {
    text.decode()
        .map(|value| value.into_owned())
        .map_err(|error| SidecarError::Workbook(error.to_string()))
}
