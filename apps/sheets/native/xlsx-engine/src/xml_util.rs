//! Zip and quick-xml helpers shared by every part reader: tolerant entry
//! lookup, attribute access and text decoding.

use std::borrow::Cow;

use super::*;

pub(crate) fn decode_uri_path(value: &str) -> Result<String, SidecarError> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            decoded.push(bytes[index]);
            index += 1;
            continue;
        }
        if index + 2 >= bytes.len() {
            return Err(SidecarError::Workbook(
                "OOXML relationship target is malformed.".into(),
            ));
        }
        let high = hex_value(bytes[index + 1]).ok_or_else(|| {
            SidecarError::Workbook("OOXML relationship target is malformed.".into())
        })?;
        let low = hex_value(bytes[index + 2]).ok_or_else(|| {
            SidecarError::Workbook("OOXML relationship target is malformed.".into())
        })?;
        decoded.push(high << 4 | low);
        index += 3;
    }
    String::from_utf8(decoded)
        .map_err(|_| SidecarError::Workbook("OOXML relationship target is not valid UTF-8.".into()))
}

fn hex_value(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

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
            crate::archive::canonical_entry_name(value).map(|name| name.to_ascii_lowercase())
        };
        let decoded = decode_uri_path(name).ok();
        let mut resolved = None;
        for candidate in [Some(name), decoded.as_deref()].into_iter().flatten() {
            let wanted = normalize(candidate);
            if let Some(found) = archive
                .file_names()
                .find(|candidate| normalize(candidate) == wanted)
            {
                resolved = Some(found.to_owned());
                break;
            }
        }
        resolved
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

/// `ZipFile::size()` is only the declared length and zip bounds just the
/// compressed input, so an entry may inflate past its claim; reading one byte
/// past `min(declared, cap)` catches that (same rule as `@genoffice/zip-gate`,
/// genoffice#781 / genoffice#1386).
pub(crate) fn copy_entry_bounded<R: Read, W: Write>(
    entry: &mut R,
    declared: u64,
    cap: u64,
    out: &mut W,
) -> Result<u64, SidecarError> {
    let allowance = declared.min(cap);
    let copied = std::io::copy(&mut entry.take(allowance.saturating_add(1)), out)?;
    if copied > allowance {
        return Err(SidecarError::Workbook(
            "ZIP entry inflates past the size it declares.".into(),
        ));
    }
    Ok(copied)
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

/// ECMA-376 §22.4.2.4: characters illegal in XML are stored as `_xHHHH_`
/// (`_x000D_` = CR); a literal `_x` is itself escaped as `_x005F_x…`, which
/// a single left-to-right pass resolves naturally.
pub(crate) fn decode_xlsx_escapes(text: &str) -> Cow<'_, str> {
    if !text.contains("_x") {
        return Cow::Borrowed(text);
    }
    let bytes = text.as_bytes();
    let mut out = String::with_capacity(text.len());
    let mut last = 0;
    let mut index = 0;
    while index + 7 <= bytes.len() {
        let escaped = bytes[index] == b'_'
            && bytes[index + 1] == b'x'
            && bytes[index + 6] == b'_'
            && bytes[index + 2..index + 6]
                .iter()
                .all(u8::is_ascii_hexdigit);
        let decoded = escaped
            .then(|| u32::from_str_radix(&text[index + 2..index + 6], 16).ok())
            .flatten()
            .and_then(char::from_u32);
        match decoded {
            Some(character) => {
                out.push_str(&text[last..index]);
                out.push(character);
                index += 7;
                last = index;
            }
            None => index += 1,
        }
    }
    if last == 0 {
        return Cow::Borrowed(text);
    }
    out.push_str(&text[last..]);
    Cow::Owned(out)
}

/// Cell text as Excel sees it. Raw CRLF is XML-level noise and must collapse
/// before the escaped CR appears, otherwise `_x000D_\r\n` (Excel's encoding
/// of CR LF) would turn into two line breaks.
pub(crate) fn normalize_cell_text(text: &mut String) {
    normalize_line_endings(text);
    if let Cow::Owned(decoded) = decode_xlsx_escapes(text) {
        *text = decoded;
        normalize_line_endings(text);
    }
}

/// Text of one `<t>` node as Excel reads it: without `xml:space="preserve"`
/// leading/trailing XML whitespace is dropped, so `<t> </t>` is an empty
/// string (a CF rule comparing it with a blank cell matches).
pub(crate) fn text_node_content(text: String, preserve: bool) -> String {
    if preserve {
        return text;
    }
    let trimmed = text.trim_matches([' ', '\t', '\r', '\n']);
    if trimmed.len() == text.len() {
        text
    } else {
        trimmed.to_owned()
    }
}

pub(crate) fn preserves_space<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &BytesStart<'_>,
) -> Result<bool, SidecarError> {
    Ok(attribute_value(reader, element, b"space")?.as_deref() == Some("preserve"))
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
