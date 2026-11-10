/// Minimal workbook stylesheet for workbooks the gateway creates from scratch.
///
/// A workbook without `xl/styles.xml` is tolerated by Excel, which falls back to
/// its own defaults — but not by the Rust sidecar's formula engine (ironcalc),
/// which fails the whole import with `specified file not found in archive`. A
/// blank workbook plus a formula write therefore produced a file the app could
/// read and the CLI could not: every cell-level result was unreachable from
/// `genoffice sheet read` and from any headless recalc.
///
/// The shape mirrors what `StylesheetEditor` expects to exist so a later style
/// edit extends this table instead of replacing it.
export const MINIMAL_STYLESHEET_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border/></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>'
