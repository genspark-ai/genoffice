/// Symbol-encoded fonts (Wingdings & co.) map Latin letters to pictographs,
/// so previewing their name in the font itself renders as gibberish. Font
/// pickers show these names in the UI font instead — same as modern Word.
/// Name-based: the Local Font Access API exposes no charset/PANOSE metadata.
const SYMBOL_FONT_RE =
  /^(webdings|wingdings( \d)?|(itc )?zapf dingbats|symbol|marlett|ms outlook|ms reference specialty|mt extra|bookshelf symbol \d+|segoe mdl2 assets|segoe fluent icons|hololens mdl2 assets|bodoni ornaments)$/

export function isSymbolFontFamily(family: string): boolean {
  return SYMBOL_FONT_RE.test(family.trim().toLowerCase())
}
