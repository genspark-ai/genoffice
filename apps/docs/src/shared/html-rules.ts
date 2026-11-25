export const HTML_RULES = [
  'The html tool input is a restricted HTML fragment. Rules:',
  '- Writable tags: h1 h2 h3 h4 h5 h6 p ul ol li strong em u s a br table thead tbody tr th td pre code blockquote formula; img is accepted only with an exact supported current media reference as described below. The special <span data-tab></span> writes a tab.',
  '- An explicit <p></p> is an empty paragraph and is preserved. <p><br></p> instead contains a real line break; empty wrapper <div> elements do not create paragraphs.',
  '- Native tables support horizontal/vertical merges using colspan/rowspan and explicit insert/delete row or column, merge/split, cell-format and cell-text operations. Table HTML writes support ordinary cell text, line breaks and referenced inline pictures; they do not reconstruct inline run formatting, formulas, nested tables or opaque OOXML. Use setTableCellText only when the current cell query advertises it; cells containing protected payloads require a narrower text edit or an explicit structure operation with checked deletion intent.',
  '- Use <pre> for code samples (monospace font + shading, line breaks preserved); use <blockquote> for quotations (indent + left bar)',
  '- Use <formula>LaTeX</formula> for math (produces native Word equations): as a top-level block it becomes its own centered paragraph; placed inside <p>/<li>/<h*> text it is an inline formula flowing with the text, e.g. <p>From <formula>E = mc^2</formula> we know…</p>; supports the common subset of \\frac \\sqrt super/subscripts \\sum \\int \\lim matrix environments Greek letters etc., the align environment is not supported; invalid LaTeX fails the whole call — fix and retry',
  '- Do not include <html>/<body>, markdown code fences, or explanatory text',
  "- <sel>…</sel> appears only in the context to mark the user's selection; it is not a tag you may write",
  '- Elements with data-doc-readonly are read-only context, not replacement content or object references. Copying them into an HTML write is refused without changes. Edit surrounding text in place or exclude those blocks; use explicit structure operations when deletion is intended.',
  '- Preserve a picture by copying its exact current <img> tag from read_blocks. The public runtime emits src="media:D:R:B.I", bound to document D and revision R; stale, malformed, missing, unsupported or unbound references reject the entire write before modification. Legacy doc:B.I is an internal coordinate projection, not a safe external write identity. New pictures go through insert_picture; other <img> sources are rejected rather than omitted.',
  '- Table cells carry colspan/rowspan for merges and zero-based data-row/data-col grid positions. Covered grid positions do not contain extra cells. Preserve spans when rewriting; grid positions are read-only descriptions.',
  '- Query/read HTML is a restricted projection, not a lossless write-back representation. Cell projection metadata and structure inventories describe protected content that HTML may omit. Item/page completeness describes truncation or enumeration, not round-trip fidelity. Technical protection is separate from host access permission; explicit checked structure deletion remains available to authorized callers.',
  '- <span style="color:#RRGGBB;font-size:Npt"> appears only in read_blocks output, on runs whose color or size differs from the rest of the block (color:auto / font-size:default = back to the block\'s normal style); it is read-only — writing <span> has no effect, change colors and sizes with the apply_ops formatting commands',
  '- Organize long content into sections with h2/h3 (unless the user only wants a single paragraph)',
  '- Keep the same language as the original document / user instruction, unless translation is requested',
].join('\n')

export const INLINE_HTML_RULES =
  'Use an inline restricted HTML fragment with text, strong, em, u, s, a, br, formula or <span data-tab></span>. ' +
  'Use the exact current media reference for any preserved img. Use insert_picture for new pictures. ' +
  'Exclude block tags, markdown fences, selection markers and data-doc-readonly content.'
