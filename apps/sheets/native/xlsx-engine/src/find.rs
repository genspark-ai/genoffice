//! `find_cells`: walks a session's chunk index for cells whose value or
//! formula text matches an Excel-style query (optional `*` `?` `~`
//! wildcards), returning coordinates in file space. Journal edits made in
//! the host are not visible here; the renderer overlays them.
//!
//! Paged, always in row-major order (the host sorts for by-column searches):
//! a response that ran out of `limit`, time budget, or indexed rows carries
//! `next_cursor` (inclusive resume position) and `complete: false`.

use super::*;
use std::time::Instant;

/// Hard cap on one page; the renderer pages past it.
pub const MAX_FIND_MATCHES: usize = 100_000;
/// One page gives up well under the host's 30 s request timeout and lets
/// the caller resume from `next_cursor`.
const FIND_TIME_BUDGET: Duration = Duration::from_secs(8);

#[derive(Clone, Copy, Debug, Default, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum FindLookIn {
    /// Cached value text only.
    Values,
    /// Excel's "Look in: Formulas": formula text for formula cells, the
    /// constant for everything else.
    #[default]
    Formulas,
    /// Formula cells only, by formula text.
    FormulasOnly,
    /// Value text or formula text.
    Both,
}

#[derive(Clone, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct FindCursor {
    pub sheet_id: String,
    pub row: usize,
    pub column: usize,
}

fn default_limit() -> usize {
    MAX_FIND_MATCHES
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FindCellsRequest {
    pub session_id: String,
    /// None searches every sheet in workbook order.
    #[serde(default)]
    pub sheet_id: Option<String>,
    pub query: String,
    #[serde(default)]
    pub match_case: bool,
    #[serde(default)]
    pub match_entire_cell: bool,
    #[serde(default)]
    pub look_in: FindLookIn,
    #[serde(default)]
    pub wildcards: bool,
    #[serde(default)]
    pub resume_at: Option<FindCursor>,
    #[serde(default = "default_limit")]
    pub limit: usize,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FindCellsMatch {
    pub sheet_id: String,
    pub row: usize,
    pub column: usize,
    pub value: Option<CellValue>,
    pub value_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula_text: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FindCellsResult {
    pub matches: Vec<FindCellsMatch>,
    /// Every requested sheet was scanned to its end.
    pub complete: bool,
    /// Where the next page starts (inclusive); present iff `complete` is false.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub next_cursor: Option<FindCursor>,
    /// False when the page stopped because worksheet indexing had not
    /// reached the next rows yet — the caller should retry after a pause.
    pub indexing_complete: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Token {
    Any,
    One,
    Literal(char),
}

/// Excel find semantics: substring or whole cell, case-folded unless
/// `match_case`, optional wildcards where `~` escapes the next character.
pub struct FindMatcher {
    tokens: Vec<Token>,
    needle: String,
    match_case: bool,
    entire_cell: bool,
    wildcards: bool,
}

impl FindMatcher {
    /// None when the query is empty after Excel's trim.
    pub fn new(query: &str, match_case: bool, entire_cell: bool, wildcards: bool) -> Option<Self> {
        let folded = if match_case {
            query.to_owned()
        } else {
            query.to_lowercase()
        };
        let needle = folded.trim().to_owned();
        if needle.is_empty() {
            return None;
        }
        let tokens = if wildcards {
            parse_wildcards(&needle)
        } else {
            Vec::new()
        };
        Some(Self {
            tokens,
            needle,
            match_case,
            entire_cell,
            wildcards,
        })
    }

    pub fn matches_text(&self, text: &str) -> bool {
        let folded;
        let haystack: &str = if self.match_case {
            text
        } else {
            folded = text.to_lowercase();
            &folded
        };
        // Whole-cell compares trim spaces only, keeping line breaks.
        let haystack = if self.entire_cell {
            haystack.trim_matches(' ')
        } else {
            haystack
        };
        if !self.wildcards {
            return if self.entire_cell {
                haystack == self.needle
            } else {
                haystack.contains(&self.needle)
            };
        }
        let chars: Vec<char> = haystack.chars().collect();
        if self.entire_cell {
            return glob_match(&self.tokens, &chars);
        }
        (0..=chars.len()).any(|start| glob_prefix_match(&self.tokens, &chars[start..]))
    }

    pub fn matches_cell(
        &self,
        look_in: FindLookIn,
        value: Option<&str>,
        formula: Option<&str>,
    ) -> bool {
        let value_hit = || value.is_some_and(|text| self.matches_text(text));
        let formula_hit = || formula.is_some_and(|text| self.matches_text(text));
        match look_in {
            FindLookIn::Values => value_hit(),
            FindLookIn::Formulas => {
                if formula.is_some() {
                    formula_hit()
                } else {
                    value_hit()
                }
            }
            FindLookIn::FormulasOnly => formula_hit(),
            FindLookIn::Both => value_hit() || formula_hit(),
        }
    }
}

fn parse_wildcards(needle: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let mut chars = needle.chars();
    while let Some(c) = chars.next() {
        match c {
            '*' => {
                if tokens.last() != Some(&Token::Any) {
                    tokens.push(Token::Any);
                }
            }
            '?' => tokens.push(Token::One),
            '~' => match chars.next() {
                Some(next) => tokens.push(Token::Literal(next)),
                None => tokens.push(Token::Literal('~')),
            },
            other => tokens.push(Token::Literal(other)),
        }
    }
    tokens
}

/// Whole-string glob match with single-star backtracking.
fn glob_match(tokens: &[Token], text: &[char]) -> bool {
    let (mut t, mut s) = (0usize, 0usize);
    let mut star: Option<(usize, usize)> = None;
    while s < text.len() {
        match tokens.get(t) {
            Some(Token::Any) => {
                star = Some((t, s));
                t += 1;
            }
            Some(Token::One) => {
                t += 1;
                s += 1;
            }
            Some(Token::Literal(c)) if *c == text[s] => {
                t += 1;
                s += 1;
            }
            _ => match star {
                Some((star_t, star_s)) => {
                    t = star_t + 1;
                    s = star_s + 1;
                    star = Some((star_t, star_s + 1));
                }
                None => return false,
            },
        }
    }
    tokens[t..].iter().all(|token| *token == Token::Any)
}

/// Does some prefix of `text` match the whole pattern?
fn glob_prefix_match(tokens: &[Token], text: &[char]) -> bool {
    match tokens.split_first() {
        None => true,
        Some((Token::Any, rest)) => {
            (0..=text.len()).any(|skip| glob_prefix_match(rest, &text[skip..]))
        }
        Some((Token::One, rest)) => !text.is_empty() && glob_prefix_match(rest, &text[1..]),
        Some((Token::Literal(c), rest)) => {
            text.first() == Some(c) && glob_prefix_match(rest, &text[1..])
        }
    }
}

/// Stringifies like the renderer (Univer's extractPureValue): JS number
/// formatting, booleans as "1"/"0".
pub fn cell_value_text(value: &CellValue) -> String {
    match value {
        CellValue::String(text) => text.clone(),
        CellValue::Boolean(flag) => (if *flag { "1" } else { "0" }).to_owned(),
        CellValue::Number(number) => js_number_text(*number),
    }
}

fn js_number_text(number: f64) -> String {
    let magnitude = number.abs();
    if magnitude != 0.0 && !(1e-6..1e21).contains(&magnitude) {
        let exp = format!("{number:e}");
        return match exp.split_once('e') {
            Some((mantissa, exponent)) if !exponent.starts_with('-') => {
                format!("{mantissa}e+{exponent}")
            }
            _ => exp,
        };
    }
    format!("{number}")
}

struct ScanOutcome {
    complete: bool,
    next: Option<(usize, usize)>,
    indexing_complete: bool,
}

impl WorkbookSessions {
    pub fn find_cells(
        &mut self,
        request: &FindCellsRequest,
        cancelled: &AtomicBool,
    ) -> Result<FindCellsResult, SidecarError> {
        let session = self
            .sessions
            .get_mut(&request.session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        let matcher = FindMatcher::new(
            &request.query,
            request.match_case,
            request.match_entire_cell,
            request.wildcards,
        )
        .ok_or_else(|| SidecarError::InvalidRequest("Find query is empty.".into()))?;
        let limit = request.limit.clamp(1, MAX_FIND_MATCHES);
        let sheet_position = |id: &str| session.sheets.iter().position(|sheet| sheet.id == id);
        let targets: Vec<usize> = match &request.sheet_id {
            Some(id) => vec![
                sheet_position(id)
                    .ok_or_else(|| SidecarError::InvalidRequest("Unknown worksheet.".into()))?,
            ],
            None => (0..session.sheets.len()).collect(),
        };
        let mut resume = match &request.resume_at {
            Some(cursor) => {
                let index = sheet_position(&cursor.sheet_id).ok_or_else(|| {
                    SidecarError::InvalidRequest("Unknown cursor worksheet.".into())
                })?;
                Some((index, (cursor.row, cursor.column)))
            }
            None => None,
        };
        let started = Instant::now();
        let mut matches = Vec::new();
        for &sheet_index in &targets {
            let start = match resume {
                Some((index, _)) if index > sheet_index => continue,
                Some((index, position)) if index == sheet_index => Some(position),
                _ => None,
            };
            resume = None;
            if cancelled.load(Ordering::Acquire) {
                return Err(SidecarError::cancelled());
            }
            session.ensure_parser(sheet_index)?;
            let outcome = session.scan_sheet(
                sheet_index,
                &matcher,
                request.look_in,
                start,
                limit,
                &mut matches,
                started + FIND_TIME_BUDGET,
                cancelled,
            )?;
            if !outcome.complete {
                let (row, column) = outcome.next.unwrap_or((0, 0));
                return Ok(FindCellsResult {
                    matches,
                    complete: false,
                    next_cursor: Some(FindCursor {
                        sheet_id: session.sheets[sheet_index].id.clone(),
                        row,
                        column,
                    }),
                    indexing_complete: outcome.indexing_complete,
                });
            }
        }
        Ok(FindCellsResult {
            matches,
            complete: true,
            next_cursor: None,
            indexing_complete: true,
        })
    }
}

impl WorkbookSession {
    /// Waits (bounded) for chunk `chunk_index` to be indexed. Ok(None) when
    /// indexing has not reached it yet; Ok(Some(None)) when it holds no data.
    fn chunk_path(
        &self,
        sheet_index: usize,
        chunk_index: usize,
    ) -> Result<Option<Option<PathBuf>>, SidecarError> {
        let runtime = &self.runtimes[sheet_index];
        let (lock, condition) = &*runtime.state;
        let last_row = self.sheets[sheet_index]
            .row_count
            .saturating_sub(1)
            .min((chunk_index + 1) * CHUNK_ROW_COUNT - 1);
        let ready = |index: &SheetIndex| {
            index.complete || index.indexed_through_row.is_some_and(|row| row >= last_row)
        };
        let index = lock
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
        let index = condition
            .wait_timeout_while(index, RANGE_WAIT, |current| {
                current.error.is_none() && !ready(current)
            })
            .map_err(|_| SidecarError::Io("Worksheet index wait failed.".into()))?
            .0;
        if let Some(error) = &index.error {
            return Err(SidecarError::Workbook(error.clone()));
        }
        if !ready(&index) {
            return Ok(None);
        }
        Ok(Some(index.chunk_files.get(&chunk_index).cloned()))
    }

    fn chunk_matches(
        &self,
        sheet_index: usize,
        path: &Path,
        matcher: &FindMatcher,
        look_in: FindLookIn,
    ) -> Result<Vec<FindCellsMatch>, SidecarError> {
        let file = File::open(path)?;
        let chunk: ChunkData = serde_json::from_reader(BufReader::new(file))?;
        let sheet_id = &self.sheets[sheet_index].id;
        Ok(chunk
            .cells
            .into_iter()
            .filter_map(|cell| {
                let value_text = cell.value.as_ref().map(cell_value_text);
                if !matcher.matches_cell(look_in, value_text.as_deref(), cell.formula.as_deref()) {
                    return None;
                }
                Some(FindCellsMatch {
                    sheet_id: sheet_id.clone(),
                    row: cell.row,
                    column: cell.column,
                    value: cell.value,
                    value_text,
                    formula_text: cell.formula,
                })
            })
            .collect())
    }

    #[allow(clippy::too_many_arguments)]
    fn scan_sheet(
        &self,
        sheet_index: usize,
        matcher: &FindMatcher,
        look_in: FindLookIn,
        start: Option<(usize, usize)>,
        limit: usize,
        out: &mut Vec<FindCellsMatch>,
        deadline: Instant,
        cancelled: &AtomicBool,
    ) -> Result<ScanOutcome, SidecarError> {
        let row_count = self.sheets[sheet_index].row_count;
        if row_count == 0 {
            return Ok(ScanOutcome {
                complete: true,
                next: None,
                indexing_complete: true,
            });
        }
        let last_chunk = (row_count - 1) / CHUNK_ROW_COUNT;
        let accepts = |hit: &FindCellsMatch| {
            start.is_none_or(|(row, column)| (hit.row, hit.column) >= (row, column))
        };
        let first_chunk = start.map_or(0, |(row, _)| row / CHUNK_ROW_COUNT);
        for chunk_index in first_chunk..=last_chunk {
            if cancelled.load(Ordering::Acquire) {
                return Err(SidecarError::cancelled());
            }
            let paused = |indexing_complete: bool| ScanOutcome {
                complete: false,
                next: Some((chunk_index * CHUNK_ROW_COUNT, 0)),
                indexing_complete,
            };
            if chunk_index > first_chunk && Instant::now() >= deadline {
                return Ok(paused(true));
            }
            let Some(path) = self.chunk_path(sheet_index, chunk_index)? else {
                return Ok(paused(false));
            };
            let Some(path) = path else { continue };
            let mut hits = self.chunk_matches(sheet_index, &path, matcher, look_in)?;
            hits.retain(|hit| accepts(hit));
            hits.sort_by_key(|hit| (hit.row, hit.column));
            for hit in hits {
                if out.len() >= limit {
                    return Ok(ScanOutcome {
                        complete: false,
                        next: Some((hit.row, hit.column)),
                        indexing_complete: true,
                    });
                }
                out.push(hit);
            }
        }
        Ok(ScanOutcome {
            complete: true,
            next: None,
            indexing_complete: true,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tests::open_fixture;

    const WORKBOOK_XML: &str = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/></sheets></workbook>"#;
    const RELS_XML: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#;
    const SST_XML: &str = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Apple Pie</t></si><si><t>apple</t></si><si><t> Pie </t></si></sst>"#;
    const SHEET1_XML: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C600"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1"><v>42</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><f>SUM(C1:C1)</f><v>42</v></c><c r="C2" t="b"><v>1</v></c></row><row r="600"><c r="A600" t="inlineStr"><is><t>pineapple</t></is></c><c r="C600"><v>420</v></c></row></sheetData></worksheet>"#;
    const SHEET2_XML: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Apple</t></is></c></row></sheetData></worksheet>"#;

    fn open() -> (tempfile::TempDir, WorkbookSessions, String) {
        let (dir, path) = open_fixture(&[
            ("xl/workbook.xml", WORKBOOK_XML),
            ("xl/_rels/workbook.xml.rels", RELS_XML),
            ("xl/sharedStrings.xml", SST_XML),
            ("xl/worksheets/sheet1.xml", SHEET1_XML),
            ("xl/worksheets/sheet2.xml", SHEET2_XML),
        ]);
        let mut sessions = WorkbookSessions::new();
        let metadata = sessions.open(&path).unwrap();
        (dir, sessions, metadata.session_id)
    }

    fn request(session_id: &str, query: &str) -> FindCellsRequest {
        FindCellsRequest {
            session_id: session_id.to_owned(),
            sheet_id: None,
            query: query.to_owned(),
            match_case: false,
            match_entire_cell: false,
            look_in: FindLookIn::Formulas,
            wildcards: false,
            resume_at: None,
            limit: MAX_FIND_MATCHES,
        }
    }

    fn run_to_completion(
        sessions: &mut WorkbookSessions,
        mut request: FindCellsRequest,
    ) -> Vec<FindCellsMatch> {
        let mut all = Vec::new();
        let not_cancelled = AtomicBool::new(false);
        loop {
            let result = sessions.find_cells(&request, &not_cancelled).unwrap();
            all.extend(result.matches);
            if result.complete {
                return all;
            }
            if !result.indexing_complete {
                std::thread::sleep(Duration::from_millis(5));
            }
            request.resume_at = result.next_cursor;
        }
    }

    fn coords(matches: &[FindCellsMatch]) -> Vec<(&str, usize, usize)> {
        matches
            .iter()
            .map(|hit| (hit.sheet_id.as_str(), hit.row, hit.column))
            .collect()
    }

    #[test]
    fn substring_search_is_case_insensitive_across_sheets_and_shared_strings() {
        let (_dir, mut sessions, session_id) = open();
        let hits = run_to_completion(&mut sessions, request(&session_id, "apple"));
        assert_eq!(
            coords(&hits),
            vec![
                ("sheet-1", 0, 0),
                ("sheet-1", 0, 1),
                ("sheet-1", 599, 0),
                ("sheet-2", 0, 0)
            ]
        );
        assert_eq!(hits[0].value_text.as_deref(), Some("Apple Pie"));
    }

    #[test]
    fn match_case_and_entire_cell() {
        let (_dir, mut sessions, session_id) = open();
        let mut cased = request(&session_id, "Apple");
        cased.match_case = true;
        assert_eq!(
            coords(&run_to_completion(&mut sessions, cased)),
            vec![("sheet-1", 0, 0), ("sheet-2", 0, 0)]
        );
        let mut whole = request(&session_id, "pie");
        whole.match_entire_cell = true;
        // " Pie " trims its spaces; "Apple Pie" is not a whole-cell hit.
        assert_eq!(
            coords(&run_to_completion(&mut sessions, whole)),
            vec![("sheet-1", 1, 0)]
        );
    }

    #[test]
    fn wildcards_follow_excel_semantics() {
        let (_dir, mut sessions, session_id) = open();
        let mut glob = request(&session_id, "p?ne*");
        glob.wildcards = true;
        assert_eq!(
            coords(&run_to_completion(&mut sessions, glob)),
            vec![("sheet-1", 599, 0)]
        );
        let mut whole = request(&session_id, "apple*");
        whole.wildcards = true;
        whole.match_entire_cell = true;
        assert_eq!(
            coords(&run_to_completion(&mut sessions, whole)),
            vec![("sheet-1", 0, 0), ("sheet-1", 0, 1), ("sheet-2", 0, 0)]
        );
        let matcher = FindMatcher::new("a~*b", false, true, true).unwrap();
        assert!(matcher.matches_text("a*b"));
        assert!(!matcher.matches_text("axb"));
        let star_tail = FindMatcher::new("*", false, true, true).unwrap();
        assert!(star_tail.matches_text("anything"));
    }

    #[test]
    fn look_in_values_versus_formulas() {
        let (_dir, mut sessions, session_id) = open();
        let mut values = request(&session_id, "42");
        values.look_in = FindLookIn::Values;
        // B2's cached value 42, C1 = 42, C600 = 420 contains "42".
        assert_eq!(
            coords(&run_to_completion(&mut sessions, values)),
            vec![("sheet-1", 0, 2), ("sheet-1", 1, 1), ("sheet-1", 599, 2)]
        );
        let formulas = request(&session_id, "42");
        // Formulas: B2 is matched by its formula text "=SUM(C1:C1)", not its value.
        assert_eq!(
            coords(&run_to_completion(&mut sessions, formulas)),
            vec![("sheet-1", 0, 2), ("sheet-1", 599, 2)]
        );
        let mut sum = request(&session_id, "sum(");
        sum.look_in = FindLookIn::FormulasOnly;
        let hits = run_to_completion(&mut sessions, sum);
        assert_eq!(coords(&hits), vec![("sheet-1", 1, 1)]);
        assert_eq!(hits[0].formula_text.as_deref(), Some("=SUM(C1:C1)"));
        let mut both = request(&session_id, "1");
        both.look_in = FindLookIn::Both;
        // C2 is boolean TRUE → "1"; B2 formula contains "1".
        assert_eq!(
            coords(&run_to_completion(&mut sessions, both)),
            vec![("sheet-1", 1, 1), ("sheet-1", 1, 2)]
        );
    }

    #[test]
    fn pages_with_limit_and_cursor() {
        let (_dir, mut sessions, session_id) = open();
        let mut paged = request(&session_id, "apple");
        paged.limit = 1;
        let not_cancelled = AtomicBool::new(false);
        let first = loop {
            let result = sessions.find_cells(&paged, &not_cancelled).unwrap();
            if result.indexing_complete || !result.matches.is_empty() {
                break result;
            }
            std::thread::sleep(Duration::from_millis(5));
        };
        assert_eq!(coords(&first.matches), vec![("sheet-1", 0, 0)]);
        assert!(!first.complete);
        assert_eq!(
            first.next_cursor,
            Some(FindCursor {
                sheet_id: "sheet-1".into(),
                row: 0,
                column: 1
            })
        );
        assert_eq!(run_to_completion(&mut sessions, paged).len(), 4);

        let mut two = request(&session_id, "apple");
        two.limit = 2;
        assert_eq!(
            coords(&run_to_completion(&mut sessions, two)),
            vec![
                ("sheet-1", 0, 0),
                ("sheet-1", 0, 1),
                ("sheet-1", 599, 0),
                ("sheet-2", 0, 0)
            ]
        );
    }

    #[test]
    fn single_sheet_scope_and_cancel() {
        let (_dir, mut sessions, session_id) = open();
        let mut scoped = request(&session_id, "apple");
        scoped.sheet_id = Some("sheet-2".into());
        assert_eq!(
            coords(&run_to_completion(&mut sessions, scoped)),
            vec![("sheet-2", 0, 0)]
        );
        let cancelled = AtomicBool::new(true);
        assert!(matches!(
            sessions.find_cells(&request(&session_id, "apple"), &cancelled),
            Err(SidecarError::Cancelled)
        ));
        assert!(matches!(
            sessions.find_cells(&request(&session_id, "   "), &AtomicBool::new(false)),
            Err(SidecarError::InvalidRequest(_))
        ));
    }

    #[test]
    fn number_text_mirrors_javascript() {
        assert_eq!(js_number_text(42.0), "42");
        assert_eq!(js_number_text(0.5), "0.5");
        assert_eq!(js_number_text(-3.25), "-3.25");
        assert_eq!(js_number_text(1e21), "1e+21");
        assert_eq!(js_number_text(1.5e-7), "1.5e-7");
        assert_eq!(js_number_text(0.0), "0");
    }
}
