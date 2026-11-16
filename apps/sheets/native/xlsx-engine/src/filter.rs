//! AutoFilter for streamed XLSX sheets. The renderer sends criteria and its
//! unsaved edits, and gets a page of value counts or a row visibility mask.
use super::*;
use base64::Engine as _;
use std::borrow::Cow;
use std::collections::VecDeque;

const MAX_ROWS: usize = 1_048_576;
const MAX_COLUMNS: usize = 16_384;
const MAX_PAGE: usize = 200;
const COLUMN_CACHE_BYTES: usize = 96 * 1024 * 1024;

#[derive(Clone, Debug, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterRowSpan {
    start_row: usize,
    end_row: usize,
    file_start_row: usize,
}

#[derive(Clone, Debug, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Override {
    start_row: usize,
    end_row: usize,
    value: Option<CellValue>,
}

#[derive(Clone, Debug, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Column {
    file_column: Option<usize>,
    #[serde(default)]
    overrides: Vec<Override>,
    predicate: Option<Predicate>,
}

#[derive(Clone, Debug, PartialEq, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
enum Predicate {
    Values {
        values: Vec<String>,
        blank: bool,
    },
    Conditions {
        conditions: Vec<Condition>,
        and: bool,
    },
}

#[derive(Clone, Debug, PartialEq, Deserialize)]
struct Condition {
    value: CellValue,
    operator: Option<Operator>,
}

#[derive(Clone, Copy, Debug, PartialEq, Deserialize)]
#[serde(rename_all = "camelCase")]
enum Operator {
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterRequest {
    session_id: String,
    sheet_id: String,
    start_row: usize,
    end_row: usize,
    spans: Vec<FilterRowSpan>,
    criteria: Vec<Column>,
    page: Option<Page>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Page {
    column: Column,
    query: String,
    offset: usize,
    limit: usize,
}

#[derive(Debug, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterResult {
    ready: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    hidden: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    page: Option<ValuePage>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ValueCount {
    text: String,
    count: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    style_index: Option<usize>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ValuePage {
    values: Vec<ValueCount>,
    total: usize,
}

#[derive(Clone, Default)]
struct FilterCell {
    text: String,
    // Conditions use the original value type, not a numeric-looking string.
    number: Option<f64>,
    exists: bool,
    style_index: Option<usize>,
}

fn filter_cell(value: Option<CellValue>, style_index: Option<usize>) -> FilterCell {
    let number = match value {
        Some(CellValue::Number(n)) => Some(n),
        _ => None,
    };
    let exists = value.is_some();
    let text = match value {
        Some(CellValue::String(text)) => text,
        Some(CellValue::Number(n)) => find::js_number_text(n),
        Some(CellValue::Boolean(b)) => (if b { "TRUE" } else { "FALSE" }).into(),
        None => String::new(),
    };
    FilterCell {
        text,
        number,
        exists,
        style_index,
    }
}

impl FilterCell {
    fn checklist_text(&self) -> Cow<'_, str> {
        if self.text.contains(['\r', '\n']) {
            Cow::Owned(self.text.replace(['\r', '\n'], ""))
        } else {
            Cow::Borrowed(&self.text)
        }
    }
}

struct Projection {
    cells: Vec<FilterCell>,
    bytes: usize,
}

type ColumnKey = (usize, usize, usize, usize);

#[derive(PartialEq)]
struct ValuesKey {
    sheet: usize,
    start: usize,
    end: usize,
    spans: Vec<FilterRowSpan>,
    criteria: Vec<Column>,
    column: Column,
}

#[derive(Default)]
pub(crate) struct ColumnCache {
    columns: VecDeque<(ColumnKey, Arc<Projection>)>,
    bytes: usize,
    values: Option<(ValuesKey, Vec<ValueCount>)>,
}

impl ColumnCache {
    fn get(&mut self, key: ColumnKey) -> Option<Arc<Projection>> {
        let index = self.columns.iter().position(|(cached, _)| *cached == key)?;
        let (_, projection) = self.columns.remove(index)?;
        self.columns.push_back((key, Arc::clone(&projection)));
        Some(projection)
    }

    fn put(&mut self, key: ColumnKey, projection: Arc<Projection>) {
        // Keep at least the last column even when all million texts are distinct.
        while self.bytes + projection.bytes > COLUMN_CACHE_BYTES {
            let Some((_, removed)) = self.columns.pop_front() else {
                break;
            };
            self.bytes -= removed.bytes;
        }
        self.bytes += projection.bytes;
        self.columns.push_back((key, projection));
    }
}

impl FilterRequest {
    fn validate(&self) -> Result<(), SidecarError> {
        let valid_span = |start, end| start <= end && end < MAX_ROWS;
        if !valid_span(self.start_row, self.end_row)
            || self.spans.len() > MAX_ROWS
            || self.spans.iter().any(|s| {
                !valid_span(s.start_row, s.end_row)
                    || s.start_row < self.start_row
                    || s.end_row > self.end_row
                    || s.file_start_row >= MAX_ROWS
                    || s.file_start_row
                        .saturating_add(s.end_row.saturating_sub(s.start_row))
                        >= MAX_ROWS
            })
            || self
                .spans
                .windows(2)
                .any(|pair| pair[0].end_row >= pair[1].start_row)
            || self.criteria.len() > MAX_COLUMNS
            || self
                .page
                .as_ref()
                .is_some_and(|p| !(1..=MAX_PAGE).contains(&p.limit))
            || self
                .criteria
                .iter()
                .chain(self.page.iter().map(|p| &p.column))
                .any(|col| {
                    col.file_column.is_some_and(|c| c >= MAX_COLUMNS)
                        || col.overrides.iter().any(|o| {
                            !valid_span(o.start_row, o.end_row)
                                || o.start_row < self.start_row
                                || o.end_row > self.end_row
                        })
                })
        {
            return Err(SidecarError::InvalidRequest(
                "Invalid AutoFilter bounds.".into(),
            ));
        }
        Ok(())
    }
}

enum Matcher {
    Values(HashSet<String>, bool),
    Conditions(Vec<ConditionMatcher>, bool),
}

enum ConditionMatcher {
    Number(Operator, f64),
    Text(find::FilterGlob, bool),
    NotText(String),
    NonBlank,
    Never,
}

impl ConditionMatcher {
    fn new(condition: &Condition) -> Self {
        let text = filter_cell(Some(condition.value.clone()), None).text;
        if let Some(operator) = condition.operator {
            if let Some(number) = condition_number(&condition.value, &text) {
                return Self::Number(operator, number);
            }
            if matches!(operator, Operator::NotEqual) {
                return match text.as_str() {
                    " " => Self::NonBlank,
                    _ if text.contains(['*', '?']) => {
                        Self::Text(find::FilterGlob::new(&text), true)
                    }
                    _ => Self::NotText(text),
                };
            }
            return Self::Never;
        }
        Self::Text(find::FilterGlob::new(&text), false)
    }

    fn keeps(&self, cell: &FilterCell) -> bool {
        match self {
            Self::Number(operator, right) => match operator {
                Operator::Equal => cell.number == Some(*right),
                Operator::NotEqual => cell.number != Some(*right),
                Operator::GreaterThan => cell.number.is_some_and(|n| n > *right),
                Operator::GreaterThanOrEqual => cell.number.is_some_and(|n| n >= *right),
                Operator::LessThan => cell.number.is_some_and(|n| n < *right),
                Operator::LessThanOrEqual => cell.number.is_some_and(|n| n <= *right),
            },
            Self::Text(glob, negate) => {
                let matched = glob.matches(&cell.text) && (!*negate || cell.exists);
                matched != *negate
            }
            Self::NotText(text) => !cell.exists || &cell.text != text,
            Self::NonBlank => cell.exists,
            Self::Never => false,
        }
    }
}

// The condition editor only coerces plain decimal strings. The cell on
// the left still has to be numeric, including when a numeric string was entered.
fn condition_number(value: &CellValue, text: &str) -> Option<f64> {
    if let CellValue::Number(number) = value {
        return Some(*number);
    }
    let text = text.strip_prefix('-').unwrap_or(text);
    let digits = |part: &str| !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit());
    if !text.split('.').all(digits) {
        return None;
    }
    match value {
        CellValue::String(text) => text.parse().ok(),
        _ => None,
    }
}

impl Matcher {
    fn new(predicate: &Predicate) -> Self {
        match predicate {
            Predicate::Values { values, blank } => {
                Self::Values(values.iter().cloned().collect(), *blank)
            }
            Predicate::Conditions { conditions, and } => Self::Conditions(
                conditions
                    .iter()
                    .take(2)
                    .map(ConditionMatcher::new)
                    .collect(),
                *and,
            ),
        }
    }
    fn keeps(&self, cell: &FilterCell) -> bool {
        match self {
            Self::Values(values, blank) => {
                let text = cell.checklist_text();
                if text.is_empty() {
                    *blank
                } else {
                    values.contains(text.as_ref())
                }
            }
            Self::Conditions(conditions, and) => {
                conditions.is_empty()
                    || if *and {
                        conditions.iter().all(|c| c.keeps(cell))
                    } else {
                        conditions.iter().any(|c| c.keeps(cell))
                    }
            }
        }
    }
}

impl WorkbookSessions {
    pub fn filter_rows(
        &mut self,
        request: FilterRequest,
        cancelled: &AtomicBool,
    ) -> Result<FilterResult, SidecarError> {
        request.validate()?;
        let session = self
            .sessions
            .get_mut(&request.session_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown workbook session.".into()))?;
        let sheet = session
            .sheets
            .iter()
            .position(|sheet| sheet.id == request.sheet_id)
            .ok_or_else(|| SidecarError::InvalidRequest("Unknown worksheet.".into()))?;
        session.ensure_parser(sheet)?;
        session.filter_rows(sheet, request, cancelled)
    }
}

fn file_bounds(spans: &[FilterRowSpan], row_count: usize) -> Option<(usize, usize)> {
    let start = spans.iter().map(|s| s.file_start_row).min()?;
    let end = spans
        .iter()
        .map(|s| s.file_start_row + s.end_row - s.start_row)
        .max()?;
    (start < row_count).then(|| (start, end.min(row_count - 1)))
}

struct ColumnView<'a> {
    base: Option<(usize, Arc<Projection>)>,
    spans: &'a [FilterRowSpan],
    overrides: Vec<Option<u32>>,
    edited: Vec<FilterCell>,
    start: usize,
}

impl ColumnView<'_> {
    fn walk<'a>(
        &'a self,
        blank: &'a FilterCell,
        mut read: impl FnMut(usize, &'a FilterCell),
        len: usize,
    ) {
        let mut spans = self.spans.iter().peekable();
        for i in 0..len {
            if let Some(index) = self.overrides.get(i).copied().flatten() {
                read(i, &self.edited[index as usize]);
                continue;
            }
            let row = self.start + i;
            while spans.peek().is_some_and(|span| span.end_row < row) {
                spans.next();
            }
            let cell = self.base.as_ref().and_then(|(first, projection)| {
                let span = spans.peek().filter(|span| span.start_row <= row)?;
                projection
                    .cells
                    .get(span.file_start_row + row - span.start_row - first)
            });
            read(i, cell.unwrap_or(blank));
        }
    }
}

impl WorkbookSession {
    fn load_filter_column(
        &mut self,
        sheet: usize,
        column: usize,
        start: usize,
        end: usize,
        cancelled: &AtomicBool,
    ) -> Result<Arc<Projection>, SidecarError> {
        let key = (sheet, column, start, end);
        if let Some(projection) = self.filter_columns.get(key) {
            return Ok(projection);
        }
        let (lock, _) = &*self.runtimes[sheet].state;
        let index = lock
            .lock()
            .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
        let paths: Vec<_> = (start / CHUNK_ROW_COUNT..=end / CHUNK_ROW_COUNT)
            .filter_map(|chunk| index.chunk_files.get(&chunk).cloned())
            .collect();
        drop(index);
        let mut cells = vec![FilterCell::default(); end - start + 1];
        for path in paths {
            if cancelled.load(Ordering::Acquire) {
                return Err(SidecarError::Cancelled);
            }
            for cell in chunks::read_chunk(&path, column..=column, false)?.cells {
                if cell.row >= start && cell.row <= end {
                    cells[cell.row - start] = filter_cell(cell.value, cell.style_index);
                }
            }
        }
        let bytes = cells.len() * std::mem::size_of::<FilterCell>()
            + cells.iter().map(|c| c.text.capacity()).sum::<usize>();
        let projection = Arc::new(Projection { cells, bytes });
        self.filter_columns.put(key, Arc::clone(&projection));
        Ok(projection)
    }

    fn filter_view<'a>(
        &mut self,
        sheet: usize,
        bounds: Option<(usize, usize)>,
        request: &'a FilterRequest,
        column: &Column,
        cancelled: &AtomicBool,
    ) -> Result<ColumnView<'a>, SidecarError> {
        let base = match (bounds, column.file_column) {
            (Some((start, end)), Some(col)) if col < self.sheets[sheet].column_count => Some((
                start,
                self.load_filter_column(sheet, col, start, end, cancelled)?,
            )),
            _ => None,
        };
        let mut overrides = Vec::new();
        let mut edited = Vec::new();
        if !column.overrides.is_empty() {
            overrides.resize(request.end_row - request.start_row + 1, None);
            for edit in &column.overrides {
                overrides[edit.start_row - request.start_row..=edit.end_row - request.start_row]
                    .fill(Some(edited.len() as u32));
                edited.push(filter_cell(edit.value.clone(), None));
            }
        }
        Ok(ColumnView {
            base,
            spans: &request.spans,
            overrides,
            edited,
            start: request.start_row,
        })
    }

    fn filter_rows(
        &mut self,
        sheet: usize,
        request: FilterRequest,
        cancelled: &AtomicBool,
    ) -> Result<FilterResult, SidecarError> {
        let bounds = file_bounds(&request.spans, self.sheets[sheet].row_count);
        if let Some((_, end)) = bounds {
            let (lock, _) = &*self.runtimes[sheet].state;
            let index = lock
                .lock()
                .map_err(|_| SidecarError::Io("Worksheet index lock was poisoned.".into()))?;
            if let Some(error) = &index.error {
                return Err(SidecarError::Workbook(error.clone()));
            }
            if !index.complete
                && index
                    .indexed_through_row
                    .is_none_or(|through| through < end)
            {
                return Ok(FilterResult::default());
            }
        }
        let values_key = request.page.as_ref().map(|page| ValuesKey {
            sheet,
            start: request.start_row,
            end: request.end_row,
            spans: request.spans.clone(),
            criteria: request.criteria.clone(),
            column: page.column.clone(),
        });
        if let (Some(key), Some(page), Some((cached, values))) =
            (&values_key, &request.page, &self.filter_columns.values)
        {
            if key == cached {
                return Ok(FilterResult {
                    ready: true,
                    hidden: None,
                    page: Some(value_page(values, page)),
                });
            }
        }
        let len = request.end_row - request.start_row + 1;
        let mut hidden = vec![0u8; len.div_ceil(8)];
        let blank = FilterCell::default();
        for column in &request.criteria {
            if cancelled.load(Ordering::Acquire) {
                return Err(SidecarError::Cancelled);
            }
            let Some(predicate) = &column.predicate else {
                continue;
            };
            let matcher = Matcher::new(predicate);
            let view = self.filter_view(sheet, bounds, &request, column, cancelled)?;
            view.walk(
                &blank,
                |i, cell| {
                    let bit = 1 << (i % 8);
                    if hidden[i / 8] & bit == 0 && !matcher.keeps(cell) {
                        hidden[i / 8] |= bit;
                    }
                },
                len,
            );
        }
        if let (Some(page), Some(key)) = (&request.page, values_key) {
            let view = self.filter_view(sheet, bounds, &request, &page.column, cancelled)?;
            let mut counts = HashMap::<Cow<str>, (usize, Option<usize>)>::new();
            // Borrow the strings owned by the projection until aggregation ends.
            view.walk(
                &blank,
                |i, cell| {
                    if hidden[i / 8] & (1 << (i % 8)) == 0 {
                        let count = counts
                            .entry(cell.checklist_text())
                            .or_insert((0, cell.style_index));
                        count.0 += 1;
                    }
                },
                len,
            );
            let mut sorted: Vec<_> = counts
                .into_iter()
                .map(|(text, (count, style_index))| {
                    let number = text.parse::<f64>().ok().filter(|n| n.is_finite());
                    (text, count, style_index, number)
                })
                .collect();
            sorted.sort_unstable_by(|a, b| match (a.3, b.3) {
                (Some(x), Some(y)) => x.total_cmp(&y).then(a.0.cmp(&b.0)),
                (Some(_), None) => std::cmp::Ordering::Less,
                (None, Some(_)) => std::cmp::Ordering::Greater,
                (None, None) => a.0.cmp(&b.0),
            });
            let values: Vec<_> = sorted
                .into_iter()
                .map(|(text, count, style_index, _)| ValueCount {
                    text: text.into_owned(),
                    count,
                    style_index,
                })
                .collect();
            let result = value_page(&values, page);
            self.filter_columns.values = Some((key, values));
            return Ok(FilterResult {
                ready: true,
                hidden: None,
                page: Some(result),
            });
        }
        Ok(FilterResult {
            ready: true,
            hidden: Some(base64::engine::general_purpose::STANDARD.encode(hidden)),
            page: None,
        })
    }
}

fn value_page(values: &[ValueCount], page: &Page) -> ValuePage {
    let query = page.query.to_lowercase();
    let words: Vec<_> = query.split_whitespace().collect();
    if words.is_empty() {
        return ValuePage {
            values: values
                .iter()
                .skip(page.offset)
                .take(page.limit)
                .cloned()
                .collect(),
            total: values.len(),
        };
    }
    let mut found = Vec::new();
    let mut total = 0;
    for entry in values {
        let folded = entry.text.to_lowercase();
        if words.iter().all(|needle| folded.contains(needle)) {
            if total >= page.offset && found.len() < page.limit {
                found.push(entry.clone());
            }
            total += 1;
        }
    }
    ValuePage {
        values: found,
        total,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tests::open_fixture;
    use serde_json::json;

    #[test]
    fn filters_and_reads_matching_rows_across_large_gaps() {
        let (_dir, path) = open_fixture(&[
            (
                "xl/workbook.xml",
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Test" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dimension ref="A1:B920001"/>
                <sheetData>
                  <row r="1"><c r="A1" t="inlineStr"><is><t>Group</t></is></c></row>
                  <row r="2"><c r="A2" t="inlineStr"><is><t>Drop</t></is></c></row>
                  <row r="500001"><c r="A500001" t="inlineStr"><is><t>Keep</t></is></c><c r="B500001"><v>11</v></c></row>
                  <row r="920001"><c r="A920001" t="inlineStr"><is><t>Keep</t></is></c><c r="B920001"><v>22</v></c></row>
                </sheetData></worksheet>"#,
            ),
        ]);
        let mut sessions = WorkbookSessions::new();
        let open = sessions.open(&path).unwrap();
        let sheet = &open.sheets[0];
        let cancelled = AtomicBool::new(false);
        let criteria = json!([{
            "fileColumn": 0,
            "predicate": { "kind": "values", "values": ["Keep"], "blank": false }
        }]);
        let mut request = json!({
            "sessionId": open.session_id,
            "sheetId": sheet.id,
            "startRow": 1,
            "endRow": 920000,
            "spans": [{ "startRow": 1, "endRow": 920000, "fileStartRow": 1 }],
            "criteria": criteria,
            "page": { "column": { "fileColumn": 1 }, "query": "", "offset": 1, "limit": 1 },
        });
        let deadline = std::time::Instant::now() + Duration::from_secs(5);
        let page = loop {
            let response = sessions
                .filter_rows(serde_json::from_value(request.clone()).unwrap(), &cancelled)
                .unwrap();
            if response.ready {
                break response.page.unwrap();
            }
            assert!(
                std::time::Instant::now() < deadline,
                "worksheet indexing stalled"
            );
            std::thread::sleep(Duration::from_millis(10));
        };
        assert_eq!(page.total, 2);
        assert_eq!(page.values[0].text, "22");
        assert_eq!(page.values[0].count, 1);

        request.as_object_mut().unwrap().remove("page");
        let result = sessions
            .filter_rows(serde_json::from_value(request).unwrap(), &cancelled)
            .unwrap();
        let mask = base64::engine::general_purpose::STANDARD
            .decode(result.hidden.unwrap())
            .unwrap();
        let matches: Vec<_> = (1..=920_000)
            .filter(|&row| mask[(row - 1) / 8] & (1 << ((row - 1) % 8)) == 0)
            .collect();
        assert_eq!(matches, [500_000, 920_000]);

        let viewport = sessions
            .read_range_with_mode(
                &open.session_id,
                &sheet.id,
                &CellRange {
                    start_row: 500_000,
                    end_row: 920_000,
                    start_column: 1,
                    end_column: 1,
                },
                RangeReadMode::Viewport,
                Some(&matches),
            )
            .unwrap();
        assert_eq!(
            viewport
                .cells
                .iter()
                .map(|cell| (cell.row, cell.value.clone()))
                .collect::<Vec<_>>(),
            [
                (500_000, Some(CellValue::Number(11.0))),
                (920_000, Some(CellValue::Number(22.0)))
            ]
        );
    }
}
