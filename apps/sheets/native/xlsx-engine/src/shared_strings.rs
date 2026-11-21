use super::*;
use std::io::SeekFrom;

const DISK_THRESHOLD: u64 = 32 * 1024 * 1024;
const PAGE_STRINGS: usize = 128;
const READER_SLOTS: usize = 1021;
const MAX_CACHED_PAGE_BYTES: usize = 16 * 1024;
const RICH: u32 = 1 << 31;

pub(crate) enum SharedStrings {
    Memory(Vec<SharedString>),
    Disk(DiskSharedStrings),
}

impl SharedStrings {
    pub(crate) fn open(
        archive: &mut ZipArchive<File>,
        workbook: &Path,
        cache: &Path,
        colors: &ColorContext,
        progressive: bool,
        cancelled: Arc<AtomicBool>,
    ) -> Result<Self, SidecarError> {
        if !progressive
            || !zip_entry(archive, "xl/sharedStrings.xml")
                .is_ok_and(|entry| entry.size() >= DISK_THRESHOLD)
        {
            return Ok(Self::Memory(read_shared_strings(archive, colors)?));
        }
        Ok(Self::Disk(DiskSharedStrings::start(
            workbook, cache, colors, cancelled,
        )?))
    }

    pub(crate) fn reader(&self) -> Result<SharedStringReader<'_>, SidecarError> {
        Ok(match self {
            Self::Memory(strings) => SharedStringReader::Memory(strings),
            Self::Disk(table) => SharedStringReader::Disk(PageReader {
                table,
                file: File::open(&table.path)?,
                cache: (0..READER_SLOTS).map(|_| None).collect(),
                large_page: None,
            }),
        })
    }
}

pub(crate) struct DiskSharedStrings {
    path: PathBuf,
    pages: Arc<(Mutex<Pages>, Condvar)>,
    cancelled: Arc<AtomicBool>,
    writer: Option<JoinHandle<()>>,
}

impl DiskSharedStrings {
    fn start(
        workbook: &Path,
        cache: &Path,
        colors: &ColorContext,
        cancelled: Arc<AtomicBool>,
    ) -> Result<Self, SidecarError> {
        let path = cache.join("shared-strings.bin");
        let file = File::create_new(&path)?;
        let pages = Arc::new((Mutex::new(Pages::default()), Condvar::new()));
        let writer = {
            let workbook = workbook.to_owned();
            let colors = colors.clone();
            let pages = Arc::clone(&pages);
            let cancelled = Arc::clone(&cancelled);
            thread::Builder::new()
                .name("xlsx-shared-strings".into())
                .spawn(move || {
                    let result = write_pages(&workbook, &colors, file, &pages, &cancelled);
                    let mut state = pages.0.lock().unwrap();
                    state.finished = true;
                    state.error = result.err().map(|error| error.to_string());
                    pages.1.notify_all();
                })?
        };
        Ok(Self {
            path,
            pages,
            cancelled,
            writer: Some(writer),
        })
    }
}

impl Drop for DiskSharedStrings {
    fn drop(&mut self) {
        self.cancelled.store(true, Ordering::Release);
        if let Some(writer) = self.writer.take() {
            let _ = writer.join();
        }
    }
}

#[derive(Default)]
struct Pages {
    entries: Vec<Page>,
    finished: bool,
    error: Option<String>,
}

#[derive(Clone, Copy)]
struct Page {
    end: usize,
    offset: u64,
    bytes: usize,
}

// Publish complete pages so a worksheet can use early strings while the rest
// of the table is still inflating. The shared index holds one entry per page.
fn write_pages(
    workbook: &Path,
    colors: &ColorContext,
    mut file: File,
    pages: &(Mutex<Pages>, Condvar),
    cancelled: &AtomicBool,
) -> Result<(), SidecarError> {
    let mut archive = ZipArchive::new(File::open(workbook)?)?;
    let mut buffer = Vec::with_capacity(MAX_CACHED_PAGE_BYTES);
    let mut offset = 0;
    let mut count = 0;
    let mut publish = |buffer: &mut Vec<u8>, end| -> Result<(), SidecarError> {
        if buffer.is_empty() {
            return Ok(());
        }
        file.write_all(buffer)?;
        let mut state = pages.0.lock().unwrap();
        state.entries.push(Page {
            end,
            offset,
            bytes: buffer.len(),
        });
        offset += buffer.len() as u64;
        buffer.clear();
        pages.1.notify_all();
        Ok(())
    };

    visit_shared_strings(&mut archive, colors, |entry| {
        if cancelled.load(Ordering::Relaxed) {
            return Err(SidecarError::cancelled());
        }
        let rich;
        let (payload, flag) = if entry.runs.is_some() {
            rich = serde_json::to_vec(&entry)?;
            (rich.as_slice(), RICH)
        } else {
            (entry.text.as_bytes(), 0)
        };
        if payload.len() >= RICH as usize {
            return Err(SidecarError::Workbook("Shared string is too large.".into()));
        }
        buffer.extend_from_slice(&(payload.len() as u32 | flag).to_le_bytes());
        buffer.extend_from_slice(payload);
        count += 1;
        if count % PAGE_STRINGS == 0 {
            publish(&mut buffer, count)?;
        }
        Ok(())
    })?;
    publish(&mut buffer, count)
}

pub(crate) enum SharedStringReader<'a> {
    Memory(&'a [SharedString]),
    Disk(PageReader<'a>),
}

impl SharedStringReader<'_> {
    pub(crate) fn get(&mut self, index: usize) -> Result<Option<SharedString>, SidecarError> {
        match self {
            Self::Memory(strings) => Ok(strings.get(index).cloned()),
            Self::Disk(reader) => reader.get(index),
        }
    }
}

pub(crate) struct PageReader<'a> {
    table: &'a DiskSharedStrings,
    file: File,
    cache: Vec<Option<CachedPage>>,
    large_page: Option<CachedPage>,
}

struct CachedPage {
    start: usize,
    end: usize,
    data: Vec<u8>,
    offsets: Vec<usize>,
}

impl PageReader<'_> {
    fn get(&mut self, index: usize) -> Result<Option<SharedString>, SidecarError> {
        let page_index = index / PAGE_STRINGS;
        let slot = page_index % READER_SLOTS;
        let matches = |page: &&CachedPage| page.start / PAGE_STRINGS == page_index;
        let page = if let Some(page) = self.cache[slot].as_ref().filter(matches) {
            page
        } else if let Some(page) = self.large_page.as_ref().filter(matches) {
            page
        } else {
            let Some(page) = self.load_page(page_index)? else {
                return Ok(None);
            };
            if page.data.len() <= MAX_CACHED_PAGE_BYTES {
                self.cache[slot].insert(page)
            } else {
                self.large_page.insert(page)
            }
        };
        if index >= page.end {
            return Ok(None);
        }
        let start = page.offsets[index - page.start];
        let length = record_length(&page.data, start);
        let payload = &page.data[start + 4..start + 4 + (length & !RICH) as usize];
        Ok(Some(if length & RICH != 0 {
            serde_json::from_slice(payload)?
        } else {
            SharedString {
                text: std::str::from_utf8(payload)
                    .map_err(|error| SidecarError::Io(error.to_string()))?
                    .to_owned(),
                runs: None,
            }
        }))
    }

    fn load_page(&mut self, page_index: usize) -> Result<Option<CachedPage>, SidecarError> {
        let (lock, ready) = &*self.table.pages;
        let page = {
            let state = ready
                .wait_while(lock.lock().unwrap(), |state| {
                    !state.finished && state.entries.len() <= page_index
                })
                .unwrap();
            if let Some(error) = &state.error {
                return Err(SidecarError::Workbook(error.clone()));
            }
            let Some(&page) = state.entries.get(page_index) else {
                return Ok(None);
            };
            page
        };

        let start = page_index * PAGE_STRINGS;
        self.file.seek(SeekFrom::Start(page.offset))?;
        let mut data = vec![0; page.bytes];
        self.file.read_exact(&mut data)?;
        let mut offsets = Vec::with_capacity(page.end - start);
        let mut cursor = 0;
        for _ in start..page.end {
            offsets.push(cursor);
            cursor += 4 + (record_length(&data, cursor) & !RICH) as usize;
        }
        Ok(Some(CachedPage {
            start,
            end: page.end,
            data,
            offsets,
        }))
    }
}

fn record_length(data: &[u8], start: usize) -> u32 {
    u32::from_le_bytes(data[start..start + 4].try_into().unwrap())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn disk_table_preserves_ids_and_rich_strings_across_pages() {
        let large_text = "a".repeat(MAX_CACHED_PAGE_BYTES);
        let padding = "<si/>".repeat(PAGE_STRINGS - 2);
        let xml = format!(
            r#"<sst><si/><si><t>{large_text}</t></si>{padding}<si><r><rPr><b/></rPr><t>A&amp;B</t></r><rPh><t>ruby</t></rPh></si><si><t xml:space="preserve">  last  </t></si></sst>"#
        );
        let (dir, workbook) = crate::tests::open_fixture(&[("xl/sharedStrings.xml", &xml)]);
        let table = SharedStrings::Disk(
            DiskSharedStrings::start(
                &workbook,
                dir.path(),
                &ColorContext::default(),
                Arc::new(AtomicBool::new(false)),
            )
            .unwrap(),
        );
        let mut first = table.reader().unwrap();
        let mut second = table.reader().unwrap();
        assert_eq!(first.get(0).unwrap().unwrap().text, "");
        let styled = second.get(PAGE_STRINGS).unwrap().unwrap();
        assert_eq!(styled.text, "A&B");
        assert!(styled.runs.unwrap()[0].bold);
        assert_eq!(
            first.get(PAGE_STRINGS + 1).unwrap().unwrap().text,
            "  last  "
        );
        assert_eq!(second.get(1).unwrap().unwrap().text, large_text);
        assert!(first.get(PAGE_STRINGS + 2).unwrap().is_none());
    }
}
