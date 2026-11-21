//! Session-local worksheet cache. Each column is a separately decodable
//! block, so a column scan never parses unrelated cell records.
use super::*;

pub(crate) fn write_chunk(
    path: &Path,
    chunk: &ChunkData,
    reserve_bytes: impl FnOnce(usize) -> Result<(), SidecarError>,
) -> Result<(), SidecarError> {
    let mut columns = std::collections::BTreeMap::<usize, Vec<&CellRecord>>::new();
    for cell in &chunk.cells {
        columns.entry(cell.column).or_default().push(cell);
    }
    let mut payload = serde_json::to_vec(&chunk.rows)?;
    let mut directory = Vec::with_capacity(columns.len() * 8);
    for (column, cells) in columns {
        directory.extend_from_slice(&(column as u32).to_le_bytes());
        let offset = u32::try_from(payload.len())
            .map_err(|_| SidecarError::Io("Worksheet cache chunk is too large.".into()))?;
        directory.extend_from_slice(&offset.to_le_bytes());
        serde_json::to_writer(&mut payload, &cells)?;
    }
    reserve_bytes(4 + directory.len() + payload.len())?;
    let mut file = BufWriter::new(File::create(path)?);
    file.write_all(&((directory.len() / 8) as u32).to_le_bytes())?;
    file.write_all(&directory)?;
    file.write_all(&payload)?;
    file.flush()?;
    Ok(())
}

pub(crate) fn read_chunk(
    path: &Path,
    columns: std::ops::RangeInclusive<usize>,
    with_rows: bool,
) -> Result<ChunkData, SidecarError> {
    let mut file = File::open(path)?;
    let mut count = [0; 4];
    file.read_exact(&mut count)?;
    let mut directory = vec![0; u32::from_le_bytes(count) as usize * 8];
    file.read_exact(&mut directory)?;
    let payload_start = 4 + directory.len() as u64;
    let payload_end = file.metadata()?.len() - payload_start;
    let offset = |index: usize| {
        directory
            .get(index * 8 + 4..index * 8 + 8)
            .map_or(payload_end, |bytes| {
                u32::from_le_bytes(bytes.try_into().unwrap()) as u64
            })
    };
    let count = directory.len() / 8;
    let column_at = |index: usize| {
        u32::from_le_bytes(directory[index * 8..index * 8 + 4].try_into().unwrap()) as usize
    };
    let mut first = 0;
    while first < count && column_at(first) < *columns.start() {
        first += 1;
    }
    let mut last = first;
    while last < count && column_at(last) <= *columns.end() {
        last += 1;
    }
    let mut chunk = ChunkData::default();
    if with_rows {
        let mut data = vec![0; offset(0) as usize];
        file.read_exact(&mut data)?;
        chunk.rows = serde_json::from_slice(&data)?;
    }
    if first < last {
        let start = offset(first);
        file.seek(std::io::SeekFrom::Start(payload_start + start))?;
        let mut data = vec![0; (offset(last) - start) as usize];
        file.read_exact(&mut data)?;
        for index in first..last {
            let cells: Vec<CellRecord> = serde_json::from_slice(
                &data[(offset(index) - start) as usize..(offset(index + 1) - start) as usize],
            )?;
            chunk.cells.extend(cells);
        }
        if last - first > 1 {
            chunk.cells.sort_by_key(|cell| (cell.row, cell.column));
        }
    }
    Ok(chunk)
}
