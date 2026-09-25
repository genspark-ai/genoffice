/**
 * Where user scripts are kept (issue #815).
 *
 * v1 keeps them in localStorage next to the other per-app settings: a script
 * library follows the user, not the file, and no new IPC or disk format is
 * involved. Moving them into the document (or the project store) is a later step
 * — the shape here is already the one such a move would serialise.
 */
export interface SavedScript {
  id: string
  name: string
  code: string
  updatedAt: number
}

export const SCRIPT_STORAGE_KEY = 'genoffice.sheets.scripts.v1'
export const SCRIPT_MAX_CHARS = 200_000
export const SCRIPT_MAX_COUNT = 100

export function newScriptId(): string {
  return `s${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`
}

/** Reads back defensively: a hand-edited or half-written entry must not break the app. */
export function loadScripts(storage: Pick<Storage, 'getItem'> = localStorage): SavedScript[] {
  try {
    const raw = storage.getItem(SCRIPT_STORAGE_KEY)
    if (!raw) return []
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed
      .filter((entry): entry is SavedScript => {
        const rec = entry as Partial<SavedScript> | null
        return !!rec && typeof rec.id === 'string' && typeof rec.code === 'string'
      })
      .map((entry) => ({
        id: entry.id,
        name: typeof entry.name === 'string' && entry.name.trim() ? entry.name : 'Script',
        code: entry.code.slice(0, SCRIPT_MAX_CHARS),
        updatedAt: typeof entry.updatedAt === 'number' ? entry.updatedAt : Date.now(),
      }))
      .slice(0, SCRIPT_MAX_COUNT)
  } catch {
    return []
  }
}

export function saveScripts(
  scripts: readonly SavedScript[],
  storage: Pick<Storage, 'setItem'> = localStorage,
): boolean {
  try {
    storage.setItem(
      SCRIPT_STORAGE_KEY,
      JSON.stringify(
        scripts
          .slice(0, SCRIPT_MAX_COUNT)
          .map((s) => ({ ...s, code: s.code.slice(0, SCRIPT_MAX_CHARS) })),
      ),
    )
    return true
  } catch {
    // quota or private mode: reported by the caller, not silently swallowed
    return false
  }
}

export const SAMPLE_SCRIPT = `// Read the used range of the active sheet and log it.
// Every API call is asynchronous — await it.
const spreadsheet = await SpreadsheetApp.getActiveSpreadsheet();
if (!spreadsheet) throw new Error('No workbook is open');
const sheet = await spreadsheet.getActiveSheet();
const range = await sheet.getDataRange();
const values = await range.getValues();

Logger.log((await sheet.getName()) + ': ' + values.length + ' rows');
for (const row of values) Logger.log(row.join(' | '));
`
