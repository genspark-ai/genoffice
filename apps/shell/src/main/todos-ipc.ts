import { randomUUID } from 'node:crypto'
import { join } from 'node:path'
import { app, ipcMain } from 'electron'
import { createJsonStore } from './json-store'
import { HOME_CHANNELS, type TodoItem } from '../shared/home-api'

/** where the todo list lives: one file the whole widget reads and writes through */
export function todosFilePath(userDataDir: string): string {
  return join(userDataDir, 'todos.json')
}

/** the list never grows past this, whatever the renderer sends (defensive, not a refusal) */
const MAX_ITEMS = 500
/** one todo's text is clamped to this before anything is stored */
const MAX_TEXT = 500

function isTodoItem(value: unknown): value is TodoItem {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const item = value as Record<string, unknown>
  return (
    typeof item.id === 'string' && typeof item.text === 'string' && typeof item.done === 'boolean'
  )
}

/**
 * The store's invariant, applied to anything read or written: shape-valid
 * entries only, at most `MAX_ITEMS` of them, each text clamped. A payload that
 * is not an array is the empty list; entries that lost their shape are dropped
 * individually so one bad line never costs the rest of the list.
 */
function normalizeTodos(raw: unknown): TodoItem[] {
  if (!Array.isArray(raw)) return []
  const out: TodoItem[] = []
  for (const entry of raw) {
    if (!isTodoItem(entry)) continue
    out.push({ id: entry.id, text: entry.text.slice(0, MAX_TEXT), done: entry.done })
    if (out.length >= MAX_ITEMS) break
  }
  return out
}

/**
 * A missing, corrupt or shapeless file is the empty list — the todos are a
 * convenience, never a reason Home fails to open. Reading is also the one
 * moment a hand-edited or foreign file is normalized into the invariant.
 * Version 1 is today's bare array body (see json-store.ts for the marker and
 * migration convention).
 */
const todosStore = (userDataDir: string) =>
  createJsonStore<TodoItem[]>({
    path: () => todosFilePath(userDataDir),
    version: 1,
    normalize: normalizeTodos,
  })

export function readTodos(userDataDir: string): TodoItem[] {
  return todosStore(userDataDir).read()
}

/** the scratch-then-rename write the imports use, so a crash mid-write costs nothing */
export async function writeTodos(userDataDir: string, items: readonly TodoItem[]): Promise<void> {
  await todosStore(userDataDir).write([...items])
}

/**
 * One new entry added the way the widget adds one (prepended, id from
 * randomUUID, text clamped to the store's per-item cap); returns the
 * normalized list as written, so the caller sees exactly what landed.
 */
export async function addTodo(userDataDir: string, text: string): Promise<TodoItem[]> {
  const items = [{ id: randomUUID(), text, done: false }, ...readTodos(userDataDir)]
  const normalized = normalizeTodos(items)
  await writeTodos(userDataDir, normalized)
  return normalized
}

/**
 * The Home todo widget's IPC: list and whole-list save. The file is the user's
 * data, so a save that does not carry the promised shape is refused outright
 * (the file stays as it was), while size is only ever clamped — never a
 * user-facing refusal. Registered from index.ts like the other feature IPC
 * modules.
 */
export function registerTodosIpc(): void {
  ipcMain.handle(HOME_CHANNELS.todosList, async (): Promise<TodoItem[]> =>
    readTodos(app.getPath('userData')),
  )

  ipcMain.handle(HOME_CHANNELS.todosSave, async (_event, raw: unknown): Promise<void> => {
    if (!Array.isArray(raw) || !raw.every(isTodoItem)) {
      throw new Error('the todo list must be an array of { id, text, done } entries')
    }
    await writeTodos(app.getPath('userData'), normalizeTodos(raw))
  })
}
