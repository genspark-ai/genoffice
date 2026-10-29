import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The todo widget's two channels against a real temp dir: the handlers are the
 * only writers of `<userData>/todos.json`, so this pins what that file is
 * allowed to look like — roundtrip, corrupt-reads-empty, shape refusal, size
 * clamps, and that a save never leaves the file half-written.
 */

const mockedRename = vi.hoisted(() => vi.fn())
/** the real rename, captured at mock time for the retry test's second attempt */
const realRename = vi.hoisted(() => ({ fn: undefined as typeof import('node:fs/promises').rename }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  realRename.fn = actual.rename
  return { ...actual, rename: mockedRename }
})

const handles = new Map<string, (...args: never[]) => unknown>()
let userData = ''

vi.mock('electron', () => ({
  app: { getPath: () => userData },
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      handles.set(channel, handler),
  },
}))

import { HOME_CHANNELS } from '../src/shared/home-api'
import { registerTodosIpc, todosFilePath } from '../src/main/todos-ipc'

const invoke = (channel: string, ...args: unknown[]) =>
  handles.get(channel)!({ sender: { id: 7 } } as never, ...args)

const listTodos = () => invoke(HOME_CHANNELS.todosList) as Promise<unknown[]>
const saveTodos = (raw: unknown) => invoke(HOME_CHANNELS.todosSave, raw) as Promise<void>

const item = (id: string, text: string, done = false) => ({ id, text, done })

/** the file's parsed content, as the next launch would read it */
const readFile = (): unknown => JSON.parse(readFileSync(todosFilePath(userData), 'utf8'))

/** scratch files a failed or finished atomic write might leave behind */
const leftovers = (): string[] =>
  readdirSync(userData).filter((name) => name.includes('.tmp') || name.startsWith('.todos'))

beforeEach(() => {
  userData = mkdtempSync(join(tmpdir(), 'genoffice-todos-'))
  // the write is the real scratch-then-rename unless a test overrides it
  mockedRename
    .mockReset()
    .mockImplementation((from: string, to: string) => realRename.fn!(from, to))
  registerTodosIpc()
})

afterEach(() => {
  mockedRename.mockReset()
  rmSync(userData, { recursive: true, force: true })
})

describe('home:todos-list', () => {
  it('reads back what a save wrote, order and flags intact', async () => {
    await saveTodos([item('b', 'second', true), item('a', 'first')])
    await expect(listTodos()).resolves.toEqual([item('b', 'second', true), item('a', 'first')])
  })

  it('answers an empty list when the file does not exist yet', async () => {
    await expect(listTodos()).resolves.toEqual([])
  })

  it('answers an empty list for a corrupt or shapeless file', async () => {
    for (const garbage of ['{oops', 'null', '"nope"', '42', '{"items":[]}', '[]]']) {
      writeFileSync(todosFilePath(userData), garbage, 'utf8')
      await expect(listTodos()).resolves.toEqual([])
    }
  })

  it('drops entries that lost their shape, keeping the rest of the list', async () => {
    writeFileSync(
      todosFilePath(userData),
      JSON.stringify([
        item('a', 'good one'),
        { id: 'b', text: 'no flag' },
        'not an object',
        item('c', 'also good', true),
        null,
      ]),
      'utf8',
    )
    await expect(listTodos()).resolves.toEqual([
      item('a', 'good one'),
      item('c', 'also good', true),
    ])
  })

  it('normalizes a hand-edited file to the store invariant: caps and no extra fields', async () => {
    const many = Array.from({ length: 600 }, (_unused, i) =>
      item(`id${i}`, 'x'.repeat(600), i % 2 === 0),
    )
    writeFileSync(todosFilePath(userData), JSON.stringify(many), 'utf8')
    const listed = await listTodos()
    expect(listed).toHaveLength(500)
    expect(listed[0]).toEqual(item('id0', 'x'.repeat(500), true))
    expect(listed[499]).toEqual(item('id499', 'x'.repeat(500), false))
  })
})

describe('home:todos-save', () => {
  it('writes a pretty-printed array to <userData>/todos.json', async () => {
    await saveTodos([item('a', 'buy milk')])
    expect(readFile()).toEqual([item('a', 'buy milk')])
    expect(readFileSync(todosFilePath(userData), 'utf8')).toContain('\n  ')
  })

  it('refuses a payload that is not an array of { id, text, done }, leaving the file untouched', async () => {
    await saveTodos([item('a', 'kept')])
    const before = readFileSync(todosFilePath(userData), 'utf8')

    for (const bad of [
      undefined,
      null,
      'a list',
      42,
      { id: 'a', text: 'x', done: false },
      [item('a', 'ok'), { id: 'b', text: 'no flag' }],
      [item('a', 'ok'), { id: 'b', done: true }],
      [item('a', 'ok'), { text: 'no id', done: true }],
      [['id', 'text', false]],
      [null],
    ]) {
      await expect(saveTodos(bad), `${JSON.stringify(bad)} should refuse`).rejects.toThrow(
        'must be an array',
      )
    }
    expect(readFileSync(todosFilePath(userData), 'utf8')).toBe(before)
    expect(leftovers()).toEqual([])
  })

  it('clamps to 500 items and 500-character texts instead of refusing', async () => {
    const many = Array.from({ length: 502 }, (_unused, i) => item(`id${i}`, 'x'.repeat(600)))
    await saveTodos(many)
    const stored = readFile() as Array<{ id: string; text: string; done: boolean }>
    expect(stored).toHaveLength(500)
    expect(stored.map((entry) => entry.id)).toEqual(many.slice(0, 500).map((entry) => entry.id))
    expect(stored[0]!.text).toBe('x'.repeat(500))
  })

  it('stores exactly the contract shape, dropping anything extra an entry carries', async () => {
    await saveTodos([{ id: 'a', text: 'x', done: false, note: 'foreign field' }])
    expect(readFile()).toEqual([{ id: 'a', text: 'x', done: false }])
  })

  it('lets an empty array empty the file, not delete it', async () => {
    await saveTodos([item('a', 'one')])
    await saveTodos([])
    await expect(listTodos()).resolves.toEqual([])
  })

  it('keeps the last of two saves when they land in order', async () => {
    await saveTodos([item('a', 'first write')])
    await saveTodos([item('b', 'second write'), item('a', 'first write')])
    await expect(listTodos()).resolves.toEqual([
      item('b', 'second write'),
      item('a', 'first write'),
    ])
  })
})

describe('the file is never half-written', () => {
  it('publishes through a scratch file and leaves none behind on success', async () => {
    await saveTodos([item('a', 'one')])
    expect(readdirSync(userData)).toEqual(['todos.json'])
    expect(readFile()).toEqual([item('a', 'one')])
  })

  it('keeps the previous list intact, and no scratch file behind, when the publish fails', async () => {
    await saveTodos([item('a', 'old')])
    const before = readFileSync(todosFilePath(userData), 'utf8')
    mockedRename.mockRejectedValue(Object.assign(new Error('busy'), { code: 'EPERM' }))
    try {
      await expect(saveTodos([item('b', 'new'), item('a', 'old')])).rejects.toMatchObject({
        code: 'EPERM',
      })
      expect(readFileSync(todosFilePath(userData), 'utf8')).toBe(before)
      expect(readdirSync(userData)).toEqual(['todos.json'])
    } finally {
      mockedRename.mockReset()
    }
  })

  it('is read the same after a write that retried a locked rename', async () => {
    await saveTodos([item('a', 'old')])
    // one EPERM (AV / cloud-sync lock), then the real rename
    mockedRename
      .mockRejectedValueOnce(Object.assign(new Error('busy'), { code: 'EPERM' }))
      .mockImplementationOnce((from: unknown, to: unknown) =>
        realRename.fn!(from as string, to as string),
      )
    try {
      await saveTodos([item('b', 'new')])
      await expect(listTodos()).resolves.toEqual([item('b', 'new')])
      expect(readdirSync(userData)).toEqual(['todos.json'])
    } finally {
      mockedRename.mockReset()
    }
  })
})
