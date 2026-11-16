import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createDocsInspection } from '../src/renderer/sdk/inspection'
import { createDocsAnnotations } from '../src/renderer/sdk/annotations'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))
function fixture() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>bad <b>va</b>lue</p>',
  })
  editors.push(editor)
  let identity = {}
  const current = () => ({ editor, identity })
  const inspection = createDocsInspection(current)
  const annotations = createDocsAnnotations(current, inspection)
  annotations.activate()
  return {
    editor,
    inspection,
    annotations,
    switchDocument() {
      identity = {}
      annotations.activate()
    },
  }
}

it('updates advisory rules and independent decorations without authoring content or undo entries', async () => {
  const { editor, annotations } = fixture()
  const before = editor.getJSON()
  annotations.setRules('host', [
    { id: 'warning', pattern: 'bad', message: 'Check this word', severity: 'warning' },
    { id: 'highlight', pattern: 'value', style: { underline: true, background: 'yellow' } },
  ])
  expect(annotations.diagnostics()).toHaveLength(1)
  expect(editor.view.dom.querySelectorAll('.sdk-range-decoration').length).toBeGreaterThan(0)
  expect(editor.getJSON()).toEqual(before)
  expect(editor.commands.undo()).toBe(false)
  editor.setEditable(false)
  expect(editor.view.dom.querySelector('.sdk-range-decoration')).not.toBeNull()
  editor.setEditable(true)
  editor.commands.selectAll()
  editor.commands.insertContent('good')
  await annotations.refresh()
  expect(annotations.diagnostics()).toEqual([])
  editor.commands.undo()
  await annotations.refresh()
  expect(annotations.diagnostics()).toHaveLength(1)
  annotations.remove('host')
  expect(editor.view.dom.querySelector('.sdk-range-decoration')).toBeNull()
})

it('suspends annotation boxes for printing and restores the current view without editing content', async () => {
  const { editor, annotations } = fixture()
  annotations.setRules('host', [{ id: 'bad', pattern: 'bad', message: 'Invalid' }])
  const before = editor.getJSON()
  const diagnostics = annotations.diagnostics()
  let transactions = 0
  editor.on('transaction', () => transactions++)
  window.dispatchEvent(new Event('beforeprint'))
  expect(editor.view.dom.querySelector('.sdk-range-decoration')).toBeNull()
  expect(annotations.diagnostics()).toEqual(diagnostics)
  window.dispatchEvent(new Event('afterprint'))
  expect(editor.view.dom.querySelector('.sdk-range-decoration')?.textContent).toBe('bad')
  expect(transactions).toBe(0)
  expect(editor.getJSON()).toEqual(before)
  expect(editor.commands.undo()).toBe(false)

  window.dispatchEvent(new Event('beforeprint'))
  annotations.setRules('host', [{ id: 'new', pattern: 'value', message: 'Current finding' }])
  await annotations.refresh()
  expect(editor.view.dom.querySelector('.sdk-range-decoration')).toBeNull()
  window.dispatchEvent(new Event('afterprint'))
  expect(editor.view.dom.querySelectorAll('.sdk-range-decoration')).toHaveLength(2)
  expect(annotations.diagnostics()[0].message).toBe('Current finding')
})

it('does not resurrect annotations reset while a print is outstanding', () => {
  const { editor, annotations } = fixture()
  annotations.setRules('host', [{ id: 'bad', pattern: 'bad', message: 'Invalid' }])
  window.dispatchEvent(new Event('beforeprint'))
  annotations.reset()
  window.dispatchEvent(new Event('afterprint'))
  expect(editor.view.dom.querySelector('.sdk-range-decoration')).toBeNull()
  expect(annotations.diagnostics()).toEqual([])
})

it('discards asynchronous stale findings and clears registrations on document switches', async () => {
  const { editor, annotations } = fixture()
  let resolve!: (value: Array<{ from: number; to: number; message: string }>) => void
  annotations.registerChecker(
    'async',
    () =>
      new Promise((done) => {
        resolve = done
      }),
  )
  const pending = annotations.refresh()
  annotations.remove('async')
  resolve([{ from: 1, to: 4, message: 'Old finding' }])
  await pending
  expect(annotations.diagnostics()).toEqual([])
  annotations.setRules('host', [{ id: 'bad', pattern: 'bad', message: 'Invalid' }])
  expect(annotations.diagnostics()).toHaveLength(1)
  const before = editor.getJSON()
  fixture() // another editor cannot consume this editor's annotations
  expect(editor.getJSON()).toEqual(before)
})

it('rejects stale explicit ranges and clears all ranges and rules on replacement', () => {
  const { annotations, inspection, switchDocument } = fixture()
  const snapshot = inspection.snapshot()
  annotations.setRanges('host', [{ ...snapshot, from: 1, to: 4, style: { underline: true } }])
  switchDocument()
  expect(annotations.diagnostics()).toEqual([])
  expect(() => annotations.setRanges('old', [{ ...snapshot, from: 1, to: 4 }])).toThrow(/stale/i)
})

it('maps explicit ranges through edits and makes checker failures observable without blocking edits', async () => {
  const { annotations, inspection, editor } = fixture()
  const snapshot = inspection.snapshot()
  annotations.setRanges('host', [{ ...snapshot, from: 5, to: 10, message: 'Mapped' }])
  editor.commands.insertContentAt(1, 'prefix ')
  await annotations.refresh()
  expect(annotations.diagnostics()[0]).toMatchObject({ from: 12, to: 17 })
  annotations.registerChecker('broken', () => {
    throw new Error('Host checker failed')
  })
  await expect(annotations.refresh()).rejects.toThrow('Host checker failed')
  expect(annotations.errors()).toEqual([{ source: 'broken', message: 'Host checker failed' }])
  expect(editor.getText()).toContain('prefix ')
  annotations.remove('broken')
  expect(annotations.errors()).toEqual([])
})

it('exposes pending and completed checks without running checkers when the snapshot is read', async () => {
  const { annotations, inspection, editor } = fixture()
  let calls = 0
  let complete!: (value: Array<{ from: number; to: number; message: string; code: string }>) => void
  annotations.registerChecker('host-review', () => {
    calls++
    return new Promise((resolve) => {
      complete = resolve
    })
  })
  const before = editor.getJSON()
  const reference = inspection.snapshot()
  expect(annotations.snapshot()).toMatchObject({
    documentId: reference.documentId,
    revision: reference.revision,
    complete: false,
    sources: [{ source: 'host-review', status: 'pending', coverage: { kind: 'document-text' } }],
    diagnostics: [],
  })
  annotations.snapshot()
  expect(calls).toBe(1)
  complete([{ from: 1, to: 4, message: 'Host finding', code: 'host.code' }])
  await Promise.resolve()
  expect(annotations.snapshot()).toMatchObject({
    complete: true,
    sources: [{ source: 'host-review', status: 'complete' }],
    diagnostics: [{ source: 'host-review', code: 'host.code', message: 'Host finding' }],
  })
  expect(editor.getJSON()).toEqual(before)
  expect(editor.commands.undo()).toBe(false)
})

it('does not relabel manually supplied diagnostics as checked after their decorations map through edits', async () => {
  const { annotations, inspection, editor } = fixture()
  const checked = inspection.snapshot()
  annotations.setRanges('manual', [{ ...checked, from: 5, to: 10, message: 'Check this value' }])
  expect(annotations.snapshot()).toMatchObject({
    complete: true,
    sources: [{ status: 'complete' }],
  })
  editor.commands.insertContentAt(1, 'prefix ')
  await annotations.refresh()
  const current = annotations.snapshot()
  expect(current.revision).not.toBe(checked.revision)
  expect(current).toMatchObject({
    complete: false,
    sources: [{ source: 'manual', status: 'stale', revision: checked.revision }],
  })
  expect(current.diagnostics[0]).toMatchObject({
    from: 12,
    to: 17,
    revision: checked.revision,
    stale: true,
  })
  annotations.setRanges('manual', [
    { ...inspection.snapshot(), from: 12, to: 17, message: 'Checked again' },
  ])
  expect(annotations.snapshot()).toMatchObject({
    complete: true,
    sources: [{ status: 'complete' }],
  })
})

it('reports failed checks separately from a clean result and clears source state on document replacement', async () => {
  const { annotations, switchDocument } = fixture()
  annotations.registerChecker('broken', () => {
    throw new Error('Unavailable checker')
  })
  await expect(annotations.refresh()).rejects.toThrow('Unavailable checker')
  expect(annotations.snapshot()).toMatchObject({
    complete: false,
    sources: [{ source: 'broken', status: 'failed', error: 'Unavailable checker' }],
    diagnostics: [],
  })
  switchDocument()
  expect(annotations.snapshot()).toMatchObject({ complete: true, sources: [], diagnostics: [] })
})

it('does not report an invalid checker result as a completed clean check', async () => {
  const { annotations } = fixture()
  annotations.registerChecker('invalid', () => [{ from: -1, to: 1000, message: 'Unlocatable' }])
  await expect(annotations.refresh()).rejects.toThrow(/invalid annotation range/i)
  expect(annotations.snapshot()).toMatchObject({
    complete: false,
    sources: [{ source: 'invalid', status: 'failed' }],
    diagnostics: [],
  })
})

it.each([
  ['{{field:known', 'unclosed'],
  ['{{field:{{field:known}}}}', 'nested'],
  ['{{field:\nknown}}', 'paragraph'],
  ['{{other:known}}', 'type'],
  ['{{field:}}', 'empty'],
  ['{{field:unknown}}', 'field'],
])(
  'runs a host-owned parser for %s without teaching the SDK its syntax',
  async (text, expected) => {
    const { editor, annotations } = fixture()
    editor.commands.setContent(
      text
        .split('\n')
        .map((part) => `<p>${part}</p>`)
        .join(''),
    )
    annotations.registerChecker('host-parser', (snapshot) => {
      const opening = snapshot.text.indexOf('{{')
      const closing = snapshot.text.indexOf('}}', opening + 2)
      const body = snapshot.text.slice(opening + 2, closing < 0 ? undefined : closing)
      const [kind, value] = body.split(':')
      const message =
        closing < 0
          ? 'unclosed'
          : body.includes('{{')
            ? 'nested'
            : body.includes('\n')
              ? 'paragraph'
              : !['field', 'ai'].includes(kind)
                ? 'type'
                : !value
                  ? 'empty'
                  : kind === 'field' && value !== 'known'
                    ? 'field'
                    : null
      if (!message) return []
      return [
        {
          from: snapshot.positions[opening]!,
          to: snapshot.positions[closing < 0 ? snapshot.text.length - 1 : closing + 1]! + 1,
          message,
        },
      ]
    })
    await annotations.refresh()
    expect(annotations.diagnostics()[0]?.message).toBe(expected)
  },
)
