// Modified for ForaOffice from GenOffice v0.12.852.
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { parseDocx, readSections } from '@genoffice/docx-engine'
import { buildDocx, TABLE_XML } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocOptions } from '../src/renderer/editor/convert'
import { collectRevisions } from '../src/renderer/editor/revisions'
import { executeTool } from '../src/renderer/ai/tools'
import {
  buildDocBytes,
  exportHtml,
  exportMarkdown,
  exportPdf,
  exportImages,
  noteDocumentSwapped,
  save,
  type FileActionContext,
} from '../src/renderer/file-actions'
import { DraftLanding } from '../src/renderer/runtime/draft-landing'
import { EMPTY_HF_VARIANTS } from '../src/renderer/doc-state'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const dirs: string[] = []
const editors: Editor[] = []
afterEach(async () => {
  for (const editor of editors.splice(0)) editor.destroy()
  for (const dir of dirs.splice(0)) await rm(dir, { recursive: true, force: true })
})

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const first = paragraph('Unchanged opening paragraph.')
const last = paragraph('Unchanged closing paragraph.')
const bodyXml = first + TABLE_XML + paragraph('Respond within forty-eight (48) hours.') + last

async function fixture(tracking: boolean, fixtureBody = bodyXml, extraStylesXml?: string) {
  noteDocumentSwapped()
  const bytes = await buildDocx({ bodyXml: fixtureBody, extraStylesXml })
  const parsed = await parseDocx(bytes)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-save-tracked-'))
  dirs.push(dir)
  const path = join(dir, 'fixture.docx')
  await writeFile(path, bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: { handleScrollToSelection: () => true },
    content: blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed)),
  })
  editors.push(editor)
  editor.storage.trackChanges.enabled = tracking
  const state = {
    editor,
    doc: { parsed, filePath: path, fileName: 'fixture.docx', hash: 'fixture' },
    dirtyRef: { current: false },
    saveInFlightRef: { current: false },
    saveIncompleteRef: { current: false },
    inkAnnotations: [],
    sections: readSections(parsed),
    sectionsDirty: [],
    pgNumDirtySections: [],
    hfVariants: EMPTY_HF_VARIANTS,
    hfVariantsDirty: [],
    sectionHfEdits: {},
    hfLinks: {},
    styleUpserts: {},
    pendingNumbering: { newDefs: [], restartNums: [] },
  }
  const noop = () => {}
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (prop === 'setDoc')
        return (updater: (prev: typeof state.doc) => typeof state.doc) => {
          target.doc = updater(target.doc)
        }
      if (prop in target) return Reflect.get(target, prop)
      if (typeof prop === 'string' && (/^set[A-Z]/.test(prop) || /^on[A-Z]/.test(prop))) return noop
      return undefined
    },
  }) as unknown as FileActionContext
  window.desktop = {
    saveDocx: async (target: string, data: ArrayBuffer) => {
      await writeFile(target, new Uint8Array(data))
      return { ok: true }
    },
  } as unknown as typeof window.desktop
  return { editor, ctx, path, bytes }
}

async function replace(editor: Editor) {
  const result = await executeTool(
    editor,
    {
      id: 'fixture-edit',
      name: 'apply_ops',
      input: {
        ops: [
          { op: 'findReplace', find: 'forty-eight (48) hours', replace: 'seventy-two (72) hours' },
        ],
      },
    },
    { bullet: null, ordered: null },
    { author: 'Fixture Agent' },
  )
  expect(result.isError).toBeFalsy()
}

async function assertUntouched(original: Uint8Array, saved: Uint8Array) {
  const before = await JSZip.loadAsync(original)
  const after = await JSZip.loadAsync(saved)
  const xml = await after.file('word/document.xml')!.async('string')
  expect(xml).toContain(first)
  expect(xml).toContain(last)
  expect(xml).toContain(TABLE_XML)
  for (const [name, part] of Object.entries(before.files)) {
    if (part.dir || name === 'word/document.xml') continue
    expect(await after.file(name)!.async('uint8array'), name).toEqual(
      await part.async('uint8array'),
    )
  }
}

describe('production save reload with track changes enabled', () => {
  it('keeps a changed table style inherited through save and a subsequent text edit', async () => {
    const styles = ['FF0000', '0000FF']
      .map(
        (fill, index) =>
          `<w:style w:type="table" w:styleId="Style${index}"><w:name w:val="Style${index}"/><w:tcPr><w:shd w:fill="${fill}"/></w:tcPr></w:style>`,
      )
      .join('')
    const { editor, ctx, path } = await fixture(
      false,
      '<w:tbl><w:tblPr><w:tblStyle w:val="Style0"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
        paragraph('ORIGINAL') +
        '</w:tc></w:tr></w:tbl><w:p/>',
      styles,
    )
    editor.storage.listNumbering.styles = ctx.doc!.parsed.styles
    const runtime = createDocsRuntime({
      current: () => ({
        editor,
        identity: editor,
        numIds: { bullet: null, ordered: null },
        autoAcceptUntracked: true,
      }),
      save: async () => {
        const ok = await save(ctx, false, true)
        return { ok, path, committed: ok, pending: !ok }
      },
    })
    try {
      const caller = runtime.connect({ id: 'saved-table-style' })
      caller.context()
      const style = await caller.execute({
        id: 'style',
        name: 'apply_ops',
        input: { ops: [{ op: 'setTableStyle', target: { blockIndexes: [0] }, styleId: 'Style1' }] },
      })
      expect(style.isError, style.output).not.toBe(true)
      expect((await caller.save({ path, overwrite: true })).ok).toBe(true)
      caller.context({ captureSelection: false })
      expect(ctx.doc!.parsed.blocks[0].table!.rows[0][0].fill).toBe('0000FF')
      const text = await caller.execute({
        id: 'text',
        name: 'apply_ops',
        input: {
          ops: [
            {
              op: 'setTableCellText',
              target: { blockIndexes: [0] },
              range: { row: 0, col: 0 },
              text: 'EDITED',
            },
          ],
        },
      })
      expect(text.isError, text.output).not.toBe(true)
      expect((await caller.save({ path, overwrite: true })).ok).toBe(true)
      const zip = await JSZip.loadAsync(await readFile(path))
      const xml = new DOMParser().parseFromString(
        await zip.file('word/document.xml')!.async('string'),
        'application/xml',
      )
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
      expect(xml.documentElement.textContent).toContain('EDITED')
      expect(xml.getElementsByTagNameNS(ns, 'shd')).toHaveLength(0)
      expect(ctx.doc!.parsed.blocks[0].table!.rows[0][0].fill).toBe('0000FF')
    } finally {
      runtime.dispose()
    }
  })
  it.each(['fill', 'orphan', 'table'] as const)(
    'preserves source properties and hidden content across save, undo and redo (%s)',
    async (kind) => {
      const row = (hidden: string) =>
        '<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>' +
        paragraph(hidden) +
        '</w:tc><w:tc>' +
        paragraph('SAME') +
        '</w:tc></w:tr>'
      const table = (text: string, band: number) =>
        `<w:tbl><w:tblPr><w:tblStyleRowBandSize w:val="${band}"/></w:tblPr>` +
        '<w:tblGrid><w:gridCol w:w="1255"/><w:gridCol w:w="974"/></w:tblGrid>' +
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="1255" w:type="dxa"/></w:tcPr>' +
        paragraph(text) +
        '</w:tc><w:tc><w:tcPr><w:tcW w:w="37" w:type="pct"/>' +
        '<w:shd w:fill="D9E2F3" w:themeFill="accent1"/></w:tcPr>' +
        paragraph('NEIGHBOR') +
        '</w:tc></w:tr></w:tbl>'
      const source =
        kind === 'orphan'
          ? '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
            row('HIDDEN_A') +
            row('HIDDEN_B') +
            '</w:tbl><w:p/>'
          : table('TABLE_A', 3) + (kind === 'table' ? table('TABLE_B', 7) : '') + '<w:p/>'
      const { editor, ctx, path, bytes } = await fixture(false, source)
      const runtime = createDocsRuntime({
        current: () => ({
          editor,
          identity: editor,
          numIds: { bullet: null, ordered: null },
          autoAcceptUntracked: true,
        }),
        save: async () => {
          const ok = await save(ctx, false, true)
          return { ok, path, committed: ok, pending: !ok }
        },
      })
      try {
        const caller = runtime.connect({ id: 'save-history-source' })
        caller.context()
        const query = await caller.execute({
          id: 'query',
          name: 'query_document',
          input: { type: 'docTable' },
        })
        const structures = (
          JSON.parse(query.output) as { items: Array<{ structures: Array<{ id: string }> }> }
        ).items[0].structures.map((item) => item.id)
        const edit = await caller.execute({
          id: 'edit',
          name: 'apply_ops',
          input: {
            ops: [
              kind === 'fill'
                ? {
                    op: 'setTableCellFormat',
                    target: { blockIndexes: [0] },
                    range: { row: 0, col: 0 },
                    fill: '#FF0000',
                  }
                : kind === 'orphan'
                  ? { op: 'deleteTableRow', target: { blockIndexes: [0] }, row: 0 }
                  : {
                      op: 'deleteBlocks',
                      target: { blockIndexes: [0] },
                      deleteStructures: structures,
                    },
            ],
          },
        })
        expect(edit.isError, edit.output).not.toBe(true)
        const saveXml = async () => {
          expect((await caller.save({ path, overwrite: true })).ok).toBe(true)
          caller.context({ captureSelection: false })
          const zip = await JSZip.loadAsync(await readFile(path))
          return new DOMParser().parseFromString(
            await zip.file('word/document.xml')!.async('string'),
            'application/xml',
          )
        }
        const edited = await saveXml()
        const undo = await caller.execute({
          id: 'undo',
          name: 'document_history',
          input: { action: 'undo' },
        })
        expect(undo.isError, undo.output).not.toBe(true)
        const restored = await saveXml()
        const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        const groups = (xml: Document, name: string) =>
          [...xml.getElementsByTagNameNS(ns, name)].map((node) => node.outerHTML)
        const initialZip = await JSZip.loadAsync(bytes)
        const initial = new DOMParser().parseFromString(
          await initialZip.file('word/document.xml')!.async('string'),
          'application/xml',
        )
        expect(restored.documentElement.textContent).toBe(initial.documentElement.textContent)
        for (const name of ['tblPr', 'tblGrid', 'tcPr'])
          expect(groups(restored, name), name).toEqual(groups(initial, name))
        const redo = await caller.execute({
          id: 'redo',
          name: 'document_history',
          input: { action: 'redo' },
        })
        expect(redo.isError, redo.output).not.toBe(true)
        const repeated = await saveXml()
        expect(repeated.documentElement.textContent).toBe(edited.documentElement.textContent)
        for (const name of ['tblPr', 'tblGrid', 'tcPr'])
          expect(groups(repeated, name), name).toEqual(groups(edited, name))
      } finally {
        runtime.dispose()
      }
    },
  )
  it('keeps a captured character selection after save rebases the document anchors', async () => {
    const { editor, ctx } = await fixture(false)
    let start = 0
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text?.startsWith('Respond')) start = pos
    })
    editor.commands.setTextSelection({ from: start, to: start + 'Respond'.length })
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      save: async ({ path }) => {
        const ok = await save(ctx, false, true)
        return { ok, path, committed: ok, pending: !ok }
      },
    })
    try {
      const scope = runtime.capture()
      const client = scope.connect({ id: 'captured-edit' })
      expect(
        (
          await client.execute({
            id: 'insert',
            name: 'insert_content',
            input: { afterBlockIndex: -1, html: '<p>Inserted before</p>' },
          })
        ).isError,
      ).not.toBe(true)
      expect((await client.save({ path: ctx.doc!.filePath!, overwrite: true })).ok).toBe(true)
      expect(
        (
          await client.execute({
            id: 'undo-insert',
            name: 'document_history',
            input: { action: 'undo' },
          })
        ).isError,
      ).not.toBe(true)
      expect(editor.getText()).not.toContain('Inserted before')
      expect(
        (
          await client.execute({
            id: 'redo-insert',
            name: 'document_history',
            input: { action: 'redo' },
          })
        ).isError,
      ).not.toBe(true)
      client.context({ captureSelection: false })
      expect(
        (
          await client.execute({
            id: 'replace',
            name: 'replace_selection',
            input: { html: 'Confirmed' },
          })
        ).isError,
      ).not.toBe(true)
      expect(editor.state.doc.childCount).toBe(5)
      expect(editor.getText()).toContain('Confirmed within forty-eight (48) hours.')
      expect(editor.getText()).toContain('Inserted before')
      expect(editor.getText()).toContain('Unchanged opening paragraph.')
      expect(editor.getText()).toContain('Unchanged closing paragraph.')
      expect((await client.save({ path: ctx.doc!.filePath!, overwrite: true })).ok).toBe(true)
      const reparsed = await parseDocx(await readFile(ctx.doc!.filePath!))
      expect(reparsed.blocks.filter((block) => block.type !== 'passthrough')).toHaveLength(5)
    } finally {
      runtime.dispose()
    }
  })
  it.each([false, true])(
    'keeps the live model aligned after an agent edit (auto=%s)',
    async (auto) => {
      const { editor, ctx, path, bytes } = await fixture(true)
      await replace(editor)
      expect(editor.state.doc.childCount).toBe(4)
      expect(collectRevisions(editor.state.doc)).toHaveLength(2)
      for (let pass = 0; pass < 2; pass++) {
        expect(await save(ctx, false, auto)).toBe(true)
        const disk = await readFile(path)
        const parsed = await parseDocx(disk)
        expect(
          parsed.blocks.filter((block) => block.type !== 'passthrough'),
          `disk save ${pass + 1}`,
        ).toHaveLength(4)
        expect(editor.state.doc.childCount, `live save ${pass + 1}`).toBe(4)
        expect(collectRevisions(editor.state.doc)).toHaveLength(2)
        expect(editor.storage.trackChanges.enabled).toBe(true)
        await assertUntouched(bytes, disk)
      }
    },
  )

  it('keeps tracking-off edits clean through the same production save', async () => {
    const { editor, ctx, path, bytes } = await fixture(false)
    // No AI tracking override: this control follows the currently disabled recorder.
    const result = await executeTool(
      editor,
      {
        id: 'off-edit',
        name: 'apply_ops',
        input: {
          ops: [
            {
              op: 'findReplace',
              find: 'forty-eight (48) hours',
              replace: 'seventy-two (72) hours',
            },
          ],
        },
      },
      { bullet: null, ordered: null },
    )
    expect(result.isError).toBeFalsy()
    expect(await save(ctx, false, true)).toBe(true)
    expect(editor.state.doc.childCount).toBe(4)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    await assertUntouched(bytes, await readFile(path))
  })

  it('keeps a manually typed tracked insertion through production save', async () => {
    const { editor, ctx } = await fixture(true)
    editor.view.dispatch(editor.state.tr.insertText('Typed ', 1))
    const count = editor.state.doc.childCount
    const revisions = collectRevisions(editor.state.doc).length
    expect(revisions).toBe(1)
    expect(await save(ctx, false, true)).toBe(true)
    expect(editor.state.doc.childCount).toBe(count)
    expect(collectRevisions(editor.state.doc)).toHaveLength(revisions)
  })

  it('keeps an untouched document clean while tracking is enabled', async () => {
    const { editor, ctx } = await fixture(true)
    expect(await save(ctx, false, true)).toBe(true)
    expect(editor.state.doc.childCount).toBe(4)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
  })

  it('refuses serialization of an unsettled draft for recovery or export', async () => {
    const { editor, ctx } = await fixture(false)
    const draft = new DraftLanding(
      editor,
      { bullet: null, ordered: null },
      { kind: 'after', index: 0 },
    )
    try {
      draft.update('<p>Unaccepted preview</p>')
      expect(await buildDocBytes(ctx)).toBeNull()
      expect(
        await Promise.all([
          exportHtml(ctx),
          exportMarkdown(ctx),
          exportPdf(ctx),
          exportImages(ctx),
        ]),
      ).toEqual([false, false, false, false])
    } finally {
      draft.finish()
    }
    expect(await buildDocBytes(ctx)).not.toBeNull()
  })

  it('passes the opened source identity to an explicit save target', async () => {
    const { ctx, path } = await fixture(false)
    const target = join(dirs[dirs.length - 1], 'explicit.docx')
    let sourcePath: string | null | undefined
    window.desktop = {
      saveDocxTo: async (
        output: string,
        bytes: ArrayBuffer,
        _overwrite: boolean,
        source?: string | null,
      ) => {
        sourcePath = source
        await writeFile(output, new Uint8Array(bytes))
        return { ok: true, path: output }
      },
    } as unknown as typeof window.desktop
    expect(await save(ctx, false, true, undefined, { path: target, overwrite: false })).toBe(true)
    expect(sourcePath).toBe(path)
  })

  it.each([false, true])(
    'preserves current callers across saved store publication (auto=%s)',
    async (auto) => {
      const { ctx, editor, bytes, path } = await fixture(false)
      let stateVersion = {}
      ctx.getDocumentStateVersion = () => stateVersion
      ctx.commitDocumentState = (operation) => {
        operation()
        stateVersion = {}
      }
      const runtime = createDocsRuntime({
        current: () => ({
          editor,
          identity: editor,
          stateVersion,
          numIds: { bullet: null, ordered: null },
        }),
      })
      const caller = runtime.connect({ id: 'save-reader' })
      caller.context()
      try {
        const edit = await caller.execute({
          id: 'before-save',
          name: 'apply_ops',
          input: {
            ops: [
              {
                op: 'findReplace',
                find: 'forty-eight (48) hours',
                replace: 'seventy-two (72) hours',
              },
            ],
          },
        })
        expect(edit.isError, edit.output).not.toBe(true)
        const oldQuery = await caller.execute({
          id: 'old-structures',
          name: 'query_document',
          input: { type: 'docTable' },
        })
        const oldTable = JSON.parse(oldQuery.output).items[0]
        expect(await save(ctx, false, auto)).toBe(true)
        expect(ctx.saveIncompleteRef.current).toBe(false)
        await assertUntouched(bytes, new Uint8Array(await readFile(path)))
        const oldIntent = await caller.execute({
          id: 'expired-deletion',
          name: 'replace_blocks',
          input: {
            startBlockIndex: oldTable.blockIndex,
            endBlockIndex: oldTable.blockIndex,
            html: '<p>Wrong</p>',
            dryRun: true,
            deleteStructures: oldTable.structures.map((item: { id: string }) => item.id),
          },
        })
        expect(oldIntent).toMatchObject({
          isError: true,
          errorCode: 'stale_target',
          committed: false,
        })
        const query = await caller.execute({
          id: 'after-save-query',
          name: 'query_document',
          input: { containsText: 'seventy-two' },
        })
        const item = JSON.parse(query.output).items[0]
        const result = await caller.execute({
          id: 'after-save',
          name: 'apply_ops',
          input: {
            ops: [
              {
                op: 'findReplace',
                target: { engine: item.target },
                find: 'seventy-two (72) hours',
                replace: 'ninety-six (96) hours',
              },
            ],
          },
        })
        expect(result.isError, result.output).not.toBe(true)
        expect(editor.getText()).toContain('ninety-six (96) hours')
      } finally {
        runtime.dispose()
      }
    },
  )

  it.each(['body', 'format', 'section'])(
    'does not carry a caller baseline across host-modified saved %s',
    async (kind) => {
      const { ctx, editor } = await fixture(false)
      let stateVersion = {}
      ctx.getDocumentStateVersion = () => stateVersion
      ctx.commitDocumentState = (operation) => {
        operation()
        stateVersion = {}
      }
      window.desktop.saveDocx = async (_target, data) => {
        const zip = await JSZip.loadAsync(data)
        const xml = await zip.file('word/document.xml')!.async('string')
        zip.file(
          'word/document.xml',
          kind === 'body'
            ? xml.replace('<w:body>', '<w:body><w:p><w:r><w:t>Host addition</w:t></w:r></w:p>')
            : kind === 'format'
              ? xml.replace('<w:body><w:p>', '<w:body><w:p><w:pPr><w:jc w:val="right"/></w:pPr>')
              : xml.replace('w:left="1440"', 'w:left="1800"'),
        )
        const changed = await zip.generateAsync({ type: 'base64' })
        return { ok: true, dataUrl: `data:application/octet-stream;base64,${changed}` }
      }
      const runtime = createDocsRuntime({
        current: () => ({
          editor,
          identity: editor,
          stateVersion,
          numIds: { bullet: null, ordered: null },
        }),
      })
      const caller = runtime.connect({ id: 'host-save-reader' })
      caller.context()
      try {
        expect(await save(ctx, false, true)).toBe(true)
        if (kind === 'body') expect(editor.getText()).toContain('Host addition')
        if (kind === 'format') expect(editor.state.doc.firstChild?.attrs.align).toBe('right')
        await caller.execute({ id: 'after-save-query', name: 'query_document', input: {} })
        const result = await caller.execute({
          id: 'after-save',
          name: 'apply_ops',
          input: {
            ops: [{ op: 'findReplace', find: 'Unchanged opening paragraph.', replace: 'Wrong' }],
          },
        })
        expect(result).toMatchObject({ isError: true, errorCode: 'stale_target', committed: false })
      } finally {
        runtime.dispose()
      }
    },
  )

  it('preserves document-side edits made while the file is being written', async () => {
    const { ctx, editor } = await fixture(false)
    const body = editor.state.doc
    let version = {}
    ctx.getDocumentStateVersion = () => version
    window.desktop.saveDocx = async (target, data) => {
      await writeFile(target, new Uint8Array(data))
      ctx.styleUpserts = {
        LaterStyle: { styleId: 'LaterStyle', name: 'Later style', type: 'paragraph' },
      }
      version = {}
      return { ok: true }
    }
    expect(await save(ctx, false, true)).toBe(true)
    expect(ctx.saveIncompleteRef.current).toBe(true)
    expect(ctx.styleUpserts).toHaveProperty('LaterStyle')
    expect(editor.state.doc).toBe(body)
  })
})
