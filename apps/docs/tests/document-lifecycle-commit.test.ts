import { createElement, useEffect, useMemo, useState } from 'react'
import { flushSync } from 'react-dom'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { inlinePictureDrawingXml, parseDocx } from '@genoffice/docx-engine'
import { buildDocx, TINY_PNG_BASE64 } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { collectRevisions } from '../src/renderer/editor/revisions'
import {
  loadFile,
  newFile,
  saveExplicit,
  type FileActionContext,
} from '../src/renderer/file-actions'
import type { DocState } from '../src/renderer/doc-state'
import { installDocsHost, type DocsEditorServices } from '../src/renderer/host/docs-host'
import { createDocsRuntime, type RuntimeDocument } from '../src/renderer/runtime/docs-runtime'

const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups
    .splice(0)
    .reverse()
    .forEach((cleanup) => cleanup())
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

async function fixture() {
  vi.stubGlobal('CSS', { escape: (text: string) => text })
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null)
  const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Old body</w:t></w:r></w:p>' })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Old body</p>',
    editorProps: { handleScrollToSelection: () => true },
  })
  cleanups.push(() => editor.destroy())
  cleanups.push(
    installDocsHost({
      kind: 'browser',
      services: {
        getRecentFiles: async () => [],
        saveDocxTo: async (path: string) => ({ ok: true, path }),
        getSystemLocale: async () => 'en-US',
        docPasswordIntentRevision: async () => 0,
        discardDocPasswordIntents: async () => {},
      } as unknown as DocsEditorServices,
    }),
  )
  // No layout or focus scheduling is needed for this state-commit test.
  const button = document.createElement('button')
  document.body.append(button)
  button.focus()
  cleanups.push(() => button.remove())

  type Metadata = Record<string, unknown> & { doc: DocState; aiPanelKey: number }
  const initial: Metadata = {
    doc: { parsed, fileName: 'Old.docx', filePath: '/old.docx', hash: 'old' },
    aiPanelKey: 0,
    trackChanges: true,
    titlePg: true,
    titlePgDirty: true,
    evenOddHf: true,
    evenOddHfDirty: true,
    mirrorMargins: true,
    mirrorMarginsDirty: true,
    watermarkStyle: { color: 'FF0000', size: 42 },
    docLoading: true,
    comments: [{ id: 'old-comment', text: 'Old comment' }],
    pageColor: 'FF0000',
    removePersonalInfo: true,
  }
  let committed = initial
  const expected = { ...initial }
  let setMetadata!: React.Dispatch<React.SetStateAction<Metadata>>
  let current: RuntimeDocument | null = null
  function Harness() {
    const [metadata, setState] = useState(initial)
    setMetadata = setState
    committed = metadata
    const identity = useMemo(() => ({}), [metadata.aiPanelKey])
    current = {
      editor,
      identity,
      stateVersion: metadata,
      numIds: { bullet: null, ordered: null },
      track: metadata.trackChanges ? { author: 'Lifecycle test' } : undefined,
      autoAcceptUntracked: true,
    }
    useEffect(() => {
      editor.storage.trackChanges.enabled = metadata.trackChanges === true
    }, [metadata.trackChanges])
    return null
  }
  const root = createRoot(document.createElement('div'))
  flushSync(() => root.render(createElement(Harness)))
  cleanups.push(() => flushSync(() => root.unmount()))
  const base = {
    editor,
    dirtyRef: { current: true },
    saveInFlightRef: { current: false },
    saveIncompleteRef: { current: false },
    commitDocumentState: (operation: () => void) => flushSync(operation),
  }
  const ctx = new Proxy(base, {
    get(target, prop) {
      if (prop in target) return Reflect.get(target, prop)
      if (prop === 'setRecent') return () => {}
      if (typeof prop === 'string' && /^set[A-Z]/.test(prop)) {
        const key = prop[3].toLowerCase() + prop.slice(4)
        return (value: unknown) => {
          const next = typeof value === 'function' ? value(expected[key]) : value
          expected[key] = next
          setMetadata((previous) => ({ ...previous, [key]: next }))
        }
      }
      if (prop === 'onWriteProtectionLoaded') return () => {}
      return committed[prop as string]
    },
  }) as unknown as FileActionContext
  const runtime = createDocsRuntime({ current: () => current })
  cleanups.push(() => runtime.dispose())
  const old = runtime.connect({ id: 'old-document' })
  const oldId = old.context().documentId
  const assertCommitted = () => {
    // Check every setter at the await continuation, before React gets another turn.
    for (const [key, value] of Object.entries(expected))
      expect(committed[key] === value, `${key} must be committed before completion`).toBe(true)
    expect(committed.doc).not.toBe(initial.doc)
    expect(current!.track).toBeUndefined()
    expect(editor.storage.trackChanges.enabled).toBe(false)
    expect(() => old.context()).toThrow(/replaced/)
    expect(runtime.connect({ id: 'new-document' }).context().documentId).not.toBe(oldId)
  }
  return {
    editor,
    ctx,
    runtime,
    assertCommitted,
    metadata: () => committed,
    status: () => expected.status,
  }
}

it('commits opened body, identity and authored stores before its promise resolves', async () => {
  const { editor, ctx, runtime, assertCommitted, metadata, status } = await fixture()
  const bytes = await buildDocx({
    bodyXml: ['A', 'B', 'C', 'D']
      .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
      .join(''),
  })
  expect(
    await loadFile(ctx, {
      path: '/new.docx',
      name: 'New.docx',
      hash: 'new',
      dataUrl: '',
      data: bytes.slice().buffer as ArrayBuffer,
    }),
    String(status()),
  ).toBe('ok')
  expect(editor.getText()).toBe('A\n\nB\n\nC\n\nD')
  assertCommitted()
  expect(metadata().doc).toMatchObject({ fileName: 'New.docx', filePath: '/new.docx', hash: 'new' })
  expect(metadata()).toMatchObject({ comments: [], pageColor: null, trackChanges: false })
  const client = runtime.connect({ id: 'immediate-replacement' })
  client.context()
  expect(
    await client.execute({
      id: 'replace',
      name: 'replace_blocks',
      input: { startBlockIndex: 1, endBlockIndex: 2, html: '<p>Replacement</p>' },
    }),
  ).toMatchObject({ committed: true })
  expect(editor.getText()).toBe('A\n\nReplacement\n\nD')
  expect(editor.state.doc.childCount).toBe(3)
  expect(collectRevisions(editor.state.doc)).toHaveLength(0)
})

it('commits a new blank document and clears prior document metadata before completion', async () => {
  const { editor, ctx, assertCommitted, metadata, status } = await fixture()
  expect(await newFile(ctx), String(status())).toBe(true)
  expect(editor.getText()).toBe('')
  assertCommitted()
  expect(metadata().doc).toMatchObject({ filePath: null, hash: '', isBlank: true })
  expect(metadata()).toMatchObject({
    titlePg: false,
    titlePgDirty: false,
    evenOddHf: false,
    evenOddHfDirty: false,
    mirrorMargins: false,
    mirrorMarginsDirty: false,
    watermarkStyle: null,
    docLoading: false,
    comments: [],
    pageColor: null,
    trackChanges: false,
    removePersonalInfo: false,
  })
})

it('commits the saved baseline before the save promise resolves', async () => {
  const { ctx, metadata, status } = await fixture()
  expect(await newFile(ctx), String(status())).toBe(true)
  const before = metadata().doc
  const result = await saveExplicit(ctx, { path: '/saved.docx', overwrite: true })
  expect(result, String(status())).toMatchObject({ ok: true, committed: true, pending: false })
  expect(metadata().doc.filePath).toBe('/saved.docx')
  expect(metadata().doc.parsed).not.toBe(before.parsed)
  expect(metadata().hfVariantsDirty).toEqual([])
  expect(metadata().sectionHfEdits).toEqual({})
})

it('keeps immediate media references current across completed saves', async () => {
  const { ctx, runtime, metadata, status } = await fixture()
  const bytes = await buildDocx({
    withImage: true,
    bodyXml: `<w:p><w:r>${inlinePictureDrawingXml({ rId: 'rId10', docPrId: 1, widthPx: 20, heightPx: 20 })}</w:r></w:p>`,
  })
  expect(
    await loadFile(ctx, {
      path: '/image.docx',
      name: 'Image.docx',
      hash: 'image',
      dataUrl: '',
      data: bytes.slice().buffer as ArrayBuffer,
    }),
    String(status()),
  ).toBe('ok')
  const caller = runtime.connect({ id: 'immediate-saved-media' })
  caller.context()
  for (let attempt = 0; attempt < 3; attempt++) {
    const path = `/saved-image-${attempt}.docx`
    expect(await saveExplicit(ctx, { path, overwrite: true }), String(status())).toMatchObject({
      ok: true,
      committed: true,
      pending: false,
    })
    expect(metadata().doc.filePath).toBe(path)
    const context = caller.context()
    const mediaRef = /media:[a-z0-9-]+:\d+:0\.0/.exec(context.text)?.[0]
    expect(mediaRef).toBeDefined()
    expect(
      await caller.execute({
        id: `media-${attempt}`,
        name: 'read_document_media',
        input: { mediaRef },
      }),
    ).toMatchObject({
      revision: context.revision,
      committed: false,
      media: { mimeType: 'image/png', data: TINY_PNG_BASE64 },
    })
    expect(caller.context().revision).toBe(context.revision)
  }
})
