import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { BLANK_BULLET_NUM_ID, BLANK_ORDERED_NUM_ID } from '@genoffice/docx-engine'
import { createNativeDocsWorkflows } from '../src/renderer/ai/native-workflows'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import { createSkillsSkill } from '../src/renderer/ai/skills-skill'
import { createDocsResourceTools } from '../src/renderer/runtime/resource-tools'

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
  vi.unstubAllGlobals()
})

async function documentWith(text: string) {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: `<p>${text}</p>`,
  })
  editors.push(editor)
  return editor
}

describe('shared Docs runtime', () => {
  it('keeps the new document generation usable while an old independent file operation settles', async () => {
    const editor = await documentWith('Original')
    let identity: object = editor
    let release!: (value: {
      ok: boolean
      code: number
      stdout: string
      stderr: string
      committed: true
    }) => void
    const runCli = vi.fn(
      () =>
        new Promise<Parameters<typeof release>[0]>((resolve) => {
          release = resolve
        }),
    )
    const resources = createDocsResourceTools({ runCli })
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity, numIds: { bullet: null, ordered: null } }),
      resources: () => resources,
    })
    const old = runtime.connect({ id: 'old-cli' })
    const before = old.context()
    const pending = old.execute({
      id: 'cli',
      name: 'run_cli',
      input: { args: ['create', '--type', 'docx', '--out', 'created.docx'] },
    })
    await vi.waitFor(() => expect(runCli).toHaveBeenCalledOnce())
    identity = {}
    const next = runtime.connect({ id: 'replacement' })
    next.context()
    const edit = next.execute({
      id: 'edit',
      name: 'insert_content',
      input: { html: '<p>New generation edit</p>' },
    })
    const early = await Promise.race([
      edit,
      new Promise<null>((resolve) => setTimeout(() => resolve(null), 150)),
    ])
    release({ ok: true, code: 0, stdout: 'Old private CLI output', stderr: '', committed: true })
    const outcome = await pending
    await edit
    expect(early).toMatchObject({ committed: true })
    expect(outcome).toMatchObject({
      isError: true,
      errorCode: 'canceled',
      committed: true,
      documentId: before.documentId,
    })
    expect(outcome.output).not.toContain('Old private CLI output')
    expect(editor.getText()).toContain('New generation edit')
    runtime.dispose()
  })

  it.each([false, true])(
    'withholds late CLI source output while preserving its independent file effect (%s)',
    async (writesFile) => {
      const editor = await documentWith('Original')
      let release!: (value: {
        ok: boolean
        code: number
        stdout: string
        stderr: string
        json?: { status: string; command: string; output_path: string }
      }) => void
      const runCli = vi.fn(
        () =>
          new Promise<Parameters<typeof release>[0]>((resolve) => {
            release = resolve
          }),
      )
      const resources = createDocsResourceTools({ runCli })
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
        resources: () => resources,
      })
      const caller = runtime.connect({ id: 'cli-caller' })
      caller.context()
      const pending = caller.execute({
        id: 'cli',
        name: 'run_cli',
        input: {
          args: writesFile
            ? ['convert', 'source.docx', '--to', 'md', '--out', 'created.md']
            : ['docs', 'read', 'source.docx'],
        },
      })
      await vi.waitFor(() => expect(runCli).toHaveBeenCalledOnce())
      editor.commands.selectAll()
      editor.commands.setMark('redaction', { label: 'private' })
      release({
        ok: true,
        code: 0,
        stdout: 'Private CLI source',
        stderr: '',
        ...(writesFile
          ? { json: { status: 'ok', command: 'convert', output_path: 'created.md' } }
          : {}),
      })
      const result = await pending
      expect(result).toMatchObject({
        isError: true,
        errorCode: 'permission_denied',
        committed: writesFile,
      })
      expect(result.output).not.toContain('Private CLI source')
      runtime.dispose()
    },
  )

  it.each(['revoke', 'withhold'] as const)(
    'refuses failed source diagnostics after %s',
    async (action) => {
      const editor = await documentWith('Original')
      let reject!: (error: Error) => void
      const read = vi.fn(
        () =>
          new Promise<{ text: string; totalChars: number; offset: number }>((_resolve, fail) => {
            reject = fail
          }),
      )
      let resources: ReturnType<typeof createDocsResourceTools> | undefined =
        createDocsResourceTools({
          knowledge: {
            listFiles: () => [
              {
                path: '/granted/source.txt',
                name: 'source.txt',
                folder: 'granted',
                ext: 'txt',
                sizeBytes: 6,
              },
            ],
            read,
            search: async () => [],
          },
        })
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
        resources: () => resources,
      })
      const caller = runtime.connect({ id: 'reader' })
      caller.context()
      const pending = caller.execute({
        id: 'read',
        name: 'read_knowledge_file',
        input: { index: 0 },
      })
      await vi.waitFor(() => expect(read).toHaveBeenCalledOnce())
      if (action === 'revoke') resources = undefined
      else {
        editor.commands.selectAll()
        editor.commands.setMark('redaction', { label: 'private' })
      }
      reject(new Error('Private source diagnostic'))
      const result = await pending
      expect(result).toMatchObject({
        isError: true,
        committed: false,
        errorCode: 'permission_denied',
      })
      expect(result.output).not.toContain('Private source diagnostic')
      runtime.dispose()
    },
  )

  it.each(['close', 'dispose', 'replace', 'revoke', 'revoke_source', 'withhold'] as const)(
    'refuses a pending source read after %s',
    async (action) => {
      const editor = await documentWith('Original')
      let identity: object = editor
      let release!: (value: { text: string; totalChars: number; offset: number }) => void
      const read = vi.fn(
        () =>
          new Promise<{ text: string; totalChars: number; offset: number }>((resolve) => {
            release = resolve
          }),
      )
      let attachments = [
        { path: '/granted/source.txt', name: 'source.txt', ext: 'txt', sizeBytes: 6 },
      ]
      let resources: ReturnType<typeof createDocsResourceTools> | undefined =
        createDocsResourceTools({
          attachments: {
            list: () => attachments,
            read,
          },
        })
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity, numIds: { bullet: null, ordered: null } }),
        resources: () => resources,
      })
      const caller = runtime.connect({ id: 'reader' })
      caller.context()
      const pending = caller.execute({ id: 'read', name: 'read_attachment', input: { index: 0 } })
      let fresh: ReturnType<typeof caller.execute> | undefined
      await vi.waitFor(() => expect(read).toHaveBeenCalledOnce())
      if (action === 'close') caller.close()
      if (action === 'dispose') runtime.dispose()
      if (action === 'replace') {
        identity = {}
        const next = runtime.connect({ id: 'new-document' })
        next.context()
        fresh = next.execute({
          id: 'new-edit',
          name: 'insert_content',
          input: { html: '<p>New document edit</p>' },
        })
      }
      if (action === 'revoke') resources = undefined
      if (action === 'revoke_source') attachments = []
      if (action === 'withhold') {
        editor.commands.selectAll()
        editor.commands.setMark('redaction', { label: 'private' })
      }
      const early =
        action === 'close' || action === 'dispose' || action === 'replace'
          ? await Promise.race([
              pending,
              new Promise<null>((resolve) => setTimeout(() => resolve(null), 150)),
            ])
          : undefined
      const freshOutcome =
        fresh &&
        (await Promise.race([
          fresh,
          new Promise<null>((resolve) => setTimeout(() => resolve(null), 150)),
        ]))
      release({ text: 'Secret source content', offset: 0, totalChars: 21 })
      const result = await pending
      expect(result).toMatchObject({ isError: true, committed: false })
      expect(result.output).not.toContain('Secret source content')
      if (action === 'revoke_source') expect(result.errorCode).toBe('permission_denied')
      if (early !== undefined) expect(early).not.toBeNull()
      if (fresh) expect(freshOutcome).toMatchObject({ committed: true })
      runtime.dispose()
    },
  )

  it.each(['cancel', 'close', 'replace'] as const)(
    'preserves acknowledged independent creation after %s without changing the current document revision',
    async (action) => {
      const editor = await documentWith('Original')
      let identity: object = editor
      let release!: (value: { ok: boolean; path: string }) => void
      const createDocument = vi.fn(
        () =>
          new Promise<{ ok: boolean; path: string }>((resolve) => {
            release = resolve
          }),
      )
      vi.stubGlobal('desktop', { createDocument })
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity, numIds: { bullet: null, ordered: null } }),
      })
      const client = runtime.connect({ id: 'creator' })
      const before = client.context()
      const controller = new AbortController()
      let settled = false
      const pending = client
        .execute(
          {
            id: 'create',
            name: 'create_document',
            input: { type: 'md', title: 'Created', content: '# Created' },
          },
          controller.signal,
        )
        .then((result) => {
          settled = true
          return result
        })
      await vi.waitFor(() => expect(createDocument).toHaveBeenCalledOnce())
      if (action === 'cancel') controller.abort()
      else if (action === 'close') client.close()
      else {
        identity = {}
        runtime.connect({ id: 'replacement' }).context()
      }
      await Promise.resolve()
      expect(settled).toBe(false)
      release({ ok: true, path: 'Created.md' })
      expect(await pending).toMatchObject({
        committed: true,
        mutated: false,
        documentId: before.documentId,
        revision: before.revision,
      })
      expect(editor.getText()).toBe('Original')
      runtime.dispose()
    },
  )

  it('reports an unknown independent creation effect when the host response is lost', async () => {
    const editor = await documentWith('Original')
    vi.stubGlobal('desktop', {
      createDocument: async () => {
        throw new Error('Host disconnected')
      },
    })
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'creator' })
    client.context()
    expect(
      await client.execute({
        id: 'create',
        name: 'create_document',
        input: { type: 'md', title: 'Created', content: '# Created' },
      }),
    ).toMatchObject({ isError: true, errorCode: 'host_failure', committed: null, revision: 0 })
    runtime.dispose()
  })

  it.each(['search', 'image', 'media'])(
    'keeps the %s CLI inference route inside the native Harness only',
    async (command) => {
      const editor = await documentWith('Original')
      const runCli = vi.fn(async () => ({ ok: true, code: 0, stdout: 'Native result', stderr: '' }))
      vi.stubGlobal('desktop', { runCli })
      const resources = createSkillsSkill()
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
        resources: () => resources,
      })
      const external = runtime.connect({ id: 'external-cli' })
      const native = createDocsSkill(
        () => editor,
        () => ({ bullet: null, ordered: null }),
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        () => runtime,
        () => resources,
      )
      external.context()
      native.buildContext?.()
      expect(
        await external.execute({
          id: 'denied',
          name: 'run_cli',
          input: { args: [command, 'test'] },
        }),
      ).toMatchObject({ isError: true, errorCode: 'unavailable', committed: false })
      expect(runCli).not.toHaveBeenCalled()
      expect(
        (
          await native.executeTool({
            id: 'native',
            name: 'run_cli',
            input: { args: [command, 'test'] },
          })
        ).isError,
      ).not.toBe(true)
      expect(runCli).toHaveBeenCalledTimes(1)
      runtime.dispose()
    },
  )

  it('enforces AI-withheld reads on auxiliary file routes and still exposes projected live reads', async () => {
    const editor = await documentWith('Secret')
    editor.commands.selectAll()
    editor.commands.setMark('redaction', { label: 'private' })
    const runCli = vi.fn()
    vi.stubGlobal('desktop', { runCli })
    const resources = createSkillsSkill()
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      resources: () => resources,
    })
    const client = runtime.connect({ id: 'protected-cli' })
    expect(client.context().text).not.toContain('Secret')
    const caps = await client.execute({
      id: 'capabilities',
      name: 'get_editor_capabilities',
      input: {},
    })
    expect(
      JSON.parse(caps.output).tools.find((tool: { name: string }) => tool.name === 'run_cli'),
    ).toMatchObject({ available: false, errorCode: 'permission_denied' })
    expect(
      await client.execute({
        id: 'raw-read',
        name: 'run_cli',
        input: { args: ['docs', 'read', 'current.docx'] },
      }),
    ).toMatchObject({ isError: true, errorCode: 'permission_denied', committed: false })
    expect(runCli).not.toHaveBeenCalled()
    runtime.dispose()
  })
  it.each(['insert_image', 'insert_picture', 'set_watermark'])(
    'refuses oversized or malformed embedded inputs to %s before decoding or accessing the host',
    async (name) => {
      const editor = await documentWith('Original')
      const fetchImage = vi.fn()
      const set = vi.fn(() => null)
      const decode = vi.fn()
      vi.stubGlobal('desktop', { fetchImage })
      vi.stubGlobal('Image', decode)
      const runtime = createDocsRuntime({
        current: () => ({
          editor,
          identity: editor,
          stateVersion: 0,
          numIds: { bullet: null, ordered: null },
          extras: { watermark: { current: () => null, set } },
        }),
      })
      const client = runtime.connect({ id: 'invalid-image-input' })
      client.context()
      for (const url of [
        'file:///private/image.png',
        'data:image/svg+xml;base64,PHN2Zz4=',
        'data:image/png;base64,iVBORw$=',
        'data:image/png;base64,' + 'A'.repeat(36 * 1024 * 1024),
      ]) {
        const result = await client.execute({
          id: url.slice(0, 30),
          name,
          input: { [name === 'set_watermark' ? 'image' : 'url']: url },
        })
        expect(result).toMatchObject({ isError: true, committed: false })
      }
      expect(editor.getText()).toBe('Original')
      expect(fetchImage).not.toHaveBeenCalled()
      expect(set).not.toHaveBeenCalled()
      expect(decode).not.toHaveBeenCalled()
      runtime.dispose()
    },
  )
  it.each(['native', 'external'] as const)(
    'imports Harness-supplied image bytes in %s mode without any provider or file access',
    async (mode) => {
      const editor = await documentWith('Original')
      const fetchImage = vi.fn(() => {
        throw new Error('No host file or network access is allowed')
      })
      vi.stubGlobal('desktop', { fetchImage })
      vi.stubGlobal(
        'Image',
        class {
          onload: (() => void) | null = null
          naturalWidth = 100
          naturalHeight = 80
          set src(_value: string) {
            queueMicrotask(() => this.onload?.())
          }
        },
      )
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      })
      const client = runtime.connect({ id: mode })
      client.context()
      const base64 =
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII='
      const result = await client.execute({
        id: 'image',
        name: 'insert_image',
        input: { url: `data:image/png;base64,${base64}` },
      })
      expect(result.isError, result.output).not.toBe(true)
      const image = editor.state.doc.child(0)
      expect(image.type.name).toBe('docProtected')
      expect(image.attrs.genImage).toMatchObject({
        base64,
        mime: 'image/png',
        widthPx: 100,
        heightPx: 80,
      })
      expect(fetchImage).not.toHaveBeenCalled()
      runtime.dispose()
    },
  )
  it.each(['web_search', 'image_search', 'analyze_media'])(
    'keeps native %s available when the document is read-only',
    async (name) => {
      const editor = await documentWith('Original')
      editor.setEditable(false)
      vi.stubGlobal('desktop', {
        webSearch: async () => ({ results: [], method: 'fixture' }),
        imageSearch: async () => ({ images: [], method: 'fixture' }),
        analyzeMedia: async () => ({ text: 'Fixture analysis' }),
      })
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      })
      const client = runtime.connect({ id: 'read-only-native' })
      const native = createNativeDocsWorkflows(client)
      client.context()
      const input =
        name === 'analyze_media'
          ? { requirements: 'Describe', mediaUrls: ['https://example.com/image.png'] }
          : { query: 'Office' }
      const result = await native.execute({ id: name, name, input })
      expect(result.isError, result.output).not.toBe(true)
      expect(result).toMatchObject({ committed: false, mutated: false })
      expect(editor.getText()).toBe('Original')
      runtime.dispose()
    },
  )
  it('uses the native provider availability predicate for discovery and direct execution', async () => {
    const editor = await documentWith('Original')
    const generate = vi.fn()
    vi.stubGlobal('desktop', { aiGenerateImage: generate })
    const skill = createDocsSkill(
      () => editor,
      () => ({ bullet: null, ordered: null }),
      undefined,
      undefined,
      undefined,
      () => false,
    )
    expect(skill.tools.some((tool) => tool.name === 'generate_image')).toBe(false)
    expect(
      await skill.executeTool({
        id: 'generate',
        name: 'generate_image',
        input: { prompt: 'A flower' },
      }),
    ).toMatchObject({ isError: true, errorCode: 'unavailable' })
    expect(generate).not.toHaveBeenCalled()
  })
  it('reports a written snapshot with newer pending changes as an incomplete save', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      save: async ({ path }) => ({
        ok: false,
        path,
        committed: true,
        pending: true,
        errorCode: 'stale_target',
        error: 'Newer document changes remain unsaved',
      }),
    })
    const router = createDocsCommandRouter(runtime)
    await expect(
      router.handle({
        requestId: 'save',
        command: 'save_document',
        callerId: 'save-client',
        payload: { path: 'snapshot.docx', overwrite: false },
      }),
    ).rejects.toMatchObject({
      result: {
        isError: true,
        committed: true,
        pending: true,
        errorCode: 'stale_target',
        revision: 0,
        documentId: expect.any(String),
      },
    })
    router.dispose()
    runtime.dispose()
  })

  it('reports the file effect when a document is replaced during a save', async () => {
    const editor = await documentWith('Original')
    let identity = {}
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity, numIds: { bullet: null, ordered: null } }),
      save: async ({ path }) => {
        identity = {}
        return { ok: true, path, committed: true, pending: false }
      },
    })
    const client = runtime.connect({ id: 'save-client' })
    const before = client.context()
    expect(await client.save({ path: 'snapshot.docx', overwrite: false })).toMatchObject({
      ok: false,
      documentId: before.documentId,
      committed: true,
      pending: true,
      errorCode: 'stale_target',
    })
    runtime.dispose()
  })

  it('reports an unknown file effect when the host loses its save response', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      save: async () => {
        throw new Error('Host disconnected')
      },
    })
    const client = runtime.connect({ id: 'save-client' })
    expect(await client.save({ path: 'snapshot.docx', overwrite: false })).toMatchObject({
      ok: false,
      committed: null,
      pending: true,
      errorCode: 'host_failure',
      error: 'Host disconnected',
    })
    runtime.dispose()
  })
  it('does not commit a native writer draft after the document becomes read-only', async () => {
    const editor = await documentWith('Original')
    let release!: () => void
    const resume = new Promise<void>((resolve) => {
      release = resolve
    })
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'writer-permission' })
    const native = createNativeDocsWorkflows(client, () => ({
      write: async (_spec, progress) => {
        progress.onProgress('<p>Preview</p>')
        await resume
        return { ok: true, html: '<p>Generated content</p>' }
      },
    }))
    client.context()
    const pending = native.execute({
      id: 'write',
      name: 'write_document',
      input: { plan: 'Write content', afterBlockIndex: 0 },
    })
    await vi.waitFor(() => expect(editor.getText()).toContain('Preview'))
    editor.setEditable(false)
    release()
    expect(await pending).toMatchObject({
      isError: true,
      errorCode: 'permission_denied',
      committed: false,
    })
    expect(editor.getText()).toBe('Original')
    runtime.dispose()
  })

  it('rejects a late watermark when document-side state changes without a body edit', async () => {
    const editor = await documentWith('Original')
    let stateVersion = {}
    const set = vi.fn(() => null)
    const runtime = createDocsRuntime({
      current: () => ({
        editor,
        identity: editor,
        stateVersion,
        numIds: { bullet: null, ordered: null },
        extras: { watermark: { current: () => null, set } },
      }),
    })
    const client = runtime.connect({ id: 'side-state-writer' })
    client.context()
    let decode!: () => void
    let started!: () => void
    const decoding = new Promise<void>((resolve) => {
      started = resolve
    })
    vi.stubGlobal(
      'Image',
      class {
        onload: (() => void) | null = null
        naturalWidth = 100
        naturalHeight = 80
        set src(_value: string) {
          decode = () => this.onload?.()
          started()
        }
      },
    )
    const pending = client.execute({
      id: 'watermark',
      name: 'set_watermark',
      input: { image: 'data:image/png;base64,iVBORw0KGgoAAAAA' },
    })
    await decoding
    stateVersion = {}
    decode()
    expect(await pending).toMatchObject({
      isError: true,
      errorCode: 'stale_target',
      committed: false,
    })
    expect(set).not.toHaveBeenCalled()
    runtime.dispose()
  })

  it('keeps an asynchronously inserted image at the action caret when the user moves elsewhere', async () => {
    const editor = await documentWith('First</p><p>Second')
    editor.commands.setTextSelection(1)
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'image-position' })
    client.context()
    let release!: (value: { base64: string }) => void
    let started!: () => void
    const downloading = new Promise<void>((resolve) => {
      started = resolve
    })
    vi.stubGlobal('desktop', {
      fetchImage: () => {
        started()
        return new Promise<{ base64: string }>((resolve) => {
          release = resolve
        })
      },
    })
    vi.stubGlobal(
      'Image',
      class {
        onload: (() => void) | null = null
        naturalWidth = 100
        naturalHeight = 80
        set src(_value: string) {
          queueMicrotask(() => this.onload?.())
        }
      },
    )
    const pending = client.execute({
      id: 'image',
      name: 'insert_image',
      input: { url: 'https://example.com/image.png' },
    })
    await downloading
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    release({ base64: 'iVBORw0KGgoAAAAA' })
    const result = await pending
    expect(result.isError, result.output).not.toBe(true)
    expect(editor.state.doc.child(0).type.name).toBe('docProtected')
    runtime.dispose()
  })

  it('returns document media bytes for Harness analysis without exposing its internal resource URL', async () => {
    const editor = await documentWith('Original')
    const base64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII='
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docProtected',
          attrs: {
            blockType: 'image',
            imageDataUrl: `data:image/png;base64,${base64}`,
          },
        },
      ],
    })
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'media-reader' })
    client.context()
    const result = await client.execute({
      id: 'image',
      name: 'read_document_media',
      input: { blockIndex: 0 },
    })
    expect(result.isError, result.output).not.toBe(true)
    expect(result).toMatchObject({
      committed: false,
      media: { mimeType: 'image/png', data: base64 },
    })
    expect(result.output).not.toContain('data:')
    runtime.dispose()
  })

  it('cancels an owned command without allowing another caller to cancel it', async () => {
    const editor = await documentWith('Original')
    const set = vi.fn(() => null)
    const runtime = createDocsRuntime({
      current: () => ({
        editor,
        identity: editor,
        stateVersion: 0,
        numIds: { bullet: null, ordered: null },
        extras: { watermark: { current: () => null, set } },
      }),
    })
    const router = createDocsCommandRouter(runtime)
    let decode!: () => void
    let started!: () => void
    const decoding = new Promise<void>((resolve) => {
      started = resolve
    })
    vi.stubGlobal(
      'Image',
      class {
        onload: (() => void) | null = null
        naturalWidth = 100
        naturalHeight = 80
        set src(_value: string) {
          decode = () => this.onload?.()
          started()
        }
      },
    )
    const pending = router.handle({
      requestId: 'owned-command',
      callerId: 'owner',
      command: 'execute_tool',
      payload: {
        name: 'set_watermark',
        input: { image: 'data:image/png;base64,iVBORw0KGgoAAAAA' },
      },
    })
    const rejected = expect(pending).rejects.toMatchObject({
      result: { errorCode: 'canceled', committed: false },
    })
    await decoding
    expect(
      await router.handle({
        requestId: 'wrong-cancel',
        callerId: 'other',
        command: 'cancel_operation',
        payload: { requestId: 'owned-command' },
      }),
    ).toEqual({ canceled: false })
    expect(
      await router.handle({
        requestId: 'right-cancel',
        callerId: 'owner',
        command: 'cancel_operation',
        payload: { requestId: 'owned-command' },
      }),
    ).toEqual({ canceled: true })
    await rejected
    decode()
    await Promise.resolve()
    expect(set).not.toHaveBeenCalled()
    router.dispose()
    runtime.dispose()
  })

  it('reports document revisions and distinct runtime error categories without committing rejected input', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'result-reader' })
    const before = client.context()
    const invalid = await client.execute({ id: 'invalid', name: 'replace_blocks', input: {} })
    expect(invalid).toMatchObject({
      documentId: before.documentId,
      committed: false,
      errorCode: 'invalid_input',
      revision: before.revision,
    })
    editor.commands.insertContentAt(1, 'Manual ')
    const stale = await client.execute({
      id: 'stale',
      name: 'insert_content',
      input: { html: '<p>Wrong</p>' },
    })
    expect(stale).toMatchObject({ committed: false, errorCode: 'stale_target' })
    expect(stale.revision).toBeGreaterThan(before.revision)
    client.context()
    const edited = await client.execute({
      id: 'edit',
      name: 'insert_content',
      input: { afterBlockIndex: 0, html: '<p>Agent</p>' },
    })
    expect(edited.committed).toBe(true)
    expect(edited.isError).not.toBe(true)
    expect(edited.revision).toBeGreaterThan(stale.revision)
    runtime.dispose()
  })

  it('rejects a decoded watermark after the bound document is replaced', async () => {
    const editor = await documentWith('Original')
    let identity = {}
    const set = vi.fn(() => null)
    const runtime = createDocsRuntime({
      current: () => ({
        editor,
        identity,
        stateVersion: 0,
        numIds: { bullet: null, ordered: null },
        extras: { watermark: { current: () => null, set } },
      }),
    })
    const client = runtime.connect({ id: 'watermark-writer' })
    client.context()
    let decode!: () => void
    let started!: () => void
    const decoding = new Promise<void>((resolve) => {
      started = resolve
    })
    vi.stubGlobal(
      'Image',
      class {
        onload: (() => void) | null = null
        naturalWidth = 100
        naturalHeight = 80
        set src(_value: string) {
          decode = () => this.onload?.()
          started()
        }
      },
    )
    const pending = client.execute({
      id: 'watermark',
      name: 'set_watermark',
      input: { image: 'data:image/png;base64,iVBORw0KGgoAAAAA' },
    })
    await decoding
    identity = {}
    decode()
    expect((await pending).isError).toBe(true)
    expect(set).not.toHaveBeenCalled()
    runtime.dispose()
  })

  it('carries only current caller baselines across trusted loading appends', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const stale = runtime.connect({ id: 'stale-reader' })
    stale.context()
    editor.commands.insertContentAt(1, 'Manual ')
    const current = runtime.connect({ id: 'current-reader' })
    current.context()
    const { appendStreamedNodes } = await import('../src/renderer/file-actions')
    appendStreamedNodes(editor, [
      { type: 'docParagraph', content: [{ type: 'text', text: 'Loaded tail' }] },
    ])
    const edit = {
      id: 'insert',
      name: 'insert_content',
      input: { afterBlockIndex: 0, html: '<p>Agent content</p>' },
    }
    expect((await stale.execute(edit)).isError).toBe(true)
    const result = await current.execute(edit)
    expect(result.isError, result.output).not.toBe(true)
    expect(editor.getText()).toBe('Manual Original\n\nAgent content\n\nLoaded tail')
    runtime.dispose()
  })

  it('reports callable capabilities and operation guidance without enabling Harness inference inside the editor', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'capabilities-reader' })
    const result = await client.execute({
      id: 'capabilities',
      name: 'get_editor_capabilities',
      input: {},
    })
    expect(result.isError, result.output).not.toBe(true)
    const capabilities = JSON.parse(result.output)
    expect(
      capabilities.operations.some(
        (operation: { name: string }) => operation.name === 'findReplace',
      ),
    ).toBe(true)
    expect(capabilities.htmlRules).toContain('<p>')
    expect(
      capabilities.tools.find((tool: { name: string }) => tool.name === 'write_document'),
    ).toBeUndefined()
    expect(
      capabilities.tools.find((tool: { name: string }) => tool.name === 'read_comments'),
    ).toMatchObject({ available: false })
    expect(
      capabilities.tools.find((tool: { name: string }) => tool.name === 'replace_blocks'),
    ).toMatchObject({ available: true })
    expect(capabilities).not.toHaveProperty('harnessWorkflows')
    runtime.dispose()
  })

  it.each(['accept_changes', 'reject_changes'])(
    'allows freshly read %s without granting stale block edits or document-side writes',
    async (name) => {
      const editor = await documentWith('Original')
      let stateVersion = {}
      const runtime = createDocsRuntime({
        current: () => ({
          editor,
          identity: editor,
          stateVersion,
          numIds: { bullet: null, ordered: null },
        }),
      })
      const client = runtime.connect({ id: 'revision-reviewer' })
      client.context()
      const { executeTool } = await import('../src/renderer/ai/tools')
      const tracked = await executeTool(
        editor,
        {
          id: 'other-edit',
          name: 'replace_blocks',
          input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Changed</p>' },
        },
        { bullet: null, ordered: null },
        { author: 'Other author' },
      )
      expect(tracked.isError).not.toBe(true)
      stateVersion = {}
      const stale = await client.execute({ id: 'stale-decision', name, input: { all: true } })
      expect(stale).toMatchObject({ isError: true, errorCode: 'stale_target', committed: false })
      expect(stale.output).toBe(
        'Document revisions or document-side state changed since this caller read them; use read_revisions before accepting or rejecting changes, or read_document or get_document_context for fresh full context.',
      )
      expect(
        (await client.execute({ id: 'read', name: 'read_revisions', input: {} })).isError,
      ).not.toBe(true)
      const accepted = await client.execute({
        id: 'accept',
        name,
        input: { all: true },
      })
      expect(accepted.isError, accepted.output).not.toBe(true)
      expect(editor.getText()).toBe(name === 'accept_changes' ? 'Changed' : 'Original')
      expect(
        (
          await client.execute({
            id: 'stale-edit',
            name: 'insert_content',
            input: { afterBlockIndex: 0, html: '<p>Wrong</p>' },
          })
        ).isError,
      ).toBe(true)
      runtime.dispose()
    },
  )

  it('does not treat a revisions-only read as fresh block context', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'revisions-reader' })
    client.context()
    editor.commands.insertContentAt(1, 'Manual ')
    expect(
      (await client.execute({ id: 'read', name: 'read_revisions', input: {} })).isError,
    ).not.toBe(true)
    const edit = await client.execute({
      id: 'replace',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Wrong replacement</p>' },
    })
    expect(edit.isError).toBe(true)
    expect(editor.getText()).toBe('Manual Original')
    runtime.dispose()
  })

  it('rejects out-of-range insertion indexes instead of changing the requested target', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'insert-caller' })
    client.context()
    for (const afterBlockIndex of [-2, 1, 99]) {
      expect(
        (
          await client.execute({
            id: 'insert',
            name: 'insert_content',
            input: { afterBlockIndex, html: '<p>Unexpected</p>' },
          })
        ).isError,
      ).toBe(true)
      expect(editor.getText()).toBe('Original')
    }
    runtime.dispose()
  })

  it('binds a later Harness connection to the scope captured by the editor action', async () => {
    const editor = await documentWith('One Two')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    editor.commands.setTextSelection({ from: 1, to: 4 })
    const scope = runtime.capture()
    editor.commands.setTextSelection({ from: 5, to: 8 })
    const harness = scope.connect({ id: 'harness-connection' })
    expect(harness.context({ captureSelection: false }).text).toContain('One')
    const replaced = await harness.execute({
      id: 'replace',
      name: 'replace_selection',
      input: { html: 'New' },
    })
    expect(replaced.isError).not.toBe(true)
    expect(editor.getText()).toBe('New Two')
    scope.close()
    runtime.dispose()
  })

  it.each(['close', 'replace'] as const)(
    'settles an uncooperative writer after %s without blocking the document queue',
    async (action) => {
      const editor = await documentWith('Original')
      let identity: object = editor
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity, numIds: { bullet: null, ordered: null } }),
      })
      let release!: () => void
      const resume = new Promise<void>((resolve) => {
        release = resolve
      })
      const client = runtime.connect({ id: 'closing' })
      const native = createNativeDocsWorkflows(client, () => ({
        write: async (_spec, progress) => {
          progress.onProgress('<p>Preview</p>')
          await resume
          return { ok: true, html: '<p>Late output</p>' }
        },
      }))
      client.context()
      const pending = native.execute({
        id: 'write',
        name: 'write_document',
        input: { plan: 'Add content', afterBlockIndex: 0 },
      })
      await vi.waitFor(() => expect(editor.getText()).toContain('Preview'))
      let fresh: ReturnType<ReturnType<typeof runtime.connect>['execute']> | undefined
      if (action === 'close') client.close()
      else {
        editor.commands.setContent('<p>Replacement</p>')
        identity = {}
        const next = runtime.connect({ id: 'replacement' })
        next.context()
        fresh = next.execute({
          id: 'insert-new',
          name: 'insert_content',
          input: { html: '<p>New edit</p>', afterBlockIndex: 0 },
        })
      }
      const outcome = await Promise.race([
        pending,
        new Promise<null>((resolve) => setTimeout(() => resolve(null), 150)),
      ])
      const freshOutcome =
        fresh &&
        (await Promise.race([
          fresh,
          new Promise<null>((resolve) => setTimeout(() => resolve(null), 150)),
        ]))
      release()
      await pending
      expect(outcome?.isError).toBe(true)
      if (action === 'replace') {
        expect(freshOutcome).toMatchObject({ committed: true })
        expect(editor.getText()).toContain('Replacement')
        expect(editor.getText()).toContain('New edit')
      } else expect(editor.getText()).toBe('Original')
      expect(editor.getText()).not.toContain('Late output')
      runtime.dispose()
    },
  )

  it('keeps text manually edited inside the preview when the rest of a draft is discarded', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'manual-draft' })
    client.context()
    const started = await client.execute({
      id: 'start',
      name: 'start_document_draft',
      input: { afterBlockIndex: 0 },
    })
    const { draftId } = JSON.parse(started.output)
    await client.execute({
      id: 'update',
      name: 'update_document_draft',
      input: { draftId, sequence: 1, html: '<p>Preview</p><p>Untouched preview</p>' },
    })
    await vi.waitFor(() => expect(editor.getText()).toContain('Untouched preview'))
    editor.view.dispatch(
      editor.state.tr.insertText(
        ' manual addition',
        editor.state.doc.child(0).nodeSize + 1 + 'Preview'.length,
      ),
    )
    const discarded = await client.execute({
      id: 'discard',
      name: 'finish_document_draft',
      input: { draftId, decision: 'discard' },
    })
    expect(discarded.isError).not.toBe(true)
    expect(editor.getText()).toBe('Original\n\nPreview manual addition')
    runtime.dispose()
  })

  it('blocks saving an unaccepted draft and lets the Harness retain stopped partial content', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      save: async ({ path }) => ({ ok: true, path, committed: true, pending: false }),
    })
    const client = runtime.connect({ id: 'partial' })
    client.context()
    const started = await client.execute({
      id: 'start',
      name: 'start_document_draft',
      input: { afterBlockIndex: 0 },
    })
    const { draftId } = JSON.parse(started.output)
    await client.execute({
      id: 'update',
      name: 'update_document_draft',
      input: { draftId, sequence: 1, html: '<p>Incomplete section</p>' },
    })
    await vi.waitFor(() => expect(editor.getText()).toContain('Incomplete section'))
    expect((await client.save({ path: 'partial.docx', overwrite: false })).ok).toBe(false)
    await client.execute({ id: 'stop', name: 'stop_document_draft', input: { draftId } })
    const late = await client.execute({
      id: 'late',
      name: 'update_document_draft',
      input: { draftId, sequence: 2, html: '<p>Late content</p>' },
    })
    expect(late.isError).toBe(true)
    const retained = await client.execute({
      id: 'retain',
      name: 'finish_document_draft',
      input: { draftId, decision: 'retain_partial' },
    })
    expect(JSON.parse(retained.output).state).toBe('retained')
    expect(editor.getText()).toBe('Original\n\nIncomplete section')
    expect((await client.save({ path: 'partial.docx', overwrite: false })).ok).toBe(true)
    runtime.dispose()
  })

  it('preserves manual changes when a native whole-document writer finishes with a stale rewrite', async () => {
    const editor = await documentWith('Original content')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    let release!: () => void
    const resume = new Promise<void>((resolve) => {
      release = resolve
    })
    const client = runtime.connect({ id: 'native-writer' })
    const native = createNativeDocsWorkflows(client, () => ({
      write: async (_spec, progress) => {
        progress.onProgress('<p>Preview</p>')
        await resume
        return { ok: true, html: '<p>Generated replacement</p>' }
      },
    }))
    client.context()
    const pending = native.execute({
      id: 'write',
      name: 'write_document',
      input: {
        plan: 'Rewrite the document',
        replaceDocument: true,
      },
    })
    await vi.waitFor(() => expect(editor.getText()).toContain('Preview'))
    editor.view.dispatch(
      editor.state.tr.insertText(' and manual addition', editor.state.doc.content.size - 1),
    )
    release()
    const result = await pending
    expect(result.isError).toBe(true)
    expect(editor.getText()).toBe('Original content and manual addition')
    runtime.dispose()
  })

  it('shows an external draft, commits it once, and makes the complete draft one undo step', async () => {
    const editor = await documentWith('Keep this')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'draft-writer' })
    client.context()
    const started = await client.execute({
      id: 'start',
      name: 'start_document_draft',
      input: { afterBlockIndex: 0 },
    })
    expect(started.isError).not.toBe(true)
    const { draftId } = JSON.parse(started.output)
    const updated = await client.execute({
      id: 'update',
      name: 'update_document_draft',
      input: {
        draftId,
        sequence: 1,
        html: '<p>First streamed paragraph</p><p>Second streamed paragraph</p>',
      },
    })
    expect(updated.isError).not.toBe(true)
    await vi.waitFor(() => expect(editor.getText()).toContain('First streamed paragraph'))
    const finish = {
      id: 'finish',
      name: 'finish_document_draft',
      input: { draftId, decision: 'finalize' },
    }
    expect((await client.execute(finish)).isError).not.toBe(true)
    expect((await client.execute(finish)).isError).not.toBe(true)
    expect(editor.getText()).toBe(
      'Keep this\n\nFirst streamed paragraph\n\nSecond streamed paragraph',
    )
    await client.execute({ id: 'undo', name: 'document_history', input: { action: 'undo' } })
    expect(editor.getText()).toBe('Keep this')
    runtime.dispose()
  })

  it('undoes and redoes a single tool edit without losing the preceding tool result', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'history' })
    client.context()
    for (const text of ['First result', 'Second result']) {
      expect(
        (
          await client.execute({
            id: text,
            name: 'replace_blocks',
            input: {
              startBlockIndex: 0,
              endBlockIndex: 0,
              html: `<p>${text}</p>`,
            },
          })
        ).isError,
      ).not.toBe(true)
    }
    expect(
      (await client.execute({ id: 'undo', name: 'document_history', input: { action: 'undo' } }))
        .isError,
    ).not.toBe(true)
    expect(editor.getText()).toBe('First result')
    expect(
      (await client.execute({ id: 'redo', name: 'document_history', input: { action: 'redo' } }))
        .isError,
    ).not.toBe(true)
    expect(editor.getText()).toBe('Second result')
    runtime.dispose()
  })

  it('does not let an old binding refresh itself onto a replacement document', async () => {
    const editor = await documentWith('First document')
    let identity = {}
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity, numIds: { bullet: null, ordered: null } }),
    })
    const old = runtime.connect({ id: 'old-run' })
    old.context()
    identity = {}
    editor.commands.setContent('<p>Replacement document</p>')
    expect(() => old.context()).toThrow(/replaced/)
    const late = await old.execute({
      id: 'late',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>Late result</p>',
      },
    })
    expect(late.isError).toBe(true)
    expect(editor.getText()).toBe('Replacement document')
    const fresh = runtime.connect({ id: 'new-run' })
    expect(fresh.context().text).toContain('Replacement document')
    runtime.dispose()
  })

  it('validates inputs against the shared schema before a direct caller can edit', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'direct' })
    client.context()
    const result = await client.execute({
      id: 'invalid',
      name: 'insert_content',
      input: { html: 123 },
    })
    expect(result.isError).toBe(true)
    expect(editor.getText()).toBe('Original')
    runtime.dispose()
  })

  it.each(['native', 'external'] as const)(
    'enforces a read-only document for a %s caller',
    async (mode) => {
      const editor = await documentWith('Protected content')
      editor.setEditable(false)
      const runtime = createDocsRuntime({
        current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
      })
      const client = runtime.connect({ id: mode })
      expect(client.context().text).toContain('Protected content')
      const result = await client.execute({
        id: 'edit',
        name: 'replace_blocks',
        input: {
          startBlockIndex: 0,
          endBlockIndex: 0,
          html: '<p>Overwrite</p>',
        },
      })
      expect(result.isError).toBe(true)
      expect(editor.getText()).toBe('Protected content')
      runtime.dispose()
    },
  )

  it('keeps the captured selection across its own formatting edit when the user moves the visible selection', async () => {
    const editor = await documentWith('One Two')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const client = runtime.connect({ id: 'selected-writer' })
    editor.commands.setTextSelection({ from: 1, to: 4 })
    client.context()
    editor.commands.setTextSelection({ from: 5, to: 8 })
    const formatted = await client.execute({
      id: 'format',
      name: 'apply_ops',
      input: {
        ops: [{ op: 'setFont', target: { scope: 'selection' }, bold: true }],
      },
    })
    expect(formatted.isError).not.toBe(true)
    const replaced = await client.execute({
      id: 'replace',
      name: 'replace_selection',
      input: { html: 'New' },
    })
    expect(replaced.isError).not.toBe(true)
    expect(editor.getText()).toBe('New Two')
    runtime.dispose()
  })

  it('refuses a pending picture after a manual edit even when another caller reads the new content', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const caller = runtime.connect({ id: 'image-writer' })
    const reader = runtime.connect({ id: 'reader' })
    caller.context()
    let release!: (value: { base64: string }) => void
    let started!: () => void
    const downloading = new Promise<void>((resolve) => {
      started = resolve
    })
    vi.stubGlobal('desktop', {
      fetchImage: () => {
        started()
        return new Promise<{ base64: string }>((resolve) => {
          release = resolve
        })
      },
    })
    vi.stubGlobal(
      'Image',
      class {
        onload: (() => void) | null = null
        naturalWidth = 100
        naturalHeight = 80
        set src(_value: string) {
          queueMicrotask(() => this.onload?.())
        }
      },
    )
    const pending = caller.execute({
      id: 'picture',
      name: 'insert_picture',
      input: {
        url: 'https://example.com/picture.png',
        afterBlockIndex: 0,
      },
    })
    await downloading
    editor.commands.setContent('<p>Manual change</p>')
    reader.context()
    release({ base64: 'iVBORw0KGgoAAAAA' })
    expect((await pending).isError).toBe(true)
    expect(editor.getText()).toBe('Manual change')
    expect(editor.state.doc.childCount).toBe(1)
    runtime.dispose()
  })

  it('uses the same caller isolation for the native Agent skill and an external caller', async () => {
    const editor = await documentWith('Shared document')
    const numIds = () => ({ bullet: null, ordered: null })
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: numIds() }),
    })
    const native = createDocsSkill(
      () => editor,
      numIds,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      () => runtime,
    )
    native.buildContext!()
    const external = runtime.connect({ id: 'external' })
    external.context()
    await external.execute({
      id: 'external-edit',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>External result</p>',
      },
    })
    external.context()
    const refused = await native.executeTool({
      id: 'native-edit',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>Stale native result</p>',
      },
    })
    expect(refused.isError).toBe(true)
    expect(editor.getText()).toBe('External result')
    runtime.dispose()
  })

  it('lets a stale caller refresh through the public read tool before editing again', async () => {
    const editor = await documentWith('Before typing')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const caller = runtime.connect({ id: 'agent' })
    caller.context()
    editor.commands.setContent('<p>Manual change</p>')
    const read = await caller.execute({ id: 'read', name: 'get_document_context', input: {} })
    expect(read.isError).not.toBe(true)
    expect(read.output).toContain('Manual change')
    const edited = await caller.execute({
      id: 'edit',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>After refresh</p>' },
    })
    expect(edited.isError).not.toBe(true)
    expect(editor.getText()).toBe('After refresh')
    runtime.dispose()
  })

  it('lets a caller recover from body changes by reading blocks', async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
    })
    const caller = runtime.connect({ id: 'body-reader' })
    caller.context()
    editor.commands.setContent('<p>Manual change</p>')
    const call = {
      id: 'edit',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>After block refresh</p>' },
    }
    const stale = await caller.execute(call)
    expect(stale).toMatchObject({ isError: true, errorCode: 'stale_target', committed: false })
    expect(stale.output).toBe(
      'The document body changed since this caller read it; use read_blocks, read_document or get_document_context before retrying.',
    )
    const read = await caller.execute({
      id: 'read-blocks',
      name: 'read_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0 },
    })
    expect(read.isError, read.output).not.toBe(true)
    const edited = await caller.execute(call)
    expect(edited.isError, edited.output).not.toBe(true)
    expect(edited.committed).toBe(true)
    expect(editor.getText()).toBe('After block refresh')
    runtime.dispose()
  })

  it('requires full context after document-side state changes even after reading blocks', async () => {
    const editor = await documentWith('Original')
    let stateVersion = {}
    const runtime = createDocsRuntime({
      current: () => ({
        editor,
        identity: editor,
        stateVersion,
        numIds: { bullet: null, ordered: null },
      }),
    })
    const caller = runtime.connect({ id: 'side-state-reader' })
    const before = caller.context()
    stateVersion = {}
    const read = await caller.execute({
      id: 'read-blocks',
      name: 'read_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0 },
    })
    expect(read.isError, read.output).not.toBe(true)
    expect(read.documentId).toBe(before.documentId)
    expect(read.revision).toBeGreaterThan(before.revision)
    for (const call of [
      {
        id: 'replace',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Wrong</p>' },
      },
      {
        id: 'ops',
        name: 'apply_ops',
        input: { ops: [{ op: 'findReplace', find: 'Original', replace: 'Wrong' }] },
      },
    ]) {
      const stale = await caller.execute(call)
      expect(stale).toMatchObject({
        isError: true,
        errorCode: 'stale_target',
        committed: false,
        documentId: before.documentId,
        revision: read.revision,
      })
      expect(stale.output).toBe(
        'Document-side state changed since this caller read it; use read_document or get_document_context for fresh full context. read_blocks does not refresh document-side state.',
      )
    }
    expect(editor.getText()).toBe('Original')
    const refreshed = await caller.execute({
      id: 'full-context',
      name: 'get_document_context',
      input: {},
    })
    expect(refreshed).toMatchObject({ documentId: before.documentId, revision: read.revision })
    expect(refreshed.isError, refreshed.output).not.toBe(true)
    const edited = await caller.execute({
      id: 'fresh-edit',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>After refresh</p>' },
    })
    expect(edited.isError, edited.output).not.toBe(true)
    expect(edited.documentId).toBe(before.documentId)
    expect(edited.committed).toBe(true)
    expect(editor.getText()).toBe('After refresh')
    runtime.dispose()
  })

  it("keeps each caller's read baseline when another caller edits and reads the document", async () => {
    const editor = await documentWith('Original')
    const runtime = createDocsRuntime({
      current: () => ({
        editor,
        identity: editor,
        numIds: { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID },
      }),
    })
    const first = runtime.connect({ id: 'first' })
    const second = runtime.connect({ id: 'second' })
    first.context()
    second.context()
    const changed = await second.execute({
      id: 'second-edit',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Second writer</p>' },
    })
    expect(changed.isError).not.toBe(true)
    second.context()
    const stale = await first.execute({
      id: 'stale-edit',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Lost update</p>' },
    })
    expect(stale.isError).toBe(true)
    expect(editor.getText()).toBe('Second writer')
    first.context()
    const refreshed = await first.execute({
      id: 'fresh-edit',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>Fresh edit</p>' },
    })
    expect(refreshed.isError).not.toBe(true)
    expect(editor.getText()).toBe('Fresh edit')
    runtime.dispose()
  })
})
