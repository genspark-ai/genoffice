import { afterEach, expect, it, vi } from 'vitest'
import {
  getDocsHost,
  installDocsHost,
  type DocsEditorServices,
} from '../src/renderer/host/docs-host'
import { newFile, save, type FileActionContext } from '../src/renderer/file-actions'
import type { DocsDocumentHost } from '../src/shared/document-host'

let uninstall: (() => void) | undefined
afterEach(() => {
  uninstall?.()
  uninstall = undefined
})
function fixture(documents: DocsDocumentHost) {
  const localOpen = vi.fn()
  const localSave = vi.fn()
  uninstall = installDocsHost({
    kind: 'browser',
    documents,
    services: { openDocx: localOpen, saveDocx: localSave } as unknown as DocsEditorServices,
  })
  return { localOpen, localSave }
}

it('opens pathless bytes and opaque identity without local fallback on cancellation or failure', async () => {
  const open = vi.fn(async () => ({
    id: 'opaque:key',
    name: 'Hosted.docx',
    data: new ArrayBuffer(0),
  }))
  const { localOpen } = fixture({ openDocument: open })
  expect(await getDocsHost().openDocx()).toMatchObject({
    path: '',
    name: 'Hosted.docx',
    hosted: { id: 'opaque:key' },
  })
  open.mockResolvedValueOnce(null as never)
  expect(await getDocsHost().openDocx()).toBeNull()
  open.mockRejectedValueOnce(new Error('Storage offline'))
  await expect(getDocsHost().openDocx()).rejects.toThrow('Storage offline')
  expect(localOpen).not.toHaveBeenCalled()
})

it('does not create blank content or persist locally when the host cancels new or disables save', async () => {
  const { localSave } = fixture({ newDocument: async () => null, capabilities: { save: false } })
  const context = { editor: {} } as FileActionContext
  expect(await newFile(context)).toBe(false)
  expect(await save(context, false)).toBe(false)
  expect(localSave).not.toHaveBeenCalled()
  expect(() => getDocsHost().saveDocx('/local.docx', new ArrayBuffer(0))).toThrow(/hosted/i)
})

it('does not consume standalone startup or AI creation when hosted storage is authoritative', async () => {
  fixture({ capabilities: { open: false, create: false } })
  expect(await getDocsHost().consumePendingOpenDocx()).toBeNull()
  expect(await getDocsHost().consumeNewBlankDoc()).toBe(false)
  expect(await getDocsHost().consumeAiDocContent()).toBeNull()
  await expect(getDocsHost().openDocx()).rejects.toThrow(/disabled/)
  await expect(
    getDocsHost().createDocument({ type: 'docx', title: 'New', content: '<p>x</p>' }),
  ).rejects.toThrow(/disabled/)
})


it('preserves the real file target and saved version supplied by a desktop host', async () => {
  fixture({ openDocument: async () => ({ id: 'work', name: 'Report', path: '/work/document.docx',
    fileHash: 'loaded-version', data: new ArrayBuffer(0) }) })
  expect(await getDocsHost().openDocx()).toMatchObject({ path: '/work/document.docx',
    hash: 'loaded-version', hosted: { id: 'work', path: '/work/document.docx', fileHash: 'loaded-version' } })
})

it('rejects an empty file target instead of treating it as a file-backed document', async () => {
  fixture({ openDocument: async () => ({ id: 'work', name: 'Report', path: '' }) })
  await expect(getDocsHost().openDocx()).rejects.toThrow(/invalid document identity/)
})
