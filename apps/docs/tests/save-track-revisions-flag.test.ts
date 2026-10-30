import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { parseDocx, readSections } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocOptions } from '../src/renderer/editor/convert'
import { noteDocumentSwapped, save, type FileActionContext } from '../src/renderer/file-actions'
import { EMPTY_HF_VARIANTS } from '../src/renderer/doc-state'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const SETTINGS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
const settingsPart = (inner: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:settings>`

const dirs: string[] = []
const editors: Editor[] = []
afterEach(async () => {
  for (const editor of editors.splice(0)) editor.destroy()
  for (const dir of dirs.splice(0)) await rm(dir, { recursive: true, force: true })
})

async function fixture(settingsInner: string, trackChanges: boolean, dirty: boolean) {
  noteDocumentSwapped()
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraParts: [
      { path: 'word/settings.xml', xml: settingsPart(settingsInner), contentType: SETTINGS_TYPE },
    ],
  })
  const parsed = await parseDocx(bytes)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-track-flag-'))
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
    trackChanges,
    trackRevisionsDirty: dirty,
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
  return { parsed, ctx, path }
}

async function savedSettings(ctx: FileActionContext, path: string): Promise<string> {
  expect(await save(ctx, false, false)).toBe(true)
  const zip = await JSZip.loadAsync(await readFile(path))
  return zip.file('word/settings.xml')!.async('string')
}

describe('Track Changes toggle <-> settings.xml w:trackRevisions', () => {
  it('a document saved with tracking on opens with the toggle on', async () => {
    const { parsed } = await fixture('<w:trackRevisions/>', true, false)
    expect(parsed.trackRevisions).toBe(true)
  })

  it('turning tracking off and saving removes the flag', async () => {
    const { ctx, path } = await fixture('<w:trackRevisions/><w:doNotTrackMoves/>', false, true)
    const xml = await savedSettings(ctx, path)
    expect(xml).not.toContain('trackRevisions')
    expect(xml).toContain('<w:doNotTrackMoves/>')
  })

  it('turning tracking on and saving writes the flag', async () => {
    const { ctx, path } = await fixture('<w:zoom w:percent="100"/>', true, true)
    expect(await savedSettings(ctx, path)).toContain('<w:zoom w:percent="100"/><w:trackRevisions/>')
  })

  it('a forced-on recorder (protection trackedChanges) does not write the flag', async () => {
    const { ctx, path } = await fixture(
      '<w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>',
      true,
      false,
    )
    expect(await savedSettings(ctx, path)).not.toContain('trackRevisions')
  })
})
