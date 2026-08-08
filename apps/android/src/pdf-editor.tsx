import { useEffect, useState } from 'react'
import { Directory, Filesystem } from '@capacitor/filesystem'
import { PDFDocument } from 'pdf-lib'
import PdfApp from '../../pdf/src/renderer/App'
import { installRendererStub, aiDefaults } from './editor-stubs'
import { savePdfSession, writePdfToDocuments, type AndroidPdfSession } from './pdf-android-platform'
import '../../pdf/src/renderer/styles.css'

let installed = false
let pendingPdf: AndroidPdfSession | null = null

async function choosePdf(): Promise<void> {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = '.pdf,application/pdf'
  const file = await new Promise<File | null>((resolve) => {
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.click()
  })
  if (!file) return
  pendingPdf = { path: `android-pdf://${crypto.randomUUID()}`, bytes: new Uint8Array(await file.arrayBuffer()), name: file.name }
  window.dispatchEvent(new Event('genoffice-android-pdf-open'))
}

async function saveCurrent(request: any, targetPath?: string): Promise<{ ok: boolean; error?: string }> {
  if (!pendingPdf) return { ok: false, error: 'No PDF is open.' }
  const result = await savePdfSession(pendingPdf, request)
  if (!result.ok || !result.bytes) return { ok: false, error: result.error ?? 'PDF save failed.' }
  const outputName = targetPath ? targetPath.split('/').pop() || pendingPdf.name : pendingPdf.name
  const saved = await writePdfToDocuments(outputName.endsWith('.pdf') ? outputName : `${outputName}.pdf`, result.bytes)
  if (!saved.ok) return { ok: false, error: saved.error }
  if (!targetPath) pendingPdf = { ...pendingPdf, bytes: result.bytes }
  return { ok: true }
}

export function PdfEditorScreen(): React.JSX.Element {
  const [generation, setGeneration] = useState(0)

  useEffect(() => {
    const onOpen = () => setGeneration((value) => value + 1)
    window.addEventListener('genoffice-android-pdf-open', onOpen)
    if (!installed) {
      installRendererStub('pdfApi', {
        consumePending: async () => pendingPdf?.path ?? null,
        readFile: async (path: string) => {
          if (pendingPdf?.path === path) return pendingPdf.bytes.buffer.slice(pendingPdf.bytes.byteOffset, pendingPdf.bytes.byteOffset + pendingPdf.bytes.byteLength)
          throw new Error('Android PDF is no longer available in this session')
        },
        save: async (request: any) => saveCurrent(request),
        extractPages: async (request: any) => {
          if (!pendingPdf) return { ok: false, error: 'No PDF is open.' }
          try {
            const source = await PDFDocument.load(pendingPdf.bytes)
            const out = await PDFDocument.create()
            const indices = Array.isArray(request.pages) ? request.pages.filter((p: number) => p >= 0 && p < source.getPageCount()) : []
            const pages = await out.copyPages(source, indices)
            pages.forEach((p) => out.addPage(p))
            const saved = await writePdfToDocuments(request.suggestedName || 'Extracted.pdf', await out.save())
            return saved.ok ? { ok: true, savedPath: saved.path } : { ok: false, error: saved.error ?? 'Export failed.' }
          } catch (error) { return { ok: false, error: error instanceof Error ? error.message : String(error) } }
        },
        insertPdf: async (request: any) => {
          if (!pendingPdf) return { ok: false, error: 'No PDF is open.' }
          const input = document.createElement('input')
          input.type = 'file'; input.accept = '.pdf,application/pdf'
          const file = await new Promise<File | null>((resolve) => { input.onchange = () => resolve(input.files?.[0] ?? null); input.click() })
          if (!file) return { ok: true, canceled: true }
          try {
            const base = await PDFDocument.load(pendingPdf.bytes)
            const add = await PDFDocument.load(await file.arrayBuffer())
            const copied = await base.copyPages(add, add.getPageIndices())
            const at = Math.max(-1, Math.min(request.afterPageIndex ?? -1, base.getPageCount() - 1))
            copied.forEach((page, i) => base.insertPage(at + 1 + i, page))
            pendingPdf = { ...pendingPdf, bytes: await base.save() }
            setGeneration((value) => value + 1)
            return { ok: true, insertedCount: copied.length }
          } catch (error) { return { ok: false, error: error instanceof Error ? error.message : String(error) } }
        },
        exportImages: async (request: any) => {
          try {
            const base = (request.baseName || 'page').replace(/[^a-zA-Z0-9_-]/g, '_')
            for (let i = 0; i < (request.images ?? []).length; i++) {
              const name = `${base}-${request.pageNumbers?.[i] ?? i + 1}.png`
              await Filesystem.writeFile({ path: `GenOffice/Exports/${name}`, directory: Directory.Documents, data: request.images[i], recursive: true })
            }
            return { ok: true, savedDir: 'GenOffice/Exports', count: request.images?.length ?? 0 }
          } catch (error) { return { ok: false, error: error instanceof Error ? error.message : String(error) } }
        },
        setDirty: () => {},
        onCloseSaveRequest: () => () => {},
        sendCloseSaveResult: () => {},
        onSaveAsRequest: (handler: (targetPath: string) => void) => { (globalThis as any).__genofficePdfSaveAs = handler; return () => { delete (globalThis as any).__genofficePdfSaveAs } },
        sendSaveAsResult: () => {},
        onSaveAsFlow: () => () => {},
        getAiSettings: async () => aiDefaults(),
        aiStream: async () => {},
        aiStreamCancel: async () => {},
        onAiStream: () => () => {},
        getLanguage: async () => 'en',
        onLanguageChanged: () => () => {},
      })
      installed = true
    }
    return () => window.removeEventListener('genoffice-android-pdf-open', onOpen)
  }, [])

  return (
    <div className="android-pdf-host">
      <div className="android-pdf-toolbar">
        <button onClick={() => void choosePdf()}>Open PDF</button>
        <button disabled={!pendingPdf} onClick={() => void saveCurrent({ markups: [], drawings: [], formValues: [], stamps: [] })}>Save</button>
        <button disabled={!pendingPdf} onClick={() => void saveCurrent({ markups: [], drawings: [], formValues: [], stamps: [] }, 'GenOffice-Edited.pdf')}>Save As</button>
        <span>{pendingPdf?.name ?? 'No PDF opened'}</span>
      </div>
      <div className="android-pdf-view"><PdfApp key={generation} /></div>
    </div>
  )
}
