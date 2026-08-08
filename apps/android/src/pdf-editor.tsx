import { useEffect, useState } from 'react'
import { App as PdfApp } from '../../pdf/src/renderer/App'
import { installRendererStub, aiDefaults } from './editor-stubs'
import '../../pdf/src/renderer/styles.css'

let installed = false
let pendingPdf: { path: string; bytes: Uint8Array; name: string } | null = null

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

export function PdfEditorScreen(): React.JSX.Element {
  const [generation, setGeneration] = useState(0)

  useEffect(() => {
    const onOpen = () => setGeneration((value) => value + 1)
    window.addEventListener('genoffice-android-pdf-open', onOpen)
    if (!installed) {
      installRendererStub('pdfApi', {
        consumePending: async () => pendingPdf?.path ?? null,
        readFile: async (path: string) => {
          if (pendingPdf?.path === path) return pendingPdf.bytes
          throw new Error('Android PDF is no longer available in this session')
        },
        setDirty: () => {},
        onSaveAsFlow: () => () => {},
        getAiSettings: async () => aiDefaults(),
        setAiSettings: async (settings: unknown) => localStorage.setItem('genoffice.android.ai.settings', JSON.stringify(settings)),
        aiGskStatus: async () => ({ loggedIn: false }),
      })
      installed = true
    }
    return () => window.removeEventListener('genoffice-android-pdf-open', onOpen)
  }, [])

  return (
    <div className="android-pdf-host">
      <div className="android-pdf-toolbar">
        <button onClick={() => void choosePdf()}>Open PDF</button>
        <span>{pendingPdf?.name ?? 'No PDF opened'}</span>
      </div>
      <div className="android-pdf-view"><PdfApp key={generation} /></div>
    </div>
  )
}
