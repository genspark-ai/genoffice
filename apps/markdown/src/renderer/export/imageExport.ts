import type { ExportResult } from '../../shared/ipc'

/** Same 192 DPI output as Docs; render and write one page at a time. */
export async function exportImages(
  html: string,
  suggestedName: string,
  onProgress: (count: number) => void,
): Promise<ExportResult & { count?: number }> {
  const target = await window.markdownApi.prepareImageExport({ html, suggestedName })
  if (!target.ok || 'canceled' in target) return target
  let task: import('pdfjs-dist').PDFDocumentLoadingTask | undefined
  let canvas: HTMLCanvasElement | undefined
  try {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const { default: workerUrl } = await import('pdfjs-dist/legacy/build/pdf.worker.min.mjs?url')
    pdfjs.GlobalWorkerOptions.workerSrc = workerUrl
    task = pdfjs.getDocument({
      data: Uint8Array.from(atob(target.pdfBase64), (char) => char.charCodeAt(0)),
      useWasm: false,
    })
    const pdf = await task.promise
    onProgress(pdf.numPages)
    canvas = document.createElement('canvas')
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i)
      const viewport = page.getViewport({ scale: 192 / 72 })
      canvas.width = Math.ceil(viewport.width)
      canvas.height = Math.ceil(viewport.height)
      await page.render({ canvas, viewport }).promise
      page.cleanup()
      const pngBase64 = canvas.toDataURL('image/png').split(',')[1]
      const result = await window.markdownApi.writeExportImage(target.id, i, pngBase64)
      if (!result.ok) throw new Error(result.error || 'Image write failed')
    }
    const result = await window.markdownApi.finishImageExport(target.id, true)
    return { ...result, count: pdf.numPages }
  } catch (err) {
    await window.markdownApi.finishImageExport(target.id, false).catch(() => {})
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  } finally {
    if (canvas) canvas.width = canvas.height = 0
    await task?.destroy()
  }
}
