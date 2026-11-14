// Modified for ForaOffice: reliable document targeting and portable inspection.
import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { ToolExecution } from '@genoffice/agent-core'
import type {
  DocsOperationResult,
  DocsErrorCode,
  DocsImageReferenceIssue,
} from '../../shared/operation-result'
import { documentImagesOf } from '../ai/protocol'
import { DOCS_IMAGE_MAX_BYTES } from '../../shared/media-policy'

type MediaResult = ToolExecution & Pick<DocsOperationResult, 'media' | 'errorCode'>
const refused = (output: string, errorCode: DocsErrorCode): MediaResult => ({
  output,
  errorCode,
  summary: 'Document image',
  isError: true,
  mutated: false,
})

/** Public HTML writes require the same document/revision identity as media reads. */
export function prepareDocumentImageWrite(
  editor: Editor | ProseMirrorNode,
  html: string,
  documentId: string,
  revision: number,
): { html: string; errors: DocsImageReferenceIssue[] } {
  const parsed = new DOMParser().parseFromString(html, 'text/html')
  const images = [...parsed.querySelectorAll('img')]
  const errors: DocsImageReferenceIssue[] = []
  const doc = 'state' in editor ? editor.state.doc : editor
  for (const image of images) {
    const reference = image.getAttribute('src') ?? ''
    const ref = /^media:([^:]+):(\d+):(\d+)\.(\d+)$/.exec(reference)
    if (!ref) {
      errors.push({
        reference,
        reason: reference.startsWith('doc:')
          ? 'unbound'
          : reference.startsWith('media:')
            ? 'malformed'
            : 'unsupported',
      })
      continue
    }
    if (ref.slice(2).some((value) => !Number.isSafeInteger(Number(value)))) {
      errors.push({ reference, reason: 'malformed' })
      continue
    }
    if (ref[1] !== documentId || Number(ref[2]) !== revision) {
      errors.push({ reference, reason: 'expired' })
      continue
    }
    const block = doc.maybeChild(Number(ref[3]))
    const picture = block && documentImagesOf(block)[Number(ref[4])]
    if (
      !picture ||
      (picture.type.name === 'docInlineImage' &&
        (typeof picture.attrs.dataUrl !== 'string' || !picture.attrs.dataUrl))
    )
      errors.push({ reference, reason: 'missing' })
    else if (
      picture.type.name !== 'docInlineImage' ||
      /<w:object[\s>]|<o:OLEObject[\s>]|<m:oMath/.test(String(picture.attrs.xml ?? ''))
    )
      errors.push({ reference, reason: 'unsupported' })
    else image.setAttribute('src', `doc:${ref[3]}.${ref[4]}`)
  }
  return { html: images.length ? parsed.body.innerHTML : html, errors }
}

/** Only generated HTML image tags are projected; authored/escaped examples stay text. */
export function bindDocumentImageHtml(html: string, documentId: string, revision: number): string {
  return html.replace(
    /<img src="doc:(\d+)\.(\d+)">/g,
    (_tag, block: string, index: string) =>
      `<img src="media:${documentId}:${revision}:${block}.${index}">`,
  )
}

export function documentMediaReferences(
  editor: Editor,
  documentId: string,
  revision: number,
  start = 0,
  end = editor.state.doc.childCount - 1,
): string {
  const refs: string[] = []
  for (let block = start; block <= end; block++) {
    const node = editor.state.doc.maybeChild(block)
    if (!node) continue
    documentImagesOf(node).forEach((_image, index) => {
      refs.push(`doc:${block}.${index} => media:${documentId}:${revision}:${block}.${index}`)
    })
  }
  return refs.length
    ? `\nAuthorized document images (mediaRef; expires when the document revision changes):\n${refs.join('\n')}`
    : ''
}

/** Reads bytes only; media inference is owned by the Harness. */
export async function readDocumentMedia(
  editor: Editor,
  blockIndex: number,
  signal: AbortSignal,
  isCurrent: () => boolean,
  imageIndex?: number,
): Promise<MediaResult> {
  if (!isCurrent())
    return refused('Read fresh document context before accessing this image.', 'stale_target')
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex >= editor.state.doc.childCount)
    return refused('The image block reference is out of range.', 'invalid_input')
  const images = documentImagesOf(editor.state.doc.child(blockIndex))
  if (images.length === 0)
    return refused(
      'Unsupported media target: this block has no authorized images.',
      'invalid_input',
    )
  if (imageIndex === undefined && images.length > 1)
    return refused(
      'This block has multiple images; specify imageIndex or mediaRef.',
      'invalid_input',
    )
  const index = imageIndex ?? 0
  if (!Number.isInteger(index) || index < 0 || index >= images.length)
    return refused('The image ordinal is out of range.', 'invalid_input')
  const image = images[index]
  const src = image.type.name === 'docProtected' ? image.attrs.imageDataUrl : image.attrs.dataUrl
  if (typeof src !== 'string' || !src)
    return refused('The referenced image resource is missing.', 'unavailable')
  try {
    const response = await fetch(src, { signal })
    if (!response.ok)
      return refused(
        `The host could not read this image (HTTP ${response.status}).`,
        response.status === 404 || response.status === 410
          ? 'unavailable'
          : response.status === 401 || response.status === 403
            ? 'permission_denied'
            : 'host_failure',
      )
    const mimeType = response.headers.get('content-type')?.split(';')[0] ?? ''
    if (!mimeType.startsWith('image/'))
      return refused('The host did not return an image resource.', 'host_failure')
    const reader = response.body?.getReader()
    if (!reader) return refused('The host returned no image bytes.', 'host_failure')
    const chunks: Uint8Array[] = []
    let size = 0
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      size += value.byteLength
      if (size > DOCS_IMAGE_MAX_BYTES) {
        await reader.cancel()
        return refused('The image exceeds the 25 MB media-read limit.', 'unavailable')
      }
      chunks.push(value)
    }
    if (size === 0) return refused('The host returned an empty image resource.', 'host_failure')
    if (signal.aborted) return refused('The image read was canceled.', 'canceled')
    if (!isCurrent())
      return refused(
        'The document changed while the image was read; read fresh context.',
        'stale_target',
      )
    let binary = ''
    for (const chunk of chunks) {
      for (let offset = 0; offset < chunk.length; offset += 32_768)
        binary += String.fromCharCode(...chunk.subarray(offset, offset + 32_768))
    }
    return {
      output: `Image ${index} from document block ${blockIndex}.`,
      summary: 'Document image',
      mutated: false,
      media: { data: btoa(binary), mimeType },
    }
  } catch {
    return refused(
      signal.aborted
        ? 'The image read was canceled.'
        : 'The host image resource could not be read.',
      signal.aborted ? 'canceled' : 'host_failure',
    )
  }
}
