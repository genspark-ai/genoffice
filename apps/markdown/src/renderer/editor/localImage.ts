import { Image } from '@tiptap/extension-image'
import { mergeAttributes } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'

/** Directory of the open .md file; relative image paths resolve against it for display */
let imageBaseDir: string | null = null

export function setImageBaseDir(dir: string | null): void {
  imageBaseDir = dir
}

/**
 * Map an authored image src to a displayable URL. Markdown keeps the authored
 * value (usually a path relative to the .md file); the editor DOM loads it via
 * the main process's md-asset:// handler — a plain file:// subresource would be
 * blocked when the renderer page itself is served over http (dev server).
 */
export function resolveImageSrc(src: string, baseDir: string | null = imageBaseDir): string {
  if (!src) return src
  // ':' is legal in URL path segments (RFC 3986) — restore it after encoding so
  // Windows drive prefixes stay `C:` instead of the `C%3A` Chromium rejects
  const encodeSegment = (seg: string) => encodeURIComponent(seg).replace(/%3A/gi, ':')
  const toAssetUrl = (path: string) =>
    `md-asset://${path.startsWith('/') ? '' : '/'}${path.replace(/\\/g, '/').split('/').map(encodeSegment).join('/')}`
  // a Windows drive path would also match the URL-scheme regex — check it first
  if (src.startsWith('/') || /^[a-zA-Z]:[\\/]/.test(src)) return toAssetUrl(src)
  if (/^[a-z][a-z0-9+.-]*:/i.test(src)) return src
  if (!baseDir) return src
  return toAssetUrl(`${baseDir.replace(/\\/g, '/').replace(/\/$/, '')}/${src}`)
}

const EXT_BY_MIME: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
}

function imageFileIn(data: DataTransfer | null): File | null {
  for (const file of data?.files ?? []) {
    if (EXT_BY_MIME[file.type]) return file
  }
  return null
}

async function persistAndInsert(
  editor: import('@tiptap/core').Editor,
  file: File,
  pos: number,
): Promise<void> {
  const bytes = new Uint8Array(await file.arrayBuffer())
  let binary = ''
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  }
  const rel = await window.markdownApi.saveImage({
    base64: btoa(binary),
    ext: EXT_BY_MIME[file.type]!,
  })
  // untitled documents have no assets/ directory yet — drop silently
  if (!rel) return
  const alt = file.name.replace(/\.[a-z0-9]+$/i, '')
  editor
    .chain()
    .focus()
    .insertContentAt(pos, { type: 'image', attrs: { src: rel, alt } })
    .run()
}

/**
 * Image node whose DOM src is display-resolved while the stored attribute (and
 * therefore the serialized markdown) keeps the authored path untouched.
 * Pasted / dropped image files are persisted into `assets/` beside the file.
 */
export const LocalImage = Image.extend({
  renderHTML({ node, HTMLAttributes }) {
    return [
      'img',
      mergeAttributes(this.options.HTMLAttributes, HTMLAttributes, {
        src: resolveImageSrc(String(node.attrs.src ?? '')),
      }),
    ]
  },

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key: new PluginKey('localImageUpload'),
        props: {
          handlePaste(view, event) {
            const file = imageFileIn(event.clipboardData)
            if (!file) return false
            void persistAndInsert(editor, file, view.state.selection.from)
            return true
          },
          handleDrop(view, event, _slice, moved) {
            if (moved) return false
            const file = imageFileIn(event.dataTransfer)
            if (!file) return false
            const pos =
              view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos ??
              view.state.selection.from
            void persistAndInsert(editor, file, pos)
            return true
          },
        },
      }),
    ]
  },
})
