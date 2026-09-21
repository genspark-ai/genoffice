import type { JSONContent } from '@tiptap/core'
import { Image } from '@tiptap/extension-image'
import { Link } from '@tiptap/extension-link'
import { Paragraph } from '@tiptap/extension-paragraph'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { t } from '../i18n/locale'
import { showToast } from '../components/toast-bus'

/** Directory of the open .md file; relative image paths resolve against it for display */
let imageBaseDir: string | null = null

export function setImageBaseDir(dir: string | null): void {
  imageBaseDir = dir
}

/**
 * Directory of a file path for resolving relative image sources. A file at
 * the filesystem root ("/note.md") resolves to "/" so its sibling images
 * keep working; paths without a separator have no directory to take.
 */
export function dirOf(path: string): string {
  const i = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  if (i > 0) return path.slice(0, i)
  if (i === 0) return path.slice(0, 1)
  return path
}

/**
 * Map an authored image src to a displayable URL. Markdown keeps the authored
 * value (usually a path relative to the .md file); the editor DOM loads it via
 * the main process's md-asset:// handler — a plain file:// subresource would be
 * blocked when the renderer page itself is served over http (dev server).
 */
/** renderer-local: double-click on a picture asks App to open the viewer */
export const VIEW_IMAGE_EVENT = 'markdown-view-image'

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

/**
 * Reverse of {@link resolveImageSrc}: map a display URL (md-asset://) back to
 * the authored path so DOM-parsed content (copy/paste inside the editor) never
 * bakes display URLs into the stored document / serialized markdown.
 */
export function unresolveImageSrc(src: string, baseDir: string | null = imageBaseDir): string {
  if (!src.startsWith('md-asset://')) return src
  let path = decodeURIComponent(src.slice('md-asset://'.length))
  // Windows drive paths were prefixed with '/' to form a valid URL path
  if (/^\/[a-zA-Z]:\//.test(path)) path = path.slice(1)
  if (baseDir) {
    const base = `${baseDir.replace(/\\/g, '/').replace(/\/$/, '')}/`
    if (path.startsWith(base)) return path.slice(base.length)
  }
  return path
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
  // untitled documents have no assets/ directory yet — tell the user to save first
  if (!rel) {
    showToast(t('imageNeedsSavedDocument'), 'error')
    return
  }
  const alt = file.name.replace(/\.[a-z0-9]+$/i, '')
  editor
    .chain()
    .focus()
    .insertContentAt(pos, { type: 'image', attrs: { src: rel, alt } })
    .run()
}

const attr = (value: unknown) =>
  String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')

function imageMarkdown(attrs: JSONContent['attrs']): string {
  const src = String(attrs?.src ?? '')
  const alt = String(attrs?.alt ?? '')
  const title = String(attrs?.title ?? '')
  // markdown has no image size syntax; a sized picture keeps its HTML form
  if (attrs?.width != null || attrs?.height != null) {
    const parts = [`src="${attr(src)}"`]
    if (alt) parts.push(`alt="${attr(alt)}"`)
    if (title) parts.push(`title="${attr(title)}"`)
    if (attrs.width != null) parts.push(`width="${attr(attrs.width)}"`)
    if (attrs.height != null) parts.push(`height="${attr(attrs.height)}"`)
    return `<img ${parts.join(' ')} />`
  }
  return title ? `![${alt}](${src} "${title}")` : `![${alt}](${src})`
}

/**
 * Inline image node (markdown places `![alt](src)` inside paragraphs, headings,
 * list items and table cells next to text; a block node there makes the
 * document invalid and ProseMirror drops the image on the next re-parse). The
 * DOM src is display-resolved while the stored attribute (and therefore the
 * serialized markdown) keeps the authored path untouched. Pasted / dropped
 * image files are persisted into `assets/` beside the file.
 */
export const LocalImage = Image.extend({
  addOptions() {
    return { ...this.parent!(), inline: true }
  },

  addAttributes() {
    return {
      ...this.parent?.(),
      // resolved at the attribute level so the renderer displays the
      // md-asset:// URL while the stored value keeps the authored path
      src: {
        default: null,
        parseHTML: (element) => unresolveImageSrc(element.getAttribute('src') ?? ''),
        renderHTML: (attrs) => ({ src: resolveImageSrc(String(attrs.src ?? '')) }),
      },
    }
  },

  // the serializer only emits marks around text nodes, so a badge-style
  // `[![alt](src)](href)` has to wrap itself
  renderMarkdown: (node) => {
    const image = imageMarkdown(node.attrs)
    const link = node.marks?.find((mark) => mark.type === 'link')
    if (!link) return image
    const href = String(link.attrs?.href ?? '')
    const title = String(link.attrs?.title ?? '')
    return title ? `[${image}](${href} "${title}")` : `[${image}](${href})`
  },

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key: new PluginKey('localImageUpload'),
        props: {
          handleDoubleClickOn(_view, _pos, node, _nodePos, event) {
            if (node.type.name !== 'image') return false
            const src = (event.target as HTMLImageElement | null)?.currentSrc
            if (!src) return false
            window.dispatchEvent(new CustomEvent(VIEW_IMAGE_EVENT, { detail: { src } }))
            return true
          },
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

function markImages(
  content: JSONContent[],
  mark: { type: string; attrs: JSONContent['attrs'] },
): JSONContent[] {
  return content.map((node) => {
    if (node.type === 'image') return { ...node, marks: [...(node.marks ?? []), mark] }
    if (node.content) return { ...node, content: markImages(node.content, mark) }
    return node
  })
}

/** Link whose markdown parse also marks the images it wraps (the manager only marks text) */
export const ImageAwareLink = Link.extend({
  parseMarkdown: (token, helpers) => {
    const attrs = { href: token.href, title: token.title || null }
    const content = markImages(helpers.parseInline(token.tokens || []), { type: 'link', attrs })
    return helpers.applyMark('link', content, attrs)
  },
})

/** Paragraph whose markdown parse keeps a lone image inside it (the stock one lifts it out as a block) */
export const ImageParagraph = Paragraph.extend({
  parseMarkdown: (token, helpers) => {
    const parsed = Paragraph.config.parseMarkdown!(token, helpers)
    return Array.isArray(parsed) ? helpers.createNode('paragraph', undefined, parsed) : parsed
  },
})
