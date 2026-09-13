import { existsSync, readFileSync } from 'node:fs'
import { dirname, isAbsolute, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ensureDom } from '../dom'
import { assertAllowed, type PathContext } from '../fs'
import { imageSize } from './image-size'

/**
 * Markdown → Word reuses the markdown app's editor: its Tiptap markdown parser
 * and its docx exporter (Tiptap JSON → docx-engine), both pure apart from the
 * DOM the editor needs. Local images next to the markdown file are embedded
 * (path policy applies); remote ones fall back to their alt text, mermaid
 * diagrams to their source.
 */
async function markdownModules() {
  await ensureDom()
  const [{ Editor }, extensions, ops, docxExport] = await Promise.all([
    import('@tiptap/core'),
    import('../../../../apps/markdown/src/renderer/editor/extensions'),
    import('../../../../apps/markdown/src/renderer/editor/ops'),
    import('../../../../apps/markdown/src/renderer/export/docxExport'),
  ])
  return { Editor, extensions, ops, docxExport }
}

const NO_SLASH_MENU = {
  slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
  slashItems: () => [],
}

async function withMarkdownEditor<T>(
  markdown: string,
  fn: (
    editor: import('@tiptap/core').Editor,
    mods: Awaited<ReturnType<typeof markdownModules>>,
  ) => Promise<T>,
): Promise<T> {
  const mods = await markdownModules()
  const editor = new mods.Editor({
    extensions: mods.extensions.buildExtensions(NO_SLASH_MENU as never),
    content: '',
  })
  try {
    const nodes = mods.ops.parseMarkdownToNodes(editor, markdown)
    editor.commands.setContent({ type: 'doc', content: nodes.map((n) => n.toJSON()) })
    return await fn(editor, mods)
  } finally {
    editor.destroy()
  }
}

export interface MarkdownImageSource {
  /** the markdown file's location, for relative image paths */
  file: string
  ctx: PathContext
}

export async function markdownToDocx(
  markdown: string,
  images?: MarkdownImageSource,
): Promise<Uint8Array> {
  return withMarkdownEditor(markdown, (editor, mods) =>
    mods.docxExport.exportDocxBytes(editor.getJSON(), async (src) => loadLocalImage(src, images)),
  )
}

async function loadLocalImage(src: string, images?: MarkdownImageSource) {
  if (!images || !src) return null
  // only real URL schemes are remote; a Windows drive letter (C:\...) is a path
  if (/^(https?|data|blob):/i.test(src)) return null
  const local = src.startsWith('file:') ? fileURLToPath(src) : src
  const candidates = [local]
  try {
    const decoded = decodeURI(local)
    if (decoded !== local) candidates.push(decoded)
  } catch {
    // a literal % in the file name is not an escape; the raw spelling is tried
  }
  const path = candidates
    .map((c) => (isAbsolute(c) ? c : resolve(dirname(images.file), c)))
    .find((c) => existsSync(c))
  if (!path) return null
  assertAllowed(path, images.ctx.env, 'read')
  const bytes = new Uint8Array(readFileSync(path))
  const size = imageSize(bytes)
  if (!size) return null
  return {
    base64: Buffer.from(bytes).toString('base64'),
    mime: size.mime,
    widthPx: size.width,
    heightPx: size.height,
  }
}

/** HTML (the docs editor's restricted subset or a plain page body) → GFM through the markdown editor's serializer. */
export async function htmlToMarkdown(html: string): Promise<string> {
  const mods = await markdownModules()
  const editor = new mods.Editor({
    extensions: mods.extensions.buildExtensions(NO_SLASH_MENU as never),
    content: html,
    contentType: 'html',
  })
  try {
    return editor.getMarkdown()
  } finally {
    editor.destroy()
  }
}

export async function markdownToHtml(markdown: string, title: string): Promise<string> {
  return withMarkdownEditor(markdown, async (editor) => {
    const body = editor.getHTML()
    return (
      '<!doctype html>\n<html><head><meta charset="utf-8">' +
      `<title>${escapeHtml(title)}</title></head>\n<body>\n${body}\n</body></html>\n`
    )
  })
}

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!,
  )
}
