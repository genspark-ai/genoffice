/// Render a standalone HTML document to PDF bytes in a hidden window —
/// shared by the AI create_document flows (docs / pdf chat creating a new
/// PDF file). The caller supplies the window so this package keeps its
/// type-only Electron dependency; use a sandboxed, scripting-disabled
/// window (`{ show: false, webPreferences: { sandbox: true, javascript: false } }`).
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { BrowserWindow } from 'electron'
import { parseFragment, serialize } from 'parse5'

export type PrintWindow = Pick<BrowserWindow, 'loadFile' | 'destroy'> & {
  webContents: Pick<BrowserWindow['webContents'], 'printToPDF'>
}

declare const printableHtmlBrand: unique symbol
export type PrintableHtml = string & { readonly [printableHtmlBrand]: true }

function escapeHtmlText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

const ALLOWED_BODY_TAGS = new Set([
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'p',
  'ul',
  'ol',
  'li',
  'strong',
  'em',
  'u',
  's',
  'a',
  'br',
  'table',
  'thead',
  'tbody',
  'tr',
  'th',
  'td',
  'pre',
  'code',
  'blockquote',
  'formula',
])

/** Elements whose contents must not be salvaged as printable text. */
const DROP_WITH_CONTENT = new Set([
  'script',
  'style',
  'iframe',
  'frame',
  'frameset',
  'object',
  'embed',
  'link',
  'meta',
  'base',
  'form',
  'template',
  'noscript',
  'svg',
  'math',
])

interface ParsedNode {
  nodeName: string
  tagName?: string
  attrs?: Array<{ name: string; value: string }>
  childNodes?: ParsedNode[]
  parentNode?: ParsedParent
}

interface ParsedParent extends ParsedNode {
  childNodes: ParsedNode[]
}

function sanitizeChildren(parent: ParsedParent): void {
  const safe: ParsedNode[] = []
  for (const node of parent.childNodes) {
    if (node.nodeName === '#text') {
      node.parentNode = parent
      safe.push(node)
      continue
    }
    if (node.nodeName === '#comment') continue

    const tag = node.tagName?.toLowerCase()
    if (!tag || DROP_WITH_CONTENT.has(tag)) continue
    if (node.childNodes) sanitizeChildren(node as ParsedParent)

    if (ALLOWED_BODY_TAGS.has(tag)) {
      node.attrs = []
      node.parentNode = parent
      safe.push(node)
    } else {
      for (const child of node.childNodes ?? []) {
        child.parentNode = parent
        safe.push(child)
      }
    }
  }
  parent.childNodes = safe
}

/** Keep only the restricted document tags and strip every attribute/URL. */
export function sanitizePrintableBody(bodyHtml: string): string {
  const fragment = parseFragment(bodyHtml)
  sanitizeChildren(fragment as unknown as ParsedParent)
  return serialize(fragment)
}

// Print stylesheet, not UI chrome: document output must render identically in
// both app themes, so the colors stay hardcoded (CLAUDE.md theming rule 4).
const PRINT_CSS = `
  body {
    font-family: 'Helvetica Neue', Helvetica, Arial, 'PingFang SC', 'Hiragino Sans GB',
      'Microsoft YaHei', 'Noto Sans CJK SC', sans-serif;
    font-size: 11pt;
    line-height: 1.5;
    color: #1a1a1a;
    margin: 0;
  }
  h1 { font-size: 20pt; line-height: 1.3; margin: 0 0 12pt; }
  h2 { font-size: 16pt; line-height: 1.3; margin: 16pt 0 8pt; }
  h3 { font-size: 13pt; line-height: 1.3; margin: 14pt 0 6pt; }
  h4, h5, h6 { font-size: 11pt; margin: 12pt 0 4pt; }
  p { margin: 0 0 8pt; }
  ul, ol { margin: 0 0 8pt; padding-left: 22pt; }
  li { margin: 0 0 3pt; }
  table { border-collapse: collapse; margin: 0 0 10pt; width: 100%; }
  th, td { border: 1pt solid #999; padding: 3pt 6pt; text-align: left; vertical-align: top; }
  th { background: #f0f0f0; }
  pre {
    font-family: Consolas, Menlo, monospace;
    font-size: 9.5pt;
    background: #f5f5f5;
    border: 1pt solid #ddd;
    padding: 6pt 8pt;
    white-space: pre-wrap;
    word-break: break-word;
    margin: 0 0 8pt;
  }
  blockquote { border-left: 3pt solid #ccc; color: #555; margin: 0 0 8pt; padding: 0 0 0 10pt; }
`

const PRINT_CSP =
  "default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'; " +
  "frame-src 'none'; object-src 'none'; connect-src 'none'"

/** Sanitize an AI-generated restricted-HTML body and wrap it in a printable page. */
export function buildPrintableHtml(title: string, bodyHtml: string): PrintableHtml {
  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '<meta charset="utf-8">',
    `<meta http-equiv="Content-Security-Policy" content="${PRINT_CSP}">`,
    `<title>${escapeHtmlText(title)}</title>`,
    `<style>${PRINT_CSS}</style>`,
    '</head>',
    `<body>${sanitizePrintableBody(bodyHtml)}</body>`,
    '</html>',
  ].join('\n') as PrintableHtml
}

/**
 * Load the HTML in the supplied hidden window and print it to A4 PDF bytes.
 * The HTML goes through a temp file (data: URLs truncate long documents);
 * the window is destroyed and the temp dir removed in every outcome.
 */
export async function printHtmlToPdf(
  html: PrintableHtml,
  createWindow: () => PrintWindow,
): Promise<Buffer> {
  const workDir = await mkdtemp(join(tmpdir(), 'genoffice-ai-doc-'))
  let printWin: PrintWindow | undefined
  try {
    printWin = createWindow()
    const htmlPath = join(workDir, 'print.html')
    await writeFile(htmlPath, html, 'utf8')
    await printWin.loadFile(htmlPath)
    return await printWin.webContents.printToPDF({
      pageSize: 'A4',
      printBackground: true,
      margins: { top: 0.6, bottom: 0.6, left: 0.6, right: 0.6 },
    })
  } finally {
    printWin?.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}
