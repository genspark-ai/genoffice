import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, extname, isAbsolute } from 'node:path'
import { z } from 'zod'
import { readPdfText } from '../../../../../pdf/src/main/read-text'
import type { McpToolDefinition } from '../mcp-server'

/**
 * PDF reading — the registry's `pdf` family exposes read-only MCP access
 * (`mcp.read`, no session, no generation): the pdf app is a viewer, so an
 * agent reads a PDF's text layer directly from disk through the same
 * main-process pdfium the editor uses. Like read_docx, this is headless and
 * always registered. Read-only is the final design — MCP does not drive the
 * pdf editor.
 */

const PDF_EXT = '.pdf'
/** hard cap on returned text so one huge PDF cannot flood the agent's context */
const MAX_TOTAL_CHARS = 80_000

/**
 * Parse a 1-based page list like "3" or "1-5,8" into sorted unique pages.
 * Throws a guided error instead of silently reading nothing.
 */
export function parsePageList(spec: string, pageCount: number): number[] {
  const pages = new Set<number>()
  for (const part of spec.split(',')) {
    const trimmed = part.trim()
    if (!trimmed) continue
    const range = trimmed.match(/^(\d+)(?:\s*-\s*(\d+))?$/)
    if (!range) throw new Error(`invalid page "${trimmed}" — use "3" or "1-5,8"`)
    const from = Number(range[1])
    const to = range[2] === undefined ? from : Number(range[2])
    if (to < from) throw new Error(`invalid page range "${trimmed}" — end must not precede start`)
    for (let p = from; p <= to; p++) {
      if (p < 1 || p > pageCount) throw new Error(`page ${p} is out of range (1-${pageCount})`)
      pages.add(p)
    }
  }
  if (pages.size === 0) throw new Error('pages must list at least one page, e.g. "1-5,8"')
  return [...pages].sort((a, b) => a - b)
}

export function createPdfTools(): McpToolDefinition[] {
  return [
    {
      name: 'read_pdf',
      description:
        'Read a PDF file: page count, metadata (title/author), and per-page text with page size. ' +
        'Text is extracted in content order, so complex multi-column layouts may scramble the ' +
        'reading order. Scanned pages have no text layer and come back with empty text and ' +
        'hasTextLayer:false. Large documents are capped; use `pages` for a bounded read.',
      inputSchema: {
        path: z.string().describe('absolute path to a .pdf file'),
        pages: z
          .string()
          .optional()
          .describe('1-based pages to read, e.g. "3" or "1-5,8"; default is the whole document'),
      },
      handler: async (args) => {
        const filePath = String(args.path ?? '')
        if (!isAbsolute(filePath)) throw new Error('path must be absolute')
        if (extname(filePath).toLowerCase() !== PDF_EXT)
          throw new Error('path must point to a .pdf file')
        if (!existsSync(filePath)) throw new Error(`file not found: ${filePath}`)

        const bytes = new Uint8Array(await readFile(filePath))
        const open = () =>
          readPdfText(bytes, { fromPage: 1, toPage: 1 }).catch(() => {
            throw new Error('could not open the PDF — it may be encrypted or corrupt')
          })

        // probe the page count first so a page list can be validated (and the
        // open error surfaced as a clean tool error) before extracting anything
        const pageCount = (await open()).pageCount

        let fromPage = 1
        let toPage = pageCount
        let wanted: number[] | null = null
        const spec = typeof args.pages === 'string' ? args.pages.trim() : ''
        if (spec) {
          wanted = parsePageList(spec, pageCount)
          fromPage = wanted[0]!
          toPage = wanted[wanted.length - 1]!
        }

        // `truncated` means the budget clipped text inside the requested span —
        // pages the caller never asked for do not count as truncation
        const doc = await readPdfText(bytes, { fromPage, toPage, charBudget: MAX_TOTAL_CHARS })
        const pages = wanted ? doc.pages.filter((p) => wanted!.includes(p.page)) : doc.pages
        return {
          path: filePath,
          name: basename(filePath),
          pageCount: doc.pageCount,
          info: doc.info,
          pages,
          truncated: doc.truncated,
        }
      },
    },
  ]
}
