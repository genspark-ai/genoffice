import type { Editor } from '@tiptap/core'
import type { DocsEngineTarget, DocsQueryItem, DocsQueryResult } from '../../shared/document-query'
import { documentImagesOf, serializeRangeToHtml } from '../ai/protocol'
import { bindDocumentImageHtml } from './document-media'
import { TableMap } from '@tiptap/pm/tables'
import type { RuntimeDocument } from './docs-runtime'
import { inventoryStructures } from '../ai/structure-inventory'
import { hasDeletionRevision } from '../ai/doc-utils'

/** Serializer coordinates remain source-grid coordinates when deleted rows are omitted. */
function cellHtml(table: HTMLTableElement | null, row: number, col: number): string | undefined {
  return [...(table?.querySelectorAll('td[data-row],th[data-row]') ?? [])].find(
    (cell) =>
      cell.closest('table') === table &&
      cell.getAttribute('data-row') === String(row) &&
      cell.getAttribute('data-col') === String(col),
  )?.outerHTML
}

export interface QueryWindow {
  /** Internal continuation window; initial public requests start at zero. */
  structureOffset?: number
  domain?: 'body' | 'cells' | 'images' | 'sections' | 'headers' | 'footers'
  startBlockIndex?: number
  endBlockIndex?: number
  type?: string
  containsText?: string
  limit?: number
  maxItemChars?: number
  cursor?: string
}

export interface QueryContinuation {
  documentId: string
  revision: number
  offset: number
  query: QueryWindow
}

export function projectedTargetHtml(editor: Editor, target: DocsEngineTarget): string | undefined {
  let result: string | undefined
  editor.state.doc.forEach((node, pos, index) => {
    if (pos !== target.range?.from || pos + node.nodeSize !== target.range?.to) return
    const html = bindDocumentImageHtml(
      serializeRangeToHtml(editor, index, index),
      target.documentId,
      target.revision,
    )
    if (!target.cell) {
      result = html
      return
    }
    if (node.type.name !== 'docTable') return
    const map = TableMap.get(node)
    const table = new DOMParser().parseFromString(html, 'text/html').querySelector('table')
    node.forEach((row, rowOffset) => {
      if (hasDeletionRevision(row)) return
      row.forEach((cell, cellOffset) => {
        if (hasDeletionRevision(cell)) return
        const rect = map.findCell(rowOffset + 1 + cellOffset)
        if (
          target.cell &&
          target.cell.row >= rect.top &&
          target.cell.row < rect.bottom &&
          target.cell.col >= rect.left &&
          target.cell.col < rect.right
        )
          result = cellHtml(table, rect.top, rect.left)
      })
    })
  })
  return result
}

/** Only the existing current-content projection participates in filtering or output. */
export function queryDocument(
  editor: Editor,
  documentId: string,
  revision: number,
  input: QueryWindow,
  sides?: RuntimeDocument,
  continuations = new Map<string, QueryContinuation>(),
): DocsQueryResult {
  let query = input
  let offset = 0
  if (input.cursor) {
    if (Object.keys(input).some((key) => key !== 'cursor'))
      throw new Error('A continuation cannot change query filters or budgets')
    const cursor = continuations.get(input.cursor)
    if (!cursor || cursor.documentId !== documentId || cursor.revision !== revision)
      throw new Error('Stale continuation; restart query_document at the current revision.')
    if (
      !Number.isSafeInteger(cursor.offset) ||
      cursor.offset < 0 ||
      !cursor.query ||
      cursor.query.cursor
    )
      throw new Error('Invalid query continuation')
    query = cursor.query
    offset = cursor.offset
  }
  const limit = query.limit ?? 20
  const maxChars = query.maxItemChars ?? 2000
  if (
    !Number.isSafeInteger(limit) ||
    limit < 1 ||
    limit > 50 ||
    !Number.isSafeInteger(maxChars) ||
    maxChars < 1 ||
    maxChars > 4000
  )
    throw new Error('Invalid query window budget')
  const items: DocsQueryItem[] = []
  const continuation = (window: QueryWindow, offset: number): string => {
    const cursor = `query:${crypto.randomUUID()}`
    continuations.set(cursor, { documentId, revision, query: { ...window }, offset })
    if (continuations.size > 128) continuations.delete(continuations.keys().next().value!)
    return cursor
  }
  let outputChars = 0
  let matched = 0
  let more = false
  const add = (item: DocsQueryItem) => {
    if (more) return
    if (
      item.blockIndex < (query.startBlockIndex ?? 0) ||
      item.blockIndex > (query.endBlockIndex ?? Infinity)
    )
      return
    const parsed = new DOMParser().parseFromString(item.html, 'text/html')
    if (query.type && item.type !== query.type) return
    if (query.containsText && !parsed.body.textContent?.includes(query.containsText)) return
    if (matched++ < offset) return
    if (items.length >= limit) {
      more = true
      return
    }
    const projected = {
      ...item,
      html: item.html.slice(0, maxChars),
      complete: item.html.length <= maxChars,
      supportedOperations: editor.isEditable
        ? item.supportedOperations
        : item.supportedOperations.filter((name) => name === 'read_document_media'),
    }
    const size = JSON.stringify(projected).length
    if (outputChars + size > 18000 && items.length) {
      more = true
      return
    }
    if (projected.structuresPage && !projected.structuresPage.complete)
      projected.structuresPage.nextCursor = continuation(
        {
          ...query,
          startBlockIndex: item.blockIndex,
          endBlockIndex: item.blockIndex,
          limit: 1,
          structureOffset: projected.structuresPage.offset + (projected.structures?.length ?? 0),
        },
        0,
      )
    outputChars += size
    items.push(projected)
  }
  const escaped = (text: string) => {
    const element = document.createElement('p')
    element.textContent = text
    return element.outerHTML
  }
  if (query.domain === 'sections') {
    for (const section of sides?.pageSetup?.list() ?? [])
      add({
        target: {
          kind: 'section',
          documentId,
          revision,
          range: { from: 0, to: 0 },
          expectedType: 'section',
          address: { section: section.index },
        },
        type: 'section',
        blockIndex: section.firstBlock,
        html: escaped(
          JSON.stringify({
            index: section.index,
            firstBlock: section.firstBlock,
            lastBlock: section.lastBlock,
            paper: section.paper,
            width: section.width,
            height: section.height,
            orientation: section.orientation,
            margins: {
              top: section.margins.top,
              right: section.margins.right,
              bottom: section.margins.bottom,
              left: section.margins.left,
              header: section.margins.header,
              footer: section.margins.footer,
              gutter: section.margins.gutter,
            },
            columns: { count: section.columns.count, spacing: section.columns.spacing },
            titlePg: section.titlePg,
            startType: section.startType,
            pageNumberStart: section.pageNumberStart,
            pageNumberFormat: section.pageNumberFormat,
          }),
        ),
        complete: true,
        supportedOperations: ['set_page_setup'],
      })
  }
  if (query.domain === 'headers' || query.domain === 'footers') {
    const state = sides?.headerFooter?.read()
    if (state) {
      const kind = query.domain === 'headers' ? 'header' : 'footer'
      const sections = state.sections ?? [
        {
          index: 0,
          titlePg: state.titlePg,
          strips: [
            {
              kind,
              view: 'default' as const,
              text: state[kind],
              textEditable: true,
              textboxes: [],
            },
          ],
        },
      ]
      for (const section of sections)
        for (const strip of section.strips) {
          if (strip.kind !== kind) continue
          const target: DocsEngineTarget = {
            kind,
            documentId,
            revision,
            range: { from: 0, to: 0 },
            expectedType: kind,
            address: { section: section.index, kind, view: strip.view },
          }
          add({
            target,
            type: kind,
            blockIndex:
              sides?.pageSetup?.list().find((s) => s.index === section.index)?.firstBlock ?? 0,
            html: escaped(strip.text),
            complete: true,
            supportedOperations: strip.textEditable === false ? [] : ['set_header_footer'],
          })
          for (const box of strip.textboxes ?? [])
            add({
              target: {
                ...target,
                kind: 'textbox',
                expectedType: 'textbox',
                address: {
                  section: section.index,
                  kind,
                  view: strip.view,
                  textboxIndex: box.index,
                },
              },
              type: 'textbox',
              blockIndex:
                sides?.pageSetup?.list().find((s) => s.index === section.index)?.firstBlock ?? 0,
              html: escaped(box.text),
              complete: true,
              supportedOperations: box.editable ? ['set_header_footer'] : [],
            })
        }
    }
  }
  editor.state.doc.forEach((node, pos, index) => {
    if (more) return
    if (index < (query.startBlockIndex ?? 0) || index > (query.endBlockIndex ?? Infinity)) return
    const html = bindDocumentImageHtml(
      serializeRangeToHtml(editor, index, index),
      documentId,
      revision,
    )
    if (!html) return
    const parsed = new DOMParser().parseFromString(html, 'text/html')
    if (query.domain === 'cells') {
      if (node.type.name !== 'docTable') return
      const map = TableMap.get(node)
      const table = parsed.querySelector('table')
      node.forEach((row, rowOffset) => {
        if (hasDeletionRevision(row)) return
        row.forEach((cell, cellOffset) => {
          if (hasDeletionRevision(cell)) return
          const rect = map.findCell(rowOffset + 1 + cellOffset)
          add({
            target: {
              kind: 'cell',
              documentId,
              revision,
              range: { from: pos, to: pos + node.nodeSize },
              expectedType: 'docTable',
              cell: { row: rect.top, col: rect.left },
            },
            type: 'docTableCell',
            blockIndex: index,
            html: cellHtml(table, rect.top, rect.left) ?? '',
            complete: true,
            supportedOperations: ['setTableCellText'],
          })
        })
      })
      return
    }
    if (query.domain === 'images') {
      documentImagesOf(node).forEach((_image, ordinal) => {
        const src = `media:${documentId}:${revision}:${index}.${ordinal}`
        add({
          target: {
            kind: 'image',
            documentId,
            revision,
            range: { from: pos, to: pos + node.nodeSize },
            expectedType: node.type.name,
          },
          type: 'image',
          blockIndex: index,
          html: `<img src="${src}">`,
          complete: true,
          mediaRef: src,
          supportedOperations: ['read_document_media'],
        })
      })
      return
    }
    if (query.domain && query.domain !== 'body') return
    const structures = inventoryStructures(editor.state.doc, pos, pos + node.nodeSize)
    add({
      target: {
        kind: 'block',
        documentId,
        revision,
        range: { from: pos, to: pos + node.nodeSize },
        expectedType: node.type.name,
      },
      type: node.type.name,
      blockIndex: index,
      html,
      complete: html.length <= maxChars,
      supportedOperations:
        node.type.name === 'docProtected'
          ? ['deleteBlocks']
          : ['setFont', 'findReplace', 'deleteBlocks'],
      structures: structures
        .slice(query.structureOffset ?? 0, (query.structureOffset ?? 0) + 32)
        .map((item) => ({ ...item, id: `structure:${documentId}:${revision}:${item.id}` })),
      structuresComplete: !query.structureOffset && structures.length <= 32,
      structuresPage: {
        offset: query.structureOffset ?? 0,
        complete: (query.structureOffset ?? 0) + 32 >= structures.length,
      },
    })
  })
  let nextCursor: string | undefined
  if (more) {
    nextCursor = continuation(query, offset + items.length)
  }
  return {
    documentId,
    revision,
    page: { complete: !more, ...(nextCursor ? { nextCursor } : {}) },
    items,
  }
}
