/** Revision-scoped address, valid only in its issuing live document generation. */
export interface DocsEngineTarget {
  kind?: 'block' | 'cell' | 'image' | 'section' | 'header' | 'footer' | 'textbox'
  documentId: string
  revision: number
  range: { from: number; to: number }
  expectedType: string
  expectedHtml?: string
  cell?: { row: number; col: number }
  address?: {
    section: number
    kind?: 'header' | 'footer'
    view?: 'default' | 'first' | 'even'
    textboxIndex?: number
  }
}

export interface DocsQueryItem {
  target: DocsEngineTarget
  type: string
  blockIndex: number
  html: string
  complete: boolean
  supportedOperations: string[]
  mediaRef?: string
  structures?: Array<{ id: string; kind: string; from: number; to: number }>
  structuresComplete?: boolean
  structuresPage?: { offset: number; complete: boolean; nextCursor?: string }
}

export interface DocsQueryResult {
  documentId: string
  revision: number
  page: { complete: boolean; nextCursor?: string }
  items: DocsQueryItem[]
}
