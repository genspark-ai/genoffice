import type { DocsEngineTarget } from './document-query'

export interface DocsRangePreconditions {
  /** The exact ordered block sequence observed for the replacement range. */
  targets?: DocsEngineTarget[]
  allowedTypes?: string[]
  /** These observed blocks must remain outside the replaced range. */
  preserve?: DocsEngineTarget[]
}

const blockTarget = {
  type: 'object',
  additionalProperties: false,
  properties: {
    kind: { type: 'string', enum: ['block'] },
    documentId: { type: 'string' },
    revision: { type: 'integer', minimum: 0 },
    range: {
      type: 'object',
      additionalProperties: false,
      properties: { from: { type: 'integer', minimum: 0 }, to: { type: 'integer', minimum: 1 } },
      required: ['from', 'to'],
    },
    expectedType: { type: 'string' },
    expectedHtml: { type: 'string' },
  },
  required: ['documentId', 'revision', 'range', 'expectedType'],
}

export const DOCS_RANGE_PRECONDITIONS_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    targets: { type: 'array', minItems: 1, maxItems: 128, items: blockTarget },
    allowedTypes: {
      type: 'array',
      minItems: 1,
      maxItems: 32,
      uniqueItems: true,
      items: { type: 'string' },
    },
    preserve: { type: 'array', maxItems: 128, items: blockTarget },
  },
}
