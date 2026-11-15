import { z } from 'zod'
import type { DocsInspection } from './inspection'
import type { DocsAnnotations } from './annotations'
import type { DocsDocumentCommands } from '../../shared/document-host'
import { parseDocsFileMenu, type DocsEditorUi } from './file-menu'
import { parseDocsRibbon } from './ribbon'

const reference = z.object({
  documentId: z.string(),
  revision: z.number().int().nonnegative(),
  from: z.number().int().nonnegative(),
  to: z.number().int().nonnegative(),
})
const options = z.object({
  matchCase: z.boolean().optional(),
  wholeWord: z.boolean().optional(),
  matchWidth: z.boolean().optional(),
  mode: z.enum(['literal', 'regex', 'wildcard']).optional(),
})
const style = z.object({
  background: z.string().max(100).optional(),
  color: z.string().max(100).optional(),
  underline: z.boolean().optional(),
})
const annotation = z.object({
  from: z.number().int().nonnegative(),
  to: z.number().int().nonnegative(),
  message: z.string().max(10000).optional(),
  code: z.string().max(1000).optional(),
  severity: z.enum(['info', 'warning', 'error']).optional(),
  style: style.optional(),
})
const rule = z.object({
  id: z.string(),
  pattern: z.string().max(10000),
  flags: z.string().max(4).optional(),
  message: z.string().max(10000).optional(),
  severity: z.enum(['info', 'warning', 'error']).optional(),
  style: style.optional(),
  allowedValues: z.array(z.string()).max(100000).optional(),
  capture: z.number().int().nonnegative().optional(),
})

/** Native callers use the same public editor methods as browser hosts. No arbitrary method evaluation. */
export async function executeDocsSdkCall(
  api: {
    inspection: DocsInspection
    annotations: DocsAnnotations
    documents: DocsDocumentCommands
    ui?: DocsEditorUi
  },
  operation: string,
  input: unknown,
): Promise<unknown> {
  switch (operation) {
    case 'ui.ribbon.set':
      if (!api.ui) throw new Error('The editor UI SDK is not ready')
      api.ui.ribbon.set(parseDocsRibbon(input))
      return { ok: true }
    case 'ui.ribbon.reset':
      if (!api.ui) throw new Error('The editor UI SDK is not ready')
      z.object({}).strict().parse(input)
      api.ui.ribbon.reset()
      return { ok: true }
    case 'ui.fileMenu.set':
      if (!api.ui) throw new Error('The editor UI SDK is not ready')
      api.ui.fileMenu.set(parseDocsFileMenu(input))
      return { ok: true }
    case 'ui.fileMenu.reset':
      if (!api.ui) throw new Error('The editor UI SDK is not ready')
      z.object({}).strict().parse(input)
      api.ui.fileMenu.reset()
      return { ok: true }
    case 'snapshot':
      return api.inspection.snapshot()
    case 'find': {
      const value = z
        .object({ query: z.string().max(10000), options: options.optional() })
        .parse(input)
      return api.inspection.find(value.query, value.options)
    }
    case 'locate': {
      const value = z.object({ reference, focus: z.boolean().optional() }).parse(input)
      api.inspection.locate(value.reference, { focus: value.focus })
      return { ok: true }
    }
    case 'replace': {
      const value = z
        .object({
          result: reference.extend({ query: z.string(), options, matches: z.array(z.unknown()) }),
          replacement: z.string(),
          index: z.number().int().nonnegative().optional(),
        })
        .parse(input)
      // Match/capture arrays are recomputed by the inspection API at the verified revision.
      return {
        count: api.inspection.replace(
          {
            ...value.result,
            options: { matchCase: false, wholeWord: false, ...value.result.options },
            matches: [],
          },
          value.replacement,
          value.index,
        ),
      }
    }
    case 'annotations.setRules': {
      const value = z.object({ source: z.string(), rules: z.array(rule).max(1000) }).parse(input)
      api.annotations.setRules(value.source, value.rules)
      return { ok: true }
    }
    case 'annotations.setRanges': {
      const value = z
        .object({ source: z.string(), ranges: z.array(reference.and(annotation)).max(10000) })
        .parse(input)
      api.annotations.setRanges(value.source, value.ranges)
      return { ok: true }
    }
    case 'annotations.remove':
      api.annotations.remove(z.object({ source: z.string() }).parse(input).source)
      return { ok: true }
    case 'annotations.diagnostics':
      return api.annotations.diagnostics()
    case 'annotations.snapshot':
      z.object({}).strict().parse(input)
      return api.annotations.snapshot()
    case 'annotations.errors':
      return api.annotations.errors()
    case 'annotations.refresh':
      await api.annotations.refresh()
      return api.annotations.diagnostics()
    case 'documents.new':
      return api.documents.new()
    case 'documents.open':
      return api.documents.open(z.object({ id: z.string().optional() }).parse(input).id)
    case 'documents.save':
      return api.documents.save(z.object({ saveAs: z.boolean().optional() }).parse(input).saveAs)
    case 'documents.state':
      return api.documents.state()
    default:
      throw new Error(`Unknown public SDK operation: ${operation}`)
  }
}
