import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type {
  DocsCurrentEditor,
  DocsInspection,
  DocsTextReference,
  DocsTextSnapshot,
} from './inspection'

export interface DocsRangeStyle {
  background?: string
  color?: string
  underline?: boolean
}
export interface DocsAnnotation {
  from: number
  to: number
  message?: string
  /** Host-defined diagnostic identifier; the editor does not interpret it. */
  code?: string
  severity?: 'info' | 'warning' | 'error'
  style?: DocsRangeStyle
}
export interface DocsDiagnostic extends DocsAnnotation, DocsTextReference {
  source: string
  message: string
  /** The decoration was mapped, but the host has not checked the new text. */
  stale?: boolean
}
export interface DocsDiagnosticSource {
  source: string
  status: 'pending' | 'complete' | 'stale' | 'failed'
  documentId: string
  revision: number
  coverage: { kind: 'document-text' | 'ranges'; from: number; to: number; rangeCount?: number }
  error?: string
}
export interface DocsDiagnosticsSnapshot {
  documentId: string
  revision: number
  /** Changes with diagnostic state even when document content is unchanged. */
  version: number
  complete: boolean
  sources: DocsDiagnosticSource[]
  diagnostics: DocsDiagnostic[]
  errors: Array<{ source: string; message: string }>
}
export interface DocsMatchRule {
  id: string
  pattern: string
  flags?: string
  message?: string
  severity?: DocsAnnotation['severity']
  style?: DocsRangeStyle
  /** When configured, report only captures outside this host-supplied catalog. */
  allowedValues?: string[]
  capture?: number
}
export type DocsChecker = (
  snapshot: DocsTextSnapshot,
  signal: AbortSignal,
) => DocsAnnotation[] | Promise<DocsAnnotation[]>
export interface DocsAnnotations {
  /** Called when the active document changes; registrations are document-local. */
  activate(): void
  reset(): void
  setRules(source: string, rules: DocsMatchRule[]): void
  setRanges(source: string, ranges: Array<DocsAnnotation & DocsTextReference>): void
  registerChecker(source: string, checker: DocsChecker): () => void
  remove(source: string): void
  diagnostics(): DocsDiagnostic[]
  /** Read current check state without running checkers or acknowledging stale findings. */
  snapshot(): DocsDiagnosticsSnapshot
  errors(): Array<{ source: string; message: string }>
  onChange(listener: () => void): () => void
  refresh(): Promise<void>
  dispose(): void
}

const annotationKey = new PluginKey<DecorationSet>('sdk-annotations')
export const DocsAnnotationsExtension = Extension.create({
  name: 'sdkAnnotations',
  addProseMirrorPlugins() {
    let printing = false
    return [
      new Plugin<DecorationSet>({
        key: annotationKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, previous) {
            const ranges = tr.getMeta(annotationKey) as Decoration[] | undefined
            return ranges ? DecorationSet.create(tr.doc, ranges) : previous.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return printing ? DecorationSet.empty : this.getState(state)
          },
        },
        view(view) {
          const window = view.dom.ownerDocument.defaultView
          // Even transparent inline boxes can change printed text shaping.
          const beforePrint = () => {
            printing = true
            view.updateState(view.state)
          }
          const afterPrint = () => {
            printing = false
            view.updateState(view.state)
          }
          window?.addEventListener('beforeprint', beforePrint)
          window?.addEventListener('afterprint', afterPrint)
          return {
            destroy() {
              window?.removeEventListener('beforeprint', beforePrint)
              window?.removeEventListener('afterprint', afterPrint)
            },
          }
        },
      }),
    ]
  },
})

function styleAttributes(style: DocsRangeStyle = {}): string {
  const colors: string[] = []
  for (const [key, value] of [
    ['background', style.background],
    ['color', style.color],
  ] as const) {
    if (value !== undefined) {
      // A color is data, not executable CSS. Reject declarations, URLs and variable injection.
      if (
        value.length > 100 ||
        !/^(#[\da-f]{3,8}|[a-z]+|(?:rgb|hsl)a?\([\d\s.,%/+-]+\))$/i.test(value)
      )
        throw new Error('Invalid annotation color')
      colors.push(`--sdk-annotation-${key}:${value}`)
    }
  }
  if (style.underline) colors.push('--sdk-annotation-underline:underline')
  return colors.join(';')
}

/** View-only extension state. Host grammars run here, never in the document schema. */
export function createDocsAnnotations(
  current: DocsCurrentEditor,
  inspection: DocsInspection,
): DocsAnnotations {
  let editor: Editor | undefined
  let identity: object | undefined
  let disposed = false
  let generation = 0
  let running: AbortController | undefined
  let scheduled = false
  let version = 0
  const checkers = new Map<string, DocsChecker>()
  const rules = new Map<string, Array<{ rule: DocsMatchRule; pattern: RegExp }>>()
  const ranges = new Map<string, DocsAnnotation[]>()
  const findings = new Map<string, DocsDiagnostic[]>()
  const errors = new Map<string, string>()
  const sourceStates = new Map<string, DocsDiagnosticSource>()
  const listeners = new Set<() => void>()
  const live = () => {
    if (disposed) throw new Error('The annotations API is disposed')
    const next = current()
    if (!next) throw new Error('The document is not ready')
    if (identity !== next.identity || editor !== next.editor) api.activate()
    return next.editor
  }
  const paint = () => {
    version++
    if (!editor || editor.isDestroyed) return
    const decorations: Decoration[] = []
    for (const entries of ranges.values())
      for (const range of entries) {
        if (
          !Number.isSafeInteger(range.from) ||
          !Number.isSafeInteger(range.to) ||
          range.from < 0 ||
          range.to <= range.from ||
          range.to > editor.state.doc.content.size
        )
          continue
        decorations.push(
          Decoration.inline(range.from, range.to, {
            class: 'sdk-range-decoration',
            style: styleAttributes(range.style ?? (range.message ? { underline: true } : {})),
            ...(range.message ? { title: range.message } : {}),
          }),
        )
      }
    editor.view.dispatch(
      editor.state.tr.setMeta(annotationKey, decorations).setMeta('addToHistory', false),
    )
    for (const listener of [...listeners]) {
      try {
        listener()
      } catch (error) {
        errors.set('onChange', error instanceof Error ? error.message : String(error))
      }
    }
  }
  const cancel = () => {
    generation++
    running?.abort()
  }
  const onTransaction = ({ transaction }: { transaction: Transaction }) => {
    if (!transaction.docChanged) return
    cancel()
    // Never display diagnostics whose old offsets now address different text.
    findings.clear()
    errors.clear()
    for (const [source, state] of sourceStates)
      sourceStates.set(source, { ...state, status: 'stale', error: undefined })
    for (const [source, values] of ranges) {
      if (rules.has(source) || checkers.has(source)) {
        ranges.delete(source)
        continue
      }
      ranges.set(
        source,
        values.flatMap((value) => {
          const from = transaction.mapping.mapResult(value.from, 1)
          const to = transaction.mapping.mapResult(value.to, -1)
          return from.deletedAcross || to.deletedAcross || from.pos >= to.pos
            ? []
            : [{ ...value, from: from.pos, to: to.pos }]
        }),
      )
      const state = sourceStates.get(source)
      if (state)
        findings.set(
          source,
          (ranges.get(source) ?? []).flatMap((value) => {
            return value.message === undefined
              ? []
              : [
                  {
                    ...value,
                    documentId: state.documentId,
                    revision: state.revision,
                    source,
                    message: value.message,
                    stale: true,
                  },
                ]
          }),
        )
    }
    paint()
    if (scheduled) return
    scheduled = true
    queueMicrotask(() => {
      scheduled = false
      if (!disposed) void api.refresh().catch(() => {})
    })
  }
  const api: DocsAnnotations = {
    reset() {
      cancel()
      checkers.clear()
      rules.clear()
      ranges.clear()
      findings.clear()
      errors.clear()
      sourceStates.clear()
      paint()
    },
    activate() {
      if (disposed) throw new Error('The annotations API is disposed')
      const next = current()
      if (!next || (next.identity === identity && next.editor === editor)) return
      cancel()
      if (editor) {
        editor.off('transaction', onTransaction)
        ranges.clear()
        paint()
      }
      checkers.clear()
      rules.clear()
      ranges.clear()
      findings.clear()
      errors.clear()
      sourceStates.clear()
      identity = next.identity
      editor = next.editor
      editor.on('transaction', onTransaction)
      paint()
    },
    setRules(source, values) {
      live()
      const compiled = values.map((rule) => {
        styleAttributes(rule.style)
        if (rule.capture !== undefined && (!Number.isSafeInteger(rule.capture) || rule.capture < 0))
          throw new Error('Invalid capture group')
        if (/[^imsu]/.test(rule.flags ?? '')) throw new Error('Invalid rule flags')
        return { rule: { ...rule }, pattern: new RegExp(rule.pattern, `g${rule.flags ?? 'u'}`) }
      })
      rules.set(source, compiled)
      checkers.delete(source)
      ranges.delete(source)
      findings.delete(source)
      errors.delete(source)
      void api.refresh().catch(() => {}) // Failure remains observable through errors()/onChange.
    },
    setRanges(source, values) {
      live()
      const snapshot = inspection.snapshot()
      for (const value of values) {
        if (value.documentId !== snapshot.documentId || value.revision !== snapshot.revision)
          throw new Error('Stale annotation reference')
        if (
          !Number.isSafeInteger(value.from) ||
          !Number.isSafeInteger(value.to) ||
          value.from < 0 ||
          value.to <= value.from ||
          value.to > editor!.state.doc.content.size
        )
          throw new Error('Invalid annotation range')
        styleAttributes(value.style)
      }
      cancel()
      rules.delete(source)
      checkers.delete(source)
      ranges.set(
        source,
        values.map((value) => ({ ...value })),
      )
      findings.set(
        source,
        values
          .filter((value) => value.message !== undefined)
          .map((value) => ({ ...value, source, message: value.message! })),
      )
      sourceStates.set(source, {
        source,
        status: 'complete',
        documentId: snapshot.documentId,
        revision: snapshot.revision,
        coverage: {
          kind: 'ranges',
          from: values.length
            ? values.reduce((from, value) => Math.min(from, value.from), Infinity)
            : 0,
          to: values.reduce((to, value) => Math.max(to, value.to), 0),
          rangeCount: values.length,
        },
      })
      paint()
      void api.refresh().catch(() => {})
    },
    registerChecker(source, checker) {
      live()
      checkers.set(source, checker)
      rules.delete(source)
      ranges.delete(source)
      findings.delete(source)
      errors.delete(source)
      sourceStates.delete(source)
      paint()
      void api.refresh().catch(() => {}) // Failure remains observable through errors()/onChange.
      return () => {
        if (checkers.get(source) === checker) api.remove(source)
      }
    },
    remove(source) {
      live()
      cancel()
      checkers.delete(source)
      rules.delete(source)
      ranges.delete(source)
      findings.delete(source)
      errors.delete(source)
      sourceStates.delete(source)
      paint()
      void api.refresh().catch(() => {})
    },
    diagnostics() {
      live()
      return [...findings.values()].flat().map((finding) => ({ ...finding }))
    },
    snapshot() {
      live()
      const reference = inspection.snapshot()
      const sources = [...sourceStates.values()].map((state) => ({
        ...state,
        coverage: { ...state.coverage },
        status:
          state.documentId !== reference.documentId || state.revision !== reference.revision
            ? ('stale' as const)
            : state.status,
      }))
      return {
        documentId: reference.documentId,
        revision: reference.revision,
        version,
        complete: sources.every((state) => state.status === 'complete'),
        sources,
        diagnostics: [...findings.values()].flat().map((finding) => ({ ...finding })),
        errors: [...errors].map(([source, message]) => ({ source, message })),
      }
    },
    errors() {
      live()
      return [...errors].map(([source, message]) => ({ source, message }))
    },
    onChange(listener) {
      live()
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    async refresh() {
      live()
      cancel()
      const token = generation
      const controller = (running = new AbortController())
      const snapshot = inspection.snapshot()
      for (const source of [...rules.keys(), ...checkers.keys()])
        sourceStates.set(source, {
          source,
          status: 'pending',
          documentId: snapshot.documentId,
          revision: snapshot.revision,
          coverage: { kind: 'document-text', from: 0, to: editor!.state.doc.content.size },
        })
      if (rules.size || checkers.size) paint()
      const apply = (source: string, values: DocsAnnotation[]) => {
        if (disposed || token !== generation || controller.signal.aborted) return
        const now = inspection.snapshot()
        if (now.documentId !== snapshot.documentId || now.revision !== snapshot.revision) return
        const accepted = values.filter(
          (value) =>
            Number.isSafeInteger(value.from) &&
            Number.isSafeInteger(value.to) &&
            value.from >= 0 &&
            value.to > value.from &&
            value.to <= editor!.state.doc.content.size,
        )
        if (accepted.length !== values.length)
          throw new Error('Host checker returned an invalid annotation range')
        for (const value of accepted) styleAttributes(value.style)
        errors.delete(source)
        const state = sourceStates.get(source)
        if (state) sourceStates.set(source, { ...state, status: 'complete', error: undefined })
        ranges.set(source, accepted)
        findings.set(
          source,
          accepted
            .filter((value) => value.message !== undefined)
            .map((value) => ({
              ...value,
              documentId: snapshot.documentId,
              revision: snapshot.revision,
              source,
              message: value.message!,
            })),
        )
        paint()
      }
      for (const [source, values] of ranges) {
        // Mapping decorations is not a new business check. Explicit findings keep
        // their original reference until the host publishes a fresh setRanges.
        const state = sourceStates.get(source)
        if (
          !rules.has(source) &&
          !checkers.has(source) &&
          state?.status === 'complete' &&
          state.documentId === snapshot.documentId &&
          state.revision === snapshot.revision
        )
          apply(source, values)
      }
      for (const [source, entries] of rules) {
        const values: DocsAnnotation[] = []
        for (const { rule, pattern } of entries) {
          pattern.lastIndex = 0
          for (const match of snapshot.text.matchAll(pattern)) {
            if (!match[0].length) continue
            if (rule.allowedValues?.includes(match[rule.capture ?? 0])) continue
            const from = snapshot.positions[match.index]
            const last = snapshot.positions[match.index + match[0].length - 1]
            if (from == null || last == null) continue
            values.push({
              from,
              to: last + 1,
              message: rule.message,
              code: rule.id,
              severity: rule.severity,
              style: rule.style,
            })
          }
        }
        apply(source, values)
      }
      await Promise.all(
        [...checkers].map(async ([source, checker]) => {
          try {
            apply(source, await checker(snapshot, controller.signal))
          } catch (error) {
            if (!controller.signal.aborted && token === generation) {
              errors.set(source, error instanceof Error ? error.message : String(error))
              const state = sourceStates.get(source)
              if (state)
                sourceStates.set(source, {
                  ...state,
                  status: 'failed',
                  error: error instanceof Error ? error.message : String(error),
                })
              ranges.delete(source)
              findings.delete(source)
              paint()
              throw error
            }
          }
        }),
      )
    },
    dispose() {
      if (disposed) return
      cancel()
      editor?.off('transaction', onTransaction)
      checkers.clear()
      rules.clear()
      ranges.clear()
      findings.clear()
      errors.clear()
      sourceStates.clear()
      paint()
      listeners.clear()
      disposed = true
    },
  }
  return api
}
