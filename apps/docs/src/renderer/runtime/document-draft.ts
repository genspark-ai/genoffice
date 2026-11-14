import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import { DraftLanding, DRAFT_META, DOC_MAX_CHARS, type WritePosition } from './draft-landing'
import {
  insertBlocksAfter,
  isBlankDocument,
  parseHtmlFragment,
  readOnlyHtmlError,
  replaceBlockRange,
  resolveImageRefs,
  inheritBlockFormatting,
  preservedImageStructureIds,
  documentImagesOf,
  type AiTrack,
  type NumIds,
} from '../ai/protocol'
import type { DocsErrorCode, DocsImageReferenceIssue } from '../../shared/operation-result'
import { prepareDocumentImageWrite } from './document-media'
import {
  preservedInlineStructures,
  preservedTableStructures,
  validateStructureDeletion,
} from '../ai/structure-inventory'
import type {
  DocsDraftInspection,
  DocsDraftSettlement,
  DocsDraftState,
} from '../../shared/document-draft'

export interface DraftAuthority {
  documentId: string
  revision: number
  isCurrent(): boolean
}

export class DraftInputError extends Error {
  constructor(
    message: string,
    readonly errorCode: DocsErrorCode = 'invalid_input',
    readonly imageReferenceErrors?: DocsImageReferenceIssue[],
  ) {
    super(message)
  }
}

export type DraftState = DocsDraftState
export type DraftSettlement = DocsDraftSettlement

export function draftPosition(editor: Editor, input: Record<string, unknown>): WritePosition {
  if (input.afterBlockIndex !== undefined && input.afterBlockIndex !== null) {
    const index = input.afterBlockIndex
    if (
      typeof index !== 'number' ||
      !Number.isInteger(index) ||
      index < -1 ||
      index >= editor.state.doc.childCount
    )
      throw new Error(`afterBlockIndex must be -1..${editor.state.doc.childCount - 1}`)
    return { kind: 'after', index }
  }
  if (isBlankDocument(editor) || input.replaceDocument === true) return { kind: 'whole' }
  throw new Error(
    'The document is not blank: supply afterBlockIndex or explicitly request replaceDocument',
  )
}

/** Shared draft ownership and landing; inference and the keep/discard decision belong to the caller. */
export class DocumentDraft {
  private readonly landing: DraftLanding
  private readonly original: ProseMirrorNode
  private readonly originallyBlank: boolean
  private html = ''
  private sequence = 0
  private state: DraftState = 'streaming'
  private settlement?: DraftSettlement
  private nodes: ReturnType<typeof parseHtmlFragment> = []
  private deleteStructures: string[] = []
  private sourceChanged = false
  private requiresSource = false
  private parseError?: string
  private validatedSequence = 0
  private validatedCharacters = 0
  private receivedIntent: string[] = []
  private readonly observeSource = ({ transaction }: { transaction: Transaction }) => {
    if (
      !transaction.getMeta(DRAFT_META) &&
      transaction.docChanged &&
      !transaction.before.eq(transaction.doc)
    ) {
      this.sourceChanged = true
      if (this.requiresSource) this.landing.pause()
    }
  }

  constructor(
    private readonly editor: Editor,
    private readonly numIds: NumIds,
    private readonly position: WritePosition,
    private readonly valid: () => boolean = () => !editor.isDestroyed,
    private readonly currentTrack?: () => AiTrack | undefined,
    private readonly authority?: DraftAuthority,
  ) {
    this.original = editor.state.doc
    this.originallyBlank = isBlankDocument(editor)
    this.landing = new DraftLanding(
      editor,
      numIds,
      position,
      () =>
        this.valid() &&
        !editor.isDestroyed &&
        editor.isEditable &&
        !(this.requiresSource && this.sourceChanged) &&
        (!this.authority || this.authority.isCurrent()),
    )
    editor.on('transaction', this.observeSource)
  }

  get pending(): boolean {
    return this.state === 'streaming' || this.state === 'awaiting_decision'
  }

  inspect(): DocsDraftInspection {
    return {
      state: this.state,
      sequence: this.sequence,
      validatedSequence: this.validatedSequence,
      validatedCharacters: this.validatedCharacters,
      buffering: !!this.parseError,
      characters: this.html.length,
      conflicted: this.landing.hasConflict,
      ...(this.authority
        ? {
            referenceSnapshot: {
              documentId: this.authority.documentId,
              revision: this.authority.revision,
            },
          }
        : {}),
      ...(this.settlement ? { result: this.settlement } : {}),
    }
  }

  update(sequence: number, html: string, deleteStructures?: unknown): void {
    if (!this.pending || this.state !== 'streaming' || !this.valid())
      throw new Error('The draft is no longer streaming')
    if (!Number.isInteger(sequence) || sequence < 1 || sequence < this.sequence)
      throw new Error('Draft updates must use increasing positive sequence numbers')
    if (html.length > DOC_MAX_CHARS)
      throw new Error(`The draft exceeds ${DOC_MAX_CHARS} characters`)
    const readOnly = readOnlyHtmlError(html)
    if (readOnly) throw new DraftInputError(readOnly)
    const intent: string[] = []
    if (deleteStructures !== undefined) {
      if (!this.authority || !Array.isArray(deleteStructures))
        throw new DraftInputError(
          'deleteStructures requires exact references from the draft source snapshot',
        )
      const prefix = `structure:${this.authority.documentId}:${this.authority.revision}:`
      for (const id of deleteStructures) {
        if (typeof id !== 'string' || !id.startsWith(prefix))
          throw new DraftInputError('Stale or unbound draft structure reference', 'stale_target')
        intent.push(id.slice(prefix.length))
      }
    }
    if (sequence === this.sequence) {
      if (html !== this.html || JSON.stringify(intent) !== JSON.stringify(this.receivedIntent))
        throw new Error('A replayed draft update has different content or deletion intent')
      return
    }
    const requiresSource =
      /<img\b/i.test(html) ||
      intent.length > 0 ||
      (this.position.kind === 'whole' && !this.originallyBlank)
    if ((requiresSource && this.sourceChanged) || (this.authority && !this.authority.isCurrent())) {
      this.landing.pause()
      throw new DraftInputError(
        'The draft source changed; discard it and read fresh context before starting again.',
        'stale_target',
      )
    }
    // Buffer only an unfinished image token. Complete preceding tokens still pass
    // the normal resolver, and a known src prefix must name an authorized image.
    const starts = [...html.matchAll(/<img\b/gi)]
    const tailStart = starts[starts.length - 1]?.index
    const tail = tailStart === undefined ? '' : html.slice(tailStart)
    let quote = ''
    let closed = false
    for (const char of tail) {
      if (quote) {
        if (char === quote) quote = ''
      } else if (char === '"' || char === "'") quote = char
      else if (char === '>') {
        closed = true
        break
      }
    }
    const pendingImage = !!tail && !closed
    const checkedHtml = pendingImage ? html.slice(0, tailStart) : html
    const images = prepareDocumentImageWrite(
      this.original,
      checkedHtml,
      this.authority?.documentId ?? '',
      this.authority?.revision ?? -1,
    )
    if (images.errors.length)
      throw new DraftInputError(
        'Draft image references must resolve in the captured source snapshot; use its bound media references.',
        images.errors.some((issue) => issue.reason === 'expired')
          ? 'stale_target'
          : 'invalid_input',
        images.errors,
      )
    if (pendingImage) {
      if (!this.authority)
        throw new DraftInputError('Draft image references require live document authority')
      // Validate complete earlier images, including their supported placement.
      resolveImageRefs(this.original, parseHtmlFragment(images.html, this.numIds), images.html)
      const prefix = new DOMParser()
        .parseFromString(`${tail}${quote}>`, 'text/html')
        .querySelector('img')
        ?.getAttribute('src')
      const authorized: string[] = []
      this.original.forEach((block, _pos, index) =>
        documentImagesOf(block).forEach((image, ordinal) => {
          if (image.type.name === 'docInlineImage')
            authorized.push(
              `media:${this.authority!.documentId}:${this.authority!.revision}:${index}.${ordinal}`,
            )
        }),
      )
      if (prefix != null && !authorized.some((reference) => reference.startsWith(prefix)))
        throw new DraftInputError(
          'The incomplete image source does not name an authorized draft snapshot image',
        )
      this.sequence = sequence
      this.html = html
      this.receivedIntent = intent
      this.requiresSource = true
      this.parseError =
        'The draft ends in an incomplete image tag; complete it or explicitly retain the identified validated prefix.'
      return
    }
    let resolved: ReturnType<typeof resolveImageRefs>
    try {
      resolved = resolveImageRefs(
        this.original,
        parseHtmlFragment(images.html, this.numIds),
        images.html,
      )
    } catch (error) {
      // Tool-less writers may stop in the middle of a formula. Only a reference-free
      // insertion/blank draft may retain its previously validated rendered prefix.
      if (requiresSource)
        throw new DraftInputError(error instanceof Error ? error.message : String(error))
      this.sequence = sequence
      this.html = html
      this.parseError = error instanceof Error ? error.message : String(error)
      this.receivedIntent = intent
      return
    }
    if (this.position.kind === 'whole') {
      const inherited = inheritBlockFormatting(
        this.original,
        0,
        this.original.childCount - 1,
        resolved.nodes,
        false,
      )
      const kept = [
        ...preservedImageStructureIds(this.original, resolved.kept),
        ...preservedTableStructures(
          this.original,
          0,
          this.original.content.size,
          inherited.map((node) => this.editor.schema.nodeFromJSON(node)),
        ),
        ...preservedInlineStructures(
          this.original,
          0,
          this.original.content.size,
          inherited.map((node) => this.editor.schema.nodeFromJSON(node)),
          intent,
        ),
      ]
      const error = validateStructureDeletion(
        this.original,
        0,
        this.original.content.size,
        intent,
        kept,
      )
      if (error) throw new DraftInputError(error)
    } else if (intent.length)
      throw new DraftInputError('An insertion draft cannot declare structure deletion')
    this.sequence = sequence
    this.html = html
    this.nodes = resolved.nodes
    this.requiresSource = requiresSource
    this.parseError = undefined
    this.deleteStructures = intent
    this.receivedIntent = intent
    this.validatedSequence = sequence
    this.validatedCharacters = html.length
    this.landing.update(html, resolved.nodes)
  }

  stop(): void {
    if (this.state === 'streaming') this.state = 'awaiting_decision'
  }

  discard(): void {
    if (!this.pending) return
    this.landing.finish(this.valid())
    this.state = 'discarded'
    this.editor.off('transaction', this.observeSource)
  }

  finish(partial: boolean, track?: AiTrack): DraftSettlement {
    if (this.settlement) return this.settlement
    if (!this.pending || !this.valid()) {
      this.discard()
      return { ok: false, error: 'The draft binding is no longer valid' }
    }
    if (
      (this.requiresSource && this.sourceChanged) ||
      (this.authority && !this.authority.isCurrent())
    ) {
      this.landing.pause()
      return {
        ok: false,
        error: 'The draft source changed; discard it and read fresh context before starting again.',
        errorCode: 'stale_target',
      }
    }
    if (!this.editor.isEditable) {
      this.landing.pause()
      return {
        ok: false,
        error: 'The document became read-only; the draft was not committed.',
        errorCode: 'permission_denied',
      }
    }
    if (this.parseError && !partial) {
      this.landing.pause()
      return { ok: false, error: this.parseError, errorCode: 'invalid_input' }
    }
    this.landing.finish()
    this.editor.off('transaction', this.observeSource)
    const fail = (error: string, errorCode: DocsErrorCode = 'invalid_input'): DraftSettlement => {
      this.state = 'failed'
      this.settlement = { ok: false, error, errorCode }
      return this.settlement
    }
    if (this.landing.hasConflict)
      return fail(
        'The draft was edited manually; those edits were kept and automatic finalization was refused',
      )
    if (!this.html.trim()) return fail('The draft is empty; the document is unchanged')
    const readOnly = readOnlyHtmlError(this.html)
    if (readOnly) return fail(readOnly)
    const nodes = this.nodes
    if (!nodes.length) return fail('The draft did not parse into any content blocks')
    const changed = !this.editor.state.doc.eq(this.original)
    if (this.position.kind === 'whole' && changed && !this.originallyBlank)
      return fail('The document changed during the draft; the whole-document rewrite was refused')
    this.editor.view.dispatch(closeHistory(this.editor.state.tr))
    const effectiveTrack = this.currentTrack ? this.currentTrack() : track
    if (this.position.kind === 'whole' && !changed) {
      const deletedSections: number[] = []
      this.original.forEach((node, pos, index) => {
        if (
          node.attrs.sectionBoundary &&
          this.deleteStructures.includes(
            `section:${pos + node.nodeSize - 1}:${pos + node.nodeSize}`,
          )
        )
          deletedSections.push(index)
      })
      replaceBlockRange(
        this.editor,
        0,
        this.editor.state.doc.childCount - 1,
        nodes,
        effectiveTrack,
        deletedSections,
      )
    } else insertBlocksAfter(this.editor, this.landing.indexBefore(), nodes, effectiveTrack)
    this.editor.view.dispatch(closeHistory(this.editor.state.tr))
    this.state = partial ? 'retained' : 'committed'
    this.settlement = {
      ok: true,
      blocks: nodes.length,
      partial,
      ...(this.parseError
        ? {
            validatedSequence: this.validatedSequence,
            validatedCharacters: this.validatedCharacters,
            omittedTail: true,
          }
        : {}),
    }
    return this.settlement
  }
}
