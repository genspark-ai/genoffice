import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { Fragment, Slice, type Mark, type Node as PmNode } from '@tiptap/pm/model'
import type { ZoteroRendererRequest } from '../../shared/ipc'
import { parseZoteroRtf, zoteroRtfToText, type ZoteroRtfParagraph, type ZoteroRtfRun } from './rtf'

const ZOTERO_INSTR_RE = /^\s*(?:ADDIN\s+)?(?:ZOTERO_|CSL_)(?:ITEM|BIBL|TEMP)\b/i

interface ZoteroFieldRange {
  id: number
  from: number
  to: number
  blockFrom: number
  blockTo: number
  mark: Mark
  text: string
  segments: ZoteroFieldSegment[]
  blocks: ZoteroFieldBlock[]
}

interface ZoteroFieldSegment {
  from: number
  to: number
  blockFrom: number
  mark: Mark
  text: string
}

interface ZoteroFieldBlock {
  from: number
  to: number
  node: PmNode
}

export interface ZoteroDocumentDataAccess {
  get(): string
  set(value: string): void
}

export function zoteroCodeFromInstruction(instruction: string): string {
  return instruction
    .trim()
    .replace(/^ADDIN\s+ZOTERO_/i, '')
    .replace(/^ZOTERO_/i, '')
    .replace(/^CSL_/i, '')
    .trim()
}

export function zoteroInstructionFromCode(code: string): string {
  const normalized = zoteroCodeFromInstruction(code)
  return `ADDIN ZOTERO_${normalized || 'TEMP'}`
}

export class ZoteroDocumentController {
  private nextFieldId = 1

  constructor(
    private readonly editor: Editor,
    private readonly documentData: ZoteroDocumentDataAccess,
  ) {}

  async handle(request: ZoteroRendererRequest): Promise<unknown> {
    const args = request.args
    switch (request.command) {
      case 'Application_getActiveDocument':
        return [3, 1]
      case 'Document_activate':
        this.editor.commands.focus()
        return null
      case 'Document_displayAlert':
        return this.displayAlert(String(args[1] ?? ''), Number(args[3] ?? 0))
      case 'Document_canInsertField':
        return this.editor.isEditable
      case 'Document_cursorInField': {
        const field = this.fieldAtSelection()
        return field ? this.fieldDescriptor(field) : null
      }
      case 'Document_getDocumentData':
        return this.documentData.get()
      case 'Document_setDocumentData':
        this.documentData.set(String(args[1] ?? ''))
        return null
      case 'Document_insertField':
        return this.insertField(Number(args[2] ?? 0))
      case 'Document_getFields': {
        const fields = this.fields()
        return [
          fields.map((field) => field.id),
          fields.map((field) => zoteroCodeFromInstruction(String(field.mark.attrs.instr))),
          fields.map(() => 0),
          fields.map((field, index) =>
            index + 1 < fields.length && field.to === fields[index + 1].from ? 0 : -1,
          ),
        ]
      }
      case 'Document_setBibliographyStyle':
      case 'Document_cleanup':
      case 'Document_complete':
        return null
      case 'Document_insertText': {
        // plain text: insertContent would parse Zotero's output as HTML
        const { state, view } = this.editor
        view.dispatch(state.tr.insertText(zoteroRtfToText(String(args[1] ?? ''))).scrollIntoView())
        this.editor.commands.focus()
        return null
      }
      case 'Document_convertPlaceholdersToFields':
        return [[], [], [], []]
      case 'Document_exportDocument':
        throw new Error('Zotero document transfer export is not supported yet')
      case 'Document_importDocument':
        throw new Error('Zotero document transfer import is not supported yet')
      case 'Field_getText':
        return this.requireField(args[1]).text
      case 'Field_setText':
        this.setFieldText(args[1], String(args[2] ?? ''), args[3] === true)
        return null
      case 'Field_getCode':
        return zoteroCodeFromInstruction(String(this.requireField(args[1]).mark.attrs.instr))
      case 'Field_setCode':
        this.setFieldCode(args[1], String(args[2] ?? ''))
        return null
      case 'Field_select':
        this.selectField(args[1])
        return null
      case 'Field_delete':
        this.deleteField(args[1])
        return null
      case 'Field_removeCode':
        this.removeFieldCode(args[1])
        return null
      case 'Field_convert':
        if (Number(args[3] ?? 0) !== 0) {
          throw new Error('Zotero footnote and endnote citations are not supported yet')
        }
        return null
      default:
        throw new Error(`Unsupported Zotero command: ${request.command}`)
    }
  }

  /** Runs that lost their id (clipboard round trips keep the instruction but not the
   *  runtime attrs) get one written back onto the mark, so the id handed to Zotero in
   *  Document_getFields still resolves in the Field_* calls that follow. Adjacent runs
   *  with the same instruction form one field again. */
  private assignMissingFieldIds(): void {
    const { state } = this.editor
    const type = state.schema.marks.instrField
    if (!type) return
    interface LostRun {
      from: number
      to: number
      mark: Mark
      blockIndex: number
      blockStart: number
      blockEnd: number
    }
    const lost: LostRun[] = []
    let maxId = 0
    state.doc.forEach((block, blockOffset, blockIndex) => {
      block.descendants((node, relativePos) => {
        if (!node.isText) return
        const mark = node.marks.find(
          (candidate) =>
            candidate.type === type && ZOTERO_INSTR_RE.test(String(candidate.attrs.instr ?? '')),
        )
        if (!mark) return
        const id = Number(mark.attrs.fieldId)
        if (Number.isSafeInteger(id) && id > 0) maxId = Math.max(maxId, id)
        else {
          const from = blockOffset + 1 + relativePos
          lost.push({
            from,
            to: from + node.nodeSize,
            mark,
            blockIndex,
            blockStart: blockOffset + 1,
            blockEnd: blockOffset + block.nodeSize - 1,
          })
        }
      })
    })
    if (lost.length === 0) return
    this.nextFieldId = Math.max(this.nextFieldId, maxId + 1)
    // only runs that touch (or end one paragraph and open the next) rejoin: two
    // copies of a citation with text between them stay two fields
    const touches = (prev: LostRun, next: LostRun) =>
      prev.blockIndex === next.blockIndex
        ? prev.to === next.from
        : next.blockIndex === prev.blockIndex + 1 &&
          prev.to === prev.blockEnd &&
          next.from === next.blockStart
    const groups: LostRun[][] = []
    for (const entry of lost) {
      const group = groups[groups.length - 1]
      const prev = group?.[group.length - 1]
      if (prev && prev.mark.attrs.instr === entry.mark.attrs.instr && touches(prev, entry)) {
        group.push(entry)
      } else groups.push([entry])
    }
    let transaction = state.tr
    for (const group of groups) {
      const fieldId = this.nextFieldId++
      group.forEach((entry, index) => {
        const fieldPart =
          group.length === 1
            ? 'single'
            : index === 0
              ? 'begin'
              : index === group.length - 1
                ? 'end'
                : 'inside'
        transaction = transaction
          .removeMark(entry.from, entry.to, type)
          .addMark(entry.from, entry.to, type.create({ ...entry.mark.attrs, fieldId, fieldPart }))
      })
    }
    this.editor.view.dispatch(transaction)
  }

  private fields(): ZoteroFieldRange[] {
    this.assignMissingFieldIds()
    const ranges = new Map<number, ZoteroFieldRange>()
    this.editor.state.doc.forEach((block, blockOffset) => {
      block.descendants((node, relativePos) => {
        if (!node.isText || !node.text) return
        const mark = node.marks.find(
          (candidate) =>
            candidate.type.name === 'instrField' &&
            ZOTERO_INSTR_RE.test(String(candidate.attrs.instr ?? '')),
        )
        if (!mark) return
        let id = Number(mark.attrs.fieldId)
        if (!Number.isSafeInteger(id) || id <= 0) id = this.nextFieldId++
        else this.nextFieldId = Math.max(this.nextFieldId, id + 1)
        const from = blockOffset + 1 + relativePos
        const to = from + node.nodeSize
        const segment: ZoteroFieldSegment = {
          from,
          to,
          blockFrom: blockOffset,
          mark,
          text: node.text,
        }
        const existing = ranges.get(id)
        if (existing) {
          const previous = existing.segments[existing.segments.length - 1]
          if (previous.blockFrom === blockOffset && previous.to === from) {
            previous.to = to
            previous.text += node.text
          } else {
            existing.segments.push(segment)
          }
          existing.to = to
          existing.blockTo = blockOffset + block.nodeSize
          if (!existing.blocks.some((entry) => entry.from === blockOffset)) {
            existing.blocks.push({
              from: blockOffset,
              to: blockOffset + block.nodeSize,
              node: block,
            })
          }
        } else {
          ranges.set(id, {
            id,
            from,
            to,
            blockFrom: blockOffset,
            blockTo: blockOffset + block.nodeSize,
            mark,
            text: '',
            segments: [segment],
            blocks: [{ from: blockOffset, to: blockOffset + block.nodeSize, node: block }],
          })
        }
      })
    })
    return [...ranges.values()]
      .sort((a, b) => a.from - b.from)
      .map((field) => {
        const textByBlock = field.blocks.map((block) =>
          field.segments
            .filter((segment) => segment.blockFrom === block.from)
            .map((segment) => segment.text)
            .join(''),
        )
        return { ...field, text: textByBlock.join('\n') }
      })
  }

  private fieldAtSelection(): ZoteroFieldRange | null {
    const { from, to } = this.editor.state.selection
    // the inline span from the first to the last field run: text typed inside a
    // bibliography counts, text sharing its first or last paragraph does not
    return this.fields().find((field) => from >= field.from && to <= field.to) ?? null
  }

  private requireField(value: unknown): ZoteroFieldRange {
    const id = Number(value)
    const field = this.fields().find((candidate) => candidate.id === id)
    if (!field) throw new Error(`Zotero field ${String(value)} no longer exists`)
    return field
  }

  private fieldDescriptor(field: ZoteroFieldRange): [number, string, number] {
    return [field.id, zoteroCodeFromInstruction(String(field.mark.attrs.instr)), 0]
  }

  private insertField(noteType: number): [number, string, number] {
    if (noteType !== 0) {
      throw new Error('Zotero footnote and endnote citations are not supported yet')
    }
    // fields() may dispatch (missing ids get assigned), so read the state after it
    this.fields()
    const { state, view } = this.editor
    const type = state.schema.marks.instrField
    if (!type) throw new Error('This document cannot store Zotero fields')
    const id = this.nextFieldId++
    const mark = type.create({
      instr: 'ADDIN ZOTERO_TEMP',
      beginXml: null,
      fieldId: id,
      fieldPart: 'single',
    })
    view.dispatch(state.tr.replaceSelectionWith(state.schema.text('{Citation}', [mark]), false))
    return [id, 'TEMP', 0]
  }

  private setFieldText(value: unknown, input: string, rich: boolean): void {
    const field = this.requireField(value)
    const paragraphs: ZoteroRtfParagraph[] = rich
      ? parseZoteroRtf(input).paragraphs
      : input.split(/\r?\n/).map((text) => ({ runs: text ? [{ text }] : [] }))
    const prepared = paragraphs.map((paragraph) => ({
      ...paragraph,
      runs: paragraph.runs.length > 0 ? paragraph.runs : [{ text: ' ' }],
    }))
    const totalRuns = prepared.reduce((total, paragraph) => total + paragraph.runs.length, 0)
    let runIndex = 0
    const markFor = (fieldPart: 'single' | 'begin' | 'inside' | 'end') =>
      field.mark.type.create({ ...field.mark.attrs, fieldId: field.id, fieldPart })
    const contentFor = (paragraph: (typeof prepared)[number]) =>
      paragraph.runs.map((run) => {
        const fieldPart =
          totalRuns === 1
            ? 'single'
            : runIndex === 0
              ? 'begin'
              : runIndex === totalRuns - 1
                ? 'end'
                : 'inside'
        runIndex++
        return this.editor.state.schema.text(run.text, [markFor(fieldPart), ...this.rtfMarks(run)])
      })
    let transaction = this.editor.state.tr
    if (prepared.length === 1 && field.blocks.length === 1) {
      transaction = transaction.replaceWith(
        field.from,
        field.to,
        Fragment.fromArray(contentFor(prepared[0])),
      )
    } else {
      const nodes = prepared.map((paragraph, index) => {
        const existing = field.blocks[index]
        const template = existing ?? field.blocks[0]
        const attrs = { ...template.node.attrs }
        if (!existing && 'docxIndex' in attrs) attrs.docxIndex = null
        for (const key of [
          'indentLeft',
          'indentRight',
          'indentFirstLine',
          'spaceBefore',
          'spaceAfter',
          'align',
        ] as const) {
          if (paragraph[key] !== undefined) attrs[key] = paragraph[key]
        }
        return template.node.type.create(attrs, Fragment.fromArray(contentFor(paragraph)))
      })
      const fillsBlocks = field.from === field.blockFrom + 1 && field.to === field.blockTo - 1
      // a field sharing its first or last paragraph with other text (a "References"
      // heading line, a citation with a paragraph break in its RTF) keeps that text:
      // the open slice joins the new first/last paragraph into the existing ones
      transaction = fillsBlocks
        ? transaction.replaceWith(field.blockFrom, field.blockTo, Fragment.fromArray(nodes))
        : transaction.replace(field.from, field.to, new Slice(Fragment.fromArray(nodes), 1, 1))
    }
    this.editor.view.dispatch(transaction)
  }

  private rtfMarks(run: ZoteroRtfRun): Mark[] {
    const marks: Mark[] = []
    const schema = this.editor.state.schema
    if (run.bold && schema.marks.bold) marks.push(schema.marks.bold.create())
    if (run.italic && schema.marks.italic) marks.push(schema.marks.italic.create())
    if (run.underline && schema.marks.underline) marks.push(schema.marks.underline.create())
    if (run.strike && schema.marks.strike) marks.push(schema.marks.strike.create())
    if ((run.vertAlign || run.caps || run.sizeHalfPoints) && schema.marks.docTextStyle) {
      marks.push(
        schema.marks.docTextStyle.create({
          vertAlign: run.vertAlign ?? null,
          caps: run.caps ?? null,
          sizeHalfPoints: run.sizeHalfPoints ?? null,
        }),
      )
    }
    return marks
  }

  private setFieldCode(value: unknown, code: string): void {
    const field = this.requireField(value)
    const type = field.mark.type
    let transaction = this.editor.state.tr
    for (const segment of field.segments) {
      transaction = transaction.removeMark(segment.from, segment.to, type).addMark(
        segment.from,
        segment.to,
        type.create({
          ...segment.mark.attrs,
          instr: zoteroInstructionFromCode(code),
          fieldId: field.id,
        }),
      )
    }
    this.editor.view.dispatch(transaction)
  }

  private selectField(value: unknown): void {
    const field = this.requireField(value)
    this.editor.view.dispatch(
      this.editor.state.tr.setSelection(
        TextSelection.create(this.editor.state.doc, field.from, field.to),
      ),
    )
    this.editor.commands.focus()
  }

  private deleteField(value: unknown): void {
    const field = this.requireField(value)
    const fillsBlocks = field.from === field.blockFrom + 1 && field.to === field.blockTo - 1
    const from = field.blocks.length > 1 && fillsBlocks ? field.blockFrom : field.from
    const to = field.blocks.length > 1 && fillsBlocks ? field.blockTo : field.to
    this.editor.view.dispatch(this.editor.state.tr.delete(from, to))
  }

  private removeFieldCode(value: unknown): void {
    const field = this.requireField(value)
    let transaction = this.editor.state.tr
    for (const segment of field.segments) {
      transaction = transaction.removeMark(segment.from, segment.to, field.mark.type)
    }
    this.editor.view.dispatch(transaction)
  }

  /** Zotero button sets: 0 OK, 1 OK/Cancel, 2 Yes/No, 3 Yes/No/Cancel; the reply is the
   *  button index counted from Cancel/No = 0, so Yes is 2 in the three-button set */
  private displayAlert(message: string, buttons: number): number {
    if (buttons > 0) return window.confirm(message) ? (buttons === 3 ? 2 : 1) : 0
    window.alert(message)
    return 0
  }
}
