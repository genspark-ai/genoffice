import type { PDFDocumentProxy } from 'pdfjs-dist'
import { VISUAL_SIGNATURE_CONTENT_PREFIX } from '../shared/ipc'

export type FormWidgetKind = 'text' | 'checkbox' | 'radio' | 'choice' | 'signature'

export interface FormWidget {
  id: string
  fieldName: string
  kind: FormWidgetKind
  pageIndex: number
  rect: [number, number, number, number]
  value: string
  checked: boolean
  multiLine: boolean
  readOnly: boolean
  required: boolean
  maxLen: number | null
  password: boolean
  comb: boolean
  textAlignment: 'left' | 'center' | 'right'
  /** radio: this button's exportValue */
  buttonValue: string
  /** choice: option list */
  options: { exportValue: string; displayValue: string }[]
  /** The /Sig field has a value or a saved visual Ink/Stamp annotation in its bounds. */
  signed: boolean
}

export interface FormField {
  name: string
  kind: FormWidgetKind
  pageIndex: number
  value: string
  checked: boolean
  required: boolean
  readOnly: boolean
  options: string[]
}

export interface FormCatalog {
  widgets: FormWidget[]
  fields: Map<string, FormField>
  byPage: Map<number, FormWidget[]>
}

/** Editable widgets in the user's current visible page order. */
export function visibleFormWidgets(
  catalog: FormCatalog | null,
  visiblePages: readonly number[],
): FormWidget[] {
  const pageOrder = new Map(visiblePages.map((pageIndex, index) => [pageIndex, index]))
  const widgets = catalog?.widgets ?? []
  const widgetOrder = new Map(widgets.map((widget, index) => [widget.id, index]))
  return widgets
    .filter((widget) => !widget.readOnly && pageOrder.has(widget.pageIndex))
    .sort(
      (a, b) =>
        pageOrder.get(a.pageIndex)! - pageOrder.get(b.pageIndex)! ||
        widgetOrder.get(a.id)! - widgetOrder.get(b.id)!,
    )
}

export interface RawFormAnnotation {
  id?: string
  subtype?: string
  fieldType?: string
  fieldName?: string
  fieldValue?: unknown
  exportValue?: string
  buttonValue?: string
  rect?: number[]
  readOnly?: boolean
  checkBox?: boolean
  radioButton?: boolean
  pushButton?: boolean
  multiLine?: boolean
  required?: boolean
  maxLen?: number
  password?: boolean
  comb?: boolean
  textAlignment?: number
  options?: { exportValue?: unknown; displayValue?: unknown }[]
  contents?: string
  contentsObj?: { str?: string }
  genOfficeFormField?: string
}

function fieldValueStr(value: unknown): string {
  const scalar = Array.isArray(value) ? value[0] : value
  return typeof scalar === 'string' ? scalar : ''
}

const textAlignment = (value: number | undefined): FormWidget['textAlignment'] =>
  value === 1 ? 'center' : value === 2 ? 'right' : 'left'

/** Convert one page of pdf.js Widget annotation data into the renderer model. */
export function widgetsFromAnnotations(annots: RawFormAnnotation[], pageIndex = 0): FormWidget[] {
  const out: FormWidget[] = []
  for (const [annotationIndex, annotation] of annots.entries()) {
    if (
      annotation.subtype !== 'Widget' ||
      !annotation.fieldName ||
      !annotation.rect ||
      annotation.pushButton
    ) {
      continue
    }
    const base = {
      id: annotation.id ?? `${pageIndex}:${annotationIndex}:${annotation.fieldName}`,
      fieldName: annotation.fieldName,
      pageIndex,
      rect: annotation.rect as FormWidget['rect'],
      checked: false,
      multiLine: false,
      readOnly: !!annotation.readOnly,
      required: !!annotation.required,
      maxLen:
        typeof annotation.maxLen === 'number' &&
        Number.isFinite(annotation.maxLen) &&
        annotation.maxLen > 0
          ? Math.floor(annotation.maxLen)
          : null,
      password: !!annotation.password,
      comb: !!annotation.comb,
      textAlignment: textAlignment(annotation.textAlignment),
      buttonValue: '',
      options: [] as FormWidget['options'],
      signed: false,
    }
    if (annotation.fieldType === 'Tx') {
      out.push({
        ...base,
        kind: 'text',
        value: fieldValueStr(annotation.fieldValue),
        multiLine: !!annotation.multiLine,
      })
    } else if (annotation.fieldType === 'Btn' && annotation.checkBox) {
      out.push({
        ...base,
        kind: 'checkbox',
        value: '',
        checked:
          typeof annotation.fieldValue === 'string' &&
          annotation.fieldValue !== 'Off' &&
          annotation.fieldValue !== '',
      })
    } else if (annotation.fieldType === 'Btn' && annotation.radioButton) {
      out.push({
        ...base,
        kind: 'radio',
        value: fieldValueStr(annotation.fieldValue),
        buttonValue: typeof annotation.buttonValue === 'string' ? annotation.buttonValue : '',
      })
    } else if (annotation.fieldType === 'Ch') {
      out.push({
        ...base,
        kind: 'choice',
        value: fieldValueStr(annotation.fieldValue),
        options: (annotation.options ?? [])
          .filter(
            (option) =>
              typeof option.exportValue === 'string' || typeof option.displayValue === 'string',
          )
          .map((option) => ({
            exportValue: String(option.exportValue ?? option.displayValue),
            displayValue: String(option.displayValue ?? option.exportValue),
          })),
      })
    } else if (annotation.fieldType === 'Sig') {
      out.push({
        ...base,
        kind: 'signature',
        value: fieldValueStr(annotation.fieldValue),
        signed:
          annotation.fieldValue !== undefined &&
          annotation.fieldValue !== null &&
          annotation.fieldValue !== '',
      })
    }
  }

  // Native Tab navigation follows DOM order. pdf.js annotation order is not reliably
  // visual, so normalize each page top-to-bottom, then left-to-right in PDF coordinates.
  return out
    .map((widget, index) => ({ widget, index }))
    .sort(
      (a, b) =>
        b.widget.rect[3] - a.widget.rect[3] ||
        a.widget.rect[0] - b.widget.rect[0] ||
        a.index - b.index,
    )
    .map(({ widget }) => widget)
}

/** Build the whole-document catalog once so ribbon, FormLayer, and AI share field semantics. */
export async function buildFormCatalog(doc: PDFDocumentProxy): Promise<FormCatalog> {
  const widgets: FormWidget[] = []
  const fields = new Map<string, FormField>()
  const byPage = new Map<number, FormWidget[]>()

  for (let pageIndex = 0; pageIndex < doc.numPages; pageIndex++) {
    const page = await doc.getPage(pageIndex + 1)
    const annotations = (await page.getAnnotations()) as RawFormAnnotation[]
    const pageWidgets = widgetsFromAnnotations(annotations, pageIndex)
    const visualSignatures = annotations.filter(
      (annotation) =>
        (annotation.subtype === 'Ink' || annotation.subtype === 'Stamp') &&
        annotation.rect?.length === 4,
    )
    for (const widget of pageWidgets) {
      if (widget.kind !== 'signature' || widget.signed) continue
      const explicitlyAssociated = visualSignatures.some(
        (annotation) => visualSignatureFieldName(annotation) === widget.fieldName,
      )
      const legacyOverlap = visualSignatures.some(
        (annotation) =>
          visualSignatureFieldName(annotation) === null &&
          rectsSubstantiallyOverlap(widget.rect, annotation.rect as number[]),
      )
      widget.signed = explicitlyAssociated || legacyOverlap
    }
    byPage.set(pageIndex, pageWidgets)
    widgets.push(...pageWidgets)

    for (const widget of pageWidgets) {
      const existing = fields.get(widget.fieldName)
      if (existing) {
        existing.required ||= widget.required
        existing.readOnly &&= widget.readOnly
        if (
          widget.kind === 'radio' &&
          widget.buttonValue &&
          !existing.options.includes(widget.buttonValue)
        ) {
          existing.options.push(widget.buttonValue)
        }
        continue
      }
      fields.set(widget.fieldName, {
        name: widget.fieldName,
        kind: widget.kind,
        pageIndex,
        value: widget.value,
        checked: widget.checked,
        required: widget.required,
        readOnly: widget.readOnly,
        options:
          widget.kind === 'radio'
            ? widget.buttonValue
              ? [widget.buttonValue]
              : []
            : widget.options.map((option) => option.exportValue),
      })
    }
  }

  return { widgets, fields, byPage }
}

function visualSignatureFieldName(annotation: RawFormAnnotation): string | null {
  if (annotation.genOfficeFormField) return annotation.genOfficeFormField
  const contents = annotation.contentsObj?.str ?? annotation.contents
  return contents?.startsWith(VISUAL_SIGNATURE_CONTENT_PREFIX)
    ? contents.slice(VISUAL_SIGNATURE_CONTENT_PREFIX.length)
    : null
}

function rectsSubstantiallyOverlap(a: number[], b: number[]): boolean {
  const ax1 = Math.min(a[0]!, a[2]!)
  const ay1 = Math.min(a[1]!, a[3]!)
  const ax2 = Math.max(a[0]!, a[2]!)
  const ay2 = Math.max(a[1]!, a[3]!)
  const bx1 = Math.min(b[0]!, b[2]!)
  const by1 = Math.min(b[1]!, b[3]!)
  const bx2 = Math.max(b[0]!, b[2]!)
  const by2 = Math.max(b[1]!, b[3]!)
  const intersection =
    Math.max(0, Math.min(ax2, bx2) - Math.max(ax1, bx1)) *
    Math.max(0, Math.min(ay2, by2) - Math.max(ay1, by1))
  const smallerArea = Math.min((ax2 - ax1) * (ay2 - ay1), (bx2 - bx1) * (by2 - by1))
  return smallerArea > 0 && intersection / smallerArea >= 0.5
}

/** Lightweight pre-parse warning for XFA or mixed AcroForm/XFA documents. */
export function hasXfaMarker(bytes: Uint8Array): boolean {
  const marker = [0x2f, 0x58, 0x46, 0x41] // /XFA
  outer: for (let index = 0; index <= bytes.length - marker.length; index++) {
    for (let offset = 0; offset < marker.length; offset++) {
      if (bytes[index + offset] !== marker[offset]) continue outer
    }
    return true
  }
  return false
}

export function documentFormFeatures(
  info: { EncryptFilterName?: string | null; IsXFAPresent?: boolean },
  bytes: Uint8Array,
): { encrypted: boolean; hasXfa: boolean } {
  return {
    encrypted: info.EncryptFilterName != null,
    // XFA entries are often compressed, so PDF.js metadata is authoritative.
    // The byte marker remains useful for malformed files with incomplete metadata.
    hasXfa: info.IsXFAPresent === true || hasXfaMarker(bytes),
  }
}
