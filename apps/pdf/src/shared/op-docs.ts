/**
 * One-line usage signatures for every canonical edit op — the single source the
 * AI surfaces consume (same contract as apps/slides/src/shared/op-docs.ts); a test
 * asserts it covers the registry exactly. Coordinates are PDF user space points
 * (y up); page indices are 0-based original indices.
 */

export interface OpDoc {
  sig: string
  group: 'annotate' | 'text' | 'image' | 'form' | 'page' | 'document'
  /** false: executable but not advertised — payload the model cannot produce, or a dedicated tool covers it */
  aiCallable?: false
  /** Model-facing signature when apply_ops rewrites the fields (1-based pages, prefixed ids) */
  aiSig?: string
  /** Registered by an in-flight branch: hidden until it lands */
  pending?: true
}

export const OP_DOCS: Record<string, OpDoc> = {
  // ── annotate ──────────────────────────────────────────────────────────
  addMarkup: {
    sig: '{markup:{pageIndex,type:"highlight"|"underline"|"strikeout",color:[r,g,b],quads:[[x1,yT,x2,yT,x1,yB,x2,yB],…]}} — quads payload; use markup_text instead',
    group: 'annotate',
    aiCallable: false,
  },
  removeMarkup: {
    sig: '{id} — a pending markup by its record id',
    aiSig:
      '{id} — a pending (unsaved) markup by its P… id from read_annotations; saved ones need delete_markup',
    group: 'annotate',
  },
  deleteSavedAnnot: {
    sig: '{annot:{pageIndex,objNum,type,rect,…}} — a markup or note already saved in the file; use delete_markup / delete_note',
    group: 'annotate',
    aiCallable: false,
  },
  editSavedNote: {
    sig: '{annot:{pageIndex,objNum,…},contents,force?} — rewrite a saved note comment in place; use edit_note',
    group: 'annotate',
    aiCallable: false,
  },
  addDrawing: {
    sig: '{drawing:{kind:"ink"|"rect"|"ellipse"|"line"|"arrow"|"image"|"note",pageIndex,…},formWidgetId?} — geometry payload; use add_note for comments',
    group: 'annotate',
    aiCallable: false,
  },
  removeDrawing: {
    sig: '{id} — a pending drawing or note by its record id; whole threads go through delete_note',
    group: 'annotate',
    aiCallable: false,
  },
  setNoteContents: {
    sig: '{id,contents} — rewrite a pending note',
    aiSig:
      '{id,contents} — rewrite a pending (unsaved) note or reply by its P… id from read_annotations; saved ones need edit_note',
    group: 'annotate',
  },
  moveDrawing: {
    sig: '{id,dx,dy} — shift a pending drawing (points, PDF user space)',
    group: 'annotate',
    aiCallable: false,
  },
  setDrawingRect: {
    sig: '{id,rect:[x1,y1,x2,y2]} — resize a pending image stamp',
    group: 'annotate',
    aiCallable: false,
  },

  // ── text ──────────────────────────────────────────────────────────────
  putTextEdit: {
    sig: '{input:TextEditInput,id?,cover?,moveBy?} — engine record; use edit_text / edit_block instead',
    group: 'text',
    aiCallable: false,
  },
  removeTextEdit: {
    sig: '{id} — a pending text edit by its record id (ids are not exposed to the model)',
    group: 'text',
    aiCallable: false,
  },
  addTextInsert: {
    sig: '{input:TextInsertInput} — positioned text block; use insert_text instead',
    group: 'text',
    aiCallable: false,
  },
  removeTextInsert: {
    sig: '{id} — a pending text insert by its record id',
    aiSig:
      '{id} — a pending inserted-text block by its T… id from insert_text / list_inserted_text',
    group: 'text',
  },
  patchTextInsert: {
    sig: '{id,input:{text?,fontSize?,color?,origin?,…}} — change fields of a pending text insert; edit_inserted_text / move_inserted_text re-measure the lines',
    group: 'text',
    aiCallable: false,
  },

  // ── image ─────────────────────────────────────────────────────────────
  addImageEdit: {
    sig: '{input:{kind:"insertImage"|"transformImage"|"replaceImage"|"deleteImage",pageIndex,…}} — PNG payload; use the image tools instead',
    group: 'image',
    aiCallable: false,
  },
  setImageEditRect: {
    sig: '{id,rect:[x1,y1,x2,y2]} — move/resize a pending image edit (ids are not exposed to the model)',
    group: 'image',
    aiCallable: false,
  },
  bakeImageEdit: {
    sig: '{id,image,rect,opacityBase?} — PNG payload',
    group: 'image',
    aiCallable: false,
  },
  patchImageEdit: {
    sig: '{id,input?:{rect?,layer?,quarterTurns?,image?},opacityBase?:string|null} — change fields of a pending image edit (ids are not exposed to the model)',
    group: 'image',
    aiCallable: false,
  },
  setStaticFillImage: {
    sig: '{id,image,staticFill} — PNG payload',
    group: 'image',
    aiCallable: false,
  },
  removeImageEdit: {
    sig: '{id} — a pending image edit by its record id (ids are not exposed to the model)',
    group: 'image',
    aiCallable: false,
  },

  // ── form ──────────────────────────────────────────────────────────────
  setFormValue: {
    sig: '{value:{name,kind:"text"|"checkbox"|"radio"|"choice",value?,checked?}} — field names from list_form_fields',
    aiSig:
      '{value:{name,value?,checked?}} — field name from list_form_fields; checkboxes take checked, everything else value (radio/choice: an option exportValue)',
    group: 'form',
  },

  // ── page ──────────────────────────────────────────────────────────────
  rotatePages: {
    sig: '{pages:[pageIndex,…],dir:90|-90|180} — clockwise degrees added to the page rotation',
    aiSig:
      '{pages:[page,…],dir:90|-90|180} — 1-based page numbers; clockwise degrees added to each page rotation',
    group: 'page',
  },
  deletePage: {
    sig: '{pageIndex} — at least one page must remain',
    aiSig: '{page} — 1-based page number; at least one page must remain',
    group: 'page',
  },
  setPageOrder: {
    sig: '{order:[pageIndex,…]|null} — every original page exactly once, deleted ones at the tail; null restores the file order',
    aiSig:
      '{order:[page,…]|null} — every current page number exactly once in the new order; null restores the file order',
    group: 'page',
  },

  // ── document ──────────────────────────────────────────────────────────
  setStamps: {
    sig: '{cfg:{wm:WatermarkConfig|null,hf:HeaderFooterConfig|null}|null} — watermark / header-footer layer; null or both layers null clears it; use set_watermark / set_header_footer',
    group: 'document',
    aiCallable: false,
  },
  setMetadata: {
    sig: '{metadata:{title?,author?,subject?,keywords?}|null} — empty string clears a field',
    group: 'document',
  },
}

const GROUP_ORDER: OpDoc['group'][] = ['annotate', 'text', 'image', 'form', 'page', 'document']

/** Names the model may call: documented, not hidden, not pending. */
export function callableOpNames(group?: OpDoc['group']): string[] {
  return Object.keys(OP_DOCS).filter((n) => {
    const d = OP_DOCS[n]!
    return (group === undefined || d.group === group) && d.aiCallable !== false && !d.pending
  })
}

/** Grouped names of the ops the model may call — the batch tool's vocabulary. */
export function opVocabulary(): string {
  return GROUP_ORDER.filter((g) => callableOpNames(g).length > 0)
    .map((g) => `- ${g}: ${callableOpNames(g).join(', ')}`)
    .join('\n')
}

/** Every AI-callable op with its model-facing signature, grouped — the apply_ops reference. */
export function opSignatureIndex(): string {
  return GROUP_ORDER.filter((g) => callableOpNames(g).length > 0)
    .map(
      (g) =>
        `## ${g}\n${callableOpNames(g)
          .map((n) => `- ${n} ${OP_DOCS[n]!.aiSig ?? OP_DOCS[n]!.sig}`)
          .join('\n')}`,
    )
    .join('\n')
}

/** One-line usage for a failing op, appended to its guided error. */
export function opUsage(name: string): string | undefined {
  const doc = OP_DOCS[name]
  if (!doc || doc.pending) return undefined
  return `Usage: ${name} ${doc.sig}`
}
