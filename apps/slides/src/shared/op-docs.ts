/**
 * One-line usage signatures for every canonical edit op — the single source
 * all AI surfaces consume, kept out of both prompt and handler so it cannot
 * drift: a test asserts the table covers the registry exactly.
 *
 * Consumers:
 * - apply_ops tool description (renderer): the grouped vocabulary of
 *   AI-callable ops — names only, so the model knows what EXISTS at a cost of
 *   a couple hundred tokens.
 * - op executor guided errors (main): a failing op's message appends its
 *   signature — the model learns the exact fields on first contact instead of
 *   needing a resident field reference, and dry_run rehearses a whole batch
 *   for free.
 *
 * Signatures list the fields beside `target`; units are document-space EMU
 * (1 px = 9525 EMU) and pt for font sizes, matching the op layer.
 */

export interface OpDoc {
  /** Compact usage signature: the fields beside target (? marks optional). */
  sig: string
  /** Vocabulary group in the tool description. */
  group: 'text' | 'element' | 'insert' | 'table' | 'slide' | 'deck'
  /**
   * false: reachable through the executor but not advertised to the model —
   * the payload is bytes/clipboard/part-path data the model cannot produce,
   * or a dedicated tool already covers it better.
   */
  aiCallable?: false
  /**
   * Registered by an in-flight branch: hidden from the vocabulary AND from
   * usage lines until it actually lands (advertising it would make the model
   * call an unknown op; a usage line on the unknown-op error would read as a
   * field problem). Drop the flag when the branch merges.
   */
  pending?: true
}

export const OP_DOCS: Record<string, OpDoc> = {
  // ── text ──────────────────────────────────────────────────────────────
  setText: {
    sig: '{paragraphs:[{runs:[{text,bold?,italic?,fontSize?,color?}],align?}]} (group children: add group:"<group id>")',
    group: 'text',
  },
  setFont: {
    sig: '{font:{fontFamily?,fontSizePt?,bold?,italic?,underline?,strike?,color?}} — merges onto every run of the element',
    group: 'text',
  },
  setParagraphFormat: {
    sig: '{format:{align?,bullet?,lineSpacingPct?,spaceBeforePt?,spaceAfterPt?}}',
    group: 'text',
  },

  // ── element ───────────────────────────────────────────────────────────
  deleteElement: { sig: '{} — no fields besides target', group: 'element' },
  setFill: {
    sig: '{fill:"#RRGGBB"|"none"|{stops:[{pos,color}],angle?}}',
    group: 'element',
  },
  setStroke: {
    sig: '{stroke:{color:"#RRGGBB",widthEmu}|null} — null removes the outline',
    group: 'element',
  },
  setTransform: {
    sig: '{box:{x,y,cx,cy},rotDeg?} (group children: absBox instead of box, plus group:"<group id>")',
    group: 'element',
  },
  setConnectorEndpoints: {
    sig: '{p1:{x,y},p2:{x,y},start?:{targetId,idx}|null,end?:{targetId,idx}|null}',
    group: 'element',
  },
  flipElements: { sig: '{els:[id,…],axis:"h"|"v"} — target.el unused', group: 'element' },
  setPictureSrcRect: {
    sig: '{srcRect:{l,t,r,b}|null} — crop fractions 0..1, null removes the crop',
    group: 'element',
  },
  setPictureOpacity: { sig: '{opacity:0..1}', group: 'element' },
  reorderElement: { sig: '{dir:"front"|"back"|"forward"|"backward"}', group: 'element' },
  groupElements: { sig: '{els:[id,id,…]} — at least two ids; target.el unused', group: 'element' },
  ungroupElement: { sig: '{} — target.el is the group', group: 'element' },
  setShapeGeometry: { sig: '{prst:"<OOXML preset geometry name>"}', group: 'element' },
  setShapeAdjust: {
    sig: '{adjust:{<gdName>:number,…}} — preset-geometry avLst values (group children: add group)',
    group: 'element',
  },
  setTextAnchor: { sig: '{anchor:"top"|"middle"|"bottom"}', group: 'element' },
  setLink: {
    sig: '{link:{kind:"url",url}|{kind:"slide",slideIndex}|null}',
    group: 'element',
  },
  setImageFill: {
    sig: '{source:{mediaPath}|{bytes,ext},tile?} — bytes payload; use the image tools instead',
    group: 'element',
    aiCallable: false,
  },

  // ── insert ────────────────────────────────────────────────────────────
  addElement: {
    sig: '{kind:"textbox"|<preset geometry>,offset:{x,y,cx,cy},paragraphs?,fill?,stroke?}',
    group: 'insert',
  },
  addPicture: {
    sig: '{bytes,ext,offset} — bytes payload; use insert_web_image instead',
    group: 'insert',
    aiCallable: false,
  },
  replacePicture: {
    sig: '{bytes,ext} — bytes payload; use replace_image instead',
    group: 'insert',
    aiCallable: false,
  },
  addTable: { sig: '{rows,cols,offset:{x,y,cx,cy}}', group: 'insert' },
  addChart: {
    sig: '{kind:"bar"|"barStacked"|"line"|"area"|"pie"|"doughnut"|"scatter"|"radar"|"comboBarLine",categories:[…],series:[{name,values:[…]}],offset,title?}',
    group: 'insert',
  },
  addSmartArt: { sig: '{layout,items:[…],offset}', group: 'insert' },
  addMedia: {
    sig: '{kind:"video"|"audio",bytes,ext,offset} — bytes payload',
    group: 'insert',
    aiCallable: false,
  },
  addModel3d: { sig: '{bytes,ext,offset} — bytes payload', group: 'insert', aiCallable: false },
  pasteElements: {
    sig: '{items,dx,dy} — clipboard payload',
    group: 'insert',
    aiCallable: false,
  },

  // ── table (target.el = the table) ─────────────────────────────────────
  setTableCell: { sig: '{row,col,paragraphs}', group: 'table' },
  tableMerge: { sig: '{kind:"merge-right"|"merge-down"|"split",row,col}', group: 'table' },
  tableStructure: {
    sig: '{kind:"insert-row"|"delete-row"|"insert-col"|"delete-col",index,before?}',
    group: 'table',
  },
  setTableRowHeight: { sig: '{row,hEmu}', group: 'table' },
  setTableCellAnchor: { sig: '{row,col,anchor:"top"|"middle"|"bottom"}', group: 'table' },
  setTableColWidth: { sig: '{col,wEmu}', group: 'table' },
  setTableStyle: {
    sig: '{edit:TableStyleEdit} — use the edit_table_style tool instead',
    group: 'table',
    aiCallable: false,
  },
  setChart: {
    sig: '{patch:ChartEdit} — use the edit_chart tool instead',
    group: 'table',
    aiCallable: false,
  },

  // ── slide (target = the slide) ────────────────────────────────────────
  deleteSlide: { sig: '{}', group: 'slide' },
  duplicateSlide: { sig: '{}', group: 'slide' },
  addBlankSlide: { sig: '{} — inserts after target.slide', group: 'slide' },
  addSlideWithLayout: {
    sig: '{layoutPath} — layout part path; not discoverable from tool results',
    group: 'slide',
    aiCallable: false,
  },
  pasteSlide: {
    sig: '{afterIndex,bundle|png,mode?} — clipboard payload',
    group: 'slide',
    aiCallable: false,
  },
  moveSlide: { sig: '{to} — 0-based destination index', group: 'slide' },
  setSlideLayout: {
    sig: '{layoutPath?} — layout part path; not discoverable from tool results',
    group: 'slide',
    aiCallable: false,
  },
  setBackground: {
    sig: '{kind:"solid"|"gradient"|"reset"|"graphics"|"image",color?,from?,to?,angleDeg?,radial?,hidden?} — image kind needs a bytes source (use set_slide_background)',
    group: 'slide',
  },
  setHidden: { sig: '{hidden:boolean}', group: 'slide' },
  setTransition: { sig: '{kind:"none"|"fade"|"push"|"wipe"|…}', group: 'slide' },
  setAdvanceTime: { sig: '{ms:number|null} — auto-advance; null clears', group: 'slide' },
  setAnimations: {
    sig: '{items:[{spid,effect,trigger,durationMs,delayMs,…}]} — spid-addressed (cNvPr id), no id translation yet',
    group: 'slide',
    aiCallable: false,
  },
  setNotes: { sig: '{text} — speaker notes', group: 'slide' },
  addComment: { sig: '{text,author}', group: 'slide' },
  deleteComment: { sig: '{authorId,idx}', group: 'slide' },

  // ── deck-wide (no target) ─────────────────────────────────────────────
  setSlideSize: { sig: '{cx,cy} — EMU page size, whole deck', group: 'deck' },
  findReplace: {
    sig: '{find,replace,matchCase?,slideIndex?} — whole deck unless slideIndex',
    group: 'deck',
  },
  applyHeaderFooter: {
    sig: '{settings:{footer?,slideNum?,date?,dateAuto?}} — every slide',
    group: 'deck',
  },
  setSections: { sig: '{sections:[{id,name,slideIndices},…]} — full replace', group: 'deck' },
  applyTheme: {
    sig: '{name,colors:{dk1?,lt1?,dk2?,lt2?,accent1?..accent6?,hlink?,folHlink?},majorFont?,minorFont?} — "#RRGGBB" slots, whole deck',
    group: 'deck',
  },
  addSection: { sig: '{atSlideIndex,name}', group: 'deck' },
  renameSection: { sig: '{id,name}', group: 'deck' },
  removeSection: { sig: '{id} — keeps the slides', group: 'deck' },
  moveSection: { sig: '{id,dir:"up"|"down"}', group: 'deck' },
}

const GROUP_ORDER: OpDoc['group'][] = ['text', 'element', 'insert', 'table', 'slide', 'deck']

/** Grouped names of AI-callable ops, one line per group — the tool description's vocabulary block. */
export function opVocabulary(): string {
  return GROUP_ORDER.map((g) => {
    const names = Object.keys(OP_DOCS).filter(
      (n) => OP_DOCS[n]!.group === g && OP_DOCS[n]!.aiCallable !== false && !OP_DOCS[n]!.pending,
    )
    return `- ${g}: ${names.join(', ')}`
  }).join('\n')
}

/** One-line usage for a failing op, appended to its guided error. */
export function opUsage(name: string): string | undefined {
  const doc = OP_DOCS[name]
  if (!doc || doc.pending) return undefined
  return `Usage: ${name} ${doc.sig}`
}
