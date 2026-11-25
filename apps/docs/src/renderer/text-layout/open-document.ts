/**
 * The document-level text-layout facts of the open document, installed once by
 * file-actions from `documentTextContext(parsed)`. The line-metrics readers
 * that used to each hold their own module flag read this record instead.
 */
export interface OpenDocumentText {
  /** balanceSingleByteDoubleByteWidth widens spaces (DocumentTextContext.balancedSpaces) */
  balancedSpaces: boolean
  /** docDefaults w:lang lays every run's Latin out with its eastAsia face */
  latinViaEaSlot: boolean
  /** the eastAsia slot as declared (Normal, then docDefaults), Latin names included */
  bodyEaSlot: string | null
  /** w:compat doNotUseHTMLParagraphAutoSpacing: adjacent paragraph spacing adds up */
  additiveSpacing: boolean
  /** docDefaults w:lang w:eastAsia as declared */
  eastAsiaLang: string | null
  /** compatibilityMode below 15: a character grid moves the right indent (w:adjustRightInd) */
  legacyLayout: boolean
}

const NONE: OpenDocumentText = {
  balancedSpaces: false,
  latinViaEaSlot: false,
  bodyEaSlot: null,
  additiveSpacing: false,
  eastAsiaLang: null,
  legacyLayout: false,
}

let current: OpenDocumentText = NONE

/** null (or a blank document) restores every switch to its off state */
export function setOpenDocumentText(ctx: Partial<OpenDocumentText> | null): void {
  current = ctx
    ? {
        balancedSpaces: ctx.balancedSpaces ?? NONE.balancedSpaces,
        latinViaEaSlot: ctx.latinViaEaSlot ?? NONE.latinViaEaSlot,
        bodyEaSlot: ctx.bodyEaSlot ?? NONE.bodyEaSlot,
        additiveSpacing: ctx.additiveSpacing ?? NONE.additiveSpacing,
        eastAsiaLang: ctx.eastAsiaLang ?? NONE.eastAsiaLang,
        legacyLayout: ctx.legacyLayout ?? NONE.legacyLayout,
      }
    : NONE
}

export function openDocumentText(): OpenDocumentText {
  return current
}
