/**
 * Word 97-2003 compatibility options (Dop Copts, MS-DOC 2.9.49) -> the
 * w:compat elements Word writes when it saves the .doc as .docx. Word 2000+
 * documents carry the 32-byte Copts at Dop offset 0x1FC (its first two bytes
 * repeat DopBase.copts60 at 0x08); Word saves a Word 97 Dop without it with
 * the Word 97 layout defaults. Mapping verified against Word's own .docx of
 * corpus files (polarity and bit positions, see sections.test.ts).
 */
const COPTS60_OFFSET = 0x08
const COPTS_OFFSET = 0x1fc
const COPTS_LENGTH = 32
/** Copts of a Word 97 document as Word saves it: every "like Word 97" switch on */
const WORD97_COPTS = [0x00, 0xf0, 0x10, 0x00, 0xdf, 0xdf, 0x03, 0x02]

/** [element, byte offset in Copts, bit, set when the flag is clear] in CT_Compat order */
const COPTS_FLAGS: Array<[string, number, number, boolean?]> = [
  ['useSingleBorderforContiguousCells', 1, 0],
  ['wpJustification', 3, 6],
  ['noTabHangInd', 0, 0],
  ['noLeading', 2, 3],
  ['spaceForUL', 2, 4, true],
  ['noColumnBalance', 0, 5],
  ['balanceSingleByteDoubleByteWidth', 1, 7, true],
  ['noExtraLineSpacing', 3, 3],
  ['doNotLeaveBackslashAlone', 1, 4, true],
  ['ulTrailSpace', 1, 6, true],
  ['doNotExpandShiftReturn', 1, 5, true],
  ['lineWrapLikeWord6', 3, 1],
  ['printBodyTextBeforeHeader', 2, 2],
  ['printColBlack', 0, 4],
  ['wpSpaceWidth', 3, 5],
  ['showBreaksInFrames', 1, 2],
  ['subFontBySize', 3, 0],
  ['suppressTopSpacing', 0, 7],
  ['suppressTopSpacingWP', 2, 0],
  ['suppressSpBfAfterPgBrk', 0, 2],
  ['swapBordersFacingPages', 1, 3],
  ['convMailMergeEsc', 0, 6],
  ['truncateFontHeightsLikeWP6', 2, 7],
  ['mwSmallCaps', 2, 5],
  ['usePrinterMetrics', 3, 7],
  ['doNotSuppressParagraphBorders', 3, 2],
  ['wrapTrailSpaces', 0, 3],
  ['footnoteLayoutLikeWW8', 4, 1],
  ['shapeLayoutLikeWW8', 4, 0],
  ['alignTablesRowByRow', 4, 6],
  ['forgetLastTabAlignment', 4, 4],
  ['adjustLineHeightInTable', 4, 3, true],
  ['autoSpaceLikeWord95', 4, 5],
  ['noSpaceRaiseLower', 0, 1],
  ['doNotUseHTMLParagraphAutoSpacing', 4, 2],
  ['layoutRawTableWidth', 4, 7],
  ['layoutTableRowsApart', 5, 0],
  ['useWord97LineBreakRules', 5, 1],
  ['doNotBreakWrappedTables', 5, 2],
  ['doNotSnapToGridInCell', 5, 3],
  ['selectFldWithFirstOrLastChar', 5, 4],
  ['applyBreakingRules', 5, 5],
  ['doNotWrapTextWithPunct', 5, 6],
  ['doNotUseEastAsianBreakRules', 5, 7],
  ['useWord2002TableStyleRules', 6, 0],
  ['growAutofit', 6, 1],
]

/** sits inside the compatibility-mode 11 tail the writer always emits */
const UNDERLINE_TAB_IN_NUM_LIST: [number, number] = [7, 1]

/** The compat element names a Dop selects, plus whether underlineTabInNumList is on. */
export function compatOptionsOf(dop: Uint8Array | undefined): string[] {
  const copts = coptsOf(dop)
  const out: string[] = []
  for (const [name, byte, bit, whenClear] of COPTS_FLAGS) {
    const set = (copts[byte]! & (1 << bit)) !== 0
    if (set !== Boolean(whenClear)) out.push(name)
  }
  if (copts[UNDERLINE_TAB_IN_NUM_LIST[0]]! & (1 << UNDERLINE_TAB_IN_NUM_LIST[1]))
    out.push('underlineTabInNumList')
  return out
}

function coptsOf(dop: Uint8Array | undefined): Uint8Array {
  const copts = new Uint8Array(COPTS_LENGTH)
  copts.set(WORD97_COPTS)
  // a Word 97 Dop (no Copts) lays out like Word 97 whatever its copts60 says
  if (!dop || dop.length < COPTS_OFFSET + COPTS_LENGTH) return copts
  if (
    dop[COPTS_OFFSET] !== dop[COPTS60_OFFSET] ||
    dop[COPTS_OFFSET + 1] !== dop[COPTS60_OFFSET + 1]
  )
    return copts
  copts.set(dop.subarray(COPTS_OFFSET, COPTS_OFFSET + COPTS_LENGTH))
  return copts
}
