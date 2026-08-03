/**
 * Pure mapping logic for slides:edit-text: paragraph/run structures sent by the editor -> model
 * paragraphs. Paragraphs/runs trace back to the original paragraph/run via the source indices the
 * editor returns (srcPara/srcRun) — splitting, inserting, or deleting paragraphs no longer shifts
 * by position (when the browser splits a div on Enter, the data attributes are copied along, so
 * both halves inherit from the same source); the ...oldRun
 * spread preserves fields the editor cannot express
 * (letterSpacing/field/outline, and hyperlinks whose rId didn't resolve).
 */
import { encodeRunLink, type Paragraph, type ParagraphFormatPatch } from '@genoffice/pptx-engine'
import type { EditParagraph } from '../shared/ipc'

/** Normalize a color to 6-digit uppercase hex (strip #/alpha), for comparison. */
function hex6(c?: string): string | undefined {
  return c?.replace(/^#/, '').slice(0, 6).toUpperCase()
}

/**
 * Font-name normalization (for comparison): the editor returns the first entry of the display
 * font stack (the English name from konva-adapter FONT_STACK); treat it as the same font as the
 * Chinese name in the model, so runs whose font was not changed are not misjudged as changed and
 * lose their latin/ea preservation markers.
 */
const FONT_ALIAS: Record<string, string> = {
  微软雅黑: 'microsoft yahei',
  苹方: 'pingfang sc',
  宋体: 'simsun',
  黑体: 'simhei',
  楷体: 'kaiti',
  仿宋: 'fangsong',
  等线: 'dengxian',
}
function fontKey(name?: string): string | undefined {
  if (!name) return undefined
  return FONT_ALIAS[name] ?? name.toLowerCase()
}

export function applyEditParagraphs(oldParas: Paragraph[], edited: EditParagraph[]): Paragraph[] {
  return edited.map((p, pi) => {
    const oldPara = oldParas[p.srcPara ?? pi]
    return {
      ...oldPara, // keep unedited paragraph attributes such as bullet/level/line spacing
      runs: p.runs.map((r, ri) => {
        const oldRun =
          r.srcRun != null ? oldPara?.runs[r.srcRun] : (oldPara?.runs[ri] ?? oldPara?.runs[0])
        const merged = {
          ...oldRun,
          text: r.text,
          bold: r.bold ?? oldRun?.bold,
          italic: r.italic ?? oldRun?.italic,
          underline: r.underline ?? oldRun?.underline,
          strike: r.strike ?? oldRun?.strike,
          fontSize: r.fontSize ?? oldRun?.fontSize,
          fontFamily: r.fontFamily ?? oldRun?.fontFamily,
          color: r.color ?? oldRun?.color,
        }
        // Super/subscript: the editor can only express three states — super/sub/none
        // (±30/-25/0) — so detect changes by sign; a different magnitude with the same sign
        // counts as unchanged (preserving exact original values like 40%)
        if (r.baseline != null && Math.sign(r.baseline) !== Math.sign(oldRun?.baseline ?? 0)) {
          if (r.baseline === 0) delete merged.baseline
          else merged.baseline = r.baseline
        }
        if (merged.strike === false) delete merged.strikeStyle
        // The editor always returns the resolved display color; treat it as an explicit color
        // only if the user actually changed it (clear colorFollowsTheme, persist as srgbClr),
        // otherwise keep the theme linkage
        if (r.color != null && hex6(r.color) !== hex6(oldRun?.color)) {
          delete merged.colorFollowsTheme
          delete merged.colorInherited
        }
        // Same for fonts: only clear the latin/ea preservation markers on a real change
        // (otherwise the original dual-font/theme references are written back as-is)
        if (r.fontFamily != null && fontKey(r.fontFamily) !== fontKey(oldRun?.fontFamily)) {
          delete merged.latinFont
          delete merged.eaFont
          delete merged.fontImplicit
        }
        // After removing underline, re-enabling starts fresh with sng;
        // do not revive the old style
        if (merged.underline === false) delete merged.underlineStyle
        // Explicit underline toggle beats hlink-derived styling
        if (r.underline != null && r.underline !== oldRun?.underline) {
          delete merged.underlineImplicit
        }
        // Hyperlink (r.link === undefined keeps the old link — see EditRun.link): the DOM
        // round-trips resolvable links, so compare against the resolved old value; a link the
        // DOM couldn't express (unresolvable/action-only rId, i.e. no oldRun.hyperlink) is
        // preserved verbatim by the ...oldRun spread instead
        const newLink = r.link ? encodeRunLink(r.link) : undefined
        if (
          r.link !== undefined &&
          newLink !== oldRun?.hyperlink &&
          (newLink || oldRun?.hyperlink)
        ) {
          if (newLink) {
            merged.hyperlink = newLink
            // Stale relationship: ensureRunLinkRels allocates a fresh rId before write-back
            delete merged.hyperlinkRId
            delete merged.hyperlinkAction
            delete merged.hyperlinkTooltip
          } else {
            delete merged.hyperlink
            delete merged.hyperlinkRId
            delete merged.hyperlinkAction
            delete merged.hyperlinkTooltip
            // Removing a link drops its derived underline (PowerPoint behavior)
            if (merged.underlineImplicit) {
              merged.underline = false
              delete merged.underlineImplicit
              delete merged.underlineStyle
            }
          }
        }
        // Font size, same as color/font: only clear the inheritance marker on a real change
        // (otherwise the rebuild writes no sz, keeping the master linkage)
        if (r.fontSize != null && r.fontSize !== oldRun?.fontSize) {
          delete merged.fontSizeImplicit
        }
        return merged
      }),
      align: p.align ?? oldPara?.align,
      // Indent level: returned by the editor after Tab adjustment; 0 resets to the default
      // (lvl attribute removed). For a paragraph with its own bullet (explicit marL + hanging
      // indent), the left indent steps with the level; inherited indents are re-resolved by
      // materialize
      ...(p.level != null && p.level !== (oldPara?.level ?? 0)
        ? {
            level: p.level || undefined,
            ...(oldPara?.pPrExplicit?.marL && oldPara.indent != null && oldPara.indent < 0
              ? { marL: -oldPara.indent * (p.level + 1) }
              : {}),
          }
        : {}),
      // Only mark alignment explicit on a real change (the editor always returns the computed
      // display value; untouched paragraphs keep inheriting)
      ...(p.align != null && oldPara && p.align !== oldPara.align
        ? { pPrExplicit: { ...oldPara.pPrExplicit, align: true } }
        : {}),
    }
  })
}

/**
 * Per-paragraph format marks set by the editor (bullets/line spacing/paragraph spacing on the
 * selection) → engine paragraph-format patches, keyed by the edited paragraph index.
 * The caller applies each through setElementParagraphFormat(…, [index]) so the pPr surgery only
 * touches the marked paragraphs.
 */
export function collectParagraphFormatPatches(
  edited: EditParagraph[],
): Array<{ index: number; patch: ParagraphFormatPatch }> {
  const out: Array<{ index: number; patch: ParagraphFormatPatch }> = []
  edited.forEach((p, i) => {
    const patch: ParagraphFormatPatch = {
      ...(p.bullet
        ? {
            bullet: p.bullet,
            ...(p.bullet === 'char' && p.bulletChar ? { bulletChar: p.bulletChar } : {}),
          }
        : {}),
      ...(p.lineSpacingPct != null ? { lineSpacingPct: p.lineSpacingPct } : {}),
      ...(p.spaceBeforePt != null ? { spaceBeforePt: p.spaceBeforePt } : {}),
      ...(p.spaceAfterPt != null ? { spaceAfterPt: p.spaceAfterPt } : {}),
    }
    if (Object.keys(patch).length) out.push({ index: i, patch })
  })
  return out
}

/** Whether the edited result differs from the original paragraphs in level (caller uses this to mark dirtyPPr.level + materialize). */
export function levelsChanged(oldParas: Paragraph[], edited: EditParagraph[]): boolean {
  return edited.some((p, pi) => {
    if (p.level == null) return false
    const oldPara = oldParas[p.srcPara ?? pi]
    return p.level !== (oldPara?.level ?? 0)
  })
}
