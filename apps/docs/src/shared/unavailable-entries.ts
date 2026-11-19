/**
 * Dialog and ribbon entries whose feature is not built yet. An entry listed
 * here is not shown; deleting its row shows it again. Menu commands are
 * recorded by `planned: true` in `commands.ts` instead.
 */
export interface UnavailableEntry {
  /** where a user would find it */
  where: string
  /** what is missing */
  why: string
}

export const UNAVAILABLE_ENTRIES: Readonly<Record<string, UnavailableEntry>> = {
  'paste-special.paste-link': {
    where: 'Edit > Paste Special… > Paste / Paste link',
    why: 'a link needs a source document to follow; the clipboard flavors read here carry none',
  },
  'paste-special.display-as-icon': {
    where: 'Edit > Paste Special… > Display as icon',
    why: 'embedded objects shown as an icon are not something the page draws',
  },
  'protect.sections': {
    where: 'Tools > Protect Document… > Sections…',
    why: 'forms protection locks the whole body; a per-section w:formProt has no editor to enforce it',
  },
  'paragraph.adjust-right-indent': {
    where: 'Format > Paragraph… > Automatically adjust right indent',
    why: 'the layout does not move the right indent onto the character grid',
  },
  'document.drawing-grid': {
    where: 'Format > Document… > Document Grid > Drawing Grid…',
    why: 'the drawing grid for shapes is not modelled',
  },
  'cross-reference.numbered-item': {
    where: 'Insert > Cross-reference… > Reference type > Numbered item',
    why: 'a numbered item is referred to by its number, which no field here can cache',
  },
  'cross-reference.separate-numbers': {
    where: 'Insert > Cross-reference… > Separate numbers with',
    why: 'it separates the numbers of a numbered item, and numbered items are not offered',
  },
  'footnote.columns': {
    where: 'Insert > Footnote… > Footnote layout > Columns',
    why: 'the page reserves a footnote area one note under the other as it places lines; columns need the height of the area as a whole',
  },
  'footnote.restart-each-page': {
    where: 'Insert > Footnote… > Numbering > Restart each page',
    why: 'marks are numbered before the page is laid out, so a page cannot restart them',
  },
  'list-styles.modify': {
    where: 'Format > Bullets and Numbering… > List Styles > Modify…',
    why: 'it needs a level editor that starts from the definition of the style, and there is none',
  },
  'autocorrect.autotext-look-in': {
    where: 'Tools > AutoCorrect Options… > AutoText > Look in',
    why: 'entries are kept on this computer, not in a template',
  },
  'autocorrect.spelling-suggestions': {
    where: 'Tools > AutoCorrect Options… > AutoCorrect > spelling checker suggestions',
    why: 'nothing here replaces text from the spelling checker',
  },
  'autocorrect.date-style': {
    where: 'Tools > AutoCorrect Options… > AutoFormat As You Type > Date style',
    why: 'typed dates are not reformatted',
  },
  'autocorrect.define-styles': {
    where:
      'Tools > AutoCorrect Options… > AutoFormat As You Type > Define styles based on your formatting',
    why: 'styles are not derived from direct formatting',
  },
  'table-properties.preferred-width-check': {
    where: 'Table > Table Properties… > Preferred width checkbox (Column and Cell tabs)',
    why: 'a grid column always carries a width, so there is no "no preferred width" to pick',
  },
  'compare.original-label': {
    where: 'Tools > Track Changes > Compare Documents… > Label changes with (original document)',
    why: 'every change of a comparison is written with the label of the revised side',
  },
  'font.snap-to-grid': {
    where: 'Format > Font… > Advanced > Snap to grid when document grid is defined',
    why: 'the page lays a character grid out per paragraph and cannot take one run off it',
  },
  'typography.kinsoku-custom': {
    where: 'Format > Paragraph… > Asian Typography > Options… > Kinsoku > Custom',
    why: 'there is no editor for the kinsoku lists',
  },
  'typography.compress-kana': {
    where:
      'Format > Paragraph… > Asian Typography > Options… > Compress punctuation and Japanese kana',
    why: 'kana are never compressed on the page',
  },
}

export function isUnavailable(id: string): boolean {
  return Object.hasOwn(UNAVAILABLE_ENTRIES, id)
}
