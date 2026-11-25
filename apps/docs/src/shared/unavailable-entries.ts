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
  'paste-special.display-as-icon': {
    where: 'Edit > Paste Special… > Display as icon',
    why: 'Word offers it for the object flavors of a paste (an embedded or linked OLE object drawn as an icon); the flavors pasted here are text and pictures, which have no icon form',
  },
  'table-properties.preferred-width-check': {
    where: 'Table > Table Properties… > Preferred width checkbox (Column and Cell tabs)',
    why: 'a column without a preferred width is sized by its content as it is typed; the page sizes columns from the grid, and reading w:tcW auto for the box changes how existing tables are laid out or saved',
  },
}

export function isUnavailable(id: string): boolean {
  return Object.hasOwn(UNAVAILABLE_ENTRIES, id)
}
