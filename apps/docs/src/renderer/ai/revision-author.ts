import { authorName } from '../author-name'

/** AI edits are attributed to the Settings user name, as Word does for add-in edits */
export function aiRevisionAuthor(): string {
  return authorName()
}
