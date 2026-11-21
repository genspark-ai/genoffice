import { beforeEach, vi } from 'vitest'
import { dirtyFieldParts } from './helpers/dirty-fields'
import { packageXmlProblems } from './helpers/strict-xml'

let dirtyFieldsAllowed = false
beforeEach(() => {
  dirtyFieldsAllowed = false
})
;(globalThis as { __allowDirtyFields?: () => void }).__allowDirtyFields = () => {
  dirtyFieldsAllowed = true
}

/**
 * Every package a test saves is read back by a strict XML parser. The app's
 * own reader is lenient, so a part with an unbound prefix reopens here and is
 * refused by Word; this is the only place such a file fails.
 *
 * A saved field marked `w:dirty` makes Word ask to update fields on open, so
 * a save that writes the mark fails too, unless the test says it is meant
 * (`allowDirtyFields`: the field came so, or its result is not computed here).
 */
vi.mock('../src/patch', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/patch')>()
  const saveDocx: typeof actual.saveDocx = async (...args) => {
    const bytes = await actual.saveDocx(...args)
    const problems = await packageXmlProblems(bytes, args[0].internal.originalBytes)
    if (problems.length > 0)
      throw new Error(`the saved package is not well-formed XML:\n${problems.join('\n')}`)
    if (!dirtyFieldsAllowed) {
      const dirty = await dirtyFieldParts(bytes)
      if (dirty.length > 0)
        throw new Error(
          `the saved package holds fields marked w:dirty (Word asks to update fields on open); ` +
            `call allowDirtyFields() if the mark is meant:\n${dirty.join('\n')}`,
        )
    }
    return bytes
  }
  return { ...actual, saveDocx }
})
