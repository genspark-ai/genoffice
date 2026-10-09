/**
 * Naming a document from its own contents.
 *
 * A model asked for a file name will happily return "Q3 / Revenue.md", a
 * sentence with a full stop, or fifty words of prose. Anything that reaches
 * the filesystem has to survive every platform's rules at once: Windows
 * forbids < > : " / \ | ? * and a trailing dot, macOS and Linux treat "/" as
 * a directory separator, and all three reserve CON, PRN, AUX and friends.
 * Cleaning here rather than at each call site keeps the "what is a valid
 * name" answer in one place.
 */

/** Long enough to be searchable, short enough to survive a file dialog. */
const MAX_NAME_LENGTH = 60

/** Windows device names, lowercased, with or without an extension. */
const RESERVED_STEMS = new Set([
  'con',
  'prn',
  'aux',
  'nul',
  'com1',
  'com2',
  'com3',
  'com4',
  'com5',
  'com6',
  'com7',
  'com8',
  'com9',
  'lpt1',
  'lpt2',
  'lpt3',
  'lpt4',
  'lpt5',
  'lpt6',
  'lpt7',
  'lpt8',
  'lpt9',
])

/**
 * Reduce a model reply to a bare file stem: no directory, no extension, no
 * character any supported platform rejects, no leading dot, no reserved name.
 * Returns '' when nothing usable is left, which the callers treat as "keep
 * the current name" rather than as an error.
 */
export function sanitizeFileStem(raw: string): string {
  let name =
    raw
      // models like to answer with a path, a bullet, or a quoted title
      .replace(/[\r\n]+/g, ' ')
      .replace(/^[\s"'`#*\-–—•>]+/, '')
      .replace(/["'`]+$/, '')
      // keep the last path segment: the stem is what lands next to the real dir
      .split(/[/\\]/)
      .pop() ?? ''
  // drop an extension the model volunteered — the caller appends the real
  // one. A name that is *only* an extension (".env") is left alone: there is
  // no stem to keep, and stripping it would empty the name.
  if (/^[^.]{2,}\.[A-Za-z0-9]{1,8}$/.test(name) || /^.{3,}\.[A-Za-z0-9]{1,8}$/.test(name)) {
    name = name.replace(/\.[A-Za-z0-9]{1,8}$/, '')
  }
  // illegal on Windows, and "/" splits a path on unix
  name = name.replace(/[<>:"|?*]/g, ' ')
  // control characters (a stray \r, a zero-width joiner) are not printable
  // and several filesystems reject them outright
  // eslint-disable-next-line no-control-regex
  name = name.replace(/[\u0000-\u001f\u007f]/g, ' ')
  name = name.replace(/\s+/g, ' ').trim()
  // a leading dot hides the file and, on Windows, is just illegal
  name = name.replace(/^\.+/, '').trim()
  if (name.length > MAX_NAME_LENGTH) {
    // cut on a space so the name never ends mid-word; a single token longer
    // than the budget is hard-cut rather than dropped
    const head = name.slice(0, MAX_NAME_LENGTH + 1)
    const lastSpace = head.lastIndexOf(' ')
    name =
      lastSpace > MAX_NAME_LENGTH / 2
        ? head.slice(0, lastSpace).trim()
        : head.slice(0, MAX_NAME_LENGTH).trim()
  }
  // a Windows device name is illegal even with an extension appended
  if (RESERVED_STEMS.has(name.toLowerCase())) name = `${name}-file`
  return name
}

/**
 * The instruction sent to the model. Kept here, rather than inline at each
 * caller, so every document type asks for the same shape of answer and the
 * language follows the reader rather than the prompt.
 */
export function buildNamingInstruction(content: string, lang: string): string {
  return [
    'Give this document a file name. Reply with the name only: no quotes, no extension, no path, no explanation.',
    `Write it in ${lang}. Use 2-6 words.`,
    'Base it on what the document is actually about, not on a generic description of its format.',
    '',
    'Document:',
    content,
  ].join('\n')
}

/**
 * Cap the text handed to the model. A first page is enough to name a
 * document, and a 3 MB transcript would cost a fortune for nothing.
 */
export function excerptForNaming(content: string, limit = 4000): string {
  const collapsed = content.replace(/\n{3,}/g, '\n\n').trim()
  if (collapsed.length <= limit) return collapsed
  return `${collapsed.slice(0, limit)}\n…`
}

/**
 * True when a document has nothing worth naming. An empty file should keep
 * the default "Untitled" name rather than get one invented for it.
 */
export function hasNameableContent(content: string): boolean {
  return content.replace(/[\s#*_\-`~>|[\](){}]/g, '').length > 0
}

/**
 * The key this preference lives under in `app-settings.json`, alongside the
 * shell's other preferences and spelled like them.
 *
 * Reader and writer both reach for this constant, because `readAppSettings` is
 * typed `Record<string, unknown>`: a key that does not match is not a type
 * error, it is a preference that writes fine and then always reads back off.
 */
export const FILE_NAMING_PREF_KEY = 'aiAutoFileNaming'

/** Whether first-save naming is switched on. Only a real `true` counts. */
export function fileNamingPrefOn(settings: Record<string, unknown>): boolean {
  return settings[FILE_NAMING_PREF_KEY] === true
}
