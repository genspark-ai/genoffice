/**
 * Reading a third-party agent skill — a `SKILL.md` — and deciding whether it is
 * worth surfacing inside GenOffice.
 *
 * ## Why this exists, and why it is not what it looks like
 *
 * A coding agent's skill is prose with YAML frontmatter: `name`, `description`,
 * and a body that tells a reader how to drive a shell. GenOffice ships one such
 * skill of its own (`skills/genoffice/SKILL.md`), and its body is written
 * against `genoffice` running in a terminal.
 *
 * The in-app agent is not that reader. Its skills are `AgentSkill` values
 * compiled into the bundle — `executeTool` is a function, not a file — so a
 * SKILL.md cannot simply be adopted as one. What *can* be adopted is its
 * instructions, as a prompt section the model reads before it acts, provided
 * the model has a way to run the commands the instructions name. That is what
 * the shell half (`run_cli`, reusing the MCP CLI runner) provides.
 *
 * ## Why the filtering is here and not in the UI
 *
 * A developer's `.agents/` and `.claude/` directories hold skills for every kind
 * of work — git, testing, deployment, databases. Most of them have nothing to do
 * with the files GenOffice opens, and listing all of them in a slash palette
 * would bury the two that are relevant. So the classification is a pure
 * function over the skill's own words, testable without touching the disk.
 */

/** the frontmatter block of a SKILL.md, plus the prose under it */
export interface SkillFrontmatter {
  name: string
  description: string
  /** everything after the closing `---`, frontmatter stripped */
  body: string
  /** the frontmatter keys we did not model, kept so nothing is silently lost */
  extra: Record<string, string>
}

/**
 * The formats GenOffice actually opens, saves and converts.
 *
 * Both the extensions and the product names are listed, because a skill about
 * "converting PDF to HTML" names the extension and a skill about "building a
 * Word report" names the product, and neither mentions the other's spelling.
 * Conversion skills are covered by construction: "pdf to html" matches both
 * tokens, so a format-pair skill is never the one that gets filtered out.
 *
 * Deliberately absent: words generic enough to describe almost any skill
 * ("office", "document", "report", "letter", "cv", "md"). When nearly every
 * skill matched, the relevance flag carried no information; a skill about
 * Git also "documents things", and `md` fires on any prose skill that
 * mentions Markdown in passing. The remaining tokens name a file type or a
 * product, which is the signal the pane is trying to show.
 */
export const SKILL_FORMAT_TOKENS: readonly string[] = [
  // extensions
  'docx',
  'doc',
  'xlsx',
  'xls',
  'xlsm',
  'csv',
  'pptx',
  'ppt',
  'pdf',
  'markdown',
  'html',
  'htm',
  'epub',
  'rtf',
  'odt',
  'ods',
  'odp',
  // product names
  'word',
  'excel',
  'powerpoint',
  // what the user calls them
  'spreadsheet',
  'workbook',
  'presentation',
  'deck',
  'slide',
  'slides',
  'handout',
  'brochure',
  'invoice',
  'resume',
  'memo',
  'newsletter',
]

export interface SkillRelevance {
  /** true when the skill talks about a file GenOffice can work with */
  relevant: boolean
  /** the tokens that matched, lowercased and de-duplicated — for the UI's "why" */
  matched: string[]
}

/**
 * Whether a skill is about file formats this app handles.
 *
 * Matching is on a word boundary: `md` must not fire on `cmd`, `doc` must not
 * fire on `doctor`, `htm` must not fire on `html` being a substring of
 * something else. So the haystack is split into tokens rather than searched as a
 * string, and an extension only counts when it stands alone or carries a dot.
 */
export function classifySkill(name: string, description: string): SkillRelevance {
  const haystack = ` ${name} ${description} `.toLowerCase()
  // a "word" for matching: letters/digits, or a run that includes a dot so
  // `.docx` and `report.pdf` survive as single tokens
  const words = haystack.match(/[a-z0-9]+(?:\.[a-z0-9]+)*/g) ?? []
  const seen = new Set<string>()
  for (const word of words) {
    // strip a leading dot and any extension: `report.pdf` also matches `pdf`
    const bare = word.replace(/^\./, '').split('.')
    for (const part of bare) {
      if (!part) continue
      if (SKILL_FORMAT_TOKENS.includes(part)) seen.add(part)
    }
  }
  return { relevant: seen.size > 0, matched: [...seen].sort() }
}

/**
 * Parse a SKILL.md.
 *
 * Only the two keys the palette needs are modelled. Every other top-level key is
 * preserved verbatim in `extra`, and an indented line is treated as a
 * continuation of the key above it — so a nested block such as
 *
 *     metadata:
 *       version: 2.67.1
 *
 * arrives as `extra.metadata === 'version: 2.67.1'`, not as a structure. That is
 * deliberate: the palette never reads `extra`, and a hand-rolled YAML parser
 * would be a worse thing to own than a documented flat one. Nothing is dropped —
 * a key we do not model is still there to look at.
 *
 * A file with no frontmatter, or with frontmatter that never closes, returns
 * null: a malformed skill is skipped, not half-read.
 */
export function parseSkillFrontmatter(source: string): SkillFrontmatter | null {
  const text = source.replace(/^\uFEFF/, '')
  // match the opening fence as a line rather than slicing three characters off:
  // `slice(3)` would turn the first key of a fence-less file into `e:` and hide
  // the very mistake the check below exists to catch
  const open = /^---[ \t]*\r?\n/.exec(text)
  if (!open) return null
  const afterOpen = text.slice(open[0].length)
  // the closing fence is a line of its own; a `---` inside a value is not one
  const close = /^---[ \t]*(\r?\n|$)/m.exec(afterOpen)
  if (!close) return null

  const block = afterOpen.slice(0, close.index)
  const body = afterOpen.slice(close.index + close[0].length)

  const flat: Record<string, string> = {}
  let lastKey: string | null = null
  for (const raw of block.split(/\r?\n/)) {
    const line = raw.replace(/\s+$/, '')
    if (!line.trim() || line.trimStart().startsWith('#')) continue
    const kv = /^([A-Za-z0-9_.-]+):[ \t]*(.*)$/.exec(line)
    // both groups always match when the test succeeds; the assertions say so to
    // a compiler checking indexed access, without changing what runs
    if (kv) {
      lastKey = kv[1]!
      flat[lastKey] = unquote(kv[2]!.trim())
    } else if (lastKey && /^[ \t]/.test(raw)) {
      // a YAML continuation line; keep it on the same value
      flat[lastKey] = `${flat[lastKey]} ${line.trim()}`.trim()
    }
  }

  const name = flat.name?.trim()
  const description = flat.description?.trim()
  if (!name || !description) return null

  const extra: Record<string, string> = {}
  for (const [k, v] of Object.entries(flat)) {
    if (k !== 'name' && k !== 'description') extra[k] = v
  }
  return { name, description, body: body.replace(/^\s*\n/, ''), extra }
}

/** single- and double-quoted YAML scalars, and a `>`/`|` folded block */
function unquote(value: string): string {
  if (
    (value.startsWith('"') && value.endsWith('"') && value.length > 1) ||
    (value.startsWith("'") && value.endsWith("'") && value.length > 1)
  ) {
    return value.slice(1, -1).replace(/''/g, "'")
  }
  return value
}
