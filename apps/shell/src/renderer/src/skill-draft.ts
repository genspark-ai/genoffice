import { parseSkillFrontmatter } from '@genoffice/agent-core'
import type { StringKey } from './strings'

/**
 * Pure helpers for the skill panes: turning a model reply into an editable
 * SKILL.md, loading an installed one back into fields, and assembling a
 * deterministic draft when no model is configured. No window, no IPC — the
 * transport-facing half lives in SkillCreator.tsx, so everything here is
 * testable straight from the test folder.
 */

/** the editable shape of a SKILL.md: the three fields the editor works with */
export interface SkillDraft {
  name: string
  description: string
  body: string
}

/** the answers the creator's interview collected, verbatim */
export interface SkillInterview {
  /** what the skill should do */
  purpose: string
  /** when the agent should use it — becomes the description */
  trigger: string
  /** steps and constraints, one step per line */
  steps: string
  /** optional worked example */
  example: string
}

/**
 * Folder-safe name rule, the renderer twin of the main process's writeSkill
 * guard: no path separators or Windows-forbidden characters, no leading dot,
 * and no trailing dot or space (Win32 strips those, so the folder would stop
 * matching its own name). Live validation in the panes reads this so the save
 * button never offers a name the main process would refuse.
 */
export function legalSkillName(name: string): boolean {
  return (
    name.length > 0 &&
    !/[/\\:*?"<>|]/.test(name) &&
    !name.startsWith('.') &&
    !/[. \t]$/.test(name) &&
    name === name.trim()
  )
}

/**
 * The live-validation verdict the editor and the creator's preview share:
 * which fields the main process's save guard would refuse. The save buttons
 * read the same booleans, so what is flagged red is exactly what is disabled.
 */
export function draftFieldErrors(draft: SkillDraft): { name: boolean; description: boolean } {
  return { name: !legalSkillName(draft.name), description: draft.description.trim() === '' }
}

/**
 * A SKILL.md into its editable fields; null when the frontmatter is unusable.
 * `parseSkillFrontmatter` (the same reader the main process lists skills with)
 * decides what usable means, so what the editor accepts and what GenOffice
 * lists can never drift apart.
 */
export function parseSkillDoc(text: string): SkillDraft | null {
  const fm = parseSkillFrontmatter(text)
  if (!fm) return null
  return { name: fm.name, description: fm.description, body: fm.body }
}

/** a fenced code block: ```info … ``` with the fence on its own lines */
const FENCED = /(?:^|\n)[ \t]*```[^\n\r`]*\r?\n([\s\S]*?)\r?\n[ \t]*```[ \t]*(?=\r?\n|$)/g

/**
 * The SKILL.md inside a model reply.
 *
 * The meta-prompt asks for exactly one fenced block and no prose around it,
 * but a reply that chats anyway is the normal case, so the fences are read out
 * first — prose before, between and after them is ignored — and only a reply
 * with no fences at all is tried as a bare document. The first candidate whose
 * frontmatter parses wins; null when none does.
 */
export function skillDocFromReply(reply: string): SkillDraft | null {
  const candidates = [...reply.matchAll(FENCED)].map((m) => m[1]!)
  candidates.push(reply.trim())
  for (const candidate of candidates) {
    const draft = parseSkillDoc(candidate)
    if (draft) return draft
  }
  return null
}

/**
 * The exact SKILL.md the main process's writeSkill writes for a draft: flat
 * frontmatter, the description collapsed onto one line, an LF-only body. The
 * editor's live preview renders this text verbatim, so what the user sees is
 * character-for-character what Save puts on disk.
 */
export function skillDocText(draft: SkillDraft): string {
  const description = draft.description.replace(/\s*\r?\n\s*/g, ' ').trim()
  const body = draft.body.replace(/\r\n/g, '\n')
  return `---\nname: ${draft.name}\ndescription: ${description}\n---\n\n${body}`
}

/**
 * The parameter vocabulary the editor's insert menu offers: one entry per form
 * `expandSkillBody` (in @genoffice/agent-core) replaces when the skill rides a
 * request. The token is written into the body verbatim; the description is the
 * menu's one-line gloss, localized like any other pane text. The `$N` entry
 * stands for the positional family — inserting `$1` and describing how `$2`,
 * `$3`, … pick later arguments.
 */
export const SKILL_PARAMS: readonly { token: string; desc: StringKey }[] = [
  { token: '$ARGUMENTS', desc: 'intgSkillParamArgumentsDesc' },
  { token: '$1', desc: 'intgSkillParamPositionalDesc' },
  { token: '${FILE_NAME}', desc: 'intgSkillParamFileNameDesc' },
  { token: '${SELECTION}', desc: 'intgSkillParamSelectionDesc' },
  { token: '${TODAY}', desc: 'intgSkillParamTodayDesc' },
  { token: '${LANGUAGE}', desc: 'intgSkillParamLanguageDesc' },
  { token: '${APP}', desc: 'intgSkillParamAppDesc' },
]

/**
 * Splice one parameter token into a body at the textarea's caret. Returns the
 * new body and the caret position that lands right after the token, so the
 * editor can restore the selection where typing would have left it. A caret
 * outside the body clamps to the nearest end.
 */
export function insertSkillParam(
  body: string,
  token: string,
  caret: number,
): { body: string; caret: number } {
  const at = Math.min(Math.max(caret, 0), body.length)
  return { body: body.slice(0, at) + token + body.slice(at), caret: at + token.length }
}

/** one step per non-empty line, list punctuation and numbering stripped */
function stepLines(steps: string): string[] {
  return steps
    .split(/\r?\n/)
    .map((line) => line.replace(/^\s*(?:[-*•]|\d+[.)])\s*/, '').trim())
    .filter((line) => line.length > 0)
}

/**
 * A kebab-case name suggestion from a phrase: latin words and digits joined
 * with dashes, capped at 40 characters on a word boundary, 'my-skill' when
 * the phrase offers nothing usable (CJK answers keep their own script and the
 * user names the skill in the preview).
 */
export function suggestSkillName(phrase: string): string {
  let slug = phrase
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
  if (slug.length > 40) slug = slug.slice(0, 40).replace(/-[a-z0-9]*$/, '')
  return legalSkillName(slug) ? slug : 'my-skill'
}

/**
 * The no-model draft: the same SKILL.md a model would be asked to write,
 * assembled mechanically from the interview answers, so the creator works
 * end-to-end before any AI is configured. The body follows the house style of
 * the skills GenOffice ships — imperative numbered workflow, constraints, and
 * what the finished output is.
 */
export function draftFromInterview(answers: SkillInterview): SkillDraft {
  const purpose = answers.purpose.trim()
  const trigger = answers.trigger.trim().replace(/\s*\r?\n\s*/g, ' ')
  const steps = stepLines(answers.steps)
  const example = answers.example.trim()
  const workflow = steps.length
    ? steps.map((step, i) => `${i + 1}. ${step}`)
    : ['1. Read the request and confirm what it asks for and which document type it concerns.']
  const lines: string[] = [
    `# ${purpose || 'Custom skill'}`,
    '',
    `You are a specialist for: ${purpose || 'the task described below'}. The agent should use this skill whenever ${trigger || 'the request matches this task'}.`,
    '',
    '## Workflow',
    ...workflow,
  ]
  if (example) {
    lines.push(
      '',
      '## Example',
      ...example
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean),
    )
  }
  lines.push(
    '',
    '## Constraints',
    '',
    '- Do exactly what the request asks; never invent facts, figures or content the user did not give.',
    '- Prefer the document tools over prose: deliver the result as a file the user can open in GenOffice.',
    '',
    '## Output',
    '',
    '- The finished file, plus one sentence saying what it is and where it was saved.',
  )
  return {
    name: suggestSkillName(purpose || trigger),
    description: trigger,
    body: `${lines.join('\n')}\n`,
  }
}

/** the interview answers as the model-facing user message (English labels, values verbatim) */
export function interviewToUserMessage(answers: SkillInterview): string {
  const sections = [
    ['Purpose — what the skill should do', answers.purpose],
    ['When to use — becomes the description', answers.trigger],
    ['Steps and constraints — one step per line', answers.steps],
    ['Example (optional)', answers.example],
  ] as const
  return sections
    .filter(([, text]) => text.trim())
    .map(([label, text]) => `## ${label}\n${text.trim()}`)
    .join('\n\n')
}

/**
 * The meta-prompt for the creator's one model call. It faces the model, so it
 * is English regardless of the UI language; the skill it asks for is the one
 * the in-app agent reads: a when-to-use description the relevance check sees,
 * and a body of imperative steps the agent executes through its document
 * tools, written with the send-time parameters the app expands ($ARGUMENTS and
 * the ${...} context variables) wherever the skill should reuse what the user
 * already wrote or the document already holds.
 */
export const SKILL_CREATOR_SYSTEM = [
  'You write SKILL.md files for GenOffice, an office suite whose in-app agent works on real documents (Word, Excel, PowerPoint, PDF, Markdown, HTML) through document tools — it never opens a shell or a terminal.',
  '',
  '## Output format',
  '',
  'Reply with exactly one fenced code block and nothing else: no prose before it, no prose after it. The block holds the whole SKILL.md:',
  '',
  '```markdown',
  '---',
  'name: a-kebab-case-id',
  'description: One sentence stating when the agent should use this skill.',
  '---',
  '',
  '<the body>',
  '```',
  '',
  '## Frontmatter rules',
  '',
  '- `name` is kebab-case: lowercase letters, digits and dashes, 40 characters at most.',
  '- `description` is one sentence that says when to use the skill — the app reads it to decide relevance, so it must name the tasks or document types the skill covers, in the same words a user would use to ask for them.',
  '',
  '## Body rules',
  '',
  '- Open with one line saying what the agent is and when it applies.',
  '- `## Workflow` holds the steps as numbered imperatives the agent executes through its document tools.',
  '- Then a short `## Constraints` (what must never happen) and `## Output` (what the user receives: which file type, what the reply besides the file says).',
  '- Under 60 lines; no code fences inside the body.',
  '- Write the skill in the language of the user answers below.',
  '',
  '## Parameters',
  '',
  'At send time the app replaces these placeholders in the body with values from the request, and a good skill uses them instead of guessing:',
  '',
  '- `$ARGUMENTS` — everything the user wrote in the request that attached this skill.',
  '- `$ARGUMENTS[N]` or `$N` — one argument by position: $1 the first, $2 the second; a quoted phrase counts as one.',
  '- `${FILE_NAME}` / `${SELECTION}` — the file name of the open document, and the text selected in it.',
  '- `${TODAY}` / `${LANGUAGE}` / `${APP}` — today’s date, the UI language, the app the request came from.',
  '',
  'Only these names are replaced; anything else stays as literal text. Prefer them wherever a step should reuse what the user already wrote or the document already holds.',
].join('\n')
