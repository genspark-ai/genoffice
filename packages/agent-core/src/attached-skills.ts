/**
 * The skills a user attached to one message from the composer.
 *
 * Skills are never offered to the model on their own: the only way a skill's
 * text reaches a turn is the user picking it, and then the whole body rides in
 * front of the request. The display text stays the user's own words; this is
 * the model-facing instruction.
 */
export interface AttachedSkill {
  name: string
  body: string
}

/**
 * Template parameters a skill body may carry, expanded once at send time.
 *
 * The vocabulary follows the Claude Code command conventions, plus context
 * variables an office suite can offer and a terminal agent cannot:
 *
 * - `$ARGUMENTS` — everything the user wrote next to the invocation
 * - `$ARGUMENTS[N]` / `$N` — one argument by position; arguments are split on
 *   whitespace with "double" or 'single' quotes grouping multi-word values.
 *   A position with no argument stays as literal text.
 * - `${NAME}` — a context variable supplied by the app (`skillContextVars`
 *   builds the standard set: FILE_NAME, SELECTION, TODAY, LANGUAGE, APP). An
 *   unknown name stays as literal text.
 *
 * `\$` directly before a placeholder escapes it to a literal `$`. Expansion
 * is a single left-to-right pass: text an argument carries in is never
 * re-scanned, so a user writing `$1` in their request gets it quoted, not
 * expanded.
 */
export type SkillVars = Readonly<Record<string, string | undefined>>

const PLACEHOLDER =
  /\\\$(?=[A-Za-z0-9_{])|\$\{([A-Za-z_][A-Za-z0-9_]*)\}|\$ARGUMENTS\[(\d+)\]|\$ARGUMENTS\b|\$(\d+)/g

/** whitespace split with "double" or 'single' quotes grouping multi-word values */
export function splitSkillArgs(args: string): string[] {
  const out: string[] = []
  for (const m of args.matchAll(/"([^"]*)"|'([^']*)'|(\S+)/g)) {
    out.push(m[1] ?? m[2] ?? m[3]!)
  }
  return out
}

export function expandSkillBody(body: string, args: string, vars?: SkillVars): string {
  if (!body.includes('$')) return body
  const positional = splitSkillArgs(args)
  return body.replace(PLACEHOLDER, (match, varName: string, indexed: string, short: string) => {
    // the escape consumed only the backslash; the token itself stays literal
    if (match.startsWith('\\')) return match.slice(1)
    if (varName !== undefined) return vars?.[varName] ?? match
    const index = indexed ?? short
    if (index !== undefined) return positional[Number(index)] ?? match
    return args
  })
}

/** the native display names of the UI languages, for `LANGUAGE` */
const LANG_NATIVE_NAME: Record<string, string> = {
  zh: '中文',
  'zh-TW': '繁體中文',
  en: 'English',
  ja: '日本語',
  ko: '한국어',
  fr: 'Français',
  de: 'Deutsch',
  es: 'Español',
  pt: 'Português',
  ru: 'Русский',
  ar: 'العربية',
  he: 'עברית',
  hi: 'हिन्दी',
  id: 'Bahasa Indonesia',
  it: 'Italiano',
  ms: 'Bahasa Melayu',
  nl: 'Nederlands',
  pl: 'Polski',
  th: 'ไทย',
  vi: 'Tiếng Việt',
  cs: 'Čeština',
}

export interface SkillContextInput {
  /** UI language tag (`zh`, `zh-TW`, ...) — drives LANGUAGE and TODAY's format */
  lang: string
  /** the product the panel lives in, in the user's language (e.g. 文档 / 表格) */
  app: string
  /** current document's file name, when the app has one */
  fileName?: string
  /** current selection text, capped by the caller (≤ 500 chars is plenty) */
  selection?: string
  /** clock override for tests */
  now?: Date
}

/** the standard context variables every panel offers its skills */
export function skillContextVars(input: SkillContextInput): SkillVars {
  const date = input.now ?? new Date()
  return {
    APP: input.app,
    LANGUAGE: LANG_NATIVE_NAME[input.lang] ?? input.lang,
    TODAY: date.toLocaleDateString(input.lang || undefined),
    ...(input.fileName ? { FILE_NAME: input.fileName } : {}),
    ...(input.selection ? { SELECTION: input.selection } : {}),
  }
}

export function withAttachedSkills(
  instruction: string,
  skills: readonly AttachedSkill[],
  vars?: SkillVars,
): string {
  if (skills.length === 0) return instruction
  const blocks = skills
    .map((s) => {
      // the user's own request doubles as the invocation's argument string
      const body = expandSkillBody(s.body, instruction, vars)
      return `<skill name="${s.name.replace(/"/g, '')}">\n${body.trim()}\n</skill>`
    })
    .join('\n\n')
  return (
    `The user attached the following skill${skills.length > 1 ? 's' : ''} to this request. ` +
    `Follow their steps and constraints with the tools you have; where a skill and general ` +
    `guidance disagree, the skill wins.\n\n${blocks}\n\nUser request:\n${instruction}`
  )
}
