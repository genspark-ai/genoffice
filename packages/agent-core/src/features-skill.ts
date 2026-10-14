import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'
import type { AgentSkill, ExecutedToolCall } from './skill'
import { SLOT_LIMITS, SLOT_PLACES, validateSlot, type SlotDeclaration } from './feature-slots'
import {
  APPEARANCE_ACCENTS,
  APPEARANCE_BACKGROUNDS,
  MD_READ_THEMES,
  MD_THEME_FONTS,
  validateAppearance,
  type FeatureAppearance,
} from './feature-appearance'
import {
  SCRIPT_API_SUMMARY,
  SCRIPT_FORMATS,
  validateScript,
  type FeatureScript,
  type ScriptRunOutcome,
} from './feature-script'
import { HOOK_EVENTS, validateHook, type FeatureHook } from './feature-hooks'
import { validateKnowledge, type FeatureKnowledge } from './feature-knowledge'
import { validateCommand, type FeatureCommand } from './feature-commands'

/**
 * The interface the mod assistant is handed: add, inspect and manage the
 * features GenOffice has, and read a document — nothing more.
 *
 * No feature kind executes yet, so the skill's whole job is to keep the registry
 * honest: record what the user asked for, show what exists, and never imply that
 * something runs when it cannot. Storage is injected, so this package stays free
 * of electron and fs and the tool surface is unit-testable.
 */

export interface FeatureRecord {
  /** stable, derived from the name at creation; the id every tool takes */
  id: string
  name: string
  description: string
  /** the kind of feature this is: "slot" is drawn by the host, others are records for now */
  kind: string
  /** the slot this feature draws, when kind is slot */
  slot?: SlotDeclaration
  /** the chrome this feature asks for, when kind is appearance */
  appearance?: FeatureAppearance
  /** the code this feature runs, when kind is script; read-only in this slice */
  script?: FeatureScript
  /** the event this feature observes, when kind is hook */
  hook?: FeatureHook
  /** the source this feature offers the assistant, when kind is knowledge */
  knowledge?: FeatureKnowledge
  /** the action the palette lists, when kind is command */
  command?: FeatureCommand
  createdAt: string
  updatedAt: string
}

export interface FeatureDraft {
  name: string
  description: string
  kind?: string
  slot?: SlotDeclaration
  appearance?: FeatureAppearance
  script?: FeatureScript
  hook?: FeatureHook
  knowledge?: FeatureKnowledge
  command?: FeatureCommand
}

export interface FeaturePatch {
  name?: string
  description?: string
  kind?: string
  /** a slot replaces the previous one; null removes it */
  slot?: SlotDeclaration | null
  /** an appearance replaces the previous one; null removes it */
  appearance?: FeatureAppearance | null
  /** a script replaces the previous one; null removes it */
  script?: FeatureScript | null
  /** a hook replaces the previous one; null removes it */
  hook?: FeatureHook | null
  /** a knowledge source replaces the previous one; null removes it */
  knowledge?: FeatureKnowledge | null
  /** a command replaces the previous one; null removes it */
  command?: FeatureCommand | null
}

/** What the skill needs from the host; async because the shell answers over IPC. */
export interface FeatureRegistryDeps {
  /** the registry as of this turn, synchronously: the panel keeps the cache buildContext reads */
  snapshot(): FeatureRecord[]
  list(): Promise<FeatureRecord[]> | FeatureRecord[]
  get(id: string): Promise<FeatureRecord | null> | FeatureRecord | null
  create(draft: FeatureDraft): Promise<FeatureRecord> | FeatureRecord
  update(id: string, patch: FeaturePatch): Promise<FeatureRecord | null> | FeatureRecord | null
  remove(id: string): Promise<boolean> | boolean
  /** parsed text of one document file, read-only */
  readDocument(path: string): Promise<{ name: string; text: string; chars: number }>
  /**
   * Run a script feature read-only against a document and report the plan.
   * The host opens the file, runs the sandbox and validates the operations;
   * nothing is written, not even to a temporary file.
   */
  runScript(
    id: string,
    request: { path: string; input?: Record<string, unknown>; timeoutMs?: number },
  ): Promise<ScriptRunOutcome> | ScriptRunOutcome
}

const TOO_LONG = 400

const FEATURES_SYSTEM_PROMPT =
  `## Features
You add features to GenOffice on the user's behalf, and manage the ones that exist.
- Before creating anything, call list_features: an existing feature may already do
  the job, or need updating instead.
- create_feature records a feature's name and description. kind is optional; leave it
  empty unless the user names one.
- update_feature and delete_feature take an id from list_features. Never invent an id.
- A script feature can be dry-run with dry_run_feature: it reads the document and
  reports the operations it would apply. Nothing is written — not the document, not
  a temporary file — so never say a script changed, fixed or updated anything.
  Every other kind is still only recorded: never say such a feature is active,
  working, installed or ready.
- You cannot change documents. read_document returns a document's text and nothing
  else; there is no tool that writes one.
- If the user is vague, ask for the name and what the feature should do before
  creating it.
- A feature that should appear in the interface is kind "slot" with a slot object.
  The host draws it; you never draw pixels yourself.
  Places: ` +
  SLOT_PLACES.join(`, `) +
  `. (The editor places are not open yet.)
  Elements: text, button, input, select, list, file-list, view, overlay, divider, markdown, group.
  An overlay is a screen-only layer over the content area: { type: "overlay", shapes: [...] }
  where a shape is frame, line, dot or label, positioned in fractions of the layer
  (0 to 1, never pixels) and toned accent, muted, success or danger (never a colour).
  It takes no pointer events and never touches a document: use it to point at
  something, not to add a control.
  A view is a real page the shell mounts over the slot: { type: "view", view: "browser",
  url: "https://…", title, minHeight }. https only (http on localhost); the shell owns
  the view, its frame and its rules, so a feature says which page, never what it may do.
  A button/input needs an action: { kind: "open", path }, { kind: "ask-assistant", prompt },
  { kind: "panel", id } (open that feature's own home.panel) or { kind: "run-feature", id }.
  A file-list is filled by the host from a named source:
  { type: "file-list", source: "recent-files" | "starred-files", limit, action }.
  Pressing an item substitutes $path, $name and $folder into the action, so a file
  sidebar or a file switcher needs no code at all.
  A feature that should be an action the user can run is kind "command" with
  { label, description, action } — the same actions a slot may use, so a command can
  open a file, ask you something, or open a feature's panel. It appears in the command
  palette (Cmd/Ctrl+K). Do not add a command for something a drawn control already
  covers.
  A feature that should give the assistant something to read is kind "knowledge" with
  { source: "folder" | "file", path: an absolute local path, label }. It registers a
  source of the knowledge base: the folder's readable files join the corpus the
  knowledge-base tools search, with citations. It is read-only and it never touches a
  document. A folder source can be narrowed with file-name rules: include and/or
  exclude, each an array of name patterns — '*' spans characters, '?' is one,
  names only, never paths. A source is indexed full-text by default; index:
  "names" keeps its files listed and readable but does not put their content in
  the search index (for a huge folder, say so when you choose it). Updating the
  feature changes the rules the source is indexed by. Never point it at a whole home
  directory: name a folder that is about the work (for example a styles folder),
  not everything the user owns. The user can turn a source off or on from the
  knowledge base's Sources section; updating or deleting the feature is how the
  source changes.
  A feature that should react to the app is kind "hook" with { on: one of ` +
  HOOK_EVENTS.join(`, `) +
  `, script: { source } }.
  The host runs it read-only when that event happens (input.event carries the name
  and the path) and shows what it logged. It cannot change anything — with one
  exception: on document.will-save it may call $.event.veto("reason"), and the host
  then refuses that save and shows the reason. Use the veto only for something the
  user would want stopped, and say so when you add such a hook.
  A feature that should compute something about a document is kind "script" with a
  script object: { source, format: one of ` +
  SCRIPT_FORMATS.join(`, `) +
  ` }. The source is the body of a function that receives
  ($, input), and format names the document the script works on — the host gives it
  that format's namespace ($.docs, $.sheets, $.text, $.deck, $.pdf) and refuses a run
  against a file of another format. ` +
  SCRIPT_API_SUMMARY +
  `
  Write small, readable code: read with the namespace the format gives ($.docs.blocks(),
  $.sheets.range("Sheet1!A2:F200"), $.text.get(), $.deck.slides(), $.pdf.pages(1, 20)),
  propose with $.docs.propose / $.sheets.propose / $.deck.propose, return the new text
  for "text" scripts, explain with $.log(). Before claiming a script works, call
  dry_run_feature with a real path and read the plan it returns.
  To change the shell's own chrome, use kind "appearance" with { background: one of
  ` +
  APPEARANCE_BACKGROUNDS.join(`, `) +
  `, accent: one of ` +
  APPEARANCE_ACCENTS.join(`, `) +
  ` }.
  To re-skin the markdown reader's paper the way a Typora theme does, add
  { markdown: ... } to the same appearance object. Either pick one of the app's
  themes — ` +
  MD_READ_THEMES.join(`, `) +
  ` (drake-dark and nord are the dark ones) — or declare your own:
  { markdown: { paper: "#f4ecd8", ink: "#43341f", accent?: "#8c5b2a",
  font?: "sans" | "serif" | "mono" } }. The validator refuses an unreadable
  pairing (paper and ink under 4.5:1), so pick real, readable combinations and
  say why you chose them. The user's own pick in the ribbon always wins; a
  feature decides the paper only while the ribbon is on default. A theme only
  changes what the user sees on screen; the document text, saving and the
  print/export output are never touched.
  Limits: at most ` +
  SLOT_LIMITS.maxElements +
  ` elements, ` +
  SLOT_LIMITS.maxDepth +
  ` levels deep,
  ` +
  SLOT_LIMITS.maxTextChars +
  ` characters of text, ` +
  SLOT_LIMITS.maxItems +
  ` list entries.
  Theme, dark mode, text size and translations come from the host: do not encode colors.`

/** What the model reads after a dry run: the plan, what was refused, and the log. */
function formatRun(id: string, path: string, outcome: ScriptRunOutcome): string {
  const total = outcome.applied + outcome.failures.length
  const lines = [
    'Dry run of ' + id + ' on ' + path + ': the document was not changed.',
    outcome.applied + ' of ' + total + ' proposed operations would apply (' + outcome.ms + 'ms).',
  ]
  if (outcome.plan.length) {
    lines.push('Plan:', ...outcome.plan.map((line) => '  ' + line))
  }
  if (outcome.failures.length) {
    lines.push('Rejected by the engine:', ...outcome.failures.map((line) => '  ' + line))
  }
  if (outcome.result) {
    lines.push('Result (what the script returned):', '  ' + outcome.result)
  }
  if (outcome.log.length) {
    lines.push('Script log:', ...outcome.log.map((line) => '  ' + line))
  }
  return lines.join('\n')
}

export function createFeaturesSkill(deps: FeatureRegistryDeps): AgentSkill {
  const tools: AgentToolDef[] = [
    {
      name: 'list_features',
      description: 'List every feature GenOffice knows about, newest last.',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
    },
    {
      name: 'get_feature',
      description: 'Read one feature in full by its id.',
      inputSchema: {
        type: 'object',
        properties: { id: { type: 'string', description: 'feature id from list_features' } },
        required: ['id'],
      },
    },
    {
      name: 'create_feature',
      description:
        'Record a new feature. kind "slot" with a slot object makes the host draw it; any other kind is recorded only, because no kind executes yet.',
      inputSchema: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'short name, as the user would say it' },
          description: { type: 'string', description: 'what the feature should do' },
          kind: {
            type: 'string',
            description: 'optional; "slot" when a slot object is given, otherwise a label',
          },
          slot: {
            type: 'object',
            description:
              'where and what to draw: { at: one of ' +
              SLOT_PLACES.join(' | ') +
              ', element: an element tree }',
            properties: {
              at: { type: 'string', enum: [...SLOT_PLACES] },
              element: { type: 'object', description: 'element tree; see the prompt for the set' },
            },
            required: ['at', 'element'],
          },
          script: {
            type: 'object',
            description:
              'the code to run: { source, format? }, the body of a function receiving ($, input); format is one of ' +
              SCRIPT_FORMATS.join(' | '),
            properties: {
              source: { type: 'string', description: 'the function body' },
              format: {
                type: 'string',
                enum: [...SCRIPT_FORMATS],
                description: 'the document format the script targets',
              },
            },
            required: ['source'],
          },
          command: {
            type: 'object',
            description:
              'an action the user invokes from the command palette: { label, description, action }, with the same actions a slot may run.',
            properties: {
              label: { type: 'string', description: 'what the palette lists' },
              description: { type: 'string' },
              action: { type: 'object', description: 'the same action shape a button uses' },
            },
            required: ['label', 'action'],
          },
          knowledge: {
            type: 'object',
            description:
              'a local source the assistant may consult: { source: "folder" | "file", path (absolute), label }. Read-only; its files join the knowledge base. Narrow a folder with include/exclude name patterns.',
            properties: {
              source: { type: 'string', enum: ['folder', 'file'] },
              path: { type: 'string', description: 'absolute path, no ~ and no address' },
              label: { type: 'string' },
              include: {
                type: 'array',
                items: { type: 'string' },
                description:
                  'name patterns (like "*.md") a file must match to join the knowledge base; absent means every readable file',
              },
              exclude: {
                type: 'array',
                items: { type: 'string' },
                description: 'name patterns (like "*-draft.md") left out even when include matched',
              },
              index: {
                type: 'string',
                enum: ['full-text', 'names'],
                description:
                  'full-text (default) indexes file contents for search; names lists the files without indexing their content',
              },
            },
            required: ['source', 'path'],
          },
          hook: {
            type: 'object',
            description:
              'an event to observe: { on: one of ' +
              HOOK_EVENTS.join(' | ') +
              ', script: { source } }. The hook runs read-only and cannot write.',
            properties: {
              on: { type: 'string', enum: [...HOOK_EVENTS] },
              script: {
                type: 'object',
                properties: { source: { type: 'string' } },
                required: ['source'],
              },
            },
            required: ['on', 'script'],
          },
          appearance: {
            type: 'object',
            description:
              'chrome to change: { background: one of ' +
              APPEARANCE_BACKGROUNDS.join(' | ') +
              ', accent: one of ' +
              APPEARANCE_ACCENTS.join(' | ') +
              ', markdown: one of ' +
              MD_READ_THEMES.join(' | ') +
              ' }. Names only: the host owns the values, so dark mode and contrast stay correct.',
            properties: {
              background: { type: 'string', enum: [...APPEARANCE_BACKGROUNDS] },
              accent: { type: 'string', enum: [...APPEARANCE_ACCENTS] },
              markdown: {
                description:
                  'a host theme name (' +
                  MD_READ_THEMES.join(' | ') +
                  ') or a custom { paper, ink, accent?, font? } with hex colors (paper-ink contrast is enforced)',
                oneOf: [
                  { type: 'string', enum: [...MD_READ_THEMES] },
                  {
                    type: 'object',
                    properties: {
                      paper: { type: 'string', description: 'hex, e.g. #f4ecd8' },
                      ink: { type: 'string', description: 'hex, readable on the paper' },
                      accent: { type: 'string', description: 'link color hex' },
                      font: { type: 'string', enum: [...MD_THEME_FONTS] },
                    },
                    required: ['paper', 'ink'],
                  },
                ],
              },
            },
          },
        },
        required: ['name', 'description'],
      },
    },
    {
      name: 'update_feature',
      description: 'Change a feature recorded earlier. Pass only the fields to change.',
      inputSchema: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'feature id from list_features' },
          name: { type: 'string' },
          description: { type: 'string' },
          kind: { type: 'string' },
          slot: {
            type: ['object', 'null'],
            description: 'replace the drawn slot, or null to stop drawing it',
          },
          appearance: {
            type: ['object', 'null'],
            description: 'replace the chrome this feature asks for, or null to give it back',
          },
          script: {
            type: ['object', 'null'],
            description: 'replace the code this feature runs, or null to remove it',
          },
          command: {
            type: ['object', 'null'],
            description: 'replace the command the palette lists, or null to withdraw it',
          },
          knowledge: {
            type: ['object', 'null'],
            description: 'replace the source the assistant may consult, or null to withdraw it',
          },
          hook: {
            type: ['object', 'null'],
            description: 'replace the event this feature observes, or null to stop observing',
          },
        },
        required: ['id'],
      },
    },
    {
      name: 'delete_feature',
      description: 'Remove a feature recorded earlier.',
      inputSchema: {
        type: 'object',
        properties: { id: { type: 'string', description: 'feature id from list_features' } },
        required: ['id'],
      },
    },
    {
      name: 'read_document',
      description:
        'Read a document file as text (docx, xlsx, pptx, pdf, md, html, txt, csv). Read-only: there is no tool that writes a document.',
      inputSchema: {
        type: 'object',
        properties: { path: { type: 'string', description: 'absolute path of the file' } },
        required: ['path'],
      },
    },
    {
      name: 'dry_run_feature',
      description:
        'Run a script feature read-only against one document and report the plan: the operations it would apply and any the engine rejects. Nothing is written.',
      inputSchema: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'script feature id from list_features' },
          path: { type: 'string', description: 'absolute path of the document to work on' },
          input: {
            type: 'object',
            description: 'optional parameters the script reads as $.input',
          },
          timeoutMs: {
            type: 'number',
            description: 'optional budget in milliseconds (max 30000)',
          },
        },
        required: ['id', 'path'],
      },
    },
  ]

  const text = (value: unknown): string => (typeof value === 'string' ? value.trim() : '')

  return {
    id: 'features',
    systemPrompt: FEATURES_SYSTEM_PROMPT,
    tools,
    buildContext: () => {
      // The listing rides along each turn so the model sees the registry without
      // spending a tool call on it. Every field is flattened and clipped: a
      // feature name is user text, and it must not carry instructions into the
      // prompt (full records stay one get_feature away).
      const features = deps.snapshot()
      if (!features.length) return ''
      const one = (value: string): string => value.replace(/\s+/g, ' ').slice(0, 120)
      const lines = features.map(
        (f) => f.id + ' | ' + one(f.name) + ' | kind: ' + (one(f.kind) || 'none'),
      )
      return 'Features recorded so far (id | name | kind):\n' + lines.join('\n')
    },
    executeTool: async (call: AgentToolCall): Promise<ToolExecution> => {
      switch (call.name) {
        case 'list_features': {
          const features = await deps.list()
          if (!features.length) {
            return { output: 'no features yet', summary: '0 features' }
          }
          return {
            output: features
              .map(
                (f) =>
                  f.id +
                  ' | ' +
                  f.name +
                  ' | kind: ' +
                  (f.kind || 'none') +
                  (f.slot ? ' | drawn at ' + f.slot.at : ''),
              )
              .join('\n'),
            summary: features.length + ' features',
          }
        }
        case 'get_feature': {
          const id = text(call.input.id)
          const feature = id ? await deps.get(id) : null
          if (!feature)
            return {
              output: 'no feature with id ' + (id || '(empty)'),
              isError: true,
              summary: 'get ' + id,
            }
          return { output: JSON.stringify(feature, null, 2), summary: feature.name }
        }
        case 'create_feature': {
          const name = text(call.input.name)
          const description = text(call.input.description)
          if (!name || !description) {
            return {
              output: 'create_feature needs both name and description',
              isError: true,
              summary: 'create (missing fields)',
            }
          }
          if (name.length > TOO_LONG || description.length > TOO_LONG * 4) {
            return {
              output: 'name or description is too long',
              isError: true,
              summary: 'create (too long)',
            }
          }
          // A slot is data the host draws; validate it here so a malformed tree
          // never reaches the registry, and name the field that is wrong.
          let slot: SlotDeclaration | undefined
          if (call.input.slot !== undefined) {
            const validated = validateSlot(call.input.slot)
            if (!validated.ok) {
              return {
                output:
                  'the slot was rejected:\n' +
                  validated.errors
                    .map((error) => '  ' + error.field + ' ' + error.message)
                    .join('\n'),
                isError: true,
                summary: 'create (bad slot)',
              }
            }
            slot = validated.slot
          }
          // The appearance a feature asks for is a choice from host-defined
          // names, so a malformed or raw value is refused here, by field.
          let appearance: FeatureAppearance | undefined
          if (call.input.appearance !== undefined) {
            const validated = validateAppearance(call.input.appearance)
            if (!validated.ok) {
              return {
                output:
                  'the appearance was rejected:\n' +
                  validated.errors
                    .map((error) => '  ' + error.field + ' ' + error.message)
                    .join('\n'),
                isError: true,
                summary: 'create (bad appearance)',
              }
            }
            appearance = validated.appearance
          }
          // A script is code the host runs in its sandbox: validate the shape and
          // size here, and let the sandbox refuse what only running can reveal.
          let script: FeatureScript | undefined
          if (call.input.script !== undefined) {
            const validated = validateScript(call.input.script)
            if (!validated.ok) {
              return {
                output:
                  'the script was rejected:\n' +
                  validated.errors
                    .map((error) => '  ' + error.field + ' ' + error.message)
                    .join('\n'),
                isError: true,
                summary: 'create (bad script)',
              }
            }
            script = validated.script
          }
          // A hook observes an event the host raises; it runs in the same
          // sandbox as a script and cannot write.
          let hook: FeatureHook | undefined
          if (call.input.hook !== undefined) {
            const validated = validateHook(call.input.hook)
            if (!validated.ok) {
              return {
                output:
                  'the hook was rejected:\n' +
                  validated.errors
                    .map((error) => '  ' + error.field + ' ' + error.message)
                    .join('\n'),
                isError: true,
                summary: 'create (bad hook)',
              }
            }
            hook = validated.hook
          }
          // A knowledge source is local and read-only: the assistant may read
          // one entry at a time out of it, and nothing else about it.
          let knowledge: FeatureKnowledge | undefined
          if (call.input.knowledge !== undefined) {
            const validated = validateKnowledge(call.input.knowledge)
            if (!validated.ok) {
              return {
                output:
                  'the knowledge source was rejected:\n' +
                  validated.errors
                    .map((error) => '  ' + error.field + ' ' + error.message)
                    .join('\n'),
                isError: true,
                summary: 'create (bad knowledge)',
              }
            }
            knowledge = validated.knowledge
          }
          // A command is an action the user invokes from the palette: the same
          // declared actions a drawn control may run, and nothing more.
          let command: FeatureCommand | undefined
          if (call.input.command !== undefined) {
            const validated = validateCommand(call.input.command)
            if (!validated.ok) {
              return {
                output:
                  'the command was rejected:\n' +
                  validated.errors
                    .map((error) => '  ' + error.field + ' ' + error.message)
                    .join('\n'),
                isError: true,
                summary: 'create (bad command)',
              }
            }
            command = validated.command
          }
          const kind =
            text(call.input.kind) ||
            (slot
              ? 'slot'
              : appearance
                ? 'appearance'
                : script
                  ? 'script'
                  : hook
                    ? 'hook'
                    : knowledge
                      ? 'knowledge'
                      : command
                        ? 'command'
                        : '')
          if (kind === 'slot' && !slot) {
            return {
              output: 'kind "slot" needs a slot object saying where and what to draw',
              isError: true,
              summary: 'create (slot missing)',
            }
          }
          if (kind === 'appearance' && !appearance) {
            return {
              output:
                'kind "appearance" needs an appearance object: background and/or accent, from the names the host knows',
              isError: true,
              summary: 'create (appearance missing)',
            }
          }
          const feature = await deps.create({
            name,
            description,
            kind,
            ...(slot ? { slot } : {}),
            ...(appearance ? { appearance } : {}),
            ...(script ? { script } : {}),
            ...(hook ? { hook } : {}),
            ...(knowledge ? { knowledge } : {}),
            ...(command ? { command } : {}),
          })
          return {
            output:
              'recorded ' +
              feature.id +
              ' (' +
              feature.name +
              ').' +
              (slot
                ? ' The host draws it at ' +
                  slot.at +
                  '; the content itself is recorded only and does not run.'
                : appearance
                  ? ' The shell applies it now: ' +
                    [
                      appearance.background ? 'background ' + appearance.background : '',
                      appearance.accent ? 'accent ' + appearance.accent : '',
                    ]
                      .filter(Boolean)
                      .join(', ') +
                    '.'
                  : command
                    ? ' The command is recorded: it appears in the command palette and runs ' +
                      command.action.kind +
                      '.'
                    : knowledge
                      ? ' The knowledge source is recorded: the assistant reads it one entry at a time.'
                      : hook
                        ? ' The hook is recorded: it runs read-only when ' +
                          hook.on +
                          ' happens and reports what it saw.'
                        : script
                          ? ' The script is recorded (' +
                            script.source.length +
                            ' characters); dry_run_feature runs it read-only against a document.'
                          : ' It is stored only: nothing runs it yet.'),
            summary: 'create ' + feature.id,
          }
        }
        case 'update_feature': {
          const id = text(call.input.id)
          if (!id)
            return {
              output: 'update_feature needs an id',
              isError: true,
              summary: 'update (no id)',
            }
          let slotPatch: SlotDeclaration | null | undefined
          if (call.input.slot !== undefined) {
            if (call.input.slot === null) {
              slotPatch = null
            } else {
              const validated = validateSlot(call.input.slot)
              if (!validated.ok) {
                return {
                  output:
                    'the slot was rejected:\n' +
                    validated.errors
                      .map((error) => '  ' + error.field + ' ' + error.message)
                      .join('\n'),
                  isError: true,
                  summary: 'update (bad slot)',
                }
              }
              slotPatch = validated.slot
            }
          }
          let appearancePatch: FeatureAppearance | null | undefined
          if (call.input.appearance !== undefined) {
            if (call.input.appearance === null) {
              appearancePatch = null
            } else {
              const validated = validateAppearance(call.input.appearance)
              if (!validated.ok) {
                return {
                  output:
                    'the appearance was rejected:\n' +
                    validated.errors
                      .map((error) => '  ' + error.field + ' ' + error.message)
                      .join('\n'),
                  isError: true,
                  summary: 'update (bad appearance)',
                }
              }
              appearancePatch = validated.appearance
            }
          }
          let scriptPatch: FeatureScript | null | undefined
          if (call.input.script !== undefined) {
            if (call.input.script === null) {
              scriptPatch = null
            } else {
              const validated = validateScript(call.input.script)
              if (!validated.ok) {
                return {
                  output:
                    'the script was rejected:\n' +
                    validated.errors
                      .map((error) => '  ' + error.field + ' ' + error.message)
                      .join('\n'),
                  isError: true,
                  summary: 'update (bad script)',
                }
              }
              scriptPatch = validated.script
            }
          }
          let commandPatch: FeatureCommand | null | undefined
          if (call.input.command !== undefined) {
            if (call.input.command === null) {
              commandPatch = null
            } else {
              const validated = validateCommand(call.input.command)
              if (!validated.ok) {
                return {
                  output:
                    'the command was rejected:\n' +
                    validated.errors
                      .map((error) => '  ' + error.field + ' ' + error.message)
                      .join('\n'),
                  isError: true,
                  summary: 'update (bad command)',
                }
              }
              commandPatch = validated.command
            }
          }
          let knowledgePatch: FeatureKnowledge | null | undefined
          if (call.input.knowledge !== undefined) {
            if (call.input.knowledge === null) {
              knowledgePatch = null
            } else {
              const validated = validateKnowledge(call.input.knowledge)
              if (!validated.ok) {
                return {
                  output:
                    'the knowledge source was rejected:\n' +
                    validated.errors
                      .map((error) => '  ' + error.field + ' ' + error.message)
                      .join('\n'),
                  isError: true,
                  summary: 'update (bad knowledge)',
                }
              }
              knowledgePatch = validated.knowledge
            }
          }
          let hookPatch: FeatureHook | null | undefined
          if (call.input.hook !== undefined) {
            if (call.input.hook === null) {
              hookPatch = null
            } else {
              const validated = validateHook(call.input.hook)
              if (!validated.ok) {
                return {
                  output:
                    'the hook was rejected:\n' +
                    validated.errors
                      .map((error) => '  ' + error.field + ' ' + error.message)
                      .join('\n'),
                  isError: true,
                  summary: 'update (bad hook)',
                }
              }
              hookPatch = validated.hook
            }
          }
          const patch = {
            ...(text(call.input.name) ? { name: text(call.input.name) } : {}),
            ...(text(call.input.description) ? { description: text(call.input.description) } : {}),
            ...(text(call.input.kind)
              ? { kind: text(call.input.kind) }
              : slotPatch
                ? { kind: 'slot' }
                : appearancePatch
                  ? { kind: 'appearance' }
                  : scriptPatch
                    ? { kind: 'script' }
                    : hookPatch
                      ? { kind: 'hook' }
                      : knowledgePatch
                        ? { kind: 'knowledge' }
                        : commandPatch
                          ? { kind: 'command' }
                          : {}),
            ...(slotPatch === undefined ? {} : { slot: slotPatch }),
            ...(appearancePatch === undefined ? {} : { appearance: appearancePatch }),
            ...(scriptPatch === undefined ? {} : { script: scriptPatch }),
            ...(hookPatch === undefined ? {} : { hook: hookPatch }),
            ...(knowledgePatch === undefined ? {} : { knowledge: knowledgePatch }),
            ...(commandPatch === undefined ? {} : { command: commandPatch }),
          }
          if (!Object.keys(patch).length) {
            return { output: 'nothing to change', isError: true, summary: 'update ' + id }
          }
          const feature = await deps.update(id, patch)
          if (!feature) {
            return { output: 'no feature with id ' + id, isError: true, summary: 'update ' + id }
          }
          return {
            output: 'updated ' + feature.id + ': ' + feature.name,
            summary: 'update ' + feature.id,
          }
        }
        case 'delete_feature': {
          const id = text(call.input.id)
          if (!id)
            return {
              output: 'delete_feature needs an id',
              isError: true,
              summary: 'delete (no id)',
            }
          const ok = await deps.remove(id)
          if (!ok)
            return { output: 'no feature with id ' + id, isError: true, summary: 'delete ' + id }
          return { output: 'removed ' + id, summary: 'delete ' + id }
        }
        case 'read_document': {
          const path = text(call.input.path)
          if (!path)
            return {
              output: 'read_document needs a path',
              isError: true,
              summary: 'read (no path)',
            }
          try {
            const doc = await deps.readDocument(path)
            return {
              output: doc.name + ': ' + doc.chars + ' characters\n\n' + doc.text,
              summary: 'read ' + doc.name,
            }
          } catch (err) {
            return {
              output: err instanceof Error ? err.message : String(err),
              isError: true,
              summary: 'read failed',
            }
          }
        }
        case 'dry_run_feature': {
          const id = text(call.input.id)
          const path = text(call.input.path)
          if (!id || !path) {
            return {
              output: 'dry_run_feature needs an id and a path',
              isError: true,
              summary: 'dry run (missing fields)',
            }
          }
          const feature = await deps.get(id)
          if (!feature) {
            return { output: 'no feature with id ' + id, isError: true, summary: 'dry run ' + id }
          }
          if (!feature.script) {
            return {
              output: 'feature ' + id + ' is not a script, so there is nothing to dry-run',
              isError: true,
              summary: 'dry run ' + id,
            }
          }
          try {
            const outcome = await deps.runScript(id, {
              path,
              ...(call.input.input && typeof call.input.input === 'object'
                ? { input: call.input.input as Record<string, unknown> }
                : {}),
              ...(typeof call.input.timeoutMs === 'number'
                ? { timeoutMs: call.input.timeoutMs }
                : {}),
            })
            return {
              output: formatRun(id, path, outcome),
              summary: 'dry run ' + id + ' (' + outcome.applied + ' ops)',
            }
          } catch (err) {
            // a script that throws, times out or proposes something invalid is a
            // normal answer, not a crash: the caller gets the reason and the fix
            return {
              output: err instanceof Error ? err.message : String(err),
              isError: true,
              summary: 'dry run ' + id,
            }
          }
        }
        default:
          return { output: 'unknown tool: ' + call.name, isError: true, summary: call.name }
      }
    },
    verifyResponse: (finalText: string, executed: readonly ExecutedToolCall[]) => {
      // The registry is the only thing this skill can change, and nothing runs a
      // feature. A reply that claims otherwise is wrong in a way the user cannot
      // see, so one corrective turn is worth it.
      const claimsRunning =
        /\b(is (now )?(active|running|working|live|installed)|now (works|runs)|has been installed)\b/i.test(
          finalText,
        )
      if (!claimsRunning) return null
      const changed = executed.some(
        (call) => call.ok && call.name !== 'list_features' && call.name !== 'get_feature',
      )
      if (changed) {
        return 'Do not describe a feature as running or installed: this build only records features, and nothing executes them yet. State what was recorded and that it does not run yet.'
      }
      return 'Nothing was created or changed in this reply. Do not claim a feature is active; call create_feature first, or say what is missing.'
    },
  }
}
