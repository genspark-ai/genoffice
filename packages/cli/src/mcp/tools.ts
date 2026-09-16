import { z } from 'zod'
import type { CommandRegistry, OptionDef } from '../registry'

/**
 * How a tool parameter reaches the command: `inline` values are written to a
 * scratch file whose path is passed to the option, so an MCP client without a
 * filesystem can still send ops, specs and Markdown.
 */
export type ParamKind =
  'string' | 'boolean' | 'integer' | 'number' | 'inline-json' | 'inline-md' | 'inline-html'

export interface ParamSpec {
  /** parameter name in the tool schema */
  key: string
  /** the command option it becomes (`--<option>`); absent when key and option match */
  option?: string
  kind?: ParamKind
  description?: string
  required?: boolean
}

export interface PositionalSpec {
  key: string
  description: string
  optional?: boolean
}

export interface ToolSpec {
  name: string
  command: string
  verb?: string
  description: string
  positionals?: PositionalSpec[]
  /** argv appended to every call (`--type docx`) */
  fixed?: string[]
  /** a bare name takes its description from the command's OptionDef */
  options?: (string | ParamSpec)[]
  /** attach the PNGs listed in detail.files as image content */
  images?: boolean
  /** the text content is the result's summary, not the JSON envelope (guides) */
  plainText?: boolean
  readOnly?: boolean
  /** the call reaches the configured cloud provider (search, image, media) */
  openWorld?: boolean
}

export interface ResolvedParam extends Required<Pick<ParamSpec, 'key' | 'option' | 'kind'>> {
  description: string
  required: boolean
}

export interface ResolvedTool extends ToolSpec {
  params: ResolvedParam[]
}

const INTEGER_OPTIONS = new Set([
  'slide',
  'page',
  'max',
  'max-chars',
  'max-rows',
  'cols',
  'pad',
  'tile',
  'block',
])
const NUMBER_OPTIONS = new Set(['scale'])

const OPS_JSON = 'inline-json' as const

const READ_THEN_APPLY =
  'Read the file with the matching *_read tool first: ops target the ids and indexes it lists. '
const GUI_OPEN =
  'A file open in a GenOffice tab is refused without force (the editor would overwrite the edit on its next save).'

export const TOOLS: ToolSpec[] = [
  {
    name: 'info',
    command: 'info',
    readOnly: true,
    description:
      'Metadata and a structure summary of a document (docx, xlsx, pptx, pdf, md, html, csv): page or slide count, sheets, block count, size. Runs locally.',
    positionals: [{ key: 'file', description: 'path of the document' }],
    options: ['password'],
  },
  {
    name: 'convert',
    command: 'convert',
    description:
      'Convert a document to another format with the GenOffice engines (pdf, docx, xlsx, pptx, md, html, csv; the `to` description lists the routes). Targets that need layout (to pdf, docx to html, html to docx) start a hidden GenOffice process for 1-6 s. Output defaults to the input name with the new extension.',
    positionals: [{ key: 'file', description: 'path of the source document' }],
    options: ['to', 'out', 'force', 'password', 'sheet'],
  },
  {
    name: 'create_docx',
    command: 'create',
    description:
      'Create a new Word document from Markdown (GFM; headings, lists, tables, images by relative path, $formulas$) or from a restricted-HTML fragment (tags per guide docs). Give exactly one of markdown, html or from. Edits to an existing .docx go through docs_apply instead.',
    fixed: ['--type', 'docx'],
    options: [
      {
        key: 'markdown',
        option: 'from',
        kind: 'inline-md',
        description: 'the document as Markdown text',
      },
      {
        key: 'html',
        option: 'from',
        kind: 'inline-html',
        description: 'the document as a restricted-HTML fragment',
      },
      { key: 'from', description: 'path of an existing .md or .html file to build from' },
      { key: 'out', description: 'output .docx path (required)', required: true },
      'force',
    ],
  },
  {
    name: 'create_xlsx',
    command: 'create',
    description:
      'Create a new workbook from data: a 2-D array of cell values, { "sheets": [{ "name", "rows" }] } for several sheets, or a .csv/.json file. Strings starting with "=" are formulas. Give exactly one of data or from. Formatting, charts and more sheets afterwards: sheet_apply.',
    fixed: ['--type', 'xlsx'],
    options: [
      {
        key: 'data',
        option: 'from',
        kind: OPS_JSON,
        description: 'rows as a 2-D array, or { "sheets": [{ "name", "rows" }] }',
      },
      { key: 'from', description: 'path of a .csv or .json file to build from' },
      'header',
      'decimal',
      { key: 'out', description: 'output .xlsx path (required)', required: true },
      'force',
    ],
  },
  {
    name: 'create_pdf',
    command: 'create',
    description:
      'Print a document (md, html, docx, xlsx, pptx) to a new PDF with the GenOffice renderer, in a hidden GenOffice process.',
    fixed: ['--type', 'pdf'],
    options: [
      { key: 'from', description: 'path of the document to print', required: true },
      { key: 'out', description: 'output .pdf path (required)', required: true },
      'force',
    ],
  },
  {
    name: 'create_pptx',
    command: 'create',
    description:
      'Create a .pptx from ops on a blank one-slide deck, or from a deck spec (pages of px-positioned text, shapes and images on a 1280x720 canvas; guide slides spec). For a presentation a person will see, do not hand-place ops or a whole spec here: use deck_start, deck_page and deck_build, which check every page against an outline and a style sheet. Give exactly one of ops or spec.',
    fixed: ['--type', 'pptx'],
    options: [
      { key: 'ops', kind: OPS_JSON, description: 'JSON array of slides ops (guide slides)' },
      {
        key: 'spec',
        kind: OPS_JSON,
        description: 'a deck spec object { "pages": [...] } (guide slides spec)',
      },
      {
        key: 'spec_dir',
        option: 'spec',
        description:
          'path of a directory of one-page spec files taken in name order (the deck_* tools write one)',
      },
      'outline',
      { key: 'out', description: 'output .pptx path (required)', required: true },
      'force',
    ],
  },
  {
    name: 'docs_read',
    command: 'docs',
    verb: 'read',
    readOnly: true,
    description:
      'Structure and text of a .docx as 0-based blocks (paragraphs, headings, tables, images, charts) with a 200-character preview each; the block indexes are what docs_apply ops target. Flags add comments, tracked changes, styles, sections, fields, notes or header/footer text. Use full for whole block text, range to narrow.',
    positionals: [{ key: 'file', description: 'path of the .docx' }],
    options: [
      'range',
      'html',
      'full',
      'max-chars',
      'comments',
      'revisions',
      'styles',
      'header-footer',
      'sections',
      'fields',
      'notes',
    ],
  },
  {
    name: 'docs_apply',
    command: 'docs',
    verb: 'apply',
    description:
      'Edit a .docx with a batch of ops (the op reference: guide docs, or the genoffice://guide/docs resource). ' +
      READ_THEN_APPLY +
      'Block indexes inside one batch are live: after an insert adds two blocks, later indexes shift by two. Atomic by default: one rejected op leaves the file untouched and comes back with its index and reason; fix that op and resend the whole batch. dry_run validates without writing. ' +
      GUI_OPEN,
    positionals: [{ key: 'file', description: 'path of the .docx' }],
    options: [
      {
        key: 'ops',
        kind: OPS_JSON,
        description: 'JSON array of ops, every entry with "op"',
        required: true,
      },
      'track',
      'author',
      'dry-run',
      'best-effort',
      'stop-on-error',
      'out',
      'force',
    ],
  },
  {
    name: 'docs_check',
    command: 'docs',
    verb: 'check',
    readOnly: true,
    description:
      'Consistency checks on a .docx: fields not evaluated, broken references, stale TOC, missing images, empty charts, empty headings; one finding per issue.',
    positionals: [{ key: 'file', description: 'path of the .docx' }],
  },
  {
    name: 'sheet_read',
    command: 'sheet',
    verb: 'read',
    readOnly: true,
    description:
      'Cells of one worksheet as raw values (0.25, date serials) with formulas keyed by A1 address, plus sheet features (panes, filter, charts, tables). Defaults to the active sheet and up to 500x100 cells; the result says when it truncated. stats for counts and the sheet list, formats for styling, column widths and row heights, where to list only cells of one kind.',
    positionals: [{ key: 'file', description: 'path of the .xlsx' }],
    options: ['sheet', 'range', 'cols', 'max-rows', 'where', 'stats', 'formats'],
  },
  {
    name: 'sheet_apply',
    command: 'sheet',
    verb: 'apply',
    description:
      'Edit an .xlsx: cells sets values, formulas and styles by address; ops runs the workbook DSL (formatting, charts, images, tables, freeze panes, filters, conditional formats, validation, hyperlinks, notes, page setup, protection, defined names, sheet order; guide sheets). Give exactly one of cells or ops. ' +
      READ_THEN_APPLY +
      'Ops run in order, each seeing the previous result; structural ops (rows, columns, sheets) need their own batch. ' +
      GUI_OPEN,
    positionals: [{ key: 'file', description: 'path of the .xlsx' }],
    options: [
      'sheet',
      {
        key: 'cells',
        kind: OPS_JSON,
        description: 'JSON array of { "cell": "B2", "sheet"?, "value"? | "formula"?, "style"? }',
      },
      { key: 'ops', kind: OPS_JSON, description: 'JSON array of workbook DSL ops' },
      'dry-run',
      'best-effort',
      'stop-on-error',
      'out',
      'force',
    ],
  },
  {
    name: 'sheet_check',
    command: 'sheet',
    verb: 'check',
    readOnly: true,
    description:
      'Consistency checks on an .xlsx: formula errors, formulas not evaluated, missing sheet references, broken defined names, chart references, number overflow; one finding per issue.',
    positionals: [{ key: 'file', description: 'path of the .xlsx' }],
  },
  {
    name: 'slides_read',
    command: 'slides',
    verb: 'read',
    readOnly: true,
    description:
      'Structure of a .pptx: slides s_<n> with their elements e_* (kind, text preview, EMU geometry), tables and with full the whole text and speaker notes. The ids are what slides_apply ops target; ids of edited or created elements change, so read again before a second batch. layouts lists the master layouts for addSlideWithLayout.',
    positionals: [{ key: 'file', description: 'path of the .pptx' }],
    options: ['slide', 'full', 'layouts', 'max-chars'],
  },
  {
    name: 'slides_apply',
    command: 'slides',
    verb: 'apply',
    description:
      'Edit a .pptx with a batch of ops (setText, addElement, addPicture, deleteElement, setTransform, addSlide ...; guide slides). ' +
      READ_THEN_APPLY +
      "setText runs inherit the element's first run style unless a field is set. Ops run in order; a later op may target a slide an earlier op added. Atomic by default. After the edit, slides_audit for overflow and slides_render to look. " +
      GUI_OPEN,
    positionals: [{ key: 'file', description: 'path of the .pptx' }],
    options: [
      { key: 'ops', kind: OPS_JSON, description: 'JSON array of ops', required: true },
      'dry-run',
      'isolation',
      'out',
      'force',
    ],
  },
  {
    name: 'slides_audit',
    command: 'slides',
    verb: 'audit',
    readOnly: true,
    description:
      'Layout audit of a .pptx: text overflow, out-of-bounds and overlapping elements per slide, with the element ids and a suggested setTransform op for slides_apply. Heuristic glyph widths; confirm with slides_render.',
    positionals: [{ key: 'file', description: 'path of the .pptx' }],
    options: ['slide'],
  },
  {
    name: 'slides_render',
    command: 'slides',
    verb: 'render',
    readOnly: true,
    images: true,
    description:
      'One PNG per slide (960x540 at scale 1 for 16:9), written to a directory and returned as images so you can look at the pages. Starts a hidden GenOffice process (a few seconds). Use after building or editing a deck; slide renders one page.',
    positionals: [{ key: 'file', description: 'path of the .pptx' }],
    options: [
      { key: 'out', description: 'directory for the PNGs (required)', required: true },
      'slide',
      'scale',
    ],
  },
  {
    name: 'slides_check',
    command: 'slides',
    verb: 'check',
    readOnly: true,
    description:
      'Check a deck-flow file on disk: an outline.json (core hook and page plan) or one page spec, which is checked against outline.json and style.md beside it or one folder up. The deck_* tools call this for you; use it directly when the files were written by other means.',
    positionals: [{ key: 'file', description: 'path of outline.json or a page .json' }],
    options: ['outline', 'page'],
  },
  {
    name: 'slides_replace',
    command: 'slides',
    verb: 'replace',
    description:
      'Rebuild one slide of a .pptx from a one-page spec file on disk (a deck the staged flow built); the other slides keep their ids and content. deck_replace does the same from an inline page.',
    positionals: [{ key: 'file', description: 'path of the .pptx' }],
    options: [
      { key: 'slide', kind: 'integer', description: '0-based slide to rebuild', required: true },
      { key: 'spec', description: 'path of the one-page spec file', required: true },
      'outline',
      'out',
      'force',
    ],
  },
  {
    name: 'render',
    command: 'render',
    readOnly: true,
    images: true,
    description:
      'One PNG per page of a document (docx, xlsx, pptx, pdf, md, html) as the GenOffice renderer lays it out, written to a directory and returned as images. Starts a hidden GenOffice process. page renders one page; grid adds a contact sheet of every page.',
    positionals: [{ key: 'file', description: 'path of the document' }],
    options: [
      { key: 'out', description: 'directory for the PNGs (required)', required: true },
      'page',
      'scale',
      'el',
      'pad',
      'grid',
      'cols',
      'tile',
    ],
  },
  {
    name: 'guide',
    command: 'guide',
    readOnly: true,
    plainText: true,
    description:
      'The op reference an agent reads before writing ops: guide("docs"), guide("sheets") or guide("slides") list the groups; topic narrows to one group or one op. guide("slides", "design") is the staged deck workflow and guide("slides", "spec") the outline and page spec format (deck_start returns both). Also available as genoffice://guide/* resources.',
    positionals: [
      { key: 'domain', description: 'slides | docs | sheets' },
      {
        key: 'topic',
        description: 'a group, an op name, or for slides: design | spec',
        optional: true,
      },
    ],
    options: ['index'],
  },
  {
    name: 'capabilities',
    command: 'capabilities',
    readOnly: true,
    description:
      'Which cloud features are configured in GenOffice: web search, image search, image generation, media understanding. Check once before planning photos or live facts; everything else runs locally.',
  },
  {
    name: 'search',
    command: 'search',
    openWorld: true,
    readOnly: true,
    description:
      'Web search, or image search with images=true, through the provider configured in GenOffice (the query leaves the machine). Image results carry imageUrl, width and height for addPicture / insert_image.',
    positionals: [{ key: 'query', description: 'the search query' }],
    options: ['images', 'max'],
  },
  {
    name: 'image',
    command: 'image',
    openWorld: true,
    description:
      'Generate an image from a prompt with the provider configured in GenOffice and save it; the result names the real format. ref images steer edits (background removal, upscale).',
    positionals: [{ key: 'prompt', description: 'what to draw' }],
    options: ['out', 'aspect', 'size', 'ref', 'model', 'force'],
  },
  {
    name: 'media',
    command: 'media',
    openWorld: true,
    readOnly: true,
    description:
      'Describe or extract from an image, audio or video file (or URL) with the configured provider; the file leaves the machine.',
    positionals: [{ key: 'source', description: 'path or URL of the media' }],
    options: ['ask'],
  },
  {
    name: 'open',
    command: 'open',
    description:
      'Open a document in the GenOffice app for the user (starts the app if needed), optionally selecting a slide, element, block, range or page. Only when the user asks to see the file: an open tab makes later *_apply calls refuse to write.',
    positionals: [{ key: 'file', description: 'path of the document' }],
    options: ['slide', 'el', 'block', 'range', 'sheet', 'page'],
  },
]

export function resolveTools(registry: CommandRegistry): ResolvedTool[] {
  return TOOLS.map((tool) => resolveTool(tool, registry))
}

export function resolveTool(tool: ToolSpec, registry: CommandRegistry): ResolvedTool {
  const def = registry.get(tool.command)
  if (!def) throw new Error(`mcp tool ${tool.name}: unknown command ${tool.command}`)
  const defs = new Map((def.options ?? []).map((o) => [o.name, o]))
  const params = (tool.options ?? []).map((entry): ResolvedParam => {
    const spec: ParamSpec = typeof entry === 'string' ? { key: entry } : entry
    const option = spec.option ?? spec.key
    const opt = defs.get(option)
    if (!opt) throw new Error(`mcp tool ${tool.name}: ${tool.command} has no --${option}`)
    return {
      key: spec.key.replace(/-/g, '_'),
      option,
      kind: spec.kind ?? defaultKind(opt),
      description: spec.description ?? stripVerbPrefix(opt.description),
      required: spec.required ?? false,
    }
  })
  return { ...tool, params }
}

function defaultKind(opt: OptionDef): ParamKind {
  if (!opt.value) return 'boolean'
  if (INTEGER_OPTIONS.has(opt.name)) return 'integer'
  if (NUMBER_OPTIONS.has(opt.name)) return 'number'
  return 'string'
}

/** Option descriptions start with the verb or type they apply to ("read: ...", "pptx --spec <dir>: ..."); the tool already says. */
export function stripVerbPrefix(description: string): string {
  const m = /^([a-z][\w .<>/-]{0,28}): /.exec(description)
  return m ? description.slice(m[0].length) : description
}

export type ZodShape = Record<string, z.ZodTypeAny>

export function toolShape(tool: ResolvedTool): ZodShape {
  const shape: ZodShape = {}
  for (const p of tool.positionals ?? []) {
    const s = z.string().describe(p.description)
    shape[p.key] = p.optional ? s.optional() : s
  }
  for (const p of tool.params) {
    const s = kindSchema(p.kind).describe(p.description)
    shape[p.key] = p.required ? s : s.optional()
  }
  return shape
}

function kindSchema(kind: ParamKind): z.ZodTypeAny {
  switch (kind) {
    case 'boolean':
      return z.boolean()
    case 'integer':
      return z.number().int().nonnegative()
    case 'number':
      return z.number()
    case 'inline-json':
      return z.union([z.array(z.unknown()), z.record(z.string(), z.unknown())])
    default:
      return z.string()
  }
}

export interface InlineFile {
  option: string
  ext: string
  text: string
}

export interface BuiltCall {
  /** argv without the trailing inline-file paths; `inline[i]` is appended as `--<option> <path>` once written */
  argv: string[]
  inline: InlineFile[]
}

/** The command line a tool call becomes; inline parameters are returned separately so the caller can place them on disk. */
export function buildArgv(tool: ResolvedTool, args: Record<string, unknown>): BuiltCall {
  const argv: string[] = [tool.command, ...(tool.verb ? [tool.verb] : []), ...(tool.fixed ?? [])]
  for (const p of tool.positionals ?? []) {
    const v = args[p.key]
    if (v === undefined || v === null || v === '') {
      if (!p.optional) throw new Error(`missing ${p.key}`)
      continue
    }
    argv.push(String(v))
  }
  const inline: InlineFile[] = []
  for (const p of tool.params) {
    const v = args[p.key]
    if (v === undefined || v === null) continue
    if (p.kind === 'boolean') {
      if (v === true) argv.push(`--${p.option}`)
      continue
    }
    if (p.kind === 'inline-json') {
      inline.push({ option: p.option, ext: '.json', text: JSON.stringify(v) })
      continue
    }
    if (p.kind === 'inline-md' || p.kind === 'inline-html') {
      inline.push({
        option: p.option,
        ext: p.kind === 'inline-md' ? '.md' : '.html',
        text: String(v),
      })
      continue
    }
    argv.push(`--${p.option}`, String(v))
  }
  return { argv, inline }
}
