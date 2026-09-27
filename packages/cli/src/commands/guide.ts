import { flagBool } from '../args'
import type { CommandDef } from '../registry'
import { docsCatalog } from '../formats/docx'
import { sheetCatalog, sheetGuideText } from '../formats/xlsx-catalog'
import {
  opLines,
  renderGroups,
  withFingerprint,
  type GuideDomain,
  type OpCatalog,
  type OpEntry,
} from '../op-catalog'
import { CliError, EXIT, type CommandResult } from '../result'
import { didYouMean } from '../suggest'

const DOMAINS: readonly GuideDomain[] = ['slides', 'docs', 'sheets']

/**
 * Every op reference is generated from the definition the executor validates against (the pptx-ops
 * markdown, the sheets zod schema, the docs op registry and tool schemas), so the guide cannot
 * drift from what `apply` accepts; `--json` returns the same catalog with each op's schema.
 */
export const guideCommand: CommandDef = {
  name: 'guide',
  summary: 'Print the op reference and design guides an agent needs before writing ops or specs.',
  usage: 'guide <slides|docs|sheets> [group|op|design|spec] [--index] [--fingerprint]',
  options: [
    { name: 'index', description: 'one-line signature of every op instead of a group guide' },
    {
      name: 'fingerprint',
      description:
        'only the catalog fingerprint (changes with any op or field); no domain = all three',
    },
  ],
  async run(args) {
    const [domain, topic] = args.positionals
    const json = flagBool(args, 'json')
    if (flagBool(args, 'fingerprint') && (domain === undefined || isDomain(domain))) {
      return fingerprints(domain, json)
    }
    if (!isDomain(domain)) {
      throw new CliError(
        EXIT.usage,
        'guides available for: slides, docs, sheets',
        { usage: 'genoffice guide slides | genoffice guide docs | genoffice guide sheets' },
        {
          reason: domain === undefined ? 'missing_argument' : 'invalid_argument',
          suggestion: 'run `genoffice guide slides|docs|sheets`',
        },
      )
    }
    if (domain === 'slides' && (topic === 'design' || topic === 'spec')) {
      const { SLIDES_GUIDES } = await import('@genoffice/pipelines/slides')
      return { summary: SLIDES_GUIDES[topic].content }
    }
    const { catalog, text } = await load(domain, topic, flagBool(args, 'index'))
    if (!topic) return { summary: text(), ...(json ? { detail: { ...catalog } } : {}) }
    const group = catalog.groups.find((g) => g.name === topic)
    if (group) {
      const entries = catalog.ops.filter((e) => e.group === topic)
      return {
        summary: text(topic),
        ...(json
          ? { detail: { domain, fingerprint: catalog.fingerprint, group, ops: entries } }
          : {}),
      }
    }
    const entry = catalog.ops.find((e) => e.op === topic)
    if (!entry) {
      const names = [...catalog.groups.map((g) => g.name), ...catalog.ops.map((e) => e.op)]
      const guess = didYouMean(topic, names)
      throw new CliError(
        EXIT.usage,
        `unknown group or op: ${topic}`,
        { groups: catalog.groups.map((g) => g.name) },
        {
          reason: 'invalid_argument',
          suggestion: guess
            ? `did you mean ${guess}?`
            : `run \`genoffice guide ${domain}\` for the groups and ops`,
        },
      )
    }
    return { summary: opText(entry), ...(json ? { detail: { ...entry } } : {}) }
  },
}

function isDomain(value: string | undefined): value is GuideDomain {
  return (DOMAINS as readonly string[]).includes(value ?? '')
}

interface Loaded {
  catalog: OpCatalog
  text: (group?: string) => string
}

/** The catalog `guide <domain> --json` prints; the MCP op schemas are built from the same object. */
export async function loadCatalog(domain: GuideDomain): Promise<OpCatalog> {
  return (await load(domain, undefined, false)).catalog
}

async function load(
  domain: GuideDomain,
  topic: string | undefined,
  index: boolean,
): Promise<Loaded> {
  if (domain === 'sheets') return { catalog: sheetCatalog(), text: sheetGuideText }
  if (domain === 'docs') {
    const { catalog, htmlRules } = await docsCatalog()
    return { catalog, text: (group) => docsGuideText(catalog, htmlRules, group) }
  }
  const docs = await import('@genoffice/pptx-ops/op-docs')
  const catalog = slidesCatalog(docs)
  return {
    catalog,
    text: (group) => {
      if (index) return docs.opSignatureIndex()
      if (group) return docs.opGuide(group)!
      return [
        'Op groups (genoffice guide slides <group> prints one; genoffice guide slides <op> prints one op):',
        docs.opGuideCatalog(),
        '',
        'Building a new deck: `genoffice guide slides design` (the staged workflow: style sheet, outline, one page file at a time, build, QC) and `genoffice guide slides spec` (the outline and page spec JSON for `genoffice slides check` and `genoffice create --type pptx --spec`).',
        '',
        'Every op: { "op": "<name>", "target": { "slide": <index|"s_n">, "el"?: "e_*" }, ...fields }.',
        'Units are EMU (914400 per inch; suffixes in/cm/mm/pt/px accepted); font sizes are pt. `genoffice slides read <file>` lists ids and geometry.',
        'Vocabulary:',
        docs.opVocabulary(),
      ].join('\n')
    },
  }
}

type SlidesOpDocs = typeof import('@genoffice/pptx-ops/op-docs')

function slidesCatalog(docs: SlidesOpDocs): OpCatalog {
  const ops: OpEntry[] = Object.entries(docs.OP_DOCS)
    .filter(([, d]) => d.aiCallable !== false && !d.pending)
    .map(([op, d]) => ({ op, group: d.group, signature: d.sig, doc: d.body, examples: d.examples }))
  const groups = docs.OP_GROUPS.map((g) => ({
    name: g,
    summary: docs.OP_GUIDES[g].summary,
    ops: ops.filter((e) => e.group === g).map((e) => e.op),
  }))
  return withFingerprint({ domain: 'slides', groups, ops })
}

function docsGuideText(catalog: OpCatalog, htmlRules: string, group?: string): string {
  const lines = [
    'Word ops (genoffice docs apply --ops): a JSON array; every entry has "op".',
    'Targets: { "blockIndexes": [..] } or { "nodeType": "docHeading"|"docParagraph"|"docListItem"|"image", "headingLevel"? }; get indexes from `genoffice docs read`.',
    'Field notation: bare = string, n = number, bool = boolean, ? = optional, a|b = one of.',
    '',
    ...renderGroups(catalog, group),
  ]
  if (!group || group === 'content') lines.push('', htmlRules)
  return lines.join('\n')
}

function opText(entry: OpEntry): string {
  const lines = opLines(entry)
  if (entry.doc) lines.push('', entry.doc)
  else if (entry.schema) lines.push('', JSON.stringify(entry.schema, null, 2))
  return lines.join('\n')
}

async function fingerprints(
  domain: GuideDomain | undefined,
  json: boolean,
): Promise<CommandResult> {
  const domains = domain ? [domain] : DOMAINS
  const out: Record<string, string> = {}
  for (const d of domains) out[d] = (await load(d, undefined, false)).catalog.fingerprint
  return {
    summary: Object.entries(out)
      .map(([d, fp]) => `${d} ${fp}`)
      .join('\n'),
    ...(json ? { detail: { fingerprints: out } } : {}),
  }
}
