#!/usr/bin/env node
/**
 * Template center generator.
 *
 * The template center ships real office files from apps/shell/build/templates.
 * Their content stays reviewable text in git: every template is authored under
 * apps/shell/build/templates/src (markdown for docx, {sheets:[{name,rows}]} JSON
 * plus a workbook-ops JSON for xlsx, a deck spec for pptx, plain .md for
 * markdown) and this script builds the binaries with the repo's own creation
 * engines — the bundled genoffice CLI, the same one the app's MCP create tools
 * use (`create --type docx --from <md>`, `create --type xlsx --from <json>` +
 * `sheet apply --ops`, `create --type pptx --spec`). Generated outputs and the
 * templates.json manifest the IPC reads are committed, like the repo's other
 * generated fixtures.
 *
 *   node tools/gen-templates.mjs
 */
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { basename, dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { spawnSync } from 'node:child_process'
import JSZip from 'jszip'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = join(HERE, '..')
const TEMPLATES_DIR = join(ROOT, 'apps', 'shell', 'build', 'templates')
const SRC_DIR = join(TEMPLATES_DIR, 'src')
const MANIFEST = 'templates.json'
const CLI_DIR = join(ROOT, 'packages', 'cli')
const CLI = join(CLI_DIR, 'dist', 'genoffice.cjs')

/** the manifest's public shape — exactly what home:templates-list hands the renderer */
export const TEMPLATE_KINDS = new Set(['docx', 'xlsx', 'pptx', 'md'])
/** the categories the skills pane also uses; the gallery groups by these */
export const TEMPLATE_CATEGORIES = [
  'office',
  'business',
  'data',
  'finance',
  'hr',
  'legal',
  'marketing',
  'education',
  'other',
]

/** one zip part that must exist for the kind to count as a real document */
const REQUIRED_ZIP_PART = {
  docx: 'word/document.xml',
  xlsx: 'xl/workbook.xml',
  pptx: 'ppt/presentation.xml',
}

export class GenTemplatesError extends Error {}

/**
 * Read and shape-check the authored manifest (src/manifest.json). The same
 * checks guard the generated templates.json, so a bad entry fails the build
 * instead of shipping.
 */
export function readSourceManifest(srcDir = SRC_DIR) {
  const path = join(srcDir, 'manifest.json')
  if (!existsSync(path)) throw new GenTemplatesError(`${path}: manifest source is missing`)
  let parsed
  try {
    parsed = JSON.parse(readFileSync(path, 'utf8'))
  } catch (err) {
    throw new GenTemplatesError(`${path}: not valid JSON (${err.message})`)
  }
  if (!Array.isArray(parsed) || parsed.length === 0) {
    throw new GenTemplatesError(`${path}: expected a non-empty array of templates`)
  }
  const seen = new Set()
  parsed.forEach((entry, index) => {
    const at = `${path} entry ${index}`
    for (const key of ['id', 'title', 'description', 'kind', 'category', 'file']) {
      if (typeof entry?.[key] !== 'string' || entry[key] === '') {
        throw new GenTemplatesError(`${at}: missing or empty "${key}"`)
      }
    }
    if (!/^[a-z0-9-]+$/.test(entry.id)) {
      throw new GenTemplatesError(`${at}: id must be lowercase letters, digits and dashes`)
    }
    if (seen.has(entry.id)) throw new GenTemplatesError(`${at}: duplicate id "${entry.id}"`)
    seen.add(entry.id)
    if (!TEMPLATE_KINDS.has(entry.kind)) {
      throw new GenTemplatesError(`${at}: kind must be one of ${[...TEMPLATE_KINDS].join('/')}`)
    }
    if (!TEMPLATE_CATEGORIES.includes(entry.category)) {
      throw new GenTemplatesError(
        `${at}: category must be one of ${TEMPLATE_CATEGORIES.join('/')} (the skills pane's set)`,
      )
    }
    // a plain file name in the same directory — never a path, so the manifest
    // cannot smuggle a traversal into the IPC's resource lookup
    if (!/^[\w][\w.-]*$/.test(entry.file) || /[\\/]/.test(entry.file)) {
      throw new GenTemplatesError(`${at}: file must be a plain file name, got "${entry.file}"`)
    }
    if (basename(entry.file) !== entry.file) {
      throw new GenTemplatesError(`${at}: file must be a plain file name, got "${entry.file}"`)
    }
    const ext = entry.file.slice(entry.file.lastIndexOf('.') + 1)
    if (ext !== entry.kind) {
      throw new GenTemplatesError(
        `${at}: file extension ".${ext}" does not match kind "${entry.kind}"`,
      )
    }
  })
  return parsed
}

/** the public manifest entry: the TemplateInfo the renderer sees, plus `file` */
export function publicEntry(entry) {
  return {
    id: entry.id,
    title: entry.title,
    description: entry.description,
    kind: entry.kind,
    category: entry.category,
    file: entry.file,
  }
}

/** the authored source file(s) a template builds from, by kind */
export function sourceFilesFor(entry, srcDir = SRC_DIR) {
  switch (entry.kind) {
    case 'docx':
      return [join(srcDir, `${entry.id}.md`)]
    case 'xlsx':
      return [join(srcDir, `${entry.id}.json`), join(srcDir, `${entry.id}.ops.json`)]
    case 'pptx':
      return [join(srcDir, `${entry.id}.json`)]
    case 'md':
      return [join(srcDir, `${entry.id}.md`)]
    default:
      return []
  }
}

/** build the CLI bundle when a fresh checkout has no dist yet */
function ensureCli() {
  if (existsSync(CLI)) return
  console.log('building the bundled CLI (packages/cli/dist is missing)…')
  for (const script of ['build.mjs', 'collect-deps.mjs']) {
    const run = spawnSync(process.execPath, [script], { cwd: CLI_DIR, stdio: 'inherit' })
    if (run.status !== 0) {
      throw new GenTemplatesError(
        `node ${script} in packages/cli failed (exit ${run.status}); the template generator needs the bundled CLI`,
      )
    }
  }
  if (!existsSync(CLI)) {
    throw new GenTemplatesError(`${CLI}: CLI bundle still missing after building packages/cli`)
  }
}

/** run the bundled CLI; a nonzero exit fails the build with its own message */
function runCli(args) {
  const run = spawnSync(process.execPath, [CLI, ...args], {
    cwd: ROOT,
    encoding: 'utf8',
    windowsHide: true,
  })
  if (run.status !== 0) {
    const out = `${run.stdout ?? ''}${run.stderr ?? ''}`.trim()
    throw new GenTemplatesError(`genoffice ${args.join(' ')} failed (exit ${run.status}):\n${out}`)
  }
  return `${run.stdout ?? ''}`
}

function generateDocx(entry) {
  runCli([
    'create',
    '--type',
    'docx',
    '--from',
    join(SRC_DIR, `${entry.id}.md`),
    '--out',
    join(TEMPLATES_DIR, entry.file),
    '--force',
  ])
}

function generateXlsx(entry) {
  const out = join(TEMPLATES_DIR, entry.file)
  const staged = `${out}.staged.xlsx`
  runCli([
    'create',
    '--type',
    'xlsx',
    '--from',
    join(SRC_DIR, `${entry.id}.json`),
    '--out',
    staged,
    '--force',
  ])
  const ops = join(SRC_DIR, `${entry.id}.ops.json`)
  // header freeze / filter / column widths / number formats / header fill
  runCli(['sheet', 'apply', staged, '--ops', ops, '--out', out, '--force'])
  unlinkSync(staged)
}

function generatePptx(entry) {
  runCli([
    'create',
    '--type',
    'pptx',
    '--spec',
    join(SRC_DIR, `${entry.id}.json`),
    '--out',
    join(TEMPLATES_DIR, entry.file),
    '--force',
  ])
}

function generateMarkdown(entry) {
  copyFileSync(join(SRC_DIR, `${entry.id}.md`), join(TEMPLATES_DIR, entry.file))
}

/** every generated byte must be a real document: zip parts + CRC for OOXML, non-lorem text for md */
async function validateOutput(entry) {
  const path = join(TEMPLATES_DIR, entry.file)
  if (!existsSync(path) || statSync(path).size === 0) {
    throw new GenTemplatesError(`${entry.file}: output is missing or empty`)
  }
  if (entry.kind === 'md') {
    const text = readFileSync(path, 'utf8')
    if (text.trim() === '') throw new GenTemplatesError(`${entry.file}: markdown template is empty`)
    if (/lorem/i.test(text)) throw new GenTemplatesError(`${entry.file}: placeholder lorem text`)
    return
  }
  const zip = await JSZip.loadAsync(readFileSync(path), { checkCRC32: true })
  for (const part of ['[Content_Types].xml', REQUIRED_ZIP_PART[entry.kind]]) {
    if (!zip.file(part)) throw new GenTemplatesError(`${entry.file}: zip is missing ${part}`)
  }
}

/** remove top-level outputs a previous run (or a removed manifest entry) left behind */
function cleanStaleOutputs(manifest) {
  const keep = new Set(manifest.map((entry) => entry.file))
  keep.add(MANIFEST)
  keep.add('README.md')
  if (!existsSync(TEMPLATES_DIR)) return
  for (const name of readdirSync(TEMPLATES_DIR)) {
    const path = join(TEMPLATES_DIR, name)
    if (statSync(path).isDirectory() || keep.has(name)) continue
    if (/\.(docx|xlsx|pptx|md)$/.test(name)) {
      console.log(`removing stale output ${name}`)
      rmSync(path)
    }
  }
}

async function main() {
  mkdirSync(TEMPLATES_DIR, { recursive: true })
  const manifest = readSourceManifest()
  for (const entry of manifest) {
    const sources = sourceFilesFor(entry)
    const missing = sources.filter((src) => !existsSync(src))
    if (missing.length > 0) {
      throw new GenTemplatesError(
        `${entry.id}: missing authored source(s): ${missing.map((m) => basename(m)).join(', ')}`,
      )
    }
  }
  ensureCli()
  cleanStaleOutputs(manifest)
  for (const entry of manifest) {
    process.stdout.write(`building ${entry.file} (${entry.kind})… `)
    if (entry.kind === 'docx') generateDocx(entry)
    else if (entry.kind === 'xlsx') generateXlsx(entry)
    else if (entry.kind === 'pptx') generatePptx(entry)
    else generateMarkdown(entry)
    await validateOutput(entry)
    console.log('ok')
  }
  const shipped = manifest.map(publicEntry)
  writeFileSync(join(TEMPLATES_DIR, MANIFEST), `${JSON.stringify(shipped, null, 2)}\n`, 'utf8')
  const kinds = {}
  for (const entry of shipped) kinds[entry.kind] = (kinds[entry.kind] ?? 0) + 1
  console.log(
    `wrote ${MANIFEST}: ${shipped.length} templates ` +
      `(${Object.entries(kinds)
        .map(([kind, n]) => `${n} ${kind}`)
        .join(', ')}) in apps/shell/build/templates`,
  )
}

const invokedDirectly = process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href
if (invokedDirectly) {
  main().catch((err) => {
    console.error(`gen-templates: ${err instanceof Error ? err.message : String(err)}`)
    process.exit(1)
  })
}
