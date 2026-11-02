#!/usr/bin/env node
/**
 * Coverage of the Word layout truths that have an executable fixture, and the
 * CI gate that keeps new truths from staying prose.
 *
 * The Word-truths reference is a Markdown file kept outside the repository
 * (one `## ` section per probe wave). Fixtures under
 * packages/docx-engine/tests/fixtures/truths/<id>/truth.json cite a section
 * through `"section": "word-truths.md#L<line>"`; a citation of a bullet inside
 * a section counts for that section.
 *
 * Two generated files live in the repo:
 *   docs/word-truths-index.json  sanitised machine index: per section its
 *                                heading line, date (first YYYY-MM-DD in the
 *                                heading), ASCII title with provenance
 *                                stripped, bullet count, `prose-only` reason
 *                                and the citing fixture ids. No corpus ids, no
 *                                paths, no non-ASCII text.
 *   docs/word-truths-coverage.md human table derived from the index + fixtures.
 *
 *   node tools/truths-coverage.mjs --truths <word-truths.md>   regenerate both
 *   node tools/truths-coverage.mjs                             regenerate the
 *                                coverage doc from the committed index
 *   node tools/truths-coverage.mjs --check [--truths <md>]     CI gate
 *
 * `--check` fails when
 *   (a) the committed index (fixture ids per section, or the whole index when
 *       --truths is given) or the coverage doc differs from what this tool
 *       would generate now, or
 *   (b) a section dated within the last 14 days and on/after GATE_SINCE has
 *       neither a fixture nor a `- prose-only: <reason>` bullet inside it, or
 *   (c) a section dated on/after GATE_SINCE whose title places it in the
 *       anchored-object / floating-table / table-width geometry family
 *       (GEOMETRY_FAMILY) has no fixture at all: those rules move page
 *       geometry and are never accepted as prose, however old.
 * Sections dated before GATE_SINCE (the day after the gate landed) are exempt
 * so the backlog that existed when the gate was introduced stays visible in
 * the coverage doc without failing CI. CI has no truths file, so rule (b)
 * reads dates and markers from the committed index; refresh it with --truths
 * after appending a section.
 */
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import prettier from 'prettier'

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const FIXTURES = join(ROOT, 'packages/docx-engine/tests/fixtures/truths')
const INDEX_PATH = join(ROOT, 'docs/word-truths-index.json')
const COVERAGE_PATH = join(ROOT, 'docs/word-truths-coverage.md')
const GATE_SINCE = '2026-10-29'
const RECENT_DAYS = 14
const DATE_RE = /\b(20\d\d-\d\d-\d\d)\b/
const PROSE_ONLY_RE = /^\s*- prose-only:\s*(.+?)\s*$/
const PROVENANCE_TOKEN = /\b(wps3way|gwc_staging|docs_agent|prod\d*[-_]|sb_\d{3}|sas2)\S*/gi
/** section titles whose truths decide anchor origins, float placement or table width */
const GEOMETRY_FAMILY =
  /\b(anchor(?:ed|s)?|float(?:ing|ed|s)?|tblp\w*|wrap(?:Square|TopAndBottom|None|Tight|ped)|obstacles?|tblW|tblw-\S*|tcW|gridCol|table[- ]width|pct table|dxa table|autofit|over-?wide)\b/i
const inGeometryFamily = (s) => GEOMETRY_FAMILY.test(s.title)

const args = process.argv.slice(2)
const opt = (name) => {
  const i = args.indexOf(name)
  return i === -1 ? undefined : args[i + 1]
}
const check = args.includes('--check')
const truthsPath = opt('--truths') ?? process.env.WORD_TRUTHS_MD
const outPath = opt('--out')
if (truthsPath && !existsSync(truthsPath)) {
  console.error(`truths file not found: ${truthsPath}`)
  process.exit(2)
}
if (!truthsPath && !existsSync(INDEX_PATH)) {
  console.error(
    'usage: node tools/truths-coverage.mjs --truths <word-truths.md> [--out <file.md>] [--check]',
  )
  process.exit(2)
}

const ascii = (s) =>
  s
    .replace(/[^\x20-\x7e]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
/** heading without the parenthesised probe provenance (paths, tool names,
 *  corpus ids); non-ASCII runs collapse so the index carries no locale text */
const sanitiseHeading = (heading) => {
  const cut = heading.search(/[(（]/)
  const head = cut > 0 ? heading.slice(0, cut) : heading
  const base = ascii(head.replace(/(^|\s)[~/]\S*/g, ' ').replace(PROVENANCE_TOKEN, ' ')).replace(
    /[\s,:;-]+$/,
    '',
  )
  return { title: base, abridged: /[^\x20-\x7e]/.test(head) }
}

function buildIndex(path) {
  const lines = readFileSync(path, 'utf8').split('\n')
  const sections = []
  for (let i = 0; i < lines.length; i++) {
    if (!lines[i].startsWith('## ')) continue
    if (sections.length) sections[sections.length - 1].end = i
    sections.push({ line: i + 1, heading: lines[i].slice(3), end: lines.length })
  }
  return {
    gateSince: GATE_SINCE,
    sections: sections.map((s) => {
      const body = lines.slice(s.line, s.end)
      const proseOnly = body.map((l) => PROSE_ONLY_RE.exec(l)?.[1]).find(Boolean)
      const { title, abridged } = sanitiseHeading(s.heading)
      return {
        line: s.line,
        date: DATE_RE.exec(s.heading)?.[1] ?? null,
        title: title || `section at line ${s.line}`,
        abridged,
        truths: body.filter((l) => /^\s*- /.test(l) && !PROSE_ONLY_RE.test(l)).length,
        proseOnly: proseOnly ? ascii(proseOnly.replace(PROVENANCE_TOKEN, ' ')) : null,
        fixtures: [],
      }
    }),
  }
}

const fixtures = readdirSync(FIXTURES, { withFileTypes: true })
  .filter((d) => d.isDirectory() && existsSync(join(FIXTURES, d.name, 'truth.json')))
  .map((d) => JSON.parse(readFileSync(join(FIXTURES, d.name, 'truth.json'), 'utf8')))
  .sort((a, b) => a.id.localeCompare(b.id))

function attachFixtures(index) {
  const unresolved = []
  const byId = new Map()
  for (const s of index.sections) {
    s.fixtures = []
    s.fixtureData = []
  }
  for (const f of fixtures) {
    byId.set(f.id, f)
    const m = /#L(\d+)/.exec(f.section ?? '')
    const line = m ? Number(m[1]) : NaN
    const s = [...index.sections].reverse().find((x) => x.line <= line)
    if (!s) unresolved.push(f.id)
    else {
      s.fixtures.push(f.id)
      s.fixtureData.push(f)
    }
  }
  return unresolved
}

const serialiseIndex = (index) =>
  JSON.stringify(
    {
      gateSince: index.gateSince,
      sections: index.sections.map(({ fixtureData: _d, ...s }) => s),
    },
    null,
    2,
  )

function renderCoverage(index, unresolved) {
  const sections = index.sections
  const title = (s) => {
    const own = s.fixtureData.find((f) => f.section_title)?.section_title
    if (own) return own
    return s.title + (s.abridged ? ' (heading abridged)' : '')
  }
  const covered = sections.filter((s) => s.fixtures.length)
  const totalCases = fixtures.reduce((n, f) => n + f.cases.length, 0)
  /** probed Word values the current code does not reproduce (skipped by the runners) */
  const failing = fixtures.flatMap((f) =>
    f.cases
      .filter((c) => f.status === 'failing' || c.status === 'failing')
      .map((c) => ({ id: f.id, section: f.section, note: c.note ?? c.kind })),
  )
  const totalBullets = sections.reduce((n, s) => n + s.truths, 0)
  const coveredBullets = covered.reduce((n, s) => n + s.truths, 0)

  const out = []
  out.push('# Word truths coverage')
  out.push('')
  out.push(
    'Generated by `node tools/truths-coverage.mjs`; do not edit by hand. A truth is _executable_ when a fixture under `packages/docx-engine/tests/fixtures/truths/` cites its section: pure fixtures run in `apps/docs/tests/truths.test.ts`, render fixtures in `e2e/docs-truths.spec.ts`. CI (`--check`) fails when a section dated on/after ' +
      `${index.gateSince} and within the last ${RECENT_DAYS} days has neither a fixture nor a \`- prose-only: <reason>\` bullet, and when a geometry-family section (anchors, floating tables, table width) dated on/after that day has no fixture at all.`,
  )
  out.push('')
  out.push(
    `Sections: **${covered.length} / ${sections.length}** covered (${totalBullets} recorded truths, ${coveredBullets} in covered sections). Fixtures: **${fixtures.length}** (${fixtures.filter((f) => f.class === 'pure').length} pure, ${fixtures.filter((f) => f.class === 'render').length} render) with ${totalCases} cases, ${failing.length} marked failing (bug backlog below).`,
  )
  out.push('')
  out.push('## Covered sections')
  out.push('')
  out.push('| Line | Section | Truths | Fixtures |')
  out.push('|---|---|---|---|')
  for (const s of covered) {
    const ids = s.fixtureData.map((f) => `\`${f.id}\` (${f.class}, ${f.cases.length})`).join(', ')
    out.push(`| L${s.line} | ${title(s)} | ${s.truths} | ${ids} |`)
  }
  out.push('')
  out.push('## Fixtures')
  out.push('')
  out.push('| Fixture | Class | Section | Probe date | Title |')
  out.push('|---|---|---|---|---|')
  for (const f of fixtures) {
    out.push(`| \`${f.id}\` | ${f.class} | ${f.section} | ${f.probe_date} | ${f.title} |`)
  }
  out.push('')
  if (failing.length) {
    out.push('## Failing cases (bug backlog)')
    out.push('')
    out.push(
      'Word values the fixtures record but the current code does not reproduce; the runners skip them until fixed.',
    )
    out.push('')
    out.push('| Fixture | Section | Case |')
    out.push('|---|---|---|')
    for (const f of failing)
      out.push(`| \`${f.id}\` | ${f.section} | ${f.note.replace(/\|/g, '\\|')} |`)
    out.push('')
  }
  out.push('## Uncovered sections')
  out.push('')
  out.push('| Line | Date | Section | Truths | Note |')
  out.push('|---|---|---|---|---|')
  for (const s of sections.filter((x) => !x.fixtures.length)) {
    const note = [
      s.proseOnly ? `prose-only: ${s.proseOnly}` : '',
      inGeometryFamily(s) ? 'geometry family: fixture required' : '',
    ]
      .filter(Boolean)
      .join('; ')
    out.push(`| L${s.line} | ${s.date ?? ''} | ${title(s)} | ${s.truths} | ${note} |`)
  }
  out.push('')
  if (unresolved.length) out.push(`Fixtures with no resolvable section: ${unresolved.join(', ')}`)
  return out.join('\n')
}

const format = async (text, path) =>
  prettier.format(text, { ...(await prettier.resolveConfig(path)), filepath: path })

const committedIndex = existsSync(INDEX_PATH) ? JSON.parse(readFileSync(INDEX_PATH, 'utf8')) : null
const index = truthsPath ? buildIndex(truthsPath) : structuredClone(committedIndex)
const unresolved = attachFixtures(index)
const coverageTarget = resolve(ROOT, outPath ?? COVERAGE_PATH)
const indexText = await format(serialiseIndex(index), INDEX_PATH)
const coverageText = await format(renderCoverage(index, unresolved), coverageTarget)
const covered = index.sections.filter((s) => s.fixtures.length).length
const rel = (p) => relative(ROOT, p)

if (!check) {
  if (truthsPath) writeFileSync(INDEX_PATH, indexText)
  writeFileSync(coverageTarget, coverageText)
  console.log(
    `${covered}/${index.sections.length} sections covered by ${fixtures.length} fixtures -> ${rel(coverageTarget)}` +
      (truthsPath ? `, ${rel(INDEX_PATH)}` : ''),
  )
  process.exit(0)
}

const problems = []
const refresh = `run \`node tools/truths-coverage.mjs --truths <word-truths.md>\` and commit the result`
if (truthsPath) {
  if (!committedIndex || JSON.stringify(committedIndex) !== JSON.stringify(JSON.parse(indexText)))
    problems.push(`${rel(INDEX_PATH)} is out of date with the truths file; ${refresh}`)
} else {
  const stale = committedIndex.sections.filter(
    (s, i) => JSON.stringify(s.fixtures) !== JSON.stringify(index.sections[i].fixtures),
  )
  if (stale.length)
    problems.push(
      `${rel(INDEX_PATH)} fixture lists are out of date (L${stale.map((s) => s.line).join(', L')}); ${refresh}`,
    )
}
if (!existsSync(coverageTarget) || readFileSync(coverageTarget, 'utf8') !== coverageText)
  problems.push(`${rel(coverageTarget)} is out of date; ${refresh}`)
if (unresolved.length) problems.push(`fixtures cite no known section: ${unresolved.join(', ')}`)

const today = new Date()
const cutoff = new Date(today.getTime() - RECENT_DAYS * 86400e3).toISOString().slice(0, 10)
const since = cutoff > index.gateSince ? cutoff : index.gateSince
const naked = index.sections.filter(
  (s) => s.date && s.date >= since && !s.fixtures.length && !s.proseOnly,
)
for (const s of naked)
  problems.push(
    `section L${s.line} (${s.date}) "${s.title}" has no fixture: add one under packages/docx-engine/tests/fixtures/truths/ citing word-truths.md#L${s.line}, or add a \`- prose-only: <reason>\` bullet to the section and refresh the index`,
  )
const nakedGeometry = index.sections.filter(
  (s) => s.date && s.date >= index.gateSince && !s.fixtures.length && inGeometryFamily(s),
)
for (const s of nakedGeometry)
  if (!naked.includes(s))
    problems.push(
      `section L${s.line} (${s.date}) "${s.title}" is a geometry-family truth (anchor / floating table / table width) and has no fixture: prose-only is not accepted here, add one under packages/docx-engine/tests/fixtures/truths/ citing word-truths.md#L${s.line}`,
    )

if (!problems.length) {
  console.log(
    `Word truths index/coverage up to date; ${covered}/${index.sections.length} sections covered, no unguarded sections since ${since}.`,
  )
  process.exit(0)
}
console.error('Word truths gate failed:')
for (const p of problems) console.error(`  - ${p}`)
process.exit(1)
