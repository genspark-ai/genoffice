#!/usr/bin/env node
/**
 * OOXML schema gate for .pptx files.
 *
 * Every PresentationML / DrawingML part is run through Markup Compatibility
 * preprocessing (no extension namespace is "understood": ignorable
 * attributes/elements are dropped, mc:AlternateContent collapses to its
 * Fallback) and validated with xmllint against the ISO/IEC 29500-4
 * Transitional schemas in ./schemas. python-pptx and well-formedness checks
 * cannot see child-order or value-range violations; PowerPoint can, and
 * answers with a repair prompt.
 *
 *   node tools/ooxml-validate/validate-pptx.mjs [--base ORIGINAL.pptx] [--json] FILE.pptx...
 *
 * With --base only problems absent from the original are reported, so a
 * foreign deck's pre-existing quirks do not mask what an edit introduced.
 */
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'
import JSZip from 'jszip'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const SCHEMA_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), 'schemas')
const PART_SCHEMA = [
  [
    /^ppt\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters|handoutMasters|comments)\/[^/]+\.xml$/,
    'pml.xsd',
  ],
  [/^ppt\/(presentation|presProps|viewProps|commentAuthors)\.xml$/, 'pml.xsd'],
  [/^ppt\/charts\/chart[^/]*\.xml$/, 'dml-chart.xsd'],
  [/^ppt\/theme\/theme[^/]*\.xml$/, 'dml-main.xsd'],
  [/^ppt\/tableStyles\.xml$/, 'dml-main.xsd'],
]

export function xmllintAvailable() {
  return spawnSync('xmllint', ['--version'], { encoding: 'utf8' }).status === 0
}

/** Markup Compatibility preprocessing against the base schema (no extension namespace understood). */
export function mcePreprocess(xml) {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  const walk = (el, inherited) => {
    const ignorable = new Set(inherited)
    const ign = el.getAttributeNS(MC, 'Ignorable')
    if (ign) {
      for (const prefix of ign.trim().split(/\s+/)) {
        const ns = el.lookupNamespaceURI(prefix)
        if (ns) ignorable.add(ns)
      }
    }
    for (const attr of Array.from(el.attributes)) {
      if (attr.namespaceURI === MC || (attr.namespaceURI && ignorable.has(attr.namespaceURI))) {
        el.removeAttributeNode(attr)
      }
    }
    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType !== 1) continue
      if (child.namespaceURI === MC && child.localName === 'AlternateContent') {
        const fallback = Array.from(child.childNodes).find(
          (n) => n.nodeType === 1 && n.namespaceURI === MC && n.localName === 'Fallback',
        )
        const spliced = fallback ? Array.from(fallback.childNodes) : []
        for (const node of spliced) el.insertBefore(node, child)
        el.removeChild(child)
        for (const node of spliced) if (node.nodeType === 1) walk(node, ignorable)
        continue
      }
      if (child.namespaceURI === MC || (child.namespaceURI && ignorable.has(child.namespaceURI))) {
        el.removeChild(child)
        continue
      }
      walk(child, ignorable)
    }
  }
  walk(doc.documentElement, new Set())
  return new XMLSerializer().serializeToString(doc)
}

/** Validate every schema-mapped part; returns [{ part, message }] (empty when the deck is clean). */
export async function validatePptx(input) {
  const bytes = typeof input === 'string' ? fs.readFileSync(input) : input
  const zip = await JSZip.loadAsync(bytes)
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ooxml-validate-'))
  try {
    const problems = []
    // Well-formedness on the raw bytes of every XML part first: the DOM used for MCE
    // preprocessing silently repairs unbalanced tags, so it must never be what xmllint sees.
    const raw = []
    for (const name of Object.keys(zip.files)) {
      if (!/\.(xml|rels)$/.test(name) || zip.files[name].dir) continue
      const file = path.join(tmp, 'raw__' + name.replace(/\//g, '__'))
      fs.writeFileSync(file, await zip.file(name).async('nodebuffer'))
      raw.push({ name, file })
    }
    const wf = spawnSync('xmllint', ['--noout', '--nonet', ...raw.map((p) => p.file)], {
      encoding: 'utf8',
      maxBuffer: 256 << 20,
    })
    if (wf.error) throw wf.error
    const malformed = new Set()
    for (const line of wf.stderr.split('\n')) {
      const m = /^(.*?):(\d+): (.*)$/.exec(line)
      if (!m) continue
      const part = raw.find((p) => p.file === m[1])?.name
      if (!part) continue
      malformed.add(part)
      problems.push({ part, message: m[3] })
    }
    const groups = new Map()
    for (const name of Object.keys(zip.files)) {
      const schema = PART_SCHEMA.find(([re]) => re.test(name))?.[1]
      if (!schema || malformed.has(name)) continue
      const file = path.join(tmp, name.replace(/\//g, '__'))
      fs.writeFileSync(file, mcePreprocess(await zip.file(name).async('string')))
      if (!groups.has(schema)) groups.set(schema, [])
      groups.get(schema).push({ name, file })
    }
    for (const [schema, parts] of groups) {
      const r = spawnSync(
        'xmllint',
        [
          '--noout',
          '--nonet',
          '--schema',
          path.join(SCHEMA_DIR, schema),
          ...parts.map((p) => p.file),
        ],
        { encoding: 'utf8', maxBuffer: 256 << 20 },
      )
      if (r.error) throw r.error
      for (const line of r.stderr.split('\n')) {
        if (!line || / validates$/.test(line) || / fails to validate$/.test(line)) continue
        const m = /^(.*?):(\d+): (.*)$/.exec(line)
        const part = m ? (parts.find((p) => p.file === m[1])?.name ?? m[1]) : parts[0]?.name
        problems.push({ part, message: m ? m[3] : line })
      }
    }
    return problems
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true })
  }
}

/**
 * Problems of `edited` not already present in `base`: a multiset difference on (part, message).
 * Parts the base did not have (a duplicated slide carries its source's quirks under a new part
 * name) only report messages the base never produced anywhere.
 */
export function newProblems(base, edited) {
  const byPart = new Map()
  for (const p of base) {
    const k = `${p.part}\n${p.message}`
    byPart.set(k, (byPart.get(k) ?? 0) + 1)
  }
  const baseParts = new Set(base.map((p) => p.part))
  const baseMessages = new Set(base.map((p) => p.message))
  return edited.filter((p) => {
    const k = `${p.part}\n${p.message}`
    const n = byPart.get(k) ?? 0
    if (n > 0) {
      byPart.set(k, n - 1)
      return false
    }
    return baseParts.has(p.part) || !baseMessages.has(p.message)
  })
}

async function main(argv) {
  const files = []
  let base = null
  let json = false
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--base') base = argv[++i]
    else if (argv[i] === '--json') json = true
    else files.push(argv[i])
  }
  if (!files.length) {
    console.error('usage: validate-pptx.mjs [--base ORIGINAL.pptx] [--json] FILE.pptx...')
    return 2
  }
  if (!xmllintAvailable()) {
    console.error('xmllint not found on PATH (install libxml2-utils)')
    return 2
  }
  const baseProblems = base ? await validatePptx(base) : []
  const report = []
  for (const file of files) {
    const all = await validatePptx(file)
    const problems = base ? newProblems(baseProblems, all) : all
    report.push({ file, problems, total: all.length })
  }
  if (json) {
    console.log(JSON.stringify(report))
  } else {
    for (const r of report) {
      console.log(
        `${r.problems.length ? 'FAIL' : 'ok  '} ${r.file} ${r.problems.length}${base ? ` new (${r.total} total)` : ''}`,
      )
      for (const p of r.problems.slice(0, 20)) console.log(`    ${p.part}: ${p.message}`)
    }
  }
  return report.some((r) => r.problems.length) ? 1 : 0
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  process.exit(await main(process.argv.slice(2)))
}
