/**
 * Builds the icon data of Docs' Insert > Icons from the curated list (list.mjs)
 * and the SVG files of the pinned @fluentui/svg-icons dev dependency:
 *
 *   node tools/docs-icon-library/build.mjs          writes the data file
 *   node tools/docs-icon-library/build.mjs --check  fails when the file is stale
 *
 * The data file is committed, so a build needs neither this script nor the
 * package. Only single-path 24 px filled icons are taken: one `d` string each.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ICON_LIST } from './list.mjs'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..')
const PACKAGE = '@fluentui/svg-icons'
const VERSION = '1.1.346'
const SIZE = 24
const OUT = join(ROOT, 'apps/docs/src/renderer/icon-library/icons.json')

const SINGLE_PATH = new RegExp(
  `^<svg xmlns="http://www\\.w3\\.org/2000/svg" width="${SIZE}" height="${SIZE}" viewBox="0 0 ${SIZE} ${SIZE}"><path d="([^"]+)"/></svg>\\s*$`,
)
const PATH_DATA = /^[MmLlHhVvCcSsQqTtAaZz0-9eE.,\s-]+$/
const ENTRY = /^([a-z0-9_]+)(?:\s+"([^"]+)")?(?::\s*(.+))?$/

function packageDir() {
  const dir = join(ROOT, 'node_modules', PACKAGE)
  const { version, license } = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'))
  if (version !== VERSION) throw new Error(`${PACKAGE} is ${version}, the list is for ${VERSION}`)
  if (license !== 'MIT') throw new Error(`${PACKAGE} licence is ${license}, expected MIT`)
  return dir
}

function build() {
  const dir = packageDir()
  const seen = new Set()
  const icons = []
  for (const [category, block] of Object.entries(ICON_LIST)) {
    for (const line of block.split('\n').map((l) => l.trim())) {
      if (!line) continue
      const m = ENTRY.exec(line)
      if (!m) throw new Error(`${category}: cannot read "${line}"`)
      const [, name, label, keywords = ''] = m
      if (seen.has(name)) throw new Error(`${name} is listed twice`)
      seen.add(name)
      let svg
      try {
        svg = readFileSync(join(dir, 'icons', `${name}_${SIZE}_filled.svg`), 'utf8')
      } catch {
        throw new Error(`${category}: ${PACKAGE} has no ${name}_${SIZE}_filled.svg`)
      }
      const d = SINGLE_PATH.exec(svg)?.[1]
      if (!d || !PATH_DATA.test(d)) throw new Error(`${name} is not a single plain path`)
      const shown = label ?? name.replace(/_/g, ' ').replace(/^./, (c) => c.toUpperCase())
      icons.push([name, category, shown, keywords.trim().toLowerCase(), d])
    }
  }
  const head = {
    source: PACKAGE,
    version: VERSION,
    license: 'MIT',
    size: SIZE,
    categories: Object.keys(ICON_LIST),
  }
  return (
    JSON.stringify(head).slice(0, -1) +
    ',"icons":[\n' +
    icons.map((icon) => JSON.stringify(icon)).join(',\n') +
    '\n]}\n'
  )
}

const json = build()
if (process.argv.includes('--check')) {
  if (readFileSync(OUT, 'utf8') !== json) {
    console.error('icons.json is stale: run node tools/docs-icon-library/build.mjs')
    process.exit(1)
  }
} else {
  writeFileSync(OUT, json)
  const icons = JSON.parse(json).icons
  for (const category of Object.keys(ICON_LIST))
    console.log(category.padEnd(14), icons.filter((icon) => icon[1] === category).length)
  console.log('total'.padEnd(14), icons.length, `${json.length} bytes`)
}
