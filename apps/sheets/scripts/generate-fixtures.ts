import { mkdir, writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'

import {
  buildCompatibilityFixture,
  buildEditFixture,
  buildStructureFixture,
  buildSheetsFixture,
  buildKitchenSinkFixture,
} from '../tests/fixture-builder'

const FIXTURES: ReadonlyArray<[string, () => Promise<Buffer>]> = [
  ['compatibility-basic.xlsx', buildCompatibilityFixture],
  ['compatibility-edit.xlsx', buildEditFixture],
  ['compatibility-structure.xlsx', buildStructureFixture],
  ['compatibility-sheets.xlsx', buildSheetsFixture],
  ['compatibility-kitchen-sink.xlsx', buildKitchenSinkFixture],
]

async function main(): Promise<void> {
  const outputDirectory = resolve('fixtures/generated')
  await mkdir(outputDirectory, { recursive: true })
  for (const [name, build] of FIXTURES) {
    await writeFile(resolve(outputDirectory, name), await build())
    process.stdout.write(`Generated fixtures/generated/${name}\n`)
  }
}

void main()
