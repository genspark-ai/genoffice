import { describe, expect, it } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { ribbonCommonZh } from '../src/common-ribbon'

const SHARD_DIRS = ['docs', 'slides'].map((app) =>
  join(__dirname, '..', '..', '..', 'apps', app, 'src', 'renderer', 'i18n', 'ribbon'),
)

describe('shared ribbon table', () => {
  it('no per-app ribbon shard re-declares a common key', () => {
    const common = new Set(Object.keys(ribbonCommonZh))
    const redeclared = SHARD_DIRS.flatMap((dir) =>
      readdirSync(dir).flatMap((file) =>
        [...readFileSync(join(dir, file), 'utf8').matchAll(/^\s{2}(\w+):/gm)]
          .map((m) => m[1]!)
          .filter((key) => common.has(key))
          .map((key) => `${dir.split('/apps/')[1]}/${file}: ${key}`),
      ),
    )
    expect(redeclared).toEqual([])
  })
})
