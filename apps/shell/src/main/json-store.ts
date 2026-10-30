import { readFileSync } from 'node:fs'
import { atomicWriteFile } from './atomic-write'

/**
 * The shared persistence core behind every one-file userData store (the
 * desktop layout, the desktop prefs, the desktop pins, the todo list): a
 * schema version, a normalize hook that enforces the store's invariant, the
 * scratch-then-rename atomic write, and a read that treats a missing or
 * corrupt file as the store's default value instead of a crash.
 *
 * ## Version and migration convention
 *
 * A store declares its current schema `version` (every store starts at 1 this
 * round). The on-disk **body** for version 1 is exactly the payload the store
 * always wrote — no wrapper, no marker — so files written before this factory
 * existed are already valid version-1 files and keep round-tripping
 * byte-identically.
 *
 * When a schema change lands, the author bumps `version` and adds a
 * `migrate(raw, fromVersion)` step: it receives the body as stored at version
 * `from` and must return the body as version `from + 1` shapes it (still in
 * the store's envelope — for the pins store that means `{ pins }`, for the
 * todo store the bare array). The reader walks the chain one step at a time
 * up to the current version, then hands the result to `unwrap`/`normalize`.
 * From version 2 on the writer stamps a `storeVersion` marker into the body
 * (which must therefore be a plain object from version 2 up; a store whose
 * version-1 body is not an object — the todo list — switches envelopes in its
 * first migration step). A body without the marker is read as version 1. A
 * body stamped with a *newer* version than the store (a downgrade) is read
 * through `normalize` like any foreign file.
 */

/** the field a version-2-or-later body carries its schema version in */
const VERSION_FIELD = 'storeVersion'

export interface JsonStoreSpec<T> {
  /** the one file the store reads and writes */
  path: () => string
  /** the schema version of the body this build writes; 1 = the marker-less legacy shape */
  version: number
  /**
   * The store's invariant, applied to anything read or written: a foreign or
   * hand-edited payload, a domain value heading to disk, or `undefined` (the
   * missing/corrupt-file case) all come back as a valid `T` — the default
   * value is whatever normalize answers for junk.
   */
  normalize: (raw: unknown) => T
  /** how the value is wrapped into the file body (default: the value is the body) */
  wrap?: (value: T) => unknown
  /** how the file body is unwrapped before normalize (default: the body is the value) */
  unwrap?: (raw: unknown) => unknown
  /**
   * One schema step: `raw` is the body as stored at version `from`; the return
   * value is the same data as version `from + 1` shapes it. Required for every
   * step below the store's declared version once that version is 2 or more.
   */
  migrate?: (raw: unknown, fromVersion: number) => unknown
}

export interface JsonStore<T> {
  /** the normalized value; a missing or corrupt file reads as the default */
  read(): T
  /** normalize, then atomically replace the file; a crash mid-write costs nothing */
  write(value: T): Promise<void>
}

/** the body plus the version it was stored at (no marker = version 1) */
function splitVersion(parsed: unknown): { body: unknown; version: number } {
  if (parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed)) {
    const marker = (parsed as Record<string, unknown>)[VERSION_FIELD]
    if (typeof marker === 'number' && Number.isInteger(marker) && marker >= 1) {
      const body: Record<string, unknown> = { ...(parsed as Record<string, unknown>) }
      delete body[VERSION_FIELD]
      return { body, version: marker }
    }
  }
  return { body: parsed, version: 1 }
}

export function createJsonStore<T>(spec: JsonStoreSpec<T>): JsonStore<T> {
  return {
    read(): T {
      let parsed: unknown
      try {
        parsed = JSON.parse(readFileSync(spec.path(), 'utf8'))
      } catch {
        // a missing or corrupt file is the store's default value, never a
        // reason the feature that reads it fails to open
        parsed = undefined
      }
      const { body, version } = splitVersion(parsed)
      let current = body
      for (let from = version; from < spec.version; from += 1) {
        if (!spec.migrate) break
        current = spec.migrate(current, from)
      }
      return spec.normalize(spec.unwrap ? spec.unwrap(current) : current)
    },

    async write(value: T): Promise<void> {
      const normalized = spec.normalize(value)
      let body = spec.wrap ? spec.wrap(normalized) : normalized
      if (spec.version > 1 && body !== null && typeof body === 'object' && !Array.isArray(body)) {
        body = { ...(body as Record<string, unknown>), [VERSION_FIELD]: spec.version }
      }
      await atomicWriteFile(spec.path(), Buffer.from(JSON.stringify(body, null, 2)))
    },
  }
}
