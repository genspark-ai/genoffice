import { hooksFor, type FeatureRecord, type HookEvent } from '@genoffice/agent-core'
import type { HookActivity } from '../shared/feature-channels'
import { runScriptInSandbox } from './script-host'

/**
 * The hook kind's host side (K2): run every feature that observes an event, and
 * keep a short record of what each one said.
 *
 * A hook runs in the same sandbox as a script, with the event as data
 * (input.event) and no document to read — an event is not a plan. It is read-only
 * by construction: the runner has no writer, and the sandbox has no bridge to one.
 */

export type { HookActivity }

/** How many entries the panel can see. */
export const HOOK_ACTIVITY_MAX = 50

export interface HookRunnerDeps {
  /** the registry as of now */
  list(): FeatureRecord[]
  /** run one hook body; returns what it logged, and any reason it gave to stop */
  run(
    hookId: string,
    source: string,
    event: HookEvent,
    path: string,
  ): { log: string[]; vetoes?: string[] }
  /** injected so the record is testable */
  now?: () => string
}

export interface HookRunner {
  /** run every hook that observes this event; returns what ran */
  emit(event: HookEvent, path: string): HookActivity[]
  /** the most recent activity, newest last */
  activity(): HookActivity[]
  dispose(): void
}

/**
 * Run the hooks for one event. Pure apart from the injected runner, so the
 * matching, the bounded log and the error handling are unit-tested without
 * electron.
 */
export function runHooks(
  features: readonly FeatureRecord[],
  event: HookEvent,
  path: string,
  run: HookRunnerDeps['run'],
  now: () => string = () => new Date().toISOString(),
): HookActivity[] {
  return hooksFor(features, event).map((feature) => {
    const source = feature.hook?.script.source ?? ''
    try {
      const outcome = run(feature.id, source, event, path)
      return {
        at: now(),
        hookId: feature.id,
        event,
        path,
        log: outcome.log.slice(0, 20),
        ...(outcome.vetoes?.length ? { vetoes: outcome.vetoes } : {}),
      }
    } catch (err) {
      return {
        at: now(),
        hookId: feature.id,
        event,
        path,
        log: [],
        error: err instanceof Error ? err.message : String(err),
      }
    }
  })
}

/** The live runner: registry in, in-memory activity out. */
export function createHookRunner(deps: HookRunnerDeps): HookRunner {
  let activity: HookActivity[] = []
  return {
    emit(event, path) {
      const ran = runHooks(deps.list(), event, path, deps.run, deps.now)
      if (ran.length) {
        activity = [...activity, ...ran].slice(-HOOK_ACTIVITY_MAX)
      }
      return ran
    },
    activity() {
      return activity
    },
    dispose() {
      activity = []
    },
  }
}

/** The sandbox-backed runner the shell uses: no snapshot, the event as input. */
export function sandboxHookRunner(): (
  hookId: string,
  source: string,
  event: HookEvent,
  path: string,
) => {
  log: string[]
} {
  return (hookId, source, event, path) =>
    runScriptInSandbox({
      scriptPath: hookId + '.js',
      source,
      snapshot: { blocks: [], text: '' },
      input: { event: { name: event, path } },
    })
}
/** A save-ish channel: the ones a hook may stop. */
const SAVE_CHANNEL = /(^|:)(save|write)(-|$)/

export function shouldGuardSave(channel: string): boolean {
  return SAVE_CHANNEL.test(channel)
}

/**
 * The document a save-ish call is about. The payloads differ per app, so this
 * reads the common shapes and falls back to the channel name: a hook gets
 * something to match on either way, and never a path that was not sent.
 */
export function saveTargetOf(args: readonly unknown[]): string {
  const first = args[0]
  if (typeof first === 'string') return first
  if (first && typeof first === 'object') {
    const record = first as Record<string, unknown>
    for (const key of ['path', 'filePath', 'target', 'documentPath']) {
      const value = record[key]
      if (typeof value === 'string' && value) return value
    }
  }
  return ''
}

/** The live registry handle, so the guard can run hooks without owning the store. */
let live: HookRunner | null = null

export function setLiveHookRunner(runner: HookRunner | null): void {
  live = runner
}

export function liveHookRunner(): HookRunner | null {
  return live
}

/** The shape this needs from electron's ipcMain; injected so it can be faked. */
export interface IpcHandleHost {
  handle(channel: string, listener: (...args: unknown[]) => unknown): void
}

/**
 * Let the hooks see a save before it happens, and stop it if one of them says so.
 *
 * The patch is installed once, early, so every channel an app registers later goes
 * through it. Only save-ish channels are inspected, and only when a hook exists for
 * document.will-save: with no hook, the wrapper calls straight through. A hook that
 * throws is ignored — a broken hook must never break a save — and a veto is turned
 * into a rejection on the invoke, which is what the app already renders as a failed
 * save, with the hook's reason in the message.
 */
export function installSaveGuard(host: IpcHandleHost): void {
  const original = host.handle.bind(host)
  host.handle = (channel: string, listener: (...args: unknown[]) => unknown): void => {
    if (!shouldGuardSave(channel)) {
      original(channel, listener)
      return
    }
    original(channel, ((event: unknown, ...args: unknown[]) => {
      const runner = live
      if (runner) {
        try {
          const outcome = runner.emit('document.will-save', saveTargetOf(args) || channel)
          const vetoes = outcome.flatMap((entry) => entry.vetoes ?? [])
          if (vetoes.length) {
            return Promise.reject(new Error('a hook stopped this save: ' + vetoes.join('; ')))
          }
        } catch {
          // a hook that cannot run must not stop a save
        }
      }
      return listener(event, ...args)
    }) as (...args: unknown[]) => unknown)
  }
}
