import type { ProjectApi, RunSummary } from '@genoffice/project-store'
import type { AppendRunEventRequest, HomeApi } from '../../shared/home-api'

/**
 * Run-record persistence for the home ask panel, over project-store's
 * first-class runs IPC: `project:appendRunEvent` / `project:listRuns` /
 * `project:deleteRun` against `<project>/runs/*.jsonl`. The panel's loop lives
 * in this renderer, so a run dies with the window: what survives is the event
 * pair written around it — a start event when the run begins, a terminal event
 * when it settles. A start with no terminal reads back after a restart as
 * "interrupted", which is exactly what happened.
 *
 * The records live under the default project (a file-less chat always resolves
 * there), scoped by run id, so they reuse the store's own fold (listRuns) and
 * caps without touching the panel's chat transcript. The panel reads the fold
 * once on mount and keeps the interrupted rows until they are dismissed.
 *
 * History: records used to ride the chat channel — JSON lines appended to a
 * reserved `kb-runs` chat, folded client-side. Those legacy lines are simply
 * ignored now: this transport folds its own files, and the reserved chat never
 * surfaces anyway (the panel's thread list matches `kb-ask` ids only). A
 * one-time migration would re-shape lines the task list can already live
 * without, so the simpler rule is "old kb-runs chats are dead weight".
 */

/** the runs transport the recorder needs: the shell preload plus the chat-side project resolver */
export interface RunsBridge {
  appendRunEvent(args: AppendRunEventRequest): Promise<void>
  listRuns(args: { projectId: string }): Promise<RunSummary[]>
  deleteRun(args: { projectId: string; runId: string }): Promise<void>
  /** the project the panel's runs live under (resolved once, at init) */
  resolveProject(): Promise<string>
}

/**
 * The default project the panel's runs live under. A file-less resolveChat
 * answers the default project without writing anything, so this is identity
 * resolution only — the runs API does the storing.
 */
export async function resolveRunsProject(
  project: Pick<ProjectApi, 'resolveChat'>,
): Promise<string> {
  const ids = await project.resolveChat({ filePath: null })
  return ids.projectId
}

/**
 * The runs half of the renderer bridges, or undefined when this build cannot
 * serve it (no project bridge, or a stale preload that predates the runs
 * methods — the dev-mode rebuild gap). Callers treat undefined as "no run
 * records": runs still work, nothing persists.
 */
export function runsBridgeOf(
  home: Pick<HomeApi, 'appendRunEvent' | 'listRuns' | 'deleteRun'> | undefined | null,
  project: Pick<ProjectApi, 'resolveChat'> | undefined | null,
): RunsBridge | undefined {
  const { appendRunEvent, listRuns, deleteRun } = home ?? {}
  if (
    !project ||
    typeof appendRunEvent !== 'function' ||
    typeof listRuns !== 'function' ||
    typeof deleteRun !== 'function'
  ) {
    return undefined
  }
  return {
    appendRunEvent: (args) => appendRunEvent(args),
    listRuns: (args) => listRuns(args),
    deleteRun: (args) => deleteRun(args),
    resolveProject: () => resolveRunsProject(project),
  }
}

/** one run's lifecycle events as they cross the bridge (projectId filled in at write time) */
type RunEventDraft = Omit<AppendRunEventRequest, 'projectId' | 'runId'>

/** the rows the background-task list shows from storage: what never finished */
export function interruptedOf(summaries: readonly RunSummary[]): RunSummary[] {
  return summaries.filter((run) => run.status === 'interrupted')
}

/**
 * The panel's run-record writer: queues events until the runs project is
 * resolved (a run can start before init answers), swallows append failures
 * (run records are telemetry — the transcript is the durable record), and
 * reads back the interrupted runs of earlier sessions.
 */
export class RunRecorder {
  private projectId: string | null = null
  private readonly queued: Array<{ runId: string; event: RunEventDraft }> = []

  constructor(private readonly bridge: RunsBridge) {}

  /** resolve the runs project, flush anything queued meanwhile, read back the interrupted runs */
  async init(): Promise<RunSummary[]> {
    this.projectId = await this.bridge.resolveProject()
    for (const pending of this.queued.splice(0)) this.write(pending.runId, pending.event)
    return interruptedOf(await this.bridge.listRuns({ projectId: this.projectId }))
  }

  /** a run began: its id, its instruction, and the transcript it is writing into */
  start(runId: string, instruction: string, chatId: string): void {
    this.write(runId, { kind: 'start', instruction, chatId })
  }

  /** a run settled: done, or cancelled (Stop / abort) */
  finish(runId: string, outcome: 'done' | 'cancelled'): void {
    this.write(runId, { kind: 'finish', outcome })
  }

  /** a run failed: the failure text the row shows */
  error(runId: string, message: string): void {
    this.write(runId, { kind: 'error', error: message })
  }

  /** the dismiss action on an interrupted row: the record is settled, not deleted */
  discard(runId: string): void {
    this.write(runId, { kind: 'finish', outcome: 'discarded' })
  }

  /**
   * The whole-list clear: drop every stored run record of the panel's project
   * (the same wipe deleting the old reserved chat performed). A live run's
   * record goes too; its terminal event then re-creates a start-less file the
   * fold ignores, so the list stays clean.
   */
  async clear(): Promise<void> {
    const projectId = this.projectId
    if (!projectId) return
    for (const run of await this.bridge.listRuns({ projectId })) {
      await this.bridge.deleteRun({ projectId, runId: run.runId })
    }
  }

  private write(runId: string, event: RunEventDraft): void {
    const projectId = this.projectId
    if (!projectId) {
      this.queued.push({ runId, event })
      return
    }
    void this.bridge.appendRunEvent({ projectId, runId, ...event }).catch(() => undefined)
  }
}
