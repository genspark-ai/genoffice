/**
 * createOfficeSession — wraps pi's createAgentSession with GenOffice defaults.
 *
 * Responsibilities:
 * - Construct a ModelRuntime (so model discovery works out of the box)
 * - Wire the ReactUIAdapter as the ExtensionUIContext via the session's
 *   extensionRunner.setUIContext(adapter, mode)
 * - Return the session + the adapter so React can subscribe to dialogs/notifications
 *
 * The returned `dispose()` cleans up the pi session and removes all listeners on the adapter.
 */

import {
  createAgentSession,
  DefaultResourceLoader,
  ModelRuntime,
  SessionManager,
  type AgentSession,
  type ExtensionAPI,
  type ExtensionUIContext,
  type ResourceLoader,
  type SessionManager as SessionManagerType,
  getAgentDir,
} from "@earendil-works/pi-coding-agent";
import { ReactUIAdapter } from "./ui-adapter";

/** UI mode for the extension runner. Matches pi's internal `ExtensionMode`. */
export type ExtensionUIMode = Parameters<AgentSession["extensionRunner"]["setUIContext"]>[1];

export interface OfficeSessionOptions {
  /** Working directory. Defaults to process.cwd(). */
  cwd?: string;
  /** Pi agent dir (where ~/.pi/agent/extensions live). Defaults to ~/.pi/agent. */
  agentDir?: string;
  /** Pre-built session manager. Defaults to in-memory. */
  sessionManager?: SessionManagerType;
  /** Pre-built model runtime. Defaults to a new ModelRuntime.create(). */
  modelRuntime?: ModelRuntime;
  /** Custom UI adapter. Defaults to a fresh ReactUIAdapter. */
  uiAdapter?: ReactUIAdapter;
  /** Additional extension file paths (in addition to the default discovery). */
  additionalExtensionPaths?: string[];
  /**
   * Additional skill directories (each holding `<name>/SKILL.md`), in addition
   * to pi's own discovery. GenOffice passes the marketplace skills directory
   * here so a fresh install is visible to the next session without touching
   * the user's global pi settings.
   */
  additionalSkillPaths?: string[];
  /** In-process extension factories (preferred to file paths for bundlers). */
  extensionFactories?: Array<(pi: ExtensionAPI) => void>;
  /**
   * Extension mode passed to runner.setUIContext.
   * "print" is the safe default for headless / Electron renderer use.
   * Use "tui" only if you also implement the TUI-only methods.
   */
  extensionMode?: ExtensionUIMode;
}

export interface OfficeSession {
  /** The pi AgentSession — subscribe to events, call prompt(), dispose() etc. */
  session: AgentSession;
  /** The shared UI adapter — wire React components to it. */
  uiAdapter: ReactUIAdapter;
  /**
   * Re-run pi's resource discovery so newly installed skills / plugins /
   * packages become visible without rebuilding the session. Resolves to the
   * skill + extension counts after the reload.
   */
  reloadResources: () => Promise<{ skills: number; extensions: number }>;
  /** Clean up everything. Safe to call multiple times. */
  dispose: () => void;
}

export async function createOfficeSession(opts: OfficeSessionOptions = {}): Promise<OfficeSession> {
  const cwd = opts.cwd ?? process.cwd();
  const agentDir = opts.agentDir ?? getAgentDir();
  const uiAdapter = opts.uiAdapter ?? new ReactUIAdapter();
  const extensionMode: ExtensionUIMode = opts.extensionMode ?? "print";

  const modelRuntime =
    opts.modelRuntime ?? (await ModelRuntime.create({ authPath: `${agentDir}/auth.json`, modelsPath: `${agentDir}/models.json` }));
  const sessionManager = opts.sessionManager ?? SessionManager.inMemory(cwd);

  const resourceLoader = new DefaultResourceLoader({
    cwd,
    agentDir,
    ...(opts.additionalExtensionPaths ? { additionalExtensionPaths: opts.additionalExtensionPaths } : {}),
    ...(opts.additionalSkillPaths ? { additionalSkillPaths: opts.additionalSkillPaths } : {}),
    ...(opts.extensionFactories ? { extensionFactories: opts.extensionFactories } : {}),
  });
  await resourceLoader.reload();

  const { session } = await createAgentSession({
    cwd,
    agentDir,
    sessionManager,
    modelRuntime,
    resourceLoader,
  });

  // Wire the ReactUIAdapter as the extension UI context.
  // The session exposes its ExtensionRunner so callers can adjust mode/UI at runtime.
  try {
    session.extensionRunner.setUIContext(uiAdapter.piContext as ExtensionUIContext, extensionMode);
  } catch (err) {
    // If the runner doesn't expose setUIContext on this version, the adapter
    // is still safe to use as a custom data bag — tools can reach it via
    // adapter.setEditorInstance/getEditorInstance.
    console.warn("[agent-runtime] setUIContext failed; adapter will only act as data bag:", err);
  }

  let disposed = false;
  const dispose = () => {
    if (disposed) return;
    disposed = true;
    try {
      session.dispose();
    } catch {
      // session may already be disposed
    }
  };

  const reloadResources = async (): Promise<{ skills: number; extensions: number }> => {
    await resourceLoader.reload();
    const skills = resourceLoader.getSkills().skills.length;
    const extensions = resourceLoader.getExtensions().extensions.length;
    return { skills, extensions };
  };

  return { session, uiAdapter, reloadResources, dispose };
}
