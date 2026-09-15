/**
 * frozen-selection extension — captures the user's editor selection scope at
 * `session_start` and stashes it in the ReactUIAdapter's customData bag so
 * subsequent tool calls reference the *frozen* range (not wherever the user's
 * live selection has wandered mid-run).
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   In the legacy `AgentLoop`, `docs-skill` froze the selection once per run
 *   inside its `buildContext()` callback. With pi's `AgentSession`, that
 *   hook doesn't fire for tool handlers — instead, the skill reads from
 *   the UI adapter's custom data. This extension owns the write side.
 *
 * Wiring:
 *   - apps/docs registers the extension once alongside the docs-skill:
 *       extensionFactories: [
 *         createFrozenSelectionExtension({ getEditor }),
 *         createDocsSkillExtension({ uiAdapter }),
 *       ]
 *   - docs-skill tool handlers read the frozen selection via
 *       uiAdapter.getCustomData<FrozenSelection>('frozenSelection')
 *
 * The extension captures scope via a caller-provided extractor so we don't
 * bind the skill to a specific editor implementation (Tiptap / ProseMirror
 * / a mock for tests).
 */

import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import type { ReactUIAdapter } from "@genoffice/agent-runtime";

/**
 * Generic "selection snapshot" — editor-agnostic shape.
 *
 * Tools reading this value get back exactly what the host app captured at
 * session_start. `scope` is whatever the editor considers the user's
 * effective range (block indices, a ProseMirror JSON slice, etc.). `capturedAt`
 * lets downstream tools detect a stale freeze if the doc has been edited
 * since capture.
 */
export interface FrozenSelection<T = unknown> {
  /** Editor-specific scope payload (block indices, PM JSON, …). */
  scope: T;
  /** Document fingerprint at capture time (e.g. block count + first/last hash). */
  docFingerprint: string;
  /** Monotonic timestamp (ms) of capture. */
  capturedAt: number;
}

export interface FrozenSelectionOptions<T> {
  /** Function returning the editor instance whose selection should be frozen. */
  getEditor: () => FrozenSelectionEditor<T> | null | undefined;
  /**
   * Compute a stable document fingerprint. Default uses blockCount + first/last html hash,
   * which is enough for Tiptap and matches GenOffice's legacy heuristic.
   */
  fingerprint?: (editor: FrozenSelectionEditor<T>) => string;
  /**
   * Optional override for the customData key. Defaults to "frozenSelection".
   */
  customDataKey?: string;
}

/**
 * Editor contract that the extension needs. Smaller than `DocsEditor` so it
 * works with sheets / slides editors that have a different shape (cell ranges,
 * slide selections). Tools downstream cast to the full editor type when they
 * need richer behaviour.
 */
export interface FrozenSelectionEditor<T> {
  /** Return the current selection scope, or null if the editor has none. */
  getSelectionScope(): T | null;
  /** Total number of addressable units in the document (blocks, cells, slides). */
  getUnitCount(): number;
  /** Read one unit's content for fingerprinting. */
  getUnitText(index: number): string;
}

/**
 * Default fingerprint: 4-line summary that catches additions/deletions/replacements
 * of leading or trailing content. Good enough for "did the user edit the doc
 * since capture?" — not a content hash.
 */
function defaultFingerprint<T>(editor: FrozenSelectionEditor<T>): string {
  const n = editor.getUnitCount();
  if (n === 0) return "empty";
  const head = editor.getUnitText(0);
  const tail = editor.getUnitText(n - 1);
  return `${n}|${head.slice(0, 64)}|${tail.slice(0, 64)}`;
}

/**
 * Create the pi extension factory.
 *
 * Returns a `(pi: ExtensionAPI) => void` suitable for `createOfficeSession`'s
 * `extensionFactories` array.
 */
export function createFrozenSelectionExtension<T = unknown>(
  opts: FrozenSelectionOptions<T>,
) {
  const key = opts.customDataKey ?? "frozenSelection";
  const fp = opts.fingerprint ?? defaultFingerprint;

  return (pi: ExtensionAPI): void => {
    pi.on("session_start", async (_event, ctx) => {
      const editor = opts.getEditor();
      if (!editor) return;
      const scope = editor.getSelectionScope();
      if (scope == null) return;

      const snapshot: FrozenSelection<T> = {
        scope,
        docFingerprint: fp(editor),
        capturedAt: Date.now(),
      };

      // The UI adapter is whatever the session wired. In Office mode this is
      // the ReactUIAdapter (which has setCustomData). If a different UI
      // adapter is wired, we fall back to a no-op rather than crashing.
      // The UI adapter is whatever the session wired. In Office mode this is
      // the ReactUIAdapter (which has setCustomData). If a different UI
      // adapter is wired, we fall back to a no-op rather than crashing.
      const ui = ctx.ui as unknown as Partial<ReactUIAdapter> | undefined;
      if (ui && typeof ui.setCustomData === "function") {
        (ui as { setCustomData: (k: string, v: unknown) => void }).setCustomData(key, snapshot);
      }
    });
  };
}
