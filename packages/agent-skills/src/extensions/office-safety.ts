/**
 * office-safety — combined façade that wires `frozen-selection` and
 * `verify-response` in a single `installOfficeSafety(pi, opts)` call.
 *
 * Use this when the host app doesn't need fine-grained control over the
 * two extensions. Drop down to the individual `createFrozenSelectionExtension`
 * / `createVerifyResponseExtension` factories when:
 *   - The host wants verify rules but not frozen selection (e.g. a chat-only app).
 *   - The host wants a custom selection scope extractor beyond what `getEditor` returns.
 *   - The host is wiring multiple editors (one per office app) and needs
 *     separate frozen-snapshot keys per editor.
 */

import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";
import {
  createFrozenSelectionExtension,
  type FrozenSelection,
  type FrozenSelectionEditor,
  type FrozenSelectionOptions,
} from "./frozen-selection";
import {
  createVerifyResponseExtension,
  type VerifyResponseOptions,
} from "./verify-response";

export interface OfficeSafetyOptions<T> {
  /**
   * Selection scope source. Pass a `getEditor` that returns whatever the host's
   * editor exposes (Tiptap, Univer, custom). Optional — if absent, frozen
   * selection is skipped but verify rules still install.
   */
  frozen?: Omit<FrozenSelectionOptions<T>, "customDataKey">;
  /** Verify-response overrides (rules text, marker). Optional. */
  verify?: VerifyResponseOptions;
}

/**
 * Wire both safety extensions into a single `pi` agent session.
 *
 * @example
 *   extensionFactories: [
 *     (pi) => installOfficeSafety(pi, {
 *       frozen: { getEditor: () => tiptapEditor },
 *       verify: { rules: "Don't claim without proof" },
 *     }),
 *     createDocsSkillExtension({ uiAdapter }),
 *   ]
 */
export function installOfficeSafety<T>(
  pi: ExtensionAPI,
  opts: OfficeSafetyOptions<T> = {},
): void {
  if (opts.frozen) {
    createFrozenSelectionExtension<T>(opts.frozen)(pi);
  }
  createVerifyResponseExtension(opts.verify ?? {})(pi);
}

export { createFrozenSelectionExtension, createVerifyResponseExtension };
export type { FrozenSelection, FrozenSelectionEditor, FrozenSelectionOptions };
export { DEFAULT_VERIFY_RULES, VERIFY_BLOCK_MARKER } from "./verify-response";
