/** Max agent turns allowed in the sheets AI panel per run. */
export const SHEETS_AGENT_MAX_TURNS = 50

/** Instruction used to resume a run whose connection was interrupted after tool
 * work (auto-sent when connectivity returns, or via the Resume button). */
export const SHEETS_CONTINUE_INSTRUCTION =
  'Continue the current task using the existing conversation and workbook state. Finish only the outstanding work. Do not repeat edits or steps already complete.'