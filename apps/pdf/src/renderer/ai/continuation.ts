/** Max agent turns allowed in the PDF AI panel per run. */
export const PDF_AGENT_MAX_TURNS = 50

/** Instruction used to resume a run whose connection was interrupted after tool
 * work (auto-sent when connectivity returns, or via the Resume button). */
export const PDF_CONTINUE_INSTRUCTION =
  'Continue the current task using the existing conversation and document state. Finish only the outstanding work. Do not repeat edits or steps already complete.'